function ambiguousColorFixture() {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2', 'IC 3'],
    ['10:00', 'Samara A', 'Ryan C', 'Veer Shah'],
    ['10:05', 'alex', 'yama', ''],
    ['10:10', 'table 1', 'table 2', ''],
    ['10:15', '', '', ''],
    ['11:00', '', '', ''],
  ]);
  for (const [column, rgb] of [
    ['B', 'A0C0E0'],
    ['C', 'BCC0E0'],
    ['D', 'AFC0E0'],
  ])
    for (const row of [2, 3, 4, 5])
      s[`${column}${row}`] = {
        ...(s[`${column}${row}`] || { t: 'z' }),
        s: { patternType: 'solid', fgColor: { rgb } },
      };
  return s;
}

test('Uncertain shades produce a bounded AI case with nearby instructor evidence', () => {
  const p = P.readSchedule(ambiguousColorFixture(), roster);
  assert(p.unresolvedColors.length === 1);
  const item = p.unresolvedColors[0];
  assert(item.kind === 'color' && item.cell === 'D2' && item.color === 'rgb:AFC0E0');
  assert(
    item.nearby.length <= 45 &&
      item.appointments.length === 1 &&
      item.appointments[0].student === 'Veer Shah',
  );
  assert(item.candidates.some((candidate) => candidate.instructor === 'yama'));
  assert(p.warnings.some((w) => w.includes('D2') && w.includes('uncertain shade')));
});

test('Validated AI shade correction preserves the booking and labels its assignment', () => {
  const s = ambiguousColorFixture(),
    first = P.readSchedule(s, roster),
    item = first.unresolvedColors[0];
  const proposal = {
    kind: 'color',
    cell: item.cell,
    text: item.text,
    color: item.color,
    time: item.time,
    incomingInstructor: 'yama',
    studentIds: item.appointments.map((row) => row.id),
  };
  const p = P.readSchedule(s, roster, { colorResolutions: [proposal] });
  const row = p.rows.find((row) => row.student === 'Veer Shah');
  assert(
    row.staff === 'yama' &&
      row.start === '10:00' &&
      row.end === '11:00' &&
      row.durationMinutes === 60,
  );
  assert(
    row.notes.includes('AI-assisted color assignment') &&
      !row.notes.some((note) => note.includes('Handoff')),
  );
  assert(
    !p.unresolvedColors.length &&
      p.aiApplied.length === 1 &&
      !p.warnings.some((w) => w.includes('D2') && w.includes('fell back')),
  );
  for (const patch of [
    { incomingInstructor: 'invented' },
    { time: '10:15' },
    { color: 'rgb:A0C0E0' },
    { text: 'stale' },
    { studentIds: ['a99'] },
    { confident: false },
  ]) {
    const bad = P.readSchedule(s, roster, { colorResolutions: [{ ...proposal, ...patch }] });
    assert(
      bad.rows.find((row) => row.student === 'Veer Shah').staff === 'ACD' &&
        bad.unresolvedColors.length === 1,
    );
  }
  const pinned = P.readSchedule(s, roster, {
    colorResolutions: [proposal],
    colorIndicators: [{ name: 'Adrian', color: '#AFC0E0', behavior: 'instructor' }],
  });
  assert(
    pinned.rows.find((row) => row.student === 'Veer Shah').staff === 'Adrian' &&
      !pinned.unresolvedColors.length &&
      !pinned.aiApplied.length,
  );
});

test('Clear near shades stay local and possible assessment shades are not AI-reassigned', () => {
  const s = ambiguousColorFixture();
  for (const row of [2, 3, 4, 5]) s[`D${row}`].s.fgColor.rgb = 'BCC1E0';
  const local = P.readSchedule(s, roster);
  assert(
    local.rows.find((row) => row.student === 'Veer Shah').staff === 'yama' &&
      !local.unresolvedColors.length,
  );
  const assessment = P.readSchedule(ambiguousColorFixture(), roster, {
    assessmentColor: '#AFC0E1',
  });
  assert(
    !assessment.unresolvedColors.length &&
      assessment.rows.find((row) => row.student === 'Veer Shah').staff === 'ACD',
  );
});

test('AI color checks preserve assessment prefixes and two-hour appointment pairing', () => {
  const prefix = ambiguousColorFixture();
  for (const r of [2, 3, 4]) prefix[`D${r}`].s.fgColor.rgb = '00FF00';
  const initial = P.readSchedule(prefix, roster, { assessmentColor: '#00FF00' });
  const item = initial.unresolvedColors[0];
  assert(item.cell === 'D5' && item.time === '10:15');
  const proposal = (entry) => ({
    kind: 'color',
    cell: entry.cell,
    text: entry.text,
    color: entry.color,
    time: entry.time,
    incomingInstructor: 'yama',
    studentIds: entry.appointments.map((row) => row.id),
  });
  const corrected = P.readSchedule(prefix, roster, {
    assessmentColor: '#00FF00',
    colorResolutions: [proposal(item)],
  });
  const student = corrected.rows.find((row) => row.student === 'Veer Shah');
  assert(
    student.start === '10:00' &&
      student.end === '11:00' &&
      student.notes.includes('Assessment: 10:00 - 10:15'),
  );
  const repeat = ambiguousColorFixture();
  repeat.D6 = { t: 's', v: 'Veer Shah', s: { patternType: 'solid', fgColor: { rgb: 'AFC0E0' } } };
  const repeatedCases = P.readSchedule(repeat, roster).unresolvedColors;
  assert(repeatedCases.length === 2);
  const paired = P.readSchedule(repeat, roster, { colorResolutions: repeatedCases.map(proposal) });
  const twoHours = paired.rows.filter((row) => row.student === 'Veer Shah');
  assert(
    twoHours.length === 1 &&
      twoHours[0].staff === 'yama' &&
      twoHours[0].durationMinutes === 120 &&
      twoHours[0].end === '12:00',
  );
});

test('Chosen director color overrides instructor detection for that color', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:05', 'alex'],
    ['10:10', 'table 1'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, { directorColor: '#ffd966' });
  assert(p.rows.length === 1 && p.rows[0].staff === 'Center director');
  assert(
    p.warnings.some(
      (w) =>
        w.includes('Manual color override #FFD966') &&
        w.includes('alex') &&
        w.includes('Center director'),
    ),
  );
});

test('Unrecognized color warning identifies the cell, student, time, color and ACD fallback', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '123456' } };
  const p = P.readSchedule(s, roster);
  assert(
    p.warnings.some(
      (w) =>
        w.includes('B2 (IC 1), 10:00') &&
        w.includes('Samara Adams') &&
        w.includes('#123456') &&
        w.includes('fell back to ACD') &&
        w.includes('Optional AI help') &&
        w.includes('insufficient color evidence'),
    ),
  );
});

test('Matching manual and spreadsheet instructor names do not produce override warnings', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:05', 'alex'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [{ name: 'Alex', color: '#FFD966', behavior: 'instructor' }],
  });
  assert(!p.warnings.some((w) => w.includes('Manual color override')));
});

test('Manual assessment overriding a lowercase instructor is reported even without a table label', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:05', 'alex'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [{ name: 'Initial check', color: '#FFD966', behavior: 'assessment' }],
  });
  assert(
    p.rows.length === 0 &&
      p.warnings.some(
        (w) => w.includes('spreadsheet indicates alex') && w.includes('Initial check (Assessment)'),
      ),
  );
});

test('Manual instructor overriding a spreadsheet Assessment color is reported', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:05', 'Assessment'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [{ name: 'Adrian', color: '#00FF00', behavior: 'instructor' }],
  });
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'Adrian' &&
      p.warnings.some(
        (w) => w.includes('spreadsheet indicates Assessment') && w.includes('assigns Adrian'),
      ),
  );
});

test('Named indicators override inferred teacher labels and keep separate colors independent', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Samara A', 'Ryan C'],
    ['10:05', 'alex', ''],
    ['10:10', 'table 1', ''],
    ['11:00', '', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  s.C2.s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [
      { name: 'Lilianna K', color: '#ffd966', behavior: 'instructor' },
      { name: 'Yama', color: '#c9daf8', behavior: 'instructor' },
    ],
  });
  assert(
    p.rows.length === 2 &&
      p.rows.some((row) => row.staff === 'Lilianna K') &&
      p.rows.some((row) => row.staff === 'Yama'),
  );
  assert(
    p.warnings.some(
      (w) =>
        w.includes('Manual color override #FFD966') &&
        w.includes('alex') &&
        w.includes('Lilianna K'),
    ),
  );
});

test('Custom assessment indicator retains the full appointment and assessment interval', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:15', ''],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  s.B3.s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [
      { name: 'Initial assessment', color: '#00FF00', behavior: 'assessment' },
      { name: 'Yama', color: '#C9DAF8', behavior: 'instructor' },
    ],
  });
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'Yama' &&
      p.rows[0].start === '10:00' &&
      p.rows[0].end === '11:00',
  );
  assert(p.rows[0].notes.includes('Assessment: 10:00 - 10:15'));
});

test('Incomplete and invalid named indicators are ignored', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [
      null,
      { name: '', color: '#FFD966', behavior: 'instructor' },
      { name: 'Yama', color: 'bad', behavior: 'instructor' },
      { name: 'Yama', color: '#FFD966', behavior: 'bad' },
    ],
  });
  assert(p.rows.length === 1 && p.rows[0].staff === 'ACD');
});

test('Conflicting custom colors are flagged rather than choosing an instructor', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, {
    colorIndicators: [
      { name: 'Yama', color: '#FFD966', behavior: 'instructor' },
      { name: 'Adrian', color: '#FFD966', behavior: 'instructor' },
    ],
  });
  assert(
    p.rows[0].staff === 'ACD' && p.warnings.some((w) => w.includes('Conflicting color indicator')),
  );
});

test('Selected assessment color works without an Assessment label', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  const p = P.readSchedule(s, roster, { assessmentColor: '#00FF00' });
  assert(p.rows.length === 0 && p.warnings.length === 0);
});

test('Assessment prefix retains full appointment under later instructor with exact note', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Samara A', 'Ryan C'],
    ['10:05', '', 'alex'],
    ['10:10', '', 'table 1'],
    ['10:15', '', ''],
    ['10:30', 'needs attention', ''],
    ['11:00', '', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4'])
    s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: '00FF00' } } };
  for (const a of ['B5', 'B6', 'C2', 'C3', 'C4'])
    s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } } };
  const p = P.readSchedule(s, roster, { assessmentColor: '#00FF00' });
  const row = p.rows.find((r) => r.student === 'Samara Adams');
  assert(
    p.rows.length === 2 && row.staff === 'alex' && row.start === '10:00' && row.end === '11:00',
  );
  assert(row.notes.includes('Assessment: 10:00 - 10:15') && row.notes.includes('Needs attention'));
});

test('Repeated name at assessment handoff does not become a duplicate or two-hour session', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Arjun'],
    ['10:15', 'Arjun'],
    ['10:20', 'yama'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  for (const a of ['B3', 'B4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster, { assessmentColor: '#00FF00' });
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'yama' &&
      p.rows[0].start === '10:00' &&
      p.rows[0].end === '11:00',
  );
  assert(
    p.rows[0].notes.includes('Assessment: 10:00 - 10:15') &&
      !p.warnings.some((w) => w.includes('Duplicate')),
  );
});

test('Assessment prefix plus next hourly appearance remains one two-hour appointment', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Arjun'],
    ['10:15', ''],
    ['10:20', 'yama'],
    ['11:00', 'Arjun'],
    ['12:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  for (const a of ['B3', 'B4', 'B5'])
    s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } } };
  const p = P.readSchedule(s, roster, { assessmentColor: '#00FF00' });
  assert(
    p.rows.length === 1 &&
      p.rows[0].end === '12:00' &&
      p.rows[0].durationMinutes === 120 &&
      p.rows[0].notes.includes('Assessment: 10:00 - 10:15'),
  );
});

test('Assessment-only entry does not claim the next unrelated student', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['10:30', 'Ryan C'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  s.B3.s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster, { assessmentColor: '#00FF00' });
  assert(p.rows.length === 1 && p.rows[0].student === 'Ryan Clark');
});

test('Invalid saved colors are ignored by the parser', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster, {
    directorColor: 'url(bad)',
    assessmentColor: 'not-a-color',
  });
  assert(p.rows.length === 1 && p.rows[0].staff === 'ACD');
});

test('Assessment label above the student sets the full appointment start', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Assessment'],
    ['10:05', 'Arjun'],
    ['10:15', ''],
    ['10:20', 'yama'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  for (const a of ['B4', 'B5'])
    s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } } };
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.length === 1 &&
      p.rows[0].start === '10:00' &&
      p.rows[0].end === '11:00' &&
      p.rows[0].notes.includes('Assessment: 10:00 - 10:15'),
  );
});

test('Explicit director color prevents unrelated colors being silently assigned to director', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Arjun'],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'CC66CC' } };
  const p = P.readSchedule(s, roster, { directorColor: '#FFD966' });
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'Unassigned instructor' &&
      p.warnings.some((w) => w.includes('selected director color')),
  );
});
