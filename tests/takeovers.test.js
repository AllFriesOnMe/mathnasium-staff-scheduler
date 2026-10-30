const takeoverRoster = P.readBinders(
  sheet([
    ['Student', 'Binder'],
    ['Raya Holt', 'A1'],
    ['Onna Kohut', 'B2'],
    ['April Lowe', 'C3'],
    ['Raya Hughes', 'D4'],
  ]),
);

function takeoverFixture() {
  const times = [
    '18:30',
    '18:35',
    '18:40',
    '18:45',
    '19:00',
    '19:15',
    '19:30',
    '19:35',
    '19:40',
    '20:00',
    '20:15',
    '20:40',
  ];
  const s = sheet([
    ['Time', 'IC 1', 'IC 2', 'IC 3', 'IC 4', 'Staff 1', 'Staff 2'],
    ...times.map((time) => [time, '', '', '', '', '', '']),
  ]);
  const at = (time) => times.indexOf(time) + 1;
  const put = (time, c, value, color) => {
    const a = XLSX.utils.encode_cell({ r: at(time), c });
    s[a] = { t: 's', v: value, s: { patternType: 'solid', fgColor: { rgb: color } } };
  };
  put('18:30', 5, 'elizabeth', 'FCE5CD');
  put('18:35', 5, 'table 3', 'FCE5CD');
  put('18:30', 6, 'abigail', 'D5A6BD');
  put('18:35', 6, 'table 5', 'D5A6BD');
  for (const time of times)
    if (time >= '18:35' && time <= '19:40') for (const c of [1, 2, 3]) put(time, c, '', 'FCE5CD');
  put('18:40', 1, 'Raya Holt', 'FCE5CD');
  return { s, put, at };
}
function takeoverRows(fixture) {
  return P.readSchedule(fixture.s, takeoverRoster);
}

test('Online flag follows a handoff and each instructor gets their own table', () => {
  const f = takeoverFixture();
  f.put('18:40', 1, 'Raya Holt OL', 'FCE5CD');
  f.put('19:00', 4, 'Send Raya Holt to Abigail at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  assert(p.rows.length === 2 && p.rows.every((row) => row.online));
  assert(p.rows.find((row) => row.staff === 'elizabeth').table === '3');
  assert(p.rows.find((row) => row.staff === 'abigail').table === '5');
});
function assertRayaHandoff(p, time = '19:30', incoming = 'abigail') {
  const rows = p.rows.filter((r) => r.student === 'Raya Holt');
  assert(rows.length === 2, JSON.stringify(p));
  assert(rows[0].staff === 'elizabeth' && rows[0].start === '18:40' && rows[0].end === time);
  assert(
    P.key(rows[1].staff) === P.key(incoming) && rows[1].start === time && rows[1].end === '19:40',
  );
  assert(
    rows[0].notes.some((n) => n.includes('Handoff to')) &&
      rows[1].notes.some((n) => n.includes('Taking over from')),
  );
  assert(rows.reduce((total, row) => total + row.durationMinutes, 0) === 60);
}

test('C21 Jaydon example transfers from Luca to Adrian through the surrounding orange block', () => {
  const times = [
    '09:00',
    '09:30',
    '10:00',
    '10:05',
    '10:10',
    '10:15',
    '10:30',
    '11:00',
    '11:05',
    '11:10',
    '11:15',
    '11:30',
    '12:00',
    '12:05',
    '12:10',
    '12:15',
    '12:30',
    '13:00',
    '13:05',
    '13:10',
    '13:15',
    '13:30',
    '13:35',
    '14:00',
  ];
  const s = sheet([
    ['Start Time', 'IC 1', 'IC 2', 'IC 3', 'IC 4', 'IC 5', 'IC 6'],
    ...times.map((t) => [t]),
  ]);
  const put = (address, value, rgb) =>
    (s[address] = { t: 's', v: value, s: { patternType: 'solid', fgColor: { rgb } } });
  put('B5', 'adrian', 'FCE5CD');
  put('B6', 'table 4', 'FCE5CD');
  put('F5', 'luca', '9FC5E8');
  put('F6', 'table 2', '9FC5E8');
  put('G19', 'Jaydon L', '9FC5E8');
  put('B21', 'attention', 'FCE5CD');
  put('E21', '', 'FCE5CD');
  put('C21', 'Take over Jaydon L from Luca at 13:30', 'FFFFFF');
  const p = P.readSchedule(
    s,
    P.readBinders(
      sheet([
        ['Student', 'Binder'],
        ['Jaydon L', 'A1'],
      ]),
    ),
  );
  assert(p.rows.length === 2, JSON.stringify(p));
  assert(p.rows[0].staff === 'luca' && p.rows[0].start === '13:00' && p.rows[0].end === '13:30');
  assert(p.rows[1].staff === 'adrian' && p.rows[1].start === '13:30' && p.rows[1].end === '14:00');
  assert(!p.warnings.some((w) => w.includes('handoff not applied')));
});

test('Validated AI fallback resolves only unresolved notes and labels both schedules', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'Take over Raya Holt from elizabeth at 19:30', 'FFFFFF');
  const initial = takeoverRows(f);
  assert(initial.unresolvedHandoffs.length === 1);
  const item = initial.unresolvedHandoffs[0];
  const proposal = {
    cell: item.cell,
    text: item.text,
    incomingInstructor: 'abigail',
    time: '19:30',
    studentIds: [item.appointments[0].id],
  };
  const p = P.readSchedule(f.s, takeoverRoster, { handoffResolutions: [proposal] });
  assertRayaHandoff(p);
  assert(p.unresolvedHandoffs.length === 0 && p.aiApplied.length === 1);
  assert(p.rows.every((row) => row.notes.includes('AI-assisted handoff')));
  for (const patch of [
    { incomingInstructor: 'invented' },
    { time: '19:35' },
    { text: 'stale note' },
    { studentIds: ['a999'] },
  ]) {
    const bad = P.readSchedule(f.s, takeoverRoster, {
      handoffResolutions: [{ ...proposal, ...patch }],
    });
    assert(bad.rows.length === 1 && bad.unresolvedHandoffs.length === 1 && !bad.aiApplied.length);
  }
});

test('AI cannot override a handoff that local rules already resolve', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'Take over Raya Holt from elizabeth at 19:30', 'D5A6BD');
  const p = P.readSchedule(f.s, takeoverRoster, {
    handoffResolutions: [
      {
        cell: 'E6',
        text: 'Take over Raya Holt from elizabeth at 19:30',
        incomingInstructor: 'elizabeth',
        time: '19:30',
        studentIds: ['a0'],
      },
    ],
  });
  assertRayaHandoff(p);
  assert(!p.aiApplied.length);
});

test('Named send instruction splits both instructors without restarting the hour', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'Send Raya Holt to Abigail at 7:30', 'D5A6BD');
  assertRayaHandoff(takeoverRows(f));
});

test('White note inside an instructor block inherits the matching colors on both sides', () => {
  for (const color of ['FFFFFF', undefined]) {
    const f = takeoverFixture();
    f.put('18:40', 1, '', 'FCE5CD');
    f.put('18:40', 4, 'Raya Holt', 'FCE5CD');
    f.put('19:00', 2, 'Take over Raya Holt from elizabeth at 19:30', color);
    if (!color) delete f.s[`C${f.at('19:00') + 1}`].s;
    f.put('19:00', 1, '', 'D5A6BD');
    f.put('19:00', 3, '', 'D5A6BD');
    // White space above and below the note must not count as conflicting colors.
    for (const time of ['18:35', '18:40', '18:45', '19:15', '19:30', '19:35', '19:40'])
      f.put(time, 2, '', 'FFFFFF');
    assertRayaHandoff(takeoverRows(f));
  }
});

test('Different instructors around a white note remain ambiguous', () => {
  const f = takeoverFixture();
  f.put('19:00', 2, 'Take over Raya Holt from elizabeth at 19:30', 'FFFFFF');
  f.put('19:00', 3, '', 'D5A6BD');
  const p = takeoverRows(f);
  assert(p.warnings.some((w) => w.includes('incoming instructor is missing or ambiguous')));
  assert(!p.rows.some((r) => r.notes.some((n) => n.includes('Handoff to'))));
});

test('One nearby instructor does not supply an unstated incoming teacher', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'Take over Raya Holt from elizabeth at 19:30', 'FFFFFF');
  f.put('19:00', 5, '', 'FFFFFF');
  f.put('19:00', 6, '', 'FFFFFF');
  const p = takeoverRows(f);
  assert(p.warnings.some((w) => w.includes('incoming instructor is missing or ambiguous')));
});
test('Flexible transfer, move, hand-over and passive wording resolve', () => {
  for (const instruction of [
    'transfer Raya Holt to abigail at 19:30',
    'please move Raya Holt to Abigail at 7:30 pm',
    'hand over Raya Holt to Abigail at 19:30',
    'Raya Holt should go to Abigail at 19:30',
    'Raya Holt is taken over by Abigail at 19:30',
    'Abigail will take over Raya Holt at 19:30',
  ]) {
    const f = takeoverFixture();
    f.put('19:00', 4, instruction, 'D5A6BD');
    assertRayaHandoff(takeoverRows(f));
  }
});
test('Time on the following cell takes priority over the note row', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'send Raya Holt to abigail', 'D5A6BD');
  f.put('19:15', 4, 'at 7:30', 'D5A6BD');
  assertRayaHandoff(takeoverRows(f));
});
test('Explicitly connected horizontal note cells are joined', () => {
  const f = takeoverFixture();
  f.put('19:00', 3, 'send Raya Holt to', 'D5A6BD');
  f.put('19:00', 4, 'Abigail at 19:30', 'D5A6BD');
  assertRayaHandoff(takeoverRows(f));
});
test('Incoming instructor takes over uses the row time and affected lane', () => {
  const f = takeoverFixture();
  f.put('19:30', 1, 'srishti takes over', 'A2C4C9');
  assertRayaHandoff(takeoverRows(f), '19:30', 'srishti');
  assert(!takeoverRows(f).rows.some((r) => r.student.toLowerCase().includes('takes over')));
});
test('Instruction inside a merged note applies to its covered lanes', () => {
  const f = takeoverFixture();
  f.put('18:40', 2, 'Onna Kohut', 'FCE5CD');
  f.put('19:30', 1, 'srishti takes over', 'A2C4C9');
  f.s['!merges'] = [XLSX.utils.decode_range(`B${f.at('19:30') + 1}:C${f.at('19:30') + 1}`)];
  const p = takeoverRows(f);
  assert(p.rows.filter((r) => r.staff === 'srishti').length === 2);
});
test('A student list preserves separate original end times', () => {
  const f = takeoverFixture();
  f.put('18:40', 1, 'Onna Kohut', 'FCE5CD');
  f.put('18:35', 2, 'April Lowe', 'FCE5CD');
  f.put('19:00', 4, 'take over Onna and April at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  const incoming = p.rows.filter((r) => r.staff === 'abigail');
  assert(
    incoming.length === 2 &&
      incoming.some((r) => r.student === 'Onna Kohut' && r.end === '19:40') &&
      incoming.some((r) => r.student === 'April Lowe' && r.end === '19:35'),
  );
});
test('Instructor group references transfer all active source appointments', () => {
  const f = takeoverFixture();
  f.put('18:40', 2, 'Onna Kohut', 'FCE5CD');
  f.put('19:00', 4, "srishti takes over elizabeth's students at 19:30", 'A2C4C9');
  const p = takeoverRows(f);
  assert(p.rows.filter((r) => r.staff === 'srishti').length === 2);
});
test('Clear color transition alone creates matching handoff notes', () => {
  const f = takeoverFixture();
  f.put('19:15', 1, '', 'D5A6BD');
  f.put('19:30', 1, '', 'D5A6BD');
  f.put('19:35', 1, '', 'D5A6BD');
  assertRayaHandoff(takeoverRows(f), '19:15');
});
test('Unknown shades do not invent incoming instructors', () => {
  const f = takeoverFixture();
  f.put('19:15', 1, '', '112233');
  const p = takeoverRows(f);
  assert(p.rows.length === 1 && p.rows[0].end === '19:40');
});
test('Repeating the name at a mid-hour color change is not a second appointment', () => {
  const f = takeoverFixture();
  f.put('19:15', 1, 'Raya Holt', 'D5A6BD');
  f.put('19:30', 1, '', 'D5A6BD');
  f.put('19:35', 1, '', 'D5A6BD');
  assertRayaHandoff(takeoverRows(f), '19:15');
});
test('Assessment prefix and subsequent handoff retain full appointment context', () => {
  const f = takeoverFixture();
  f.put('18:40', 1, 'Raya Holt', '00FF00');
  f.put('19:00', 4, 'send Raya Holt to abigail at 19:30', 'D5A6BD');
  const p = P.readSchedule(f.s, takeoverRoster, { assessmentColor: '#00FF00' });
  assertRayaHandoff(p);
  assert(p.rows.every((r) => r.notes.includes('Assessment: 18:40 - 18:45')));
});
test('Two-hour appointment handoff preserves its original end', () => {
  const f = takeoverFixture();
  f.put('19:40', 1, 'Raya Holt', 'FCE5CD');
  f.put('19:00', 4, 'send Raya Holt to abigail at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  assert(
    p.rows.length === 2 &&
      p.rows[1].end === '20:40' &&
      p.rows.reduce((sum, r) => sum + r.durationMinutes, 0) === 120,
  );
  assert(p.rows.every((r) => r.notes.some((n) => n.includes('Part of 2-hour appointment'))));
});
test('Ambiguous first name is flagged without applying a partial handoff', () => {
  const f = takeoverFixture();
  f.put('18:40', 2, 'Raya Hughes', 'FCE5CD');
  f.put('19:00', 4, 'send Raya to abigail at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  assert(
    p.rows.length === 2 &&
      p.rows.every((r) => r.staff === 'elizabeth') &&
      p.warnings.some((w) => w.includes('student reference is missing or ambiguous')),
  );
});
test('Handoff beyond appointment end is flagged without extending it', () => {
  const f = takeoverFixture();
  f.put('20:00', 4, 'send Raya Holt to abigail at 20:00', 'D5A6BD');
  const p = takeoverRows(f);
  assert(
    p.rows.length === 1 &&
      p.rows[0].end === '19:40' &&
      p.warnings.some((w) => w.includes('no matching appointment')),
  );
});
test('Contradictory explicit handoffs at the same time are flagged', () => {
  const f = takeoverFixture();
  f.put('19:00', 3, 'send Raya Holt to abigail at 19:30', 'D5A6BD');
  f.put('19:15', 4, 'send Raya Holt to Jordan at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'elizabeth' &&
      p.warnings.some((w) => w.includes('Conflicting handoffs')),
  );
});
test('Delayed-start text is not misread as a takeover', () => {
  const f = takeoverFixture();
  f.put('18:45', 1, '4:15 start', 'FCE5CD');
  const p = takeoverRows(f);
  assert(p.rows.length === 1 && !p.warnings.some((w) => w.includes('handoff')));
});

test('Multiple chronological takeovers retain notes on the intermediate instructor', () => {
  const f = takeoverFixture();
  f.put('19:15', 3, 'send Raya Holt to abigail at 19:00', 'D5A6BD');
  f.put('19:00', 4, "Jordan takes over abigail's students at 19:30", '112233');
  const p = takeoverRows(f);
  const rows = p.rows.filter((r) => r.student === 'Raya Holt');
  assert(
    rows.length === 3 &&
      rows[1].staff === 'abigail' &&
      rows[1].start === '19:00' &&
      rows[1].end === '19:30' &&
      P.key(rows[2].staff) === 'jordan',
  );
  assert(
    rows[1].notes.some((n) => n.includes('Taking over from')) &&
      rows[1].notes.some((n) => n.includes('Handoff to')),
  );
});

test('Hyphenated hand-off and takeover wording is supported', () => {
  for (const instruction of [
    'hand-off Raya Holt to Abigail at 19:30',
    'Abigail takes-over Raya Holt at 19:30',
    'handoff to Abigail at 19:30',
  ]) {
    const f = takeoverFixture();
    f.put('19:30', 1, instruction, 'D5A6BD');
    assertRayaHandoff(takeoverRows(f));
  }
});

test('Explicit hour-only PM and starting-at wording use the stated time', () => {
  for (const instruction of [
    'send Raya Holt to Abigail at 7pm',
    'send Raya Holt to Abigail starting at 19:00',
  ]) {
    const f = takeoverFixture();
    f.put('19:30', 4, instruction, 'D5A6BD');
    assertRayaHandoff(takeoverRows(f), '19:00');
  }
});

test('Unresolved pronouns are not invented instructor names', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'send Raya Holt to someone at 19:30', 'D5A6BD');
  const p = takeoverRows(f);
  assert(
    p.rows.length === 1 &&
      p.warnings.some((w) => w.includes('incoming instructor is missing or ambiguous')),
  );
});

test('An incomplete explicit time is flagged instead of using the note row', () => {
  const f = takeoverFixture();
  f.put('19:00', 4, 'send Raya Holt to Abigail at 7', 'D5A6BD');
  const p = takeoverRows(f);
  assert(p.rows.length === 1 && p.warnings.some((w) => w.includes('time is missing or ambiguous')));
});

test('Incoming names in handoff notes do not relabel earlier ACD bookings', () => {
  for (const instruction of [
    'Send Raya Holt to Jordan at 19:30',
    'Jordan takes over Raya Holt at 19:30',
  ]) {
    const s = sheet([
      ['Time', 'IC 1', 'Notes'],
      ['18:40', 'Raya Holt', ''],
      ['19:00', '', ''],
      ['19:30', '', ''],
      ['19:40', '', ''],
    ]);
    for (const a of ['B2', 'B3', 'B4', 'C3'])
      s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: 'FFD966' } } };
    s.C3.v = instruction;
    s.C3.t = 's';
    const p = P.readSchedule(s, takeoverRoster);
    assert(
      p.rows.length === 2 &&
        p.rows[0].staff === 'ACD' &&
        p.rows[0].start === '18:40' &&
        p.rows[0].end === '19:30' &&
        P.key(p.rows[1].staff) === 'jordan',
    );
  }
});
