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
