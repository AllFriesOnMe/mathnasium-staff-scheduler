test('Find binder headers below title and blank rows', () =>
  assert(roster.get('samara adams').location === 'A1'));

test('NA marks attention inline or below a student without becoming a student or instructor', () => {
  for (const [name, annotation] of [
    ['Samara A NA', ''],
    ['Samara A (NA)', ''],
    ['Samara A', 'NA'],
    ['Samara A', 'na'],
  ]) {
    const p = P.readSchedule(
      sheet([
        ['Time', 'HS1'],
        ['10:00', name],
        ['10:05', annotation],
        ['11:00', 'Ryan C'],
      ]),
      roster,
    );
    assert(
      p.rows.length === 2 &&
        p.rows[0].student === 'Samara Adams' &&
        p.rows[0].notes.includes('Needs attention'),
    );
    assert(!p.rows[1].notes.includes('Needs attention'));
    assert(!p.warnings.some((warning) => warning.includes('unmatched text')));
  }
});

test('High-school make-up and additional notes cannot erase table-linked instructor assignments', () => {
  const s = sheet([
    ['Time', 'HS1', 'HS2'],
    ['15:00', 'Student Alpha', 'Student Beta'],
    ['15:05', 'tutor one', 'tutor two'],
    ['15:10', 'table 7', 'table 1'],
    ['15:15', 'make up', 'makeup'],
    ['15:20', 'last one', 'additional'],
    ['16:00', '', ''],
  ]);
  for (const [col, rgb] of [
    ['B', '00FFFF'],
    ['C', 'B4A7D6'],
  ])
    for (const r of [2, 3, 4, 5, 6]) s[`${col}${r}`].s = { patternType: 'solid', fgColor: { rgb } };
  const p = P.readSchedule(s, new Map());
  assert(
    p.rows.length === 2 &&
      p.rows.some((row) => row.staff === 'tutor one' && row.table === '7') &&
      p.rows.some((row) => row.staff === 'tutor two' && row.table === '1'),
  );
  assert(
    !p.warnings.some((warning) => /conflicting instructor|fell back|unmatched text/.test(warning)),
  );
});

test('Unexpected lowercase text cannot replace an instructor explicitly linked to a table', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Student Alpha'],
    ['10:05', 'tutor one'],
    ['10:10', 'table 2'],
    ['10:15', 'follow up'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4', 'B5'])
    s[a].s = { patternType: 'solid', fgColor: { rgb: '00FFFF' } };
  const p = P.readSchedule(s, new Map());
  assert(p.rows.length === 1 && p.rows[0].staff === 'tutor one');
});

test('Two different table-linked instructors sharing a color remain flagged as conflicting', () => {
  const s = sheet([
    ['Time', 'HS1', 'HS2'],
    ['10:00', 'Student Alpha', 'Student Beta'],
    ['10:05', 'tutor one', 'tutor two'],
    ['10:10', 'table 1', 'table 2'],
    ['11:00', '', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4', 'C2', 'C3', 'C4'])
    s[a].s = { patternType: 'solid', fgColor: { rgb: '00FFFF' } };
  const p = P.readSchedule(s, new Map());
  assert(
    p.rows.every((row) => row.staff === 'ACD') &&
      p.warnings.some((warning) => warning.includes('conflicting instructor labels')),
  );
});

test('Online column headings and inline OL markers retain clean student names', () => {
  for (const title of ['OL', 'OL1', 'Online 1', 'ONLINE']) {
    const p = P.readSchedule(
      sheet([
        ['Time', title],
        ['10:00', 'Samara A'],
        ['11:00', ''],
      ]),
      roster,
    );
    assert(p.rows.length === 1 && p.rows[0].online && p.rows[0].student === 'Samara Adams');
  }
  for (const name of ['Samara A OL', 'Samara A (ONLINE)', 'OL Samara A', 'Samara A\nOL']) {
    const p = P.readSchedule(
      sheet([
        ['Time', 'Alex'],
        ['10:00', name],
        ['11:00', ''],
      ]),
      roster,
    );
    assert(
      p.rows.length === 1 && p.rows[0].online && p.rows[0].student === 'Samara Adams',
      JSON.stringify(p),
    );
  }
});

test('OL annotations above and below apply only to their adjacent student', () => {
  for (const rows of [
    [
      ['10:00', 'OL'],
      ['10:05', 'Samara A'],
      ['10:30', ''],
      ['11:00', 'Ryan C'],
    ],
    [
      ['10:00', 'Samara A'],
      ['10:05', 'needs attention OL'],
      ['10:30', ''],
      ['11:00', 'Ryan C'],
    ],
  ]) {
    const p = P.readSchedule(sheet([['Time', 'Alex'], ...rows]), roster);
    assert(p.rows.find((row) => row.student === 'Samara Adams').online);
    assert(!p.rows.find((row) => row.student === 'Ryan Clark').online);
    assert(!p.warnings.some((w) => w.includes('unmatched text')));
  }
});

test('Online annotations between adjacent students are flagged rather than marking both', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'Alex'],
      ['10:00', 'Samara A'],
      ['10:30', 'OL'],
      ['11:00', 'Ryan C'],
    ]),
    roster,
  );
  assert(
    p.rows.every((row) => !row.online) &&
      p.warnings.some((w) => w.includes('online marker is between two students')),
  );
});

test('Inline online marker in a shared cell does not mark the other student online', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'Alex'],
      ['10:00', 'Samara A OL; Ryan C'],
      ['11:00', ''],
    ]),
    roster,
  );
  assert(
    p.rows.find((row) => row.student === 'Samara Adams').online &&
      !p.rows.find((row) => row.student === 'Ryan Clark').online,
  );
});

test('Online markers in flat schedules and two-hour repeats preserve matching and duration', () => {
  const flat = P.readSchedule(
    sheet([
      ['Staff', 'Start', 'End', 'Student'],
      ['Alex', '10:00', '11:00', 'Samara A (OL)'],
    ]),
    roster,
  );
  assert(flat.rows[0].student === 'Samara Adams' && flat.rows[0].online);
  const grid = P.readSchedule(
    sheet([
      ['Time', 'Alex'],
      ['10:00', 'Samara A'],
      ['11:00', 'Samara A ONLINE'],
      ['12:00', ''],
    ]),
    roster,
  );
  assert(grid.rows.length === 1 && grid.rows[0].online && grid.rows[0].durationMinutes === 120);
});
test('Separate binder row and column', () =>
  assert(
    P.readBinders(
      sheet([
        ['Name', 'Row', 'Column'],
        ['Jane Doe', 'C', 4],
      ]),
    ).get('jane doe').location === 'C4',
  ));
test('Duplicate binder names rejected', () =>
  rejects(() =>
    P.readBinders(
      sheet([
        ['Name', 'Binder'],
        ['Jane', 'A1'],
        ['Jane', 'B2'],
      ]),
    ),
  ));
test('Invalid binder location rejected', () =>
  rejects(() =>
    P.readBinders(
      sheet([
        ['Name', 'Binder'],
        ['Jane', 'F1'],
      ]),
    ),
  ));
const grid = sheet([
  ['', 1, 2, 3, 4, 5],
  ['A', 'Jane Doe\nJohn Doe'],
  ['B'],
  ['C'],
  ['D'],
  ['E', '', '', '', '', 'Veer Shah'],
]);
test('Physical binder grid and multiple names', () => {
  const r = P.readBinders(grid);
  assert(
    r.size === 3 && r.get('john doe').location === 'A1' && r.get('veer shah').location === 'E5',
  );
});
test('Transposed binder grid', () =>
  assert(
    P.readBinders(sheet([['', 'A', 'B', 'C', 'D', 'E'], [1, 'Jane'], [2], [3], [4], [5]])).get(
      'jane',
    ).location === 'A1',
  ));
test('Numeric Excel times and AM/PM', () =>
  assert(P.minutes(0.625) === 900 && P.minutes('12:00 AM') === 0 && P.minutes('3:30 PM') === 930));
test('Invalid times rejected', () =>
  assert(P.minutes('24:00') === null && P.minutes('3:99') === null));
const schedule = sheet([
  ['Center schedule'],
  ['Start Time', 'IC 1', 'IC 2', '', 'HS1', 'Start Time'],
  ['15:30', 'Samara A', '', '', 'Veer S', '15:30'],
  ['15:35', '', '', '', 'g9 ap', '15:35'],
  ['15:40', '', '', '', '', '15:40'],
  ['16:00', 'Ryan C', '', '', '', '16:00'],
  ['16:30', '', '', '', '', '16:30'],
  ['17:00', '', '', '', '', '17:00'],
]);
schedule['!merges'] = [XLSX.utils.decode_range('B3:B5'), XLSX.utils.decode_range('B6:B7')];
for (let r = 2; r <= 4; r++) {
  const a = XLSX.utils.encode_cell({ r, c: 4 });
  schedule[a] = {
    ...(schedule[a] || { t: 'z' }),
    s: { patternType: 'solid', fgColor: { rgb: '00FFFF' } },
  };
}
test('Merged blocks use start labels and default one-hour durations', () => {
  const p = P.readSchedule(schedule, roster);
  assert(p.rows[0].start === '15:30' && p.rows[0].end === '16:30');
  assert(p.rows.find((row) => row.student === 'Ryan Clark').end === '17:00');
});
test('Repeated right-hand time column', () => {
  const p = P.readSchedule(schedule, roster);
  assert(p.rows.find((r) => r.section === 'HS1').start === '15:30');
});
test('Colored blocks and notes do not shorten default sessions', () => {
  const p = P.readSchedule(schedule, roster);
  assert(p.rows.length === 3 && p.rows.find((r) => r.section === 'HS1').end === '16:30');
  assert(!p.warnings.some((w) => w.includes('duration inferred')));
});
test('Abbreviated names resolved with warning', () => {
  const p = P.readSchedule(schedule, roster);
  assert(
    p.rows[0].student === 'Samara Adams' && p.warnings.some((w) => w.includes('matched uniquely')),
  );
});
test('Actual XLSX write/read retains merge timing', () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, schedule, 'Monday');
  const bytes = XLSX.write(book, { type: 'array', bookType: 'xlsx' });
  const read = XLSX.read(bytes, { type: 'array', cellStyles: true, sheetStubs: true });
  assert(P.readSchedule(read.Sheets.Monday, roster).rows[0].end === '16:30');
});
test('Truncated sheet still gives a one-hour session', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['13:00', 'Samara A'],
      ['13:30', ''],
    ]),
    roster,
  );
  assert(p.rows[0].end === '14:00' && !p.warnings.some((w) => w.includes('end time unknown')));
});
test('Unrecognized text is explicitly reported', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['15:30', 'Unknown Student'],
    ]),
    roster,
  );
  assert(p.rows.length === 0 && p.warnings.some((w) => w.includes('was not imported')));
});
test('Ambiguous abbreviations are not silently matched', () => {
  const r = new Map(roster);
  r.set('samara allen', { name: 'Samara Allen', location: 'A2' });
  const p = P.readSchedule(schedule, r);
  assert(
    !p.rows.some((r) => r.student.startsWith('Samara')) &&
      p.warnings.some((w) => w.includes('Samara A') && w.includes('not imported')),
  );
});
test('Flat appointment lists still supported', () => {
  const p = P.readSchedule(
    sheet([
      ['Staff', 'Start', 'End', 'Student'],
      ['Alex', 0.625, '16:30', 'Samara A'],
    ]),
    roster,
  );
  assert(p.rows[0].end === '16:30' && p.rows[0].student === 'Samara Adams');
});
test('Flat invalid end rejected', () =>
  rejects(() =>
    P.readSchedule(
      sheet([
        ['Staff', 'Start', 'End', 'Student'],
        ['Alex', '16:30', '15:30', 'Samara A'],
      ]),
      roster,
    ),
  ));
test('Non-monotonic time labels rejected', () =>
  rejects(() =>
    P.readSchedule(
      sheet([
        ['Time', 'IC 1'],
        ['16:00', 'Samara A'],
        ['15:30', ''],
      ]),
      roster,
    ),
  ));
test('CSV quoting handled by SheetJS', () => {
  const b = XLSX.read('Student,Binder\r\n"Doe, Jane",A3', { type: 'string' });
  assert(P.readBinders(b.Sheets[b.SheetNames[0]]).get('doe jane').location === 'A3');
});
test('Cancelled appointments omitted with warning', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['15:30', 'Samara A'],
    ['15:35', 'cancelled'],
    ['16:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFFF00' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 0 && p.warnings.some((w) => w.includes('marked cancelled')));
});
test('Merged staff headings cover all lanes', () => {
  const s = sheet([
    ['Time', 'Alex', ''],
    ['15:30', 'Samara A', 'Ryan C'],
    ['16:00', '', ''],
  ]);
  s['!merges'] = [XLSX.utils.decode_range('B1:C1')];
  assert(P.readSchedule(s, roster).rows.length === 2);
});
test('High-school students need not appear in binder roster', () => {
  const s = sheet([
    ['Start Time', 'HS1', 'High School 2'],
    ['15:30', 'Arjun', 'Nylah S'],
    ['15:35', 'g12', 'grade 11'],
    ['16:30', 'Srishti', ''],
    ['17:30', '', ''],
  ]);
  s['!merges'] = [XLSX.utils.decode_range('C2:C3')];
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.length === 3 &&
      p.rows.some((r) => r.student === 'Arjun' && r.end === '16:30') &&
      p.rows.some((r) => r.student === 'Nylah S' && r.end === '16:30'),
  );
  assert(!p.warnings.some((w) => w.includes('not imported')));
});
test('High-school grade and table notes are not students', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'HS 2'],
      ['15:30', 'Arjun'],
      ['15:35', 'g9 ap'],
      ['15:40', 'table 1'],
      ['16:30', ''],
    ]),
    roster,
  );
  assert(p.rows.length === 1 && p.rows[0].student === 'Arjun');
});
test('Side-by-side split-name binder lists include both halves', () => {
  const r = P.readBinders(
    sheet([
      ['', '', 'First Name', 'Last Name', '', '', '', 'First Name', 'Last Name'],
      ['A', 1, 'Alice', 'Smith', '', 'E', 5, 'Bob', 'Jones'],
      ['B', 2, 'Clara', 'Silva Urzua', '', '', '', '', ''],
    ]),
  );
  assert(
    r.size === 3 &&
      r.get('alice smith').location === 'A1' &&
      r.get('bob jones').location === 'E5' &&
      r.get('clara silva urzua').location === 'B2',
  );
});
test('Table colors group lanes and teacher annotations are excluded', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2', 'HS1'],
    ['10:00', 'Samara A', 'Ryan C', 'New Student g12'],
    ['10:05', 'alex', '', 'physics'],
    ['10:10', 'table 4', '', ''],
    ['11:00', '', '', ''],
  ]);
  for (let r = 1; r <= 3; r++)
    for (let c = 1; c <= 3; c++) {
      const a = XLSX.utils.encode_cell({ r, c });
      s[a] = { ...(s[a] || { t: 'z' }), s: { patternType: 'solid', fgColor: { rgb: 'FCE5CD' } } };
    }
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.length === 3 &&
      p.rows.every((r) => r.staff === 'alex' && r.table === '4' && r.end === '11:00'),
  );
  assert(p.rows.some((r) => r.student === 'New Student' && r.section === 'HS1'));
});
test('A color change does not shorten the one-hour session', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:30', 'New Student'],
    ['11:00', 'Physics'],
    ['11:30', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  s.B3.s = { patternType: 'solid', fgColor: { rgb: 'CFE2F3' } };
  assert(P.readSchedule(s, roster).rows[0].end === '11:30');
});
test('Assessment-colored students are skipped without role warnings', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['11:00', 'Assessment'],
    ['11:05', 'New Student'],
    ['12:00', ''],
  ]);
  for (const a of ['B2', 'B3']) s[a].s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 0 && p.warnings.length === 0);
});
test('Assessment color excludes students in other columns and separate blocks', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Assessment', 'Samara A'],
    ['11:00', '', 'Ryan C'],
    ['12:00', '', ''],
  ]);
  for (const a of ['B2', 'C2', 'C3'])
    s[a].s = { patternType: 'solid', fgColor: { rgb: 'FF00FF00' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 0 && p.warnings.length === 0);
});
test('Unassigned yellow and other colors belong to ACD', () => {
  const s = sheet([
    ['Time', 'IC 1', 'HS1'],
    ['10:00', 'Samara A', 'New Student'],
    ['11:00', '', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  s.C2.s = { patternType: 'solid', fgColor: { rgb: 'CFE2F3' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 2 && p.rows.every((r) => r.staff === 'ACD'));
  assert(p.warnings.filter((w) => w.includes('fell back to ACD')).length === 2);
});
test('Named instructor on yellow overrides ACD fallback', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Samara A', 'Ryan C'],
    ['10:05', 'alex', ''],
    ['10:10', 'table 4', ''],
    ['11:00', '', ''],
  ]);
  for (let r = 1; r <= 3; r++)
    for (let c = 1; c <= 2; c++) {
      const a = XLSX.utils.encode_cell({ r, c });
      s[a] = {
        ...(s[a] || { t: 'z' }),
        s: { patternType: 'solid', fgColor: { rgb: c === 1 ? 'FFFFD966' : 'FFD966' } },
      };
    }
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 2 && p.rows.every((r) => r.staff === 'alex'));
});
test('Unique single first names resolve while ambiguous names remain unmatched', () => {
  const p = P.readSchedule(
    sheet([
      ['Staff', 'Start', 'End', 'Student'],
      ['Alex', '10:00', '11:00', 'Samara'],
    ]),
    roster,
  );
  assert(p.rows[0].student === 'Samara Adams');
});
test('Repeated student becomes one two-hour session', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['10:00', 'Samara A'],
      ['11:00', 'Samara Adams'],
    ]),
    roster,
  );
  assert(p.rows.length === 1 && p.rows[0].start === '10:00' && p.rows[0].end === '12:00');
});
test('Repeated names in different lanes of one instructor merge', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Samara A', ''],
    ['10:05', 'alex', ''],
    ['10:10', 'table 4', ''],
    ['11:00', '', 'Samara Adams'],
  ]);
  for (const a of ['B2', 'B3', 'B4', 'C5'])
    s[a].s = { patternType: 'solid', fgColor: { rgb: 'FCE5CD' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 1 && p.rows[0].staff === 'alex' && p.rows[0].end === '12:00');
});
test('High-school repeated name needs no binder and gets two hours', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'HS1'],
      ['10:00', 'Arjun'],
      ['11:00', 'Arjun'],
    ]),
    roster,
  );
  assert(p.rows.length === 1 && p.rows[0].student === 'Arjun' && p.rows[0].end === '12:00');
});
test('Different students sharing a binder remain separate one-hour sessions', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['10:00', 'Samara A'],
      ['11:00', 'Ryan C'],
    ]),
    roster,
  );
  assert(p.rows.length === 2 && p.rows[0].end === '11:00' && p.rows[1].end === '12:00');
});
test('Cancelled second entry does not extend the first session', () => {
  const s = sheet([
    ['Time', 'IC 1'],
    ['10:00', 'Samara A'],
    ['11:00', 'Samara Adams'],
    ['11:05', 'cancelled'],
  ]);
  for (const a of ['B3', 'B4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 1 && p.rows[0].end === '11:00');
});
test('Nearby pale-blue shades map Nylah and Charlotte to Yama', () => {
  const s = sheet([
    ['Time', 'HS1', 'HS2', 'HS3'],
    ['10:00', 'Nylah S', 'Charlotte', 'New Student'],
    ['10:05', '', '', 'yama'],
    ['10:10', '', '', 'table 1'],
    ['11:00', '', '', ''],
  ]);
  for (const a of ['B2', 'C2']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'CFE2F3' } };
  for (const a of ['D2', 'D3', 'D4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster);
  assert(p.rows.length === 3 && p.rows.every((r) => r.staff === 'yama'));
});
test('Close shades do not override exact instructor labels', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2'],
    ['10:00', 'Samara A', 'Ryan C'],
    ['10:05', 'alex', 'jordan'],
    ['10:10', 'table 1', 'table 2'],
    ['11:00', '', ''],
  ]);
  for (let r = 1; r <= 3; r++)
    for (let c = 1; c <= 2; c++) {
      const a = XLSX.utils.encode_cell({ r, c });
      s[a] = {
        ...(s[a] || { t: 'z' }),
        s: { patternType: 'solid', fgColor: { rgb: c === 1 ? 'CFE2F3' : 'C9DAF8' } },
      };
    }
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.some((r) => r.student === 'Samara Adams' && r.staff === 'alex') &&
      p.rows.some((r) => r.student === 'Ryan Clark' && r.staff === 'jordan'),
  );
});
test('Ambiguous close shade stays ACD instead of guessing an instructor', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2', 'HS1'],
    ['10:00', 'Samara A', 'Ryan C', 'New Student'],
    ['10:05', 'alex', 'jordan', ''],
    ['10:10', 'table 1', 'table 2', ''],
    ['11:00', '', '', ''],
  ]);
  for (let r = 1; r <= 3; r++)
    for (let c = 1; c <= 2; c++) {
      const a = XLSX.utils.encode_cell({ r, c });
      s[a] = {
        ...(s[a] || { t: 'z' }),
        s: { patternType: 'solid', fgColor: { rgb: c === 1 ? 'CFE2F3' : 'C9DAF8' } },
      };
    }
  s.D2.s = { patternType: 'solid', fgColor: { rgb: 'CCDDF5' } };
  assert(P.readSchedule(s, roster).rows.find((r) => r.student === 'New Student').staff === 'ACD');
});
test('Yellow ACD and assessment blocks remain separate from pale-blue instructor', () => {
  const s = sheet([
    ['Time', 'IC 1', 'IC 2', 'HS1'],
    ['10:00', 'Samara A', 'Assessment', 'New Student'],
    ['10:05', '', 'Ryan C', 'yama'],
    ['10:10', '', '', 'table 1'],
    ['11:00', '', '', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'FFD966' } };
  for (const a of ['C2', 'C3']) s[a].s = { patternType: 'solid', fgColor: { rgb: '00FF00' } };
  for (const a of ['D2', 'D3', 'D4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.length === 2 &&
      p.rows.find((r) => r.student === 'Samara Adams').staff === 'ACD' &&
      p.rows.find((r) => r.student === 'New Student').staff === 'yama',
  );
});
test('Split needs/attention notes stay with their own student', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['10:00', 'Samara A'],
      ['10:05', 'needs'],
      ['10:10', 'attention'],
      ['11:00', 'Ryan C'],
    ]),
    roster,
  );
  assert(p.rows[0].notes.join('; ') === 'Needs attention' && p.rows[1].notes.length === 0);
});
test('Physics notes below names and inline normalize to Physics', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'HS1', 'HS2'],
      ['10:00', 'Arjun phy', 'Nylah S'],
      ['10:05', '', 'physics'],
      ['11:00', '', ''],
    ]),
    roster,
  );
  assert(
    p.rows.length === 2 &&
      p.rows.every((r) => r.notes.join('; ') === 'Physics') &&
      p.rows.some((r) => r.student === 'Arjun'),
  );
});
test('Lowercase name without table is instructor, not a student', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Arjun'],
    ['10:05', 'yama'],
    ['10:10', 'physics'],
    ['11:00', ''],
  ]);
  for (const a of ['B2', 'B3', 'B4']) s[a].s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  const p = P.readSchedule(s, roster);
  assert(
    p.rows.length === 1 &&
      p.rows[0].staff === 'yama' &&
      p.rows[0].student === 'Arjun' &&
      p.rows[0].notes[0] === 'Physics',
  );
});
test('Repeated two-hour session combines notes from both appearances', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'IC 1'],
      ['10:00', 'Samara A'],
      ['10:05', 'needs attention'],
      ['11:00', 'Samara Adams phy'],
    ]),
    roster,
  );
  assert(
    p.rows.length === 1 &&
      p.rows[0].end === '12:00' &&
      p.rows[0].notes.includes('Needs attention') &&
      p.rows[0].notes.includes('Physics'),
  );
});
test('Notes follow a student across nearby fill shades', () => {
  const s = sheet([
    ['Time', 'HS1'],
    ['10:00', 'Arjun'],
    ['10:05', 'physics'],
    ['11:00', ''],
  ]);
  s.B2.s = { patternType: 'solid', fgColor: { rgb: 'C9DAF8' } };
  s.B3.s = { patternType: 'solid', fgColor: { rgb: 'CFE2F3' } };
  assert(P.readSchedule(s, roster).rows[0].notes[0] === 'Physics');
});
test('Names repeated 30, 90, or 120 minutes apart do not make two-hour sessions', () => {
  for (const time of ['10:30', '11:30', '12:00']) {
    const p = P.readSchedule(
      sheet([
        ['Time', 'IC 1'],
        ['10:00', 'Samara A'],
        [time, 'Samara Adams'],
      ]),
      roster,
    );
    assert(p.rows.length === 2 && p.rows.every((r) => r.durationMinutes === 60));
  }
});
test('Same-time duplicate does not make a two-hour session', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'Alex', 'Alex'],
      ['10:00', 'Samara A', 'Samara Adams'],
    ]),
    roster,
  );
  assert(
    p.rows.length === 1 &&
      p.rows[0].durationMinutes === 60 &&
      p.warnings.some((w) => w.includes('Duplicate entry')),
  );
});
test('Two-hour pairing works in time order across lanes', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'Alex', 'Alex'],
      ['10:00', '', 'Samara A'],
      ['11:00', 'Samara Adams', ''],
    ]),
    roster,
  );
  assert(
    p.rows.length === 1 &&
      p.rows[0].start === '10:00' &&
      p.rows[0].end === '12:00' &&
      p.rows[0].durationMinutes === 120,
  );
});
test('Repeated student with different instructors remains separate', () => {
  const p = P.readSchedule(
    sheet([
      ['Time', 'Alex', 'Jordan'],
      ['10:00', 'Samara A', ''],
      ['11:00', '', 'Samara Adams'],
    ]),
    roster,
  );
  assert(p.rows.length === 2 && p.rows.every((r) => r.durationMinutes === 60));
});
