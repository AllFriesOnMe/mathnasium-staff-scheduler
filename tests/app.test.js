function finish() {
  document.getElementById('testResults').textContent =
    `${passed.length} passed; ${failures.length} failed\n` +
    [...passed.map((n) => 'PASS ' + n), ...failures.map((n) => 'FAIL ' + n)].join('\n');
  document.title = failures.length ? 'FAIL' : 'PASS';
}
const frame = document.createElement('iframe');
frame.src = '../test-app.html';
frame.width = 750;
frame.height = 800;
frame.onload = async () => {
  try {
    const d = frame.contentDocument,
      w = frame.contentWindow;
    const book = (...sheets) => {
      const b = XLSX.utils.book_new();
      sheets.forEach(([s, n]) => XLSX.utils.book_append_sheet(b, s, n));
      return XLSX.write(b, { type: 'array', bookType: 'xlsx' });
    };
    const upload = async (id, bytes) => {
      const dt = new w.DataTransfer();
      dt.items.add(new w.File([bytes], id + '.xlsx'));
      d.getElementById(id).files = dt.files;
      d.getElementById(id).dispatchEvent(new w.Event('change'));
      const status = d.getElementById(id === 'binders' ? 'binderStatus' : 'scheduleStatus');
      for (let tries = 0; tries < 1500 && status.textContent === 'Reading spreadsheet...'; tries++)
        await new Promise((resolve) => setTimeout(resolve, 10));
      assert(status.textContent.includes('loaded'), status.textContent);
    };
    await upload(
      'binders',
      book([
        sheet([
          ['Student', 'Binder'],
          ['Samara Adams', 'A1'],
          ['Ryan Clark', 'B2'],
          ['Veer Shah', 'E5'],
        ]),
        'Binders',
      ]),
    );
    await upload(
      'schedule',
      book(
        [schedule, 'Monday'],
        [
          sheet([
            ['Staff', 'Start', 'End', 'Student'],
            ['Alex', '15:30', '16:30', 'Veer Shah'],
          ]),
          'Tuesday',
        ],
      ),
    );
    test('Upload controls parse actual XLSX files', () =>
      assert(d.getElementById('binderStatus').textContent.includes('loaded')));
    test('Multiple workbook sheets selectable', () =>
      assert(
        d.getElementById('scheduleSheet').options.length === 2 &&
          !d.getElementById('scheduleSheet').hidden,
      ));
    await w.generate();
    test('Upload-to-output flow joins binders and schedules', () =>
      assert(
        !d.getElementById('results').hidden &&
          d.getElementById('staffLists').textContent.includes('Samara Adams (A1)'),
      ));
    test('Removed review and template UI stays removed', () =>
      assert(
        !d.getElementById('appointments') &&
          !d.getElementById('binderTemplate') &&
          !d.getElementById('scheduleTemplate'),
      ));
    d.getElementById('scheduleSheet').value = 'Tuesday';
    d.getElementById('scheduleSheet').dispatchEvent(new w.Event('change'));
    test('Switching sheets invalidates previous results', () =>
      assert(d.getElementById('results').hidden));
    await w.generate();
    test('Selected day generates its own staff schedule', () =>
      assert(
        d.getElementById('staffLists').textContent.includes('Alex') &&
          d.getElementById('staffLists').textContent.includes('Veer Shah (E5)') &&
          !d.getElementById('staffLists').textContent.includes('Samara'),
      ));
    await upload(
      'schedule',
      book([
        sheet([
          ['Staff', 'Start', 'End', 'Student'],
          ['Alex', '15:30', '16:30', 'Missing Student'],
        ]),
        'Unknown',
      ]),
    );
    await w.generate();
    test('Missing binders omit parentheses and retain non-HS warning', () =>
      assert(
        d.getElementById('staffLists').textContent.includes('Missing Student') &&
          !d.getElementById('staffLists').textContent.includes('Missing Student (') &&
          !d.getElementById('staffLists').textContent.includes('CHECK BINDER') &&
          !d.getElementById('warnings').hidden,
      ));
    const hs = sheet([
      ['Start Time', 'HS1'],
      ['15:30', 'Arjun'],
      ['15:35', 'g12'],
      ['16:30', ''],
    ]);
    hs['!merges'] = [XLSX.utils.decode_range('B2:B3')];
    await upload('schedule', book([hs, 'High school']));
    await w.generate();
    test('High-school student without binder renders normally', () =>
      assert(
        !d.getElementById('results').hidden &&
          d.getElementById('staffLists').textContent.includes('15:30 - 16:30: Arjun') &&
          !d.getElementById('staffLists').textContent.includes('Arjun (') &&
          d.getElementById('warnings').hidden,
      ));
    let exportedBlob;
    w.URL.createObjectURL = (blob) => {
      exportedBlob = blob;
      return 'blob:test';
    };
    w.HTMLAnchorElement.prototype.click = () => {};
    d.getElementById('export').click();
    const csv = await exportedBlob.text();
    test('CSV leaves binder field blank for high-school student', () =>
      assert(csv.includes('"HS1","15:30","16:30","Arjun",""') && !csv.includes('CHECK BINDER')));
    const marked = sheet([
      ['Time', 'IC 1'],
      ['10:00', 'Samara A'],
      ['10:05', 'needs'],
      ['10:10', 'attention'],
      ['10:15', 'phy'],
      ['11:00', ''],
    ]);
    await upload('schedule', book([marked, 'Notes']));
    await w.generate();
    test('Generated staff list shows attention and physics', () =>
      assert(
        d
          .getElementById('staffLists')
          .textContent.includes('Samara Adams (A1) [Needs attention; Physics]'),
      ));
    d.getElementById('export').click();
    const noteCsv = await exportedBlob.text();
    test('CSV includes a Notes column and student flags', () =>
      assert(noteCsv.includes('"Notes"') && noteCsv.includes('"Needs attention; Physics"')));
    const extended = sheet([
      ['Time', 'Alex', 'Alex', 'Alex'],
      ['10:00', 'Samara A', 'Ryan C', 'Veer S'],
      ['11:00', 'Samara Adams', 'Ryan Clark', ''],
      ['12:00', '', '', ''],
    ]);
    await upload('schedule', book([extended, 'Two hours']));
    await w.generate();
    const lines = [...d.querySelectorAll('#staffLists li')].map((li) => li.textContent);
    test('Two-hour sessions have a labeled timestamp and their own line', () =>
      assert(
        lines.length === 3 &&
          lines.filter((t) => t.startsWith('10:00 - 12:00 (2 hours):')).length === 2 &&
          lines.some((t) => t.includes('Samara Adams') && !t.includes('Ryan Clark')) &&
          lines.some((t) => t.includes('Ryan Clark') && !t.includes('Samara Adams')),
      ));
    test('One-hour student remains separate from two-hour students', () =>
      assert(lines.some((t) => t === '10:00 - 11:00: Veer Shah (E5)')));
    d.getElementById('export').click();
    const durationCsv = await exportedBlob.text();
    test('CSV exports explicit two-hour duration', () =>
      assert(durationCsv.includes('"Duration"') && durationCsv.includes('"2 hours"')));
    const untrustedNames = [
      '=SUM(1,1)',
      '+SUM(1,1)',
      '@SUM(1,1)',
      '-1+2',
      '<img src=x onerror=alert(1)>',
    ];
    const untrusted = sheet([
      ['Staff', 'Start', 'End', 'Student'],
      ...untrustedNames.map((name) => ['HS1', '10:00', '11:00', name]),
    ]);
    await upload('schedule', book([untrusted, 'Untrusted text']));
    await w.generate();
    test('Spreadsheet HTML is displayed as text without creating elements', () =>
      assert(
        d.getElementById('staffLists').textContent.includes('<img src=x onerror=alert(1)>') &&
          !d.querySelector('#staffLists img'),
      ));
    d.getElementById('export').click();
    const safeCsv = await exportedBlob.text();
    const safeBook = XLSX.read(safeCsv, { type: 'string' });
    const safeRows = XLSX.utils.sheet_to_json(safeBook.Sheets[safeBook.SheetNames[0]], {
      header: 1,
    });
    test('CSV export neutralizes spreadsheet formula prefixes', () =>
      assert(
        untrustedNames.slice(0, 4).every((name) => safeRows.some((row) => row[3] === "'" + name)),
      ));
  } catch (error) {
    failures.push('Browser integration: ' + error.message);
  }
  finish();
};
document.body.append(frame);
