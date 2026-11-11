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
    test('12-hour output handles midnight, noon, evening, and assessment notes', () => {
      assert(
        w.displayTime('00:00') === '12:00 AM' &&
          w.displayTime('12:00') === '12:00 PM' &&
          w.displayTime('19:30') === '7:30 PM',
      );
      assert(w.displayNote('Assessment: 11:45 - 12:00') === 'Assessment: 11:45 AM - 12:00 PM');
      assert(w.displayNote('Handoff at 7:30 PM') === 'Handoff at 7:30 PM');
    });
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
          d.getElementById('staffLists').textContent.includes('3:30 PM - 4:30 PM: Arjun') &&
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
      assert(
        csv.includes('"HS1","3:30 PM","4:30 PM","Arjun",""') && !csv.includes('CHECK BINDER'),
      ));
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
        d.getElementById('staffLists').textContent.includes('Samara Adams (A1) [Physics]') &&
          d.querySelector('#staffLists strong.needs-attention')?.textContent === 'Samara Adams' &&
          w
            .getComputedStyle(d.querySelector('#staffLists strong.needs-attention'))
            .textDecorationLine.includes('underline') &&
          !d.getElementById('staffLists').textContent.includes('Needs attention'),
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
    const lines = [...d.querySelectorAll('#staffLists li')].map((li) => {
      const content = li.cloneNode(true);
      content.querySelectorAll('.student-controls').forEach((controls) => controls.remove());
      return content.textContent;
    });
    test('Two-hour sessions have a labeled timestamp and their own line', () =>
      assert(
        lines.length === 3 &&
          lines.filter((t) => t.startsWith('10:00 AM - 12:00 PM (2 hours):')).length === 2 &&
          lines.some((t) => t.includes('Samara Adams') && !t.includes('Ryan Clark')) &&
          lines.some((t) => t.includes('Ryan Clark') && !t.includes('Samara Adams')),
      ));
    test('One-hour student remains separate from two-hour students', () =>
      assert(lines.some((t) => t === '10:00 AM - 11:00 AM: Veer Shah (E5)')));
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
    const handoffGrid = sheet([
      ['Time', 'Alex', 'Abigail', 'Notes'],
      ['10:00', 'Samara A', '', ''],
      ['10:30', '', '', 'Send Samara to Abigail at 10:45'],
      ['10:45', '', '', ''],
      ['11:00', '', '', ''],
    ]);
    await upload('schedule', book([handoffGrid, 'Handoff']));
    await w.generate();
    const staffSections = [...d.querySelectorAll('#staffLists section')];
    test('Both involved instructors receive the student and their own handoff note', () =>
      assert(
        staffSections.some(
          (section) =>
            section.querySelector('h3').textContent === 'Alex' &&
            section.textContent.includes('10:00 AM - 10:45 AM: Samara Adams (A1)') &&
            section.textContent.includes('Handoff to Abigail at 10:45 AM'),
        ) &&
          staffSections.some(
            (section) =>
              section.querySelector('h3').textContent === 'Abigail' &&
              section.textContent.includes('10:45 AM - 11:00 AM: Samara Adams (A1)') &&
              section.textContent.includes('Taking over from Alex at 10:45 AM'),
          ),
      ));
    test('Normal handoff segments do not trigger an overlapping-booking warning', () =>
      assert(!d.getElementById('warningList').textContent.includes('Overlapping bookings')));
    d.getElementById('export').click();
    const handoffCsv = await exportedBlob.text();
    const handoffBook = XLSX.read(handoffCsv, { type: 'string', raw: true });
    const handoffRows = XLSX.utils.sheet_to_json(handoffBook.Sheets[handoffBook.SheetNames[0]], {
      header: 1,
    });
    test('CSV contains both instructor segments, binders, and handoff notes', () =>
      assert(
        handoffRows.some(
          (row) =>
            row[0] === 'Alex' &&
            row[1] === '10:00 AM' &&
            row[2] === '10:45 AM' &&
            row[4] === 'A1' &&
            row[5].includes('Handoff to Abigail'),
        ) &&
          handoffRows.some(
            (row) =>
              row[0] === 'Abigail' &&
              row[1] === '10:45 AM' &&
              row[2] === '11:00 AM' &&
              row[4] === 'A1' &&
              row[5].includes('Taking over from Alex'),
          ),
      ));
    const storageKey = 'mathnasium.schedule-colors.v1';
    test('Fixed director controls are gone and Assessment is an ordinary removable indicator', () => {
      const row = d.querySelector('#colorIndicators .custom-indicator');
      assert(!d.getElementById('directorColor') && !d.getElementById('assessmentColor'));
      assert(
        row.querySelector('input[type=text]').value === 'Assessment' &&
          row.querySelector('select').value === 'assessment',
      );
      assert(
        row.textContent.includes('Not set') && row.querySelector('button').textContent === 'Remove',
      );
    });
    d.querySelector('#colorIndicators button').click();
    test('Removing the default Assessment invalidates output and saves an empty list', () =>
      assert(
        d.getElementById('results').hidden &&
          JSON.parse(w.localStorage.getItem(storageKey)).colorIndicators.length === 0,
      ));
    async function openSettingsCopy() {
      const copy = document.createElement('iframe');
      const loaded = new Promise((resolve) => (copy.onload = resolve));
      copy.src = '../test-app.html';
      document.body.append(copy);
      await loaded;
      return copy;
    }
    const emptyCopy = await openSettingsCopy();
    test('Removed Assessment stays removed on the next visit', () =>
      assert(!emptyCopy.contentDocument.querySelector('#colorIndicators .custom-indicator')));
    emptyCopy.remove();
    w.localStorage.setItem(
      storageKey,
      JSON.stringify({ directorColor: '#FFD966', assessmentColor: '#00FF00' }),
    );
    const legacyCopy = await openSettingsCopy();
    test('Previously saved fixed colors migrate to removable indicators', () => {
      const rows = [
        ...legacyCopy.contentDocument.querySelectorAll('#colorIndicators .custom-indicator'),
      ];
      assert(
        rows.length === 2 &&
          rows.some(
            (row) =>
              row.querySelector('input[type=text]').value === 'Center director' &&
              row.querySelector('input[type=color]').value === '#ffd966',
          ),
      );
      assert(
        rows.some(
          (row) =>
            row.querySelector('select').value === 'assessment' &&
            row.querySelector('input[type=color]').value === '#00ff00',
        ),
      );
    });
    legacyCopy.remove();
    w.localStorage.setItem(storageKey, JSON.stringify({ colorIndicators: [] }));
    d.getElementById('addColorIndicator').click();
    let indicator = d.querySelector('#colorIndicators .custom-indicator');
    test('Add color indicator creates an editable row with no selected color', () =>
      assert(
        indicator &&
          indicator.textContent.includes('Not set') &&
          indicator.querySelector('input[type=text]').value === '',
      ));
    const nameInput = indicator.querySelector('input[type=text]');
    nameInput.value = 'Lilianna K';
    nameInput.dispatchEvent(new w.Event('input'));
    const customColor = indicator.querySelector('input[type=color]');
    customColor.value = '#ffd966';
    customColor.dispatchEvent(new w.Event('change'));
    test('Named color indicators are persisted and included in parser options', () => {
      const saved = JSON.parse(w.localStorage.getItem(storageKey));
      assert(
        saved.colorIndicators[0].name === 'Lilianna K' &&
          saved.colorIndicators[0].color === '#FFD966',
      );
    });
    const customReload = document.createElement('iframe');
    const customLoaded = new Promise((resolve) => (customReload.onload = resolve));
    customReload.src = '../test-app.html';
    document.body.append(customReload);
    await customLoaded;
    test('Custom indicator names and colors return on the next visit', () => {
      const restored = customReload.contentDocument.querySelector(
        '#colorIndicators .custom-indicator',
      );
      assert(
        restored.querySelector('input[type=text]').value === 'Lilianna K' &&
          restored.querySelector('input[type=color]').value === '#ffd966',
      );
    });
    customReload.remove();
    d.getElementById('addColorIndicator').click();
    const secondIndicator = d.querySelectorAll('#colorIndicators .custom-indicator')[1];
    const secondColor = secondIndicator.querySelector('input[type=color]');
    secondColor.value = '#ffd966';
    secondColor.dispatchEvent(new w.Event('change'));
    test('Duplicate custom colors are rejected without replacing the existing assignment', () =>
      assert(
        d.getElementById('colorPreferenceStatus').textContent.includes('different colors') &&
          secondIndicator.textContent.includes('Not set'),
      ));
    const role = d.querySelector('#colorIndicators .custom-indicator select');
    role.value = 'assessment';
    role.dispatchEvent(new w.Event('change'));
    test('Indicator behavior can be changed to assessment', () =>
      assert(
        JSON.parse(w.localStorage.getItem(storageKey)).colorIndicators[0].behavior === 'assessment',
      ));
    while (d.querySelector('#colorIndicators button'))
      d.querySelector('#colorIndicators button').click();
    test('Removing indicators also removes their saved mappings', () =>
      assert(
        !d.querySelector('#colorIndicators .custom-indicator') &&
          JSON.parse(w.localStorage.getItem(storageKey)).colorIndicators.length === 0,
      ));
    w.localStorage.removeItem(storageKey);
    const aiGrid = sheet([
      ['Time', 'elizabeth', 'abigail', 'Notes'],
      ['18:30', '', 'Onna Kohut', ''],
      ['18:40', 'Raya Holt', '', ''],
      ['19:00', '', '', 'Take over Raya Holt from elizabeth at 19:30'],
      ['19:30', '', '', ''],
      ['19:40', '', '', ''],
    ]);
    await upload(
      'binders',
      book([
        sheet([
          ['Student', 'Binder'],
          ['Raya Holt', 'A1'],
          ['Onna Kohut', 'B2'],
        ]),
        'Binders',
      ]),
    );
    await upload('schedule', book([aiGrid, 'AI example']));
    const originalFetch = w.fetch;
    let aiCalls = 0,
      sentContext;
    w.fetch = async (url, options) => {
      aiCalls++;
      sentContext = JSON.parse(options.body);
      assert(
        url === '/api/handoffs' && options.headers.Authorization === 'Bearer synthetic-center-code',
      );
      const item = sentContext.cases[0];
      return Response.json({
        resolutions: [
          {
            cell: item.cell,
            text: item.text,
            time: '19:30',
            incomingInstructor: 'abigail',
            studentIds: [item.appointments[0].id],
          },
        ],
      });
    };
    try {
      await w.generate();
      test('AI is off by default and sends no unresolved schedule data', () =>
        assert(
          !d.getElementById('aiEnabled').checked &&
            aiCalls === 0 &&
            !d.getElementById('results').hidden,
        ));
      d.getElementById('aiEnabled').click();
      d.getElementById('aiAccessCode').value = 'synthetic-center-code';
      await w.generate();
      test('Optional AI result flows through real worker validation to both schedules', () => {
        const text = d.getElementById('staffLists').textContent;
        assert(
          aiCalls === 1 &&
            text.includes('Handoff to abigail') &&
            text.includes('Taking over from elizabeth'),
          JSON.stringify({ aiCalls, text, status: d.getElementById('aiStatus').textContent }),
        );
        assert(text.split('AI-assisted handoff').length === 3);
        assert(d.getElementById('aiStatus').textContent.includes('AI-assisted handoff'));
      });
      test('AI request contains bounded context without binder roster or credentials in its body', () => {
        assert(sentContext.cases.length === 1 && sentContext.cases[0].nearby.length <= 45);
        assert(
          !JSON.stringify(sentContext).includes('synthetic-center-code') && !sentContext.binders,
        );
      });
      w.fetch = async () => Response.json({ error: 'Test unavailable' }, { status: 503 });
      await w.generate();
      test('AI outage leaves usable local output and its unresolved handoff warning', () =>
        assert(
          !d.getElementById('results').hidden &&
            d.getElementById('warningList').textContent.includes('handoff not applied') &&
            d.getElementById('aiStatus').textContent.includes('Test unavailable') &&
            !d.getElementById('staffLists').textContent.includes('AI-assisted handoff'),
        ));
      d.getElementById('aiEnabled').click();
      test('Disabling AI clears the session access code', () =>
        assert(
          d.getElementById('aiAccessCode').value === '' &&
            d.getElementById('aiAccessSettings').hidden,
        ));
    } finally {
      w.fetch = originalFetch;
    }
    // SheetJS's writer omits styles; supply a styled sheet to the real worker
    // after upload so this exercises color parsing as well as browser opt-in.
    const colorGrid = ambiguousColorFixture();
    colorGrid.B2.v = 'Samara A (OL)';
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
    const colorWorker = w.Worker;
    w.Worker = class extends colorWorker {
      postMessage(task, ...rest) {
        return super.postMessage(
          task.type === 'generate' ? { ...task, schedule: colorGrid } : task,
          ...rest,
        );
      }
    };
    let colorCalls = 0;
    w.fetch = async (url, options) => {
      colorCalls++;
      const item = JSON.parse(options.body).cases[0];
      assert(
        item.kind === 'color' &&
          item.color === 'rgb:AFC0E0' &&
          item.candidates.some((candidate) => candidate.instructor === 'yama'),
      );
      return Response.json({
        resolutions: [
          {
            kind: 'color',
            color: item.color,
            cell: item.cell,
            text: item.text,
            time: item.time,
            incomingInstructor: 'yama',
            studentIds: item.appointments.map((row) => row.id),
          },
        ],
      });
    };
    try {
      await w.generate();
      test('Uncertain color stays local with a warning when AI is disabled', () =>
        assert(
          colorCalls === 0 &&
            d.getElementById('staffLists').textContent.includes('ACD') &&
            d.getElementById('warningList').textContent.includes('uncertain shade'),
        ));
      test('Table subheaders replace instructor table suffixes and retain online markers', () => {
        const alex = [...d.querySelectorAll('#staffLists section')].find(
          (section) => section.querySelector('h3').textContent === 'alex',
        );
        assert(
          alex &&
            alex.querySelector('h4').textContent === 'Table 1' &&
            alex.textContent.includes('Samara Adams (ONLINE) (A1)'),
        );
        assert(
          [...alex.querySelectorAll('strong')].some((label) => label.textContent === '(ONLINE)'),
        );
      });
      d.getElementById('aiEnabled').click();
      d.getElementById('aiAccessCode').value = 'synthetic-center-code';
      await w.generate();
      test('AI color-only request updates the correct instructor without a handoff', () => {
        const section = [...d.querySelectorAll('#staffLists section')].find(
          (section) => section.querySelector('h3').textContent === 'yama',
        );
        assert(
          colorCalls === 1 &&
            section.textContent.includes('Veer Shah (E5)') &&
            section.textContent.includes('AI-assisted color assignment'),
        );
        assert(
          !d.getElementById('warningList').textContent.includes('D2') &&
            !d.getElementById('staffLists').textContent.includes('Handoff'),
        );
        assert(d.getElementById('aiStatus').textContent.includes('AI-assisted color assignment'));
      });
      d.getElementById('export').click();
      const colorCsv = await exportedBlob.text();
      test('CSV retains online labels and exports Table and Online columns', () => {
        const workbook = XLSX.read(colorCsv, { type: 'string', raw: true });
        const entries = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        const samara = entries.find((row) => row.Student === 'Samara Adams (ONLINE)');
        assert(
          samara &&
            String(samara.Table) === '1' &&
            samara.Online === 'Yes' &&
            samara.Binder === 'A1',
        );
      });
      test('AI color labels are included in CSV output', () =>
        assert(
          colorCsv.includes('AI-assisted color assignment') && colorCsv.includes('Veer Shah'),
        ));
      w.fetch = async () => Response.json({ resolutions: [] });
      await w.generate();
      test('Uncertain AI color response preserves the ACD fallback and warning', () =>
        assert(
          d.getElementById('staffLists').textContent.includes('ACD') &&
            d.getElementById('warningList').textContent.includes('D2') &&
            !d.getElementById('staffLists').textContent.includes('AI-assisted color assignment'),
        ));
    } finally {
      w.fetch = originalFetch;
      w.Worker = colorWorker;
      if (d.getElementById('aiEnabled').checked) d.getElementById('aiEnabled').click();
    }
    await upload(
      'schedule',
      book([
        sheet([
          ['Time', 'Alex', 'Jordan'],
          ['10:00', 'Samara A', 'Ryan C'],
          ['11:00', '', ''],
        ]),
        'Manual edits',
      ]),
    );
    await w.generate();
    const findStaff = (staff) =>
      [...d.querySelectorAll('#staffLists section')].find(
        (section) => section.querySelector('h3').textContent === staff,
      );
    const clickButton = (parent, text) =>
      [...parent.querySelectorAll('button')].find((button) => button.textContent === text).click();
    clickButton(findStaff('Alex'), 'Add student');
    let editor = d.querySelector('dialog');
    editor.querySelector('[name="studentname"]').value = 'Veer Shah';
    editor.querySelector('[name="starttime"]').value = '12:00';
    editor.querySelector('[name="starttime"]').dispatchEvent(new w.Event('change'));
    test('Add student defaults to one hour and supplies binder suggestions', () => {
      assert(editor.querySelector('[name="endtime"]').value === '13:00');
      assert(
        [...editor.querySelectorAll('datalist option')].some(
          (option) => option.value === 'Veer Shah',
        ),
      );
    });
    editor.querySelector('[name="online"]').checked = true;
    editor.querySelector('[name="additionallabels"]').value = 'Physics; New binder';
    clickButton(editor, 'Add handoff');
    editor.querySelector('[name="handoffStaff"]').value = 'Jordan';
    editor.querySelector('[name="handoffTime"]').value = '12:25';
    editor.querySelector('form').requestSubmit();
    test('Manual add appears on both instructors with binder and online status', () => {
      assert(
        findStaff('Alex').textContent.includes('12:00 PM - 12:25 PM: Veer Shah (ONLINE) (E5)'),
      );
      assert(
        findStaff('Jordan').textContent.includes('12:25 PM - 1:00 PM: Veer Shah (ONLINE) (E5)'),
      );
    });
    test('Additional labels appear on both instructors and prefill the edit form', () => {
      assert(findStaff('Alex').textContent.includes('Physics; New binder'));
      assert(findStaff('Jordan').textContent.includes('Physics; New binder'));
    });
    findStaff('Alex').querySelector('.transfer-time').click();
    editor = d.querySelector('dialog');
    test('Additional labels can be changed in the edit prompt', () => {
      assert(editor.querySelector('[name="additionallabels"]').value === 'Physics\nNew binder');
      editor.querySelector('[name="additionallabels"]').value = 'Bring workbook';
    });
    editor.querySelector('[name="handoffTime"]').value = '13:15';
    editor.querySelector('form').requestSubmit();
    test('Out-of-range handoff input keeps the dialog open and existing schedule intact', () => {
      assert(
        editor.open &&
          editor.querySelector('[role="alert"]').textContent.includes('strictly inside'),
      );
      assert(findStaff('Jordan').textContent.includes('12:25 PM'));
    });
    editor.querySelector('[name="handoffTime"]').value = '12:17';
    editor.querySelector('form').requestSubmit();
    test('Clicking the transfer box adjusts both schedules', () => {
      assert(findStaff('Alex').textContent.includes('12:00 PM - 12:17 PM'));
      assert(findStaff('Jordan').textContent.includes('12:17 PM - 1:00 PM'));
      assert(!d.getElementById('staffLists').textContent.includes('12:25 PM'));
      assert(findStaff('Alex').textContent.includes('Bring workbook'));
      assert(findStaff('Jordan').textContent.includes('Bring workbook'));
      assert(!d.getElementById('staffLists').textContent.includes('New binder'));
    });
    const dragged = [...findStaff('Alex').querySelectorAll('.editable-student')].find((node) =>
      node.textContent.includes('Veer Shah'),
    );
    const dragData = new w.DataTransfer();
    dragged.dispatchEvent(new w.DragEvent('dragstart', { bubbles: true, dataTransfer: dragData }));
    findStaff('Jordan').dispatchEvent(
      new w.DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dragData }),
    );
    test('Drag and drop moves a portion and merges adjacent portions for the same instructor', () => {
      assert(!findStaff('Alex').textContent.includes('Veer Shah'));
      assert(findStaff('Jordan').textContent.includes('12:00 PM - 1:00 PM: Veer Shah'));
      assert(!findStaff('Jordan').querySelector('.transfer-time'));
    });
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    test('Undo restores both sides of a transfer after drag', () =>
      assert(findStaff('Alex').textContent.includes('12:17 PM')));
    d.getElementById('export').click();
    const editedCsv = await exportedBlob.text();
    test('CSV reflects manually adjusted handoff times', () =>
      assert(editedCsv.includes('12:17 PM') && !editedCsv.includes('12:25 PM')));
    test('CSV includes edited additional labels', () =>
      assert(editedCsv.includes('Bring workbook') && !editedCsv.includes('New binder')));
    const removeStudent = [...findStaff('Alex').querySelectorAll('.editable-student')].find(
      (node) => node.textContent.includes('Veer Shah'),
    );
    removeStudent.querySelector('[aria-label^="Remove "]').click();
    test('Remove deletes a student from all involved instructors', () =>
      assert(!d.getElementById('staffLists').textContent.includes('Veer Shah')));
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    test('Undo restores a removed student and handoffs', () =>
      assert(findStaff('Jordan').textContent.includes('Veer Shah')));
    clickButton(findStaff('Alex'), 'Add table');
    editor = d.querySelector('dialog');
    editor.querySelector('[name="tablenumber"]').value = '7';
    editor.querySelector('form').requestSubmit();
    test('Add table creates a table subheader under the instructor', () => {
      assert(findStaff('Alex').querySelector('[data-table="7"] h4').textContent === 'Table 7');
      assert(findStaff('Alex').querySelector('h3').textContent === 'Alex');
    });
    const veer = [...findStaff('Alex').querySelectorAll('.editable-student')].find((node) =>
      node.textContent.includes('Veer Shah'),
    );
    clickButton(veer, 'Edit');
    editor = d.querySelector('dialog');
    editor.querySelector('[name="table"]').value = '7';
    editor.querySelector('[name="handoffTable"]').value = '9';
    editor.querySelector('form').requestSubmit();
    test('Edit assigns each handoff portion its own table subheader', () => {
      assert(findStaff('Alex').querySelector('[data-table="7"]').textContent.includes('Veer Shah'));
      assert(
        findStaff('Jordan').querySelector('[data-table="9"]').textContent.includes('Veer Shah'),
      );
    });
    d.getElementById('export').click();
    const tableCsv = await exportedBlob.text();
    test('CSV exports each edited student table number', () => {
      const book = XLSX.read(tableCsv, { type: 'string', raw: true });
      const entries = XLSX.utils
        .sheet_to_json(book.Sheets[book.SheetNames[0]])
        .filter((row) => row.Student === 'Veer Shah (ONLINE)');
      assert(
        entries.length === 2 &&
          entries.some((row) => row.Staff === 'Alex' && String(row.Table) === '7') &&
          entries.some((row) => row.Staff === 'Jordan' && String(row.Table) === '9'),
      );
    });
    const tableDrag = new w.DataTransfer();
    findStaff('Alex')
      .querySelector('[data-table="7"] .editable-student')
      .dispatchEvent(new w.DragEvent('dragstart', { bubbles: true, dataTransfer: tableDrag }));
    findStaff('Jordan')
      .querySelector('[data-table="9"]')
      .dispatchEvent(
        new w.DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: tableDrag }),
      );
    test('Dropping on a table moves the student to that specific table', () => {
      assert(!findStaff('Alex').textContent.includes('Veer Shah'));
      assert(
        findStaff('Jordan')
          .querySelector('[data-table="9"]')
          .textContent.includes('12:00 PM - 1:00 PM: Veer Shah'),
      );
    });
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    test('Undo restores student tables and removes a manually added empty table', () => {
      assert(!findStaff('Alex').querySelector('[data-table="7"]'));
      assert(findStaff('Alex').querySelector('[data-table=""]').textContent.includes('Veer Shah'));
    });
    await w.generate();
    test('Regenerating clears manual edits and undo history', () => {
      assert(!d.getElementById('staffLists').textContent.includes('Veer Shah'));
      assert(
        [...d.getElementById('editToolbar').querySelectorAll('button')].find(
          (button) => button.textContent === 'Undo last edit',
        ).disabled,
      );
    });
    const samara = [...findStaff('Alex').querySelectorAll('.editable-student')].find((node) =>
      node.textContent.includes('Samara Adams'),
    );
    clickButton(samara, 'Edit');
    editor = d.querySelector('dialog');
    editor.querySelector('[name="table"]').value = '1';
    clickButton(editor, 'Add handoff');
    editor.querySelector('[name="handoffStaff"]').value = 'Alex';
    editor.querySelector('[name="handoffTable"]').value = '2';
    editor.querySelector('[name="handoffTime"]').value = '10:30';
    editor.querySelector('form').requestSubmit();
    test('A mid-session table switch splits the displayed time and shows the switch note', () => {
      const first = findStaff('Alex').querySelector('[data-table="1"]');
      const second = findStaff('Alex').querySelector('[data-table="2"]');
      assert(
        first.textContent.includes('10:00 AM - 10:30 AM: Samara Adams (A1) (switch to table 2)'),
      );
      assert(second.textContent.includes('10:30 AM - 11:00 AM: Samara Adams (A1)'));
      assert(first.querySelector('.transfer-time').textContent === '(switch to table 2)');
    });
    d.getElementById('export').click();
    const switchCsv = await exportedBlob.text();
    test('CSV includes both time portions and the table switch note', () => {
      assert(switchCsv.includes('(switch to table 2)') && switchCsv.includes('10:30 AM'));
    });
    clickButton(findStaff('Alex').querySelector('[data-table="2"]'), 'Edit table');
    editor = d.querySelector('dialog');
    test('Edit table pre-fills the current number', () =>
      assert(editor.querySelector('[name="tablenumber"]').value === '2'));
    editor.querySelector('[name="tablenumber"]').value = '4';
    editor.querySelector('form').requestSubmit();
    test('Renaming a table updates its header, student assignment and switch note', () => {
      assert(!findStaff('Alex').querySelector('[data-table="2"]'));
      assert(
        findStaff('Alex').querySelector('[data-table="4"]').textContent.includes('Samara Adams'),
      );
      assert(
        findStaff('Alex')
          .querySelector('[data-table="1"]')
          .textContent.includes('(switch to table 4)'),
      );
    });
    d.getElementById('export').click();
    const renamedCsv = await exportedBlob.text();
    test('Renamed table is reflected in CSV', () => {
      const book = XLSX.read(renamedCsv, { type: 'string', raw: true });
      const entries = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]]);
      assert(entries.some((row) => row.Student === 'Samara Adams' && String(row.Table) === '4'));
      assert(renamedCsv.includes('(switch to table 4)'));
    });
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    test('Undo restores the original table number and switch note', () => {
      assert(
        findStaff('Alex').querySelector('[data-table="2"]').textContent.includes('Samara Adams'),
      );
      assert(
        findStaff('Alex')
          .querySelector('[data-table="1"]')
          .textContent.includes('(switch to table 2)'),
      );
    });
    clickButton(findStaff('Alex'), 'Add a note');
    editor = d.querySelector('dialog');
    editor.querySelector('[name="staffnote"]').value = 'Check binders\nUse <b>paper</b>.';
    editor.querySelector('form').requestSubmit();
    test('Instructor notes appear beneath only their own schedule and render as text', () => {
      const notes = findStaff('Alex').querySelector('.instructor-notes');
      assert(
        notes.textContent.includes('Check binders\nUse <b>paper</b>.') && !notes.querySelector('b'),
      );
      assert(!findStaff('Jordan').textContent.includes('Check binders'));
    });
    clickButton(findStaff('Alex').querySelector('.instructor-notes'), 'Edit note');
    editor = d.querySelector('dialog');
    test('Edit note pre-fills the instructor note', () =>
      assert(
        editor.querySelector('[name="staffnote"]').value === 'Check binders\nUse <b>paper</b>.',
      ));
    editor.querySelector('[name="staffnote"]').value = 'Return calculators';
    editor.querySelector('form').requestSubmit();
    d.getElementById('export').click();
    const notesCsv = await exportedBlob.text();
    test('CSV exports instructor notes as separate rows without assigning a student', () => {
      const book = XLSX.read(notesCsv, { type: 'string', raw: true });
      const entries = XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]]);
      const note = entries.find((row) => row.Notes === 'Return calculators');
      assert(note && note.Staff === 'Alex' && !note.Student && !note.Start);
    });
    clickButton(findStaff('Alex').querySelector('.instructor-notes'), 'Remove note');
    test('Instructor notes can be removed', () =>
      assert(!findStaff('Alex').textContent.includes('Return calculators')));
    clickButton(d.getElementById('editToolbar'), 'Undo last edit');
    test('Undo restores instructor notes', () =>
      assert(findStaff('Alex').textContent.includes('Return calculators')));
    await w.generate();
    test('Regeneration clears instructor notes along with other manual edits', () =>
      assert(!d.getElementById('staffLists').textContent.includes('Return calculators')));
    findStaff('Alex').querySelector('.student-name').click();
    editor = d.querySelector('dialog');
    test('Clicking a student opens recorded assignment evidence with cell references', () => {
      assert(editor.open && editor.textContent.includes('Why this assignment?'));
      assert(editor.textContent.includes('B2') && editor.textContent.includes('A2'));
      assert(
        editor.textContent.includes('one-hour') && editor.textContent.includes('column heading'),
      );
      assert(editor.textContent.includes('No takeover was applied'));
    });
    clickButton(editor, 'Close');
    const currentSamara = [...findStaff('Alex').querySelectorAll('.editable-student')].find(
      (node) => node.textContent.includes('Samara Adams'),
    );
    clickButton(currentSamara, 'Edit');
    editor = d.querySelector('dialog');
    editor.querySelector('[name="table"]').value = '6';
    editor.querySelector('form').requestSubmit();
    findStaff('Alex').querySelector('.student-name').click();
    editor = d.querySelector('dialog');
    test('Manual explanation distinguishes user edits from original spreadsheet evidence', () => {
      assert(
        editor.textContent.includes('Manual edit history') &&
          editor.textContent.includes('Table 6'),
      );
      assert(editor.textContent.includes('Original import evidence (before manual edits)'));
    });
    clickButton(editor, 'Close');
  } catch (error) {
    failures.push('Browser integration: ' + error.message);
  }
  finish();
};
document.body.append(frame);
