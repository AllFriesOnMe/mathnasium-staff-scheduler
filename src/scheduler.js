'use strict';
const $ = (id) => document.getElementById(id);
const books = { binders: null, schedule: null };
const versions = { binders: 0, schedule: 0 };
const importControllers = { binders: null, schedule: null };
let generationController = null;
let exportedRows = [];
const durationOf = (row) =>
  row.durationMinutes ?? SpreadsheetParser.minutes(row.end) - SpreadsheetParser.minutes(row.start);
function invalidate() {
  generationController?.abort();
  generationController = null;
  $('generate').disabled = false;
  $('results').hidden = true;
  $('importStatus').textContent = '';
  exportedRows = [];
}
function sheetId(kind) {
  return kind === 'binders' ? 'binderSheet' : 'scheduleSheet';
}
async function importFile(kind) {
  const version = ++versions[kind];
  importControllers[kind]?.abort();
  const controller = new AbortController();
  importControllers[kind] = controller;
  invalidate();
  books[kind] = null;
  const select = $(sheetId(kind));
  select.replaceChildren();
  select.hidden = true;
  $(sheetId(kind) + 'Label').hidden = true;
  const status = $(kind === 'binders' ? 'binderStatus' : 'scheduleStatus');
  const file = $(kind).files[0];
  if (!file) {
    status.textContent = 'No file selected';
    return;
  }
  status.textContent = 'Reading spreadsheet...';
  try {
    if (!/\.(xlsx|csv)$/i.test(file.name)) throw new Error('Choose an .xlsx or .csv spreadsheet.');
    if (file.size > SpreadsheetJobs.MAX_FILE_BYTES)
      throw new Error('This file exceeds the 5 MB limit. Choose a smaller spreadsheet.');
    const book = await SpreadsheetJobs.run({ type: 'read', file }, { signal: controller.signal });
    if (version !== versions[kind]) return;
    if (!book.SheetNames.length) throw new Error('This workbook has no sheets.');
    books[kind] = book;
    book.SheetNames.forEach((name) => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      select.append(option);
    });
    select.hidden = book.SheetNames.length === 1;
    $(sheetId(kind) + 'Label').hidden = select.hidden;
    status.textContent = `${file.name}: ${book.SheetNames.length} sheet(s) loaded.`;
  } catch (error) {
    if (version === versions[kind] && error.name !== 'AbortError')
      status.textContent = error.message;
  }
}
async function generate() {
  invalidate();
  const controller = new AbortController();
  generationController = controller;
  $('generate').disabled = true;
  try {
    if (!books.binders || !books.schedule) throw new Error('Upload both spreadsheets first.');
    const parsed = await SpreadsheetJobs.run(
      {
        type: 'generate',
        binders: books.binders.Sheets[$('binderSheet').value],
        schedule: books.schedule.Sheets[$('scheduleSheet').value],
      },
      { signal: controller.signal },
    );
    if (controller.signal.aborted) return;
    const rows = parsed.rows,
      warnings = [...parsed.warnings];
    if (!rows.length) throw new Error(`No appointments could be identified. ${warnings.join(' ')}`);
    rows.sort(
      (a, b) =>
        a.staff.localeCompare(b.staff, undefined, { numeric: true }) ||
        a.start.localeCompare(b.start) ||
        (a.end || '').localeCompare(b.end || '') ||
        a.student.localeCompare(b.student),
    );
    const groups = new Map();
    rows.forEach((row, index) => {
      if (!groups.has(row.staff)) groups.set(row.staff, new Map());
      const twoHours = durationOf(row) === 120;
      const slots = groups.get(row.staff),
        time = `${row.start} - ${row.end || 'END UNKNOWN'}${twoHours ? ' (2 hours)' : ''}`;
      const slot = JSON.stringify([time, twoHours ? index : null]);
      if (!slots.has(slot)) slots.set(slot, { time, students: [] });
      slots
        .get(slot)
        .students.push(
          row.student +
            (row.binder ? ` (${row.binder})` : '') +
            (row.notes?.length ? ` [${row.notes.join('; ')}]` : ''),
        );
    });
    $('staffLists').replaceChildren();
    groups.forEach((slots, staff) => {
      const block = document.createElement('section'),
        h3 = document.createElement('h3'),
        list = document.createElement('ul');
      block.className = 'staff-block';
      list.className = 'staff-list';
      h3.textContent = staff;
      slots.forEach(({ students, time }) => {
        const li = document.createElement('li');
        li.textContent = `${time}: ${students.join(', ')}`;
        list.append(li);
      });
      block.append(h3, list);
      $('staffLists').append(block);
    });
    const uniqueWarnings = [...new Set(warnings)];
    $('warningList').replaceChildren();
    uniqueWarnings.forEach((message) => {
      const li = document.createElement('li');
      li.textContent = message;
      $('warningList').append(li);
    });
    $('warnings').hidden = !uniqueWarnings.length;
    $('warnings').open = !!uniqueWarnings.length;
    $('warningSummary').textContent =
      `${uniqueWarnings.length} import warning(s) - check before use`;
    $('resultSummary').textContent =
      `${$('scheduleSheet').value}: ${groups.size} staff/sections, ${rows.length} appointments. ${uniqueWarnings.length ? 'Some entries need checking; see warnings below.' : 'Binder locations shown where available.'}`;
    exportedRows = rows;
    $('results').hidden = false;
  } catch (error) {
    if (!controller.signal.aborted) $('importStatus').textContent = error.message;
  } finally {
    if (generationController === controller) {
      generationController = null;
      $('generate').disabled = false;
    }
  }
}
function download() {
  const cells = [
    ['Staff', 'Start', 'End', 'Student', 'Binder', 'Notes', 'Duration'],
    ...exportedRows.map((r) => [
      r.staff,
      r.start,
      r.end || 'END UNKNOWN',
      r.student,
      r.binder,
      (r.notes || []).join('; '),
      durationOf(r) === 120 ? '2 hours' : `${durationOf(r)} minutes`,
    ]),
  ];
  const csv =
    '\uFEFF' +
    cells
      .map((row) =>
        row
          .map(
            (value) =>
              '"' +
              String(value)
                .replace(/^(\s*[=+@\-\t\r])/, "'$1")
                .replace(/"/g, '""') +
              '"',
          )
          .join(','),
      )
      .join('\r\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'staff-schedules.csv';
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
$('binders').onchange = () => importFile('binders');
$('schedule').onchange = () => importFile('schedule');
$('binderSheet').onchange = invalidate;
$('scheduleSheet').onchange = invalidate;
$('generate').onclick = generate;
$('print').onclick = () => window.print();
$('export').onclick = download;
