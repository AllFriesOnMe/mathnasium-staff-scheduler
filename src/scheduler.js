'use strict';
const $ = (id) => document.getElementById(id);
const books = { binders: null, schedule: null };
const versions = { binders: 0, schedule: 0 };
const importControllers = { binders: null, schedule: null };
let generationController = null;
let exportedRows = [];
const durationOf = (row) =>
  row.durationMinutes ?? SpreadsheetParser.minutes(row.end) - SpreadsheetParser.minutes(row.start);
function displayTime(value) {
  const total = SpreadsheetParser.minutes(value);
  if (total === null) return value;
  const hour = Math.floor(total / 60);
  return `${hour % 12 || 12}:${String(total % 60).padStart(2, '0')} ${hour < 12 ? 'AM' : 'PM'}`;
}
function displayNote(value) {
  return String(value).replace(/\b(?:[01]\d|2[0-3]):[0-5]\d\b(?!\s*(?:am|pm)\b)/gi, displayTime);
}
function invalidate() {
  generationController?.abort();
  generationController = null;
  $('generate').disabled = false;
  $('results').hidden = true;
  $('importStatus').textContent = '';
  $('aiStatus').textContent = '';
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
    const task = {
      type: 'generate',
      binders: books.binders.Sheets[$('binderSheet').value],
      schedule: books.schedule.Sheets[$('scheduleSheet').value],
      options: ColorPreferences.get(),
    };
    let parsed = await SpreadsheetJobs.run(task, { signal: controller.signal });
    if (controller.signal.aborted) return;
    if ($('aiEnabled').checked) {
      let requested = false;
      const messages = [];
      try {
        // Correct uncertain colors first, then rebuild handoff evidence and IDs.
        for (const [field, option, label] of [
          ['unresolvedColors', 'colorResolutions', 'uncertain colors'],
          ['unresolvedHandoffs', 'handoffResolutions', 'unresolved handoffs'],
        ]) {
          const cases = parsed[field];
          if (!cases?.length) continue;
          requested = true;
          $('aiStatus').textContent = 'Checking ' + label + ' with AI...';
          const resolutions = await AiFallback.resolve(
            cases,
            $('aiAccessCode').value,
            controller.signal,
          );
          if (controller.signal.aborted) return;
          if (resolutions.length) {
            task.options = { ...task.options, [option]: resolutions };
            parsed = await SpreadsheetJobs.run(task, { signal: controller.signal });
            if (controller.signal.aborted) return;
          }
        }
        if (!requested)
          messages.push('No eligible unresolved handoffs or colors; no AI request was sent.');
        else if (!parsed.aiApplied?.length)
          messages.push(
            'AI could not resolve any additional assignments confidently. See import warnings.',
          );
      } catch (error) {
        if (controller.signal.aborted) return;
        messages.push('AI check unavailable: ' + error.message);
      }
      $('aiStatus').textContent = [...(parsed.aiApplied || []), ...messages]
        .map(displayNote)
        .join(' ');
    }
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
        time = `${displayTime(row.start)} - ${row.end ? displayTime(row.end) : 'END UNKNOWN'}${twoHours ? ' (2 hours)' : ''}`;
      const slot = JSON.stringify([time, twoHours ? index : null]);
      if (!slots.has(slot)) slots.set(slot, { time, students: [] });
      slots.get(slot).students.push(row);
    });
    $('staffLists').replaceChildren();
    groups.forEach((slots, staff) => {
      const block = document.createElement('section'),
        h3 = document.createElement('h3'),
        list = document.createElement('ul');
      block.className = 'staff-block';
      list.className = 'staff-list';
      const tables = [
        ...new Set(
          rows
            .filter((row) => row.staff === staff)
            .flatMap((row) => String(row.table || '').split(/,\s*/))
            .filter(Boolean),
        ),
      ].sort((a, b) => +a - +b);
      h3.textContent = staff + (tables.length ? ` (Table ${tables.join(', ')})` : '');
      slots.forEach(({ students, time }) => {
        const li = document.createElement('li');
        li.append(`${time}: `);
        students.forEach((row, index) => {
          if (index) li.append(', ');
          const attention = (note) => /^needs?\s+attention$/i.test(note);
          const needsAttention = row.notes?.some(attention);
          const name = document.createElement(needsAttention ? 'strong' : 'span');
          name.textContent = row.student;
          if (needsAttention) name.className = 'needs-attention';
          const notes = (row.notes || []).filter((note) => !attention(note)).map(displayNote);
          li.append(
            name,
            (row.online ? ' (online)' : '') +
              (row.binder ? ` (${row.binder})` : '') +
              (notes.length ? ` [${notes.join('; ')}]` : ''),
          );
        });
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
    ['Staff', 'Start', 'End', 'Student', 'Binder', 'Notes', 'Duration', 'Table', 'Online'],
    ...exportedRows.map((r) => [
      r.staff,
      displayTime(r.start),
      r.end ? displayTime(r.end) : 'END UNKNOWN',
      r.student + (r.online ? ' (online)' : ''),
      r.binder,
      (r.notes || []).map(displayNote).join('; '),
      durationOf(r) === 120 ? '2 hours' : `${durationOf(r)} minutes`,
      r.table || '',
      r.online ? 'Yes' : '',
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
ColorPreferences.onChange(invalidate);
$('aiEnabled').onchange = () => {
  $('aiAccessSettings').hidden = !$('aiEnabled').checked;
  if (!$('aiEnabled').checked) $('aiAccessCode').value = '';
  invalidate();
};
$('aiAccessCode').oninput = invalidate;
