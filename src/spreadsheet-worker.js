'use strict';
importScripts('../vendor/xlsx.full.min.js', 'takeovers.js', 'color-checks.js', 'spreadsheet.js');
const MAX_FILE_BYTES = 5 * 1024 * 1024;
function validateBook(book) {
  if (!book.SheetNames?.length) throw new Error('This workbook has no sheets.');
  if (book.SheetNames.length > 20) throw new Error('Choose a workbook with at most 20 sheets.');
  let totalArea = 0,
    totalCells = 0;
  for (const name of book.SheetNames) {
    const sheet = book.Sheets[name],
      ref = sheet['!fullref'] || sheet['!ref'];
    if (!ref) continue;
    const bounds = XLSX.utils.decode_range(ref);
    if (bounds.s.r < 0 || bounds.s.c < 0 || bounds.e.r >= 10000 || bounds.e.c >= 300)
      throw new Error(
        'A sheet exceeds the limit of 10,000 rows or 300 columns. Remove unused formatted cells.',
      );
    totalArea += (bounds.e.r - bounds.s.r + 1) * (bounds.e.c - bounds.s.c + 1);
    totalCells += Object.keys(sheet).filter((key) => !key.startsWith('!')).length;
    if (totalArea > 250000 || totalCells > 100000)
      throw new Error('This workbook contains too many cells. Choose a smaller spreadsheet.');
    if ((sheet['!merges'] || []).length > 5000)
      throw new Error('This sheet has too many merged cells.');
  }
}
self.onmessage = async (event) => {
  try {
    const task = event.data;
    let result;
    if (task.type === 'read') {
      const file = task.file;
      if (!file || !/\.(xlsx|csv)$/i.test(file.name))
        throw new Error('Choose an .xlsx or .csv spreadsheet.');
      if (!file.size || file.size > MAX_FILE_BYTES)
        throw new Error('Choose a nonempty spreadsheet of 5 MB or smaller.');
      const data = await file.arrayBuffer();
      if (/\.xlsx$/i.test(file.name)) {
        const signature = new Uint8Array(data, 0, Math.min(4, data.byteLength));
        if (signature.join(',') !== '80,75,3,4')
          throw new Error(
            'This file is not a valid .xlsx workbook. Export it from your spreadsheet application again.',
          );
      }
      result = XLSX.read(data, {
        type: 'array',
        cellStyles: true,
        sheetStubs: true,
        cellHTML: false,
        cellFormula: false,
        bookVBA: false,
        sheetRows: 10001,
      });
      validateBook(result);
    } else if (task.type === 'generate') {
      const roster = SpreadsheetParser.readBinders(task.binders);
      result = SpreadsheetParser.readSchedule(task.schedule, roster, task.options);
      if (result.rows.length > 5000)
        throw new Error('This schedule exceeds the limit of 5,000 appointments.');
      const seen = new Set(),
        byStudent = new Map();
      result.rows.forEach((row) => {
        row.binder = roster.get(SpreadsheetParser.key(row.student))?.location || '';
        if (!row.binder && !SpreadsheetParser.isHighSchool(row.section || row.staff))
          result.warnings.push(`No binder location for ${row.student}.`);
        const id = JSON.stringify([row.staff, row.start, row.end, row.student]);
        if (seen.has(id))
          result.warnings.push(
            `Duplicate appointment: ${row.student}, ${row.staff}, ${row.start}.`,
          );
        seen.add(id);
        const name = SpreadsheetParser.key(row.student);
        if (!byStudent.has(name)) byStudent.set(name, []);
        byStudent.get(name).push(row);
      });
      byStudent.forEach((rows) => {
        rows.sort((a, b) => a.start.localeCompare(b.start));
        let latestEnd = '';
        rows.forEach((row) => {
          if (row.start < latestEnd)
            result.warnings.push(`Overlapping bookings for ${row.student} at ${row.start}.`);
          if (row.end > latestEnd) latestEnd = row.end;
        });
      });
    } else throw new Error('Unknown spreadsheet operation.');
    self.postMessage({ ok: true, result });
  } catch (error) {
    self.postMessage({ ok: false, error: error.message || 'Could not process this spreadsheet.' });
  }
};
