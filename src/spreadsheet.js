/* Local spreadsheet parsing using vendored SheetJS 0.20.3 (Apache-2.0). */
'use strict';
const SpreadsheetParser = (() => {
  const key = (value) =>
    String(value ?? '')
      .toLowerCase()
      .replace(/[^\p{L}\p{N}]+/gu, ' ')
      .trim()
      .replace(/\s+/g, ' ');
  const isHighSchool = (label) => /\bhs\s*\d*\b|\bhigh school\b/.test(key(label));
  const isOnline = (value) => /\b(?:online|ol)\s*\d*\b/i.test(String(value ?? ''));
  const withoutOnline = (value) =>
    String(value ?? '')
      .replace(/\b(?:online|ol)\b/gi, ' ')
      .replace(/\(\s*\)|\[\s*\]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  const onlineOnly = (value) =>
    isOnline(value) && !withoutOnline(value).replace(/[\s()[\].,:;-]/g, '');
  const text = (cell) => String(cell?.v ?? '').trim();
  const address = (r, c) => XLSX.utils.encode_cell({ r, c });
  const cell = (sheet, r, c) => sheet[address(r, c)];
  function range(sheet) {
    if (!sheet?.['!ref']) throw new Error('The selected sheet is empty.');
    const bounds = XLSX.utils.decode_range(sheet['!ref']);
    if (bounds.e.r > 10000 || bounds.e.c > 300)
      throw new Error(
        'The sheet is too large. Remove unused formatted rows/columns or select a smaller sheet.',
      );
    return bounds;
  }
  function minutes(value) {
    if (typeof value === 'number' && value >= 0 && value < 1)
      return Math.round(value * 1440) % 1440;
    const m = String(value ?? '')
      .trim()
      .match(/^(\d{1,2}):(\d{2})(?::00)?\s*(am|pm)?$/i);
    if (!m || +m[2] > 59) return null;
    let h = +m[1];
    if (m[3]) {
      if (h < 1 || h > 12) return null;
      h = (h % 12) + (/pm/i.test(m[3]) ? 12 : 0);
    } else if (h > 23) return null;
    return h * 60 + +m[2];
  }
  const clock = (n) =>
    `${String(Math.floor(n / 60)).padStart(2, '0')}:${String(n % 60).padStart(2, '0')}`;
  function sessionNotes(values) {
    const combined = values.join(' '),
      notes = [];
    if (/\bneeds?\s+attention\b|\battention\b|\bna\b/i.test(combined))
      notes.push('Needs attention');
    if (/\b(?:phy|physics)\b/i.test(combined)) notes.push('Physics');
    return notes;
  }
  const aliases = {
    student: ['student', 'student name', 'name', 'students'],
    binder: ['binder', 'binder location', 'binder position', 'location', 'position'],
    row: ['row', 'binder row'],
    col: ['column', 'col', 'binder column'],
    staff: ['staff', 'staff name', 'instructor', 'worker', 'teacher'],
    start: ['start', 'start time'],
    end: ['end', 'end time'],
  };
  function header(sheet, accepts) {
    const bounds = range(sheet);
    for (let r = bounds.s.r; r <= Math.min(bounds.e.r, bounds.s.r + 40); r++) {
      const found = {};
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        const value = key(text(cell(sheet, r, c)));
        for (const field of Object.keys(aliases))
          if (aliases[field].includes(value) && found[field] === undefined) found[field] = c;
      }
      if (accepts(found)) return { headerRow: r, ...found };
    }
    return null;
  }
  function readBinders(sheet) {
    const bounds = range(sheet),
      roster = new Map();
    function add(name, location, source) {
      name = name.trim();
      location = location.toUpperCase().replace(/[\s-]/g, '');
      if (!name) return;
      if (!/^[A-E][1-5]$/.test(location))
        throw new Error(`${source}: invalid binder location for ${name}; expected A1 to E5.`);
      const id = key(name);
      if (roster.has(id))
        throw new Error(`${source}: duplicate student name ${name}; use unique names.`);
      roster.set(id, { name, location });
    }
    // Center exports may repeat [row, column, First Name, Last Name] side by side.
    const splitLists = [];
    for (let r = bounds.s.r; r <= Math.min(bounds.e.r, bounds.s.r + 40); r++) {
      for (let c = bounds.s.c + 2; c < bounds.e.c; c++) {
        if (
          key(text(cell(sheet, r, c))) === 'first name' &&
          key(text(cell(sheet, r, c + 1))) === 'last name'
        )
          splitLists.push({ r, c });
      }
      if (splitLists.length) break;
    }
    const h = header(
      sheet,
      (h) =>
        h.student !== undefined &&
        (h.binder !== undefined || (h.row !== undefined && h.col !== undefined)),
    );
    if (splitLists.length) {
      splitLists.forEach(({ r: heading, c }) => {
        for (let r = heading + 1; r <= bounds.e.r; r++) {
          const name = [text(cell(sheet, r, c)), text(cell(sheet, r, c + 1))]
            .filter(Boolean)
            .join(' ');
          const location = text(cell(sheet, r, c - 2)) + text(cell(sheet, r, c - 1));
          if (!name && !location) continue;
          if (!name) throw new Error(`${address(r, c)}: binder location has no student name.`);
          add(name, location, address(r, c));
        }
      });
    } else if (h) {
      for (let r = h.headerRow + 1; r <= bounds.e.r; r++) {
        const name = text(cell(sheet, r, h.student));
        const location =
          h.binder !== undefined
            ? text(cell(sheet, r, h.binder))
            : text(cell(sheet, r, h.row)) + text(cell(sheet, r, h.col));
        if (!name && !location) continue;
        if (!name) throw new Error(`Row ${r + 1}: binder location has no student name.`);
        add(name, location, `Row ${r + 1}`);
      }
    } else {
      // Physical binder grid: 1..5 columns and A..E rows, or the transpose.
      let layout = null;
      for (let r = bounds.s.r; r <= Math.min(bounds.e.r, bounds.s.r + 40) && !layout; r++) {
        for (const transpose of [false, true]) {
          const columns = [];
          for (let c = bounds.s.c; c <= bounds.e.c; c++) {
            const value = text(cell(sheet, r, c)).toUpperCase();
            if ((transpose ? /^[A-E]$/ : /^[1-5]$/).test(value)) columns.push({ c, value });
          }
          if (columns.length !== 5 || new Set(columns.map((c) => c.value)).size !== 5) continue;
          for (let c = bounds.s.c; c < columns[0].c && !layout; c++) {
            const rows = [];
            for (let rr = r + 1; rr <= bounds.e.r; rr++) {
              const value = text(cell(sheet, rr, c)).toUpperCase();
              if ((transpose ? /^[1-5]$/ : /^[A-E]$/).test(value)) rows.push({ r: rr, value });
            }
            if (rows.length === 5 && new Set(rows.map((r) => r.value)).size === 5)
              layout = { rows, columns, transpose };
          }
        }
      }
      if (!layout)
        throw new Error(
          'Could not identify binder locations. Expected a student/location list or an A-E by 1-5 grid.',
        );
      layout.rows.forEach((r) =>
        layout.columns.forEach((c) => {
          const location = layout.transpose ? c.value + r.value : r.value + c.value;
          text(cell(sheet, r.r, c.c))
            .split(/\r?\n|;/)
            .filter(Boolean)
            .forEach((name) => add(name, location, address(r.r, c.c)));
        }),
      );
    }
    if (!roster.size) throw new Error('No student binder locations found.');
    return roster;
  }
  function matchName(value, roster) {
    const id = key(value);
    if (roster.has(id)) return { entry: roster.get(id), abbreviated: false };
    const parts = id.split(' ');
    const matches = [...roster.entries()].filter(([name]) => {
      const full = name.split(' ');
      return (
        full.length >= parts.length &&
        full[0] === parts[0] &&
        parts.slice(1).every((part, i) => full[i + 1].startsWith(part))
      );
    });
    return matches.length === 1 ? { entry: matches[0][1], abbreviated: true } : null;
  }
  function fill(cell) {
    const style = cell?.s?.fill || cell?.s;
    if (!style || style.patternType === 'none') return null;
    const color = style.fgColor;
    if (!color || /^(?:FF)?FFFFFF$/i.test(color.rgb || '') || color.indexed === 64) return null;
    // XLSX may represent the same RGB color with or without its alpha prefix.
    return color.rgb ? `rgb:${color.rgb.slice(-6).toUpperCase()}` : JSON.stringify(color);
  }
  function colorDistance(a, b) {
    if (!/^rgb:[0-9A-F]{6}$/.test(a || '') || !/^rgb:[0-9A-F]{6}$/.test(b || '')) return Infinity;
    return Math.hypot(
      ...[4, 6, 8].map(
        (offset) =>
          parseInt(a.slice(offset, offset + 2), 16) - parseInt(b.slice(offset, offset + 2), 16),
      ),
    );
  }
  function readSchedule(sheet, roster, options = {}) {
    const bounds = range(sheet),
      rows = [],
      warnings = [];
    const h = header(sheet, (h) =>
      ['staff', 'start', 'end', 'student'].every((k) => h[k] !== undefined),
    );
    if (h) {
      for (let r = h.headerRow + 1; r <= bounds.e.r; r++) {
        const values = ['staff', 'start', 'end', 'student'].map((k) => cell(sheet, r, h[k])?.v);
        if (values.every((v) => v === undefined || v === '')) continue;
        const [staff, rawStart, rawEnd, rawStudent] = values,
          start = minutes(rawStart),
          end = minutes(rawEnd);
        if (!staff || !rawStudent || start === null || end === null || end <= start)
          throw new Error(`Row ${r + 1}: missing staff/student or invalid start/end times.`);
        const studentName = withoutOnline(rawStudent);
        if (!studentName) throw new Error(`Row ${r + 1}: missing student name.`);
        const match = matchName(studentName, roster);
        if (match?.abbreviated)
          warnings.push(`${rawStudent} matched uniquely to ${match.entry.name}.`);
        rows.push({
          staff: String(staff).trim(),
          student: match?.entry.name || studentName,
          online: isOnline(rawStudent) || isOnline(staff),
          start: clock(start),
          end: clock(end),
        });
      }
      return { rows, warnings };
    }
    let heading = null,
      timeColumns = [];
    for (let r = bounds.s.r; r <= Math.min(bounds.e.r, bounds.s.r + 40); r++) {
      const cols = [];
      for (let c = bounds.s.c; c <= bounds.e.c; c++)
        if (/^(start ?time|time|times)$/.test(key(text(cell(sheet, r, c))))) cols.push(c);
      if (cols.length) {
        heading = r;
        timeColumns = cols;
        break;
      }
    }
    if (heading === null)
      throw new Error(
        'Could not find a time-grid header (Start Time or Time), or a Staff/Start/End/Student list.',
      );
    const merges = sheet['!merges'] || [];
    const handoffs = ScheduleTakeovers.collect({ sheet, bounds, heading, cell, text, address });
    const ignoredNote =
      /^(?:needs?\b|attention\b|na\b|table\s*\d|binder\b|cancel(?:led|ed)?\b|assess?ments?\b|(?:ol|online)\b|g(?:r(?:ade)?)?\s*\d|adv\b|move\b|make\s*-?\s*up\b|last\s+one\b|additional\b|physics?\b|phy\b|\d)/i;
    const teacherCells = new Set(),
      assessmentCells = new Set(),
      assessmentFills = new Set(),
      spreadsheetAssessmentFills = new Set(),
      instructorByFill = new Map(),
      tableInstructorFills = new Set(),
      tablesByTeacher = new Map();
    function rememberTable(teacher, table) {
      if (!table) return;
      if (!tablesByTeacher.has(key(teacher))) tablesByTeacher.set(key(teacher), new Set());
      tablesByTeacher.get(key(teacher)).add(table);
    }
    const preferenceFill = (value) =>
      typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value)
        ? `rgb:${value.slice(1).toUpperCase()}`
        : null;
    const directorFill = preferenceFill(options?.directorColor);
    const selectedAssessmentFill = preferenceFill(options?.assessmentColor);
    if (selectedAssessmentFill) assessmentFills.add(selectedAssessmentFill);
    const namedColors = new Map();
    if (Array.isArray(options?.colorIndicators)) {
      for (const indicator of options.colorIndicators.slice(0, 50)) {
        const color = preferenceFill(indicator?.color);
        const name = typeof indicator?.name === 'string' ? indicator.name.trim() : '';
        if (
          !color ||
          !name ||
          name.length > 80 ||
          !['instructor', 'assessment'].includes(indicator.behavior)
        )
          continue;
        if (color === directorFill || color === selectedAssessmentFill || namedColors.has(color)) {
          warnings.push(
            `Conflicting color indicator ${indicator.color}; custom assignment was ignored.`,
          );
          if (namedColors.has(color)) namedColors.set(color, null);
          continue;
        }
        namedColors.set(color, { teacher: name, table: '', behavior: indicator.behavior });
      }
      for (const [color, assignment] of namedColors)
        if (assignment?.behavior === 'assessment') assessmentFills.add(color);
    }
    for (let r = heading + 1; r <= bounds.e.r; r++) {
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        if (/^assess?ments?$/i.test(text(cell(sheet, r, c)))) {
          assessmentCells.add(address(r, c));
          const color = fill(cell(sheet, r, c));
          if (color) {
            assessmentFills.add(color);
            spreadsheetAssessmentFills.add(color);
          }
        }
        const table = text(cell(sheet, r, c)).match(/^table\s*(\d+)$/i);
        if (!table) continue;
        const teacher = text(cell(sheet, r - 1, c));
        if (!teacher || ignoredNote.test(teacher) || handoffs.noteCells.has(address(r - 1, c)))
          continue;
        rememberTable(teacher, table[1]);
        teacherCells.add(address(r - 1, c));
        const color = fill(cell(sheet, r, c));
        if (!color) continue;
        const assignment = { teacher, table: table[1] };
        tableInstructorFills.add(color);
        if (
          instructorByFill.has(color) &&
          key(instructorByFill.get(color)?.teacher) !== key(teacher)
        )
          instructorByFill.set(color, null);
        else if (!instructorByFill.has(color)) instructorByFill.set(color, assignment);
      }
    }
    // A lowercase name is an instructor annotation, even without a table label.
    const lowercaseInstructor = (value) =>
      /^\p{Ll}[\p{L}\p{M}'\u2019 .-]*$/u.test(value) &&
      !ignoredNote.test(value) &&
      !ScheduleTakeovers.isNote(value);
    for (let r = heading + 1; r <= bounds.e.r; r++) {
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        const name = text(cell(sheet, r, c)),
          color = fill(cell(sheet, r, c));
        if (
          !lowercaseInstructor(name) ||
          (assessmentFills.has(color) &&
            (!namedColors.has(color) || spreadsheetAssessmentFills.has(color))) ||
          handoffs.noteCells.has(address(r, c))
        )
          continue;
        teacherCells.add(address(r, c));
        if (!color) continue;
        // Names explicitly above table labels outrank incidental lowercase text.
        // Conflicting table labels still remain ambiguous from the first pass.
        if (tableInstructorFills.has(color)) continue;
        if (!instructorByFill.has(color)) instructorByFill.set(color, { teacher: name, table: '' });
        else if (key(instructorByFill.get(color)?.teacher) !== key(name))
          instructorByFill.set(color, null);
      }
    }
    // Device preferences deliberately override inferred workbook color labels.
    const manualAssignments = new Map(namedColors);
    if (directorFill)
      manualAssignments.set(directorFill, { teacher: 'Center director', behavior: 'instructor' });
    if (selectedAssessmentFill)
      manualAssignments.set(selectedAssessmentFill, {
        teacher: 'Assessment',
        behavior: 'assessment',
      });
    for (const [color, assignment] of manualAssignments) {
      if (!assignment) continue;
      const detected = instructorByFill.get(color);
      const wasAssessment = spreadsheetAssessmentFills.has(color);
      const conflict = wasAssessment
        ? assignment.behavior !== 'assessment'
        : instructorByFill.has(color) &&
          (!detected ||
            assignment.behavior === 'assessment' ||
            key(detected.teacher) !== key(assignment.teacher));
      if (conflict) {
        const original = wasAssessment
          ? 'Assessment'
          : detected?.teacher || 'conflicting instructor labels';
        const replacement =
          assignment.behavior === 'assessment'
            ? `${assignment.teacher} (Assessment)`
            : assignment.teacher;
        warnings.push(
          `Manual color override ${color.replace(/^rgb:/, '#')}: spreadsheet indicates ${original}; your indicator assigns ${replacement}. Your indicator takes priority; verify this setting.`,
        );
      }
    }
    for (const [color, assignment] of namedColors) {
      if (!assignment || assignment.behavior === 'assessment') continue;
      assessmentFills.delete(color);
      const table = instructorByFill.get(color)?.table || '';
      instructorByFill.set(color, { teacher: assignment.teacher, table });
      rememberTable(assignment.teacher, table);
    }
    // Handoff wording names an instructor at a particular time. It must not
    // relabel every earlier booking that happens to share the note's color.
    function instructorForColor(color) {
      if (!color || assessmentFills.has(color)) return null;
      if (color === directorFill) return { teacher: 'Center director', table: '' };
      // Explicit labels always win, even when two instructors use nearby shades.
      if (instructorByFill.has(color)) return instructorByFill.get(color);
      const candidates = [...instructorByFill].map(([knownColor, assignment]) => ({
        assignment,
        distance: colorDistance(color, knownColor),
      }));
      if (directorFill)
        candidates.push({
          assignment: { teacher: 'Center director', table: '' },
          distance: colorDistance(color, directorFill),
        });
      // Assessment colors also compete, so a nearby assessment shade cannot
      // silently become an instructor's student through approximate matching.
      assessmentFills.forEach((knownColor) =>
        candidates.push({ assignment: null, distance: colorDistance(color, knownColor) }),
      );
      candidates.sort((a, b) => a.distance - b.distance);
      const best = candidates[0];
      if (!best?.assignment || best.distance > 26) return null;
      const rival = candidates.find(
        (candidate) =>
          !candidate.assignment ||
          key(candidate.assignment.teacher) !== key(best.assignment.teacher),
      );
      // Require a close match and a clear lead over any other instructor/color.
      if (rival && rival.distance - best.distance < 10) return null;
      return best.assignment;
    }
    const unresolvedColors = [],
      colorApplied = [];
    const colorPalette = [...instructorByFill]
      .filter(([color, assignment]) => assignment && !assessmentFills.has(color))
      .map(([color, assignment]) => ({ color, instructor: assignment.teacher }));
    if (directorFill) colorPalette.push({ color: directorFill, instructor: 'Center director' });
    for (let c = bounds.s.c; c <= bounds.e.c; c++) {
      if (timeColumns.includes(c)) continue;
      const headerMerge = merges.find(
        (m) => heading >= m.s.r && heading <= m.e.r && c >= m.s.c && c <= m.e.c,
      );
      const staff =
        text(cell(sheet, heading, c)) ||
        (headerMerge ? text(cell(sheet, headerMerge.s.r, headerMerge.s.c)) : '');
      if (!staff) {
        if (
          Object.keys(sheet).some(
            (a) => !a.startsWith('!') && XLSX.utils.decode_cell(a).c === c && text(sheet[a]),
          )
        )
          warnings.push(
            `Column ${XLSX.utils.encode_col(c)} has content but no staff/section heading; it was not imported.`,
          );
        continue;
      }
      const tc = [...timeColumns].sort((a, b) => Math.abs(a - c) - Math.abs(b - c) || a - b)[0];
      const times = new Map();
      let previous = -1;
      for (let r = heading + 1; r <= bounds.e.r; r++) {
        const t = minutes(cell(sheet, r, tc)?.v);
        if (t === null) continue;
        if (t <= previous)
          throw new Error(
            `${address(r, tc)}: time labels must increase within a schedule. Put different days on separate sheets.`,
          );
        times.set(r, t);
        previous = t;
      }
      if (!times.size) continue;
      const starts = [];
      for (let r = heading + 1; r <= bounds.e.r; r++) {
        const raw = text(cell(sheet, r, c));
        if (
          !raw ||
          teacherCells.has(address(r, c)) ||
          assessmentCells.has(address(r, c)) ||
          handoffs.noteCells.has(address(r, c))
        )
          continue;
        if (
          merges.some(
            (m) =>
              r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c && (r !== m.s.r || c !== m.s.c),
          )
        )
          continue;
        const students = [],
          onlineStudents = new Set(),
          nameWarnings = [];
        raw
          .split(/\r?\n|;/)
          .filter(Boolean)
          .forEach((rawValue) => {
            // Inline course notes describe the session, not part of a student's name.
            const value = withoutOnline(rawValue)
              .replace(
                /\s+(?:g(?:r(?:ade)?)?\s*\d+(?:\s+ap)?|physics?|phy|needs?\s+attention|na|\(na\)|\[na\])\s*$/i,
                '',
              )
              .trim();
            if (!value) return;
            if (ScheduleTakeovers.isNote(value)) return;
            if (lowercaseInstructor(value)) return;
            const match = matchName(value, roster);
            if (match) {
              students.push(match.entry.name);
              if (isOnline(rawValue)) onlineStudents.add(match.entry.name);
              if (match.abbreviated)
                nameWarnings.push(
                  `${address(r, c)}: ${value} matched uniquely to ${match.entry.name}.`,
                );
            } else if (
              isHighSchool(staff) &&
              !ignoredNote.test(value) &&
              /^[\p{L}\p{M}][\p{L}\p{M}'’ .-]*$/u.test(value.trim())
            ) {
              // High-school students need not have a binder. Keep their schedule name.
              students.push(value.trim());
              if (isOnline(rawValue)) onlineStudents.add(value.trim());
            } else if (!ignoredNote.test(value) && !assessmentFills.has(fill(cell(sheet, r, c))))
              warnings.push(
                `${address(r, c)} (${staff}): unmatched text "${value}" was not imported. It may be a student or a note.`,
              );
          });
        if (students.length) {
          if (raw.split(/\r?\n|;/).some(onlineOnly))
            students.forEach((name) => onlineStudents.add(name));
          starts.push({ r, students, nameWarnings, onlineStudents });
        }
      }
      for (let r = heading + 1; r <= bounds.e.r; r++) {
        if (
          !isOnline(text(cell(sheet, r, c))) ||
          starts.some((entry) => entry.r === r) ||
          teacherCells.has(address(r, c)) ||
          handoffs.noteCells.has(address(r, c))
        )
          continue;
        const adjacent = starts.filter((entry) => Math.abs(entry.r - r) === 1);
        if (adjacent.length === 1)
          adjacent[0].students.forEach((name) => adjacent[0].onlineStudents.add(name));
        else if (adjacent.length > 1)
          warnings.push(
            `${address(r, c)}: online marker is between two students; add OL to the intended student's name cell.`,
          );
      }
      const consumedContinuations = new Set();
      starts.forEach((entry, index) => {
        let startRow = entry.r;
        if (assessmentFills.has(fill(cell(sheet, entry.r, c)))) {
          for (let r = entry.r - 1; r > heading; r--) {
            if (
              !assessmentFills.has(fill(cell(sheet, r, c))) ||
              starts.some((candidate) => candidate.r === r)
            )
              break;
            if (assessmentCells.has(address(r, c))) {
              startRow = r;
              break;
            }
          }
        }
        const start = times.get(startRow);
        if (start === undefined) {
          warnings.push(
            `${address(entry.r, c)}: no start-time label for ${entry.students.join(', ')}; appointment omitted.`,
          );
          return;
        }
        const students = entry.students.filter(
          (student) => !consumedContinuations.has(`${entry.r}:${key(student)}`),
        );
        if (!students.length) return;
        let teachingRow = entry.r;
        const hourEndRow =
          [...times].find(([r, time]) => r > entry.r && time >= start + 60)?.[0] ?? bounds.e.r + 1;
        if (assessmentFills.has(fill(cell(sheet, entry.r, c)))) {
          // Exclude only the assessment portion, not the student's later
          // instructor-colored cells. Keep the original appointment end time.
          const nextEntry = starts[index + 1];
          const sameStudents =
            nextEntry &&
            nextEntry.students.every((student) =>
              students.some((name) => key(name) === key(student)),
            );
          const limit = Math.min(
            hourEndRow,
            nextEntry ? nextEntry.r + (sameStudents ? 1 : 0) : bounds.e.r + 1,
          );
          teachingRow = -1;
          for (let r = entry.r + 1; r < limit; r++) {
            const color = fill(cell(sheet, r, c));
            if (color && !assessmentFills.has(color)) {
              teachingRow = r;
              break;
            }
          }
          if (teachingRow < 0) return; // Assessment-only booking.
          if (nextEntry?.r === teachingRow && sameStudents) {
            students.forEach((student) =>
              consumedContinuations.add(`${teachingRow}:${key(student)}`),
            );
          }
        }
        const teachingStart = times.get(teachingRow);
        if (teachingStart === undefined) {
          warnings.push(
            `${address(teachingRow, c)}: no time label for the instructor portion of ${students.join(', ')}.`,
          );
          return;
        }
        const next = starts.find((candidate) => candidate.r > teachingRow)?.r;
        const color = fill(cell(sheet, teachingRow, c));
        let assignment = instructorForColor(color);
        let aiColor = false;
        let colorCase = null;
        if (color && !assignment && !instructorByFill.has(color) && !manualAssignments.has(color)) {
          colorCase = ScheduleColorChecks.create({
            sheet,
            r: teachingRow,
            c,
            students,
            start,
            teachingStart,
            staff,
            color,
            bounds,
            heading,
            times,
            palette: colorPalette,
            assessmentFills,
            distance: colorDistance,
            cell,
            fill,
            address,
            clock,
            instructorForColor,
          });
          if (colorCase) {
            const suggested = ScheduleColorChecks.resolve(colorCase, options.colorResolutions);
            if (suggested) {
              assignment = suggested;
              aiColor = true;
            }
          }
        }
        const instructor =
          assignment?.teacher || (color ? (directorFill ? 'Unassigned instructor' : 'ACD') : staff);
        if (color && directorFill && !assignment)
          warnings.push(
            `${address(teachingRow, c)}: ${students.join(', ')} has a color that matches neither an instructor nor the selected director color.`,
          );
        // Notes belong to the preceding student in this lane, up to the next
        // student or the one-hour session boundary, regardless of fill shade.
        const noteEnd = Math.min(next ?? bounds.e.r + 1, hourEndRow);
        const notes = [];
        for (let r = teachingRow; r < noteEnd; r++) {
          if (!teacherCells.has(address(r, c)) && !assessmentFills.has(fill(cell(sheet, r, c))))
            notes.push(text(cell(sheet, r, c)));
        }
        if (notes.some((note) => /\bcancel(?:led|ed)?\b/i.test(note))) {
          warnings.push(
            `${address(entry.r, c)} (${staff}): ${entry.students.join(', ')} omitted because the block is marked cancelled.`,
          );
          return;
        }
        if (color && !assignment) {
          const reason = instructorByFill.has(color)
            ? 'conflicting instructor labels'
            : 'no clear instructor match';
          warnings.push(
            `${address(teachingRow, c)} (${staff}), ${clock(start)}: ${students.join(', ')} has color ${color.replace(/^rgb:/, '#')} with ${reason}; fell back to ${instructor}. Add or correct a color indicator.${colorCase ? ' Enable Optional AI help to check this uncertain shade using nearby instructor context.' : ' There is insufficient color evidence for an AI color check; Optional AI help can still check unresolved handoff notes.'}`,
          );
          if (colorCase && unresolvedColors.length < 10) unresolvedColors.push(colorCase);
        }
        if (aiColor)
          colorApplied.push(
            `${address(teachingRow, c)}: AI-assisted color assignment to ${instructor}.`,
          );
        if (notes.some((note) => /\bmove\b/i.test(note)))
          warnings.push(
            `${address(entry.r, c)} (${staff}): appointment contains a move note; verify its time.`,
          );
        warnings.push(...entry.nameWarnings);
        students.forEach((student) =>
          rows.push({
            staff: instructor,
            table: assignment?.table || '',
            section: staff,
            student,
            online: isOnline(staff) || entry.onlineStudents.has(student),
            start: clock(start),
            sessionStart: clock(start),
            sourceColumn: c,
            sourceRow: entry.r,
            teachingStart: clock(teachingStart),
            sources: [
              {
                column: c,
                row: teachingRow,
                teachingStart: clock(teachingStart),
                stopRow:
                  starts.find(
                    (candidate) =>
                      candidate.r > teachingRow &&
                      !candidate.students.some((name) => key(name) === key(student)),
                  )?.r ?? null,
              },
            ],
            notes: [
              ...sessionNotes(notes),
              ...(aiColor ? ['AI-assisted color assignment'] : []),
              ...(teachingStart > start
                ? [`Assessment: ${clock(start)} - ${clock(teachingStart)}`]
                : []),
            ],
          }),
        );
      });
    }
    // Only appearances exactly one hour apart for the same student/instructor
    // form a two-hour session. Other appearances remain separate bookings.
    const sessions = new Map();
    rows.forEach((row) => {
      const id = JSON.stringify([key(row.student), key(row.staff)]);
      if (!sessions.has(id)) sessions.set(id, new Map());
      const appearances = sessions.get(id);
      if (!appearances.has(row.sessionStart))
        appearances.set(row.sessionStart, { ...row, notes: [...row.notes] });
      else {
        const original = appearances.get(row.sessionStart);
        original.notes = [...new Set([...original.notes, ...row.notes])];
        original.sources.push(...row.sources);
        original.online ||= row.online;
        warnings.push(
          `Duplicate entry for ${row.student}, ${row.staff}, ${row.start}; kept one occurrence.`,
        );
      }
    });
    const appointments = [];
    sessions.forEach((appearances) => {
      const entries = [...appearances.values()].sort((a, b) =>
        a.sessionStart.localeCompare(b.sessionStart),
      );
      for (let i = 0; i < entries.length; i++) {
        const row = entries[i],
          next = entries[i + 1];
        const twoHours = next && minutes(next.sessionStart) - minutes(row.sessionStart) === 60;
        const bookedDurationMinutes = twoHours ? 120 : 60;
        const sessionEnd = minutes(row.sessionStart) + bookedDurationMinutes;
        const durationMinutes = sessionEnd - minutes(row.start);
        const notes = twoHours ? [...new Set([...row.notes, ...next.notes])] : row.notes;
        appointments.push({
          ...row,
          online: row.online || (twoHours && next.online) || false,
          notes,
          durationMinutes,
          bookedDurationMinutes,
          sources: twoHours ? [...row.sources, ...next.sources] : row.sources,
          end: clock(sessionEnd),
        });
        if (twoHours) i++;
      }
    });
    const instructors = [...instructorByFill.values()]
      .filter(Boolean)
      .map((assignment) => assignment.teacher);
    if (directorFill) instructors.push('Center director');
    const unresolvedHandoffs = [],
      aiApplied = [...colorApplied];
    const segments = ScheduleTakeovers.apply(appointments, {
      sheet,
      notes: handoffs.notes,
      bounds,
      heading,
      timeColumns,
      instructorForColor,
      assessmentFills,
      cell,
      fill,
      address,
      key,
      minutes,
      clock,
      instructors,
      warnings,
      unresolvedHandoffs,
      aiApplied,
      resolutions: options.handoffResolutions || [],
    });
    for (const row of segments) {
      const tables = tablesByTeacher.get(key(row.staff));
      row.table = tables?.size ? [...tables].sort((a, b) => +a - +b).join(', ') : row.table || '';
    }
    return { rows: segments, warnings, unresolvedHandoffs, unresolvedColors, aiApplied };
  }
  return { key, isHighSchool, readBinders, readSchedule, minutes };
})();
