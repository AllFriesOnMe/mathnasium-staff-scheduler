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
    if (/\bneeds?\s+attention\b|\battention\b/i.test(combined)) notes.push('Needs attention');
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
  function readSchedule(sheet, roster) {
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
        const match = matchName(rawStudent, roster);
        if (match?.abbreviated)
          warnings.push(`${rawStudent} matched uniquely to ${match.entry.name}.`);
        rows.push({
          staff: String(staff).trim(),
          student: match?.entry.name || String(rawStudent).trim(),
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
    const ignoredNote =
      /^(?:needs?\b|attention\b|table\s*\d|binder\b|cancel(?:led|ed)?\b|assess?ments?\b|g(?:r(?:ade)?)?\s*\d|adv\b|move\b|physics?\b|phy\b|\d)/i;
    const teacherCells = new Set(),
      assessmentCells = new Set(),
      assessmentFills = new Set(),
      instructorByFill = new Map();
    for (let r = heading + 1; r <= bounds.e.r; r++) {
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        if (/^assess?ments?$/i.test(text(cell(sheet, r, c)))) {
          assessmentCells.add(address(r, c));
          const color = fill(cell(sheet, r, c));
          if (color) assessmentFills.add(color);
        }
        const table = text(cell(sheet, r, c)).match(/^table\s*(\d+)$/i);
        if (!table) continue;
        const teacher = text(cell(sheet, r - 1, c));
        if (!teacher || ignoredNote.test(teacher)) continue;
        teacherCells.add(address(r - 1, c));
        const color = fill(cell(sheet, r, c));
        if (!color) continue;
        const assignment = { teacher, table: table[1] };
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
      /^\p{Ll}[\p{L}\p{M}'\u2019 .-]*$/u.test(value) && !ignoredNote.test(value);
    for (let r = heading + 1; r <= bounds.e.r; r++) {
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        const name = text(cell(sheet, r, c)),
          color = fill(cell(sheet, r, c));
        if (!lowercaseInstructor(name) || assessmentFills.has(color)) continue;
        teacherCells.add(address(r, c));
        if (!color) continue;
        if (!instructorByFill.has(color)) instructorByFill.set(color, { teacher: name, table: '' });
        else if (key(instructorByFill.get(color)?.teacher) !== key(name))
          instructorByFill.set(color, null);
      }
    }
    function instructorForColor(color) {
      // Explicit labels always win, even when two instructors use nearby shades.
      if (instructorByFill.has(color)) return instructorByFill.get(color);
      if (!color || assessmentFills.has(color)) return null;
      const candidates = [...instructorByFill].map(([knownColor, assignment]) => ({
        assignment,
        distance: colorDistance(color, knownColor),
      }));
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
          assessmentFills.has(fill(cell(sheet, r, c)))
        )
          continue;
        if (
          merges.some(
            (m) =>
              r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c && (r !== m.s.r || c !== m.s.c),
          )
        )
          continue;
        const students = [];
        raw
          .split(/\r?\n|;/)
          .filter(Boolean)
          .forEach((rawValue) => {
            // Inline course notes describe the session, not part of a student's name.
            const value = rawValue
              .replace(
                /\s+(?:g(?:r(?:ade)?)?\s*\d+(?:\s+ap)?|physics?|phy|needs?\s+attention)\s*$/i,
                '',
              )
              .trim();
            if (!value) return;
            if (lowercaseInstructor(value)) return;
            const match = matchName(value, roster);
            if (match) {
              students.push(match.entry.name);
              if (match.abbreviated)
                warnings.push(
                  `${address(r, c)}: ${value} matched uniquely to ${match.entry.name}.`,
                );
            } else if (
              isHighSchool(staff) &&
              !ignoredNote.test(value) &&
              /^[\p{L}\p{M}][\p{L}\p{M}'’ .-]*$/u.test(value.trim())
            ) {
              // High-school students need not have a binder. Keep their schedule name.
              students.push(value.trim());
            } else if (!ignoredNote.test(value))
              warnings.push(
                `${address(r, c)} (${staff}): unmatched text "${value}" was not imported. It may be a student or a note.`,
              );
          });
        if (students.length) starts.push({ r, students });
      }
      starts.forEach((entry, index) => {
        const start = times.get(entry.r);
        if (start === undefined) {
          warnings.push(
            `${address(entry.r, c)}: no start-time label for ${entry.students.join(', ')}; appointment omitted.`,
          );
          return;
        }
        const next = starts[index + 1]?.r;
        const color = fill(cell(sheet, entry.r, c));
        const assignment = instructorForColor(color);
        const instructor = assignment?.teacher || (color ? 'ACD' : staff);
        if (color && instructorByFill.has(color) && !assignment)
          warnings.push(
            `${address(entry.r, c)}: this color has conflicting instructor names; ${entry.students.join(', ')} grouped under ACD.`,
          );
        // Notes belong to the preceding student in this lane, up to the next
        // student or the one-hour session boundary, regardless of fill shade.
        const hourEndRow =
          [...times].find(([r, time]) => r > entry.r && time >= start + 60)?.[0] ?? bounds.e.r + 1;
        const noteEnd = Math.min(next ?? bounds.e.r + 1, hourEndRow);
        const notes = [];
        for (let r = entry.r; r < noteEnd; r++) {
          if (!teacherCells.has(address(r, c)) && !assessmentFills.has(fill(cell(sheet, r, c))))
            notes.push(text(cell(sheet, r, c)));
        }
        if (notes.some((note) => /\bcancel(?:led|ed)?\b/i.test(note))) {
          warnings.push(
            `${address(entry.r, c)} (${staff}): ${entry.students.join(', ')} omitted because the block is marked cancelled.`,
          );
          return;
        }
        if (notes.some((note) => /\bmove\b/i.test(note)))
          warnings.push(
            `${address(entry.r, c)} (${staff}): appointment contains a move note; verify its time.`,
          );
        entry.students.forEach((student) =>
          rows.push({
            staff: instructor,
            table: assignment?.table || '',
            section: staff,
            student,
            start: clock(start),
            notes: sessionNotes(notes),
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
      if (!appearances.has(row.start))
        appearances.set(row.start, { ...row, notes: [...row.notes] });
      else {
        const original = appearances.get(row.start);
        original.notes = [...new Set([...original.notes, ...row.notes])];
        warnings.push(
          `Duplicate entry for ${row.student}, ${row.staff}, ${row.start}; kept one occurrence.`,
        );
      }
    });
    const appointments = [];
    sessions.forEach((appearances) => {
      const entries = [...appearances.values()].sort((a, b) => a.start.localeCompare(b.start));
      for (let i = 0; i < entries.length; i++) {
        const row = entries[i],
          next = entries[i + 1];
        const twoHours = next && minutes(next.start) - minutes(row.start) === 60;
        const durationMinutes = twoHours ? 120 : 60;
        const notes = twoHours ? [...new Set([...row.notes, ...next.notes])] : row.notes;
        appointments.push({
          ...row,
          notes,
          durationMinutes,
          end: clock(minutes(row.start) + durationMinutes),
        });
        if (twoHours) i++;
      }
    });
    return { rows: appointments, warnings };
  }
  return { key, isHighSchool, readBinders, readSchedule, minutes };
})();
