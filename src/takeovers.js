'use strict';

const ScheduleTakeovers = (() => {
  const actionPattern =
    /\b(?:take[\s-]*over|takes[\s-]*over|taking[\s-]*over|taken[\s-]*over|hand[\s-]*(?:over|off)|handed\s+over|handover|handoff|send|transfers?|transferred|pass)\b|\b(?:move|give)\b.+\bto\b|\b(?:goes|going|should go|will go)\s+to\b/i;
  const isNote = (value) => actionPattern.test(String(value));
  const clean = (value) =>
    String(value)
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/\s+/g, ' ')
      .trim();
  const timePattern =
    /\b(?:(\d{1,2})[:.](\d{2})(?::00)?\s*([ap]\.?m\.?)?|(\d{1,2})\s*([ap]\.?m\.?))\b/gi;
  const timeOnly = /^(?:at\s*|@\s*)?\d{1,2}[:.]\d{2}(?::00)?(?:\s*[ap]\.?m\.?)?\s*$/i;
  const withoutTime = (value) =>
    clean(value)
      .replace(timePattern, '')
      .replace(/\s+(?:starting|beginning|effective|from|after|at|@)(?:\s+(?:at|from))?\s*$/i, '')
      .replace(/\s*[.!:]\s*$/, '')
      .trim();
  const looksLikeName = (value) =>
    /^[\p{L}][\p{L}\p{M}' .-]{0,70}$/u.test(value) &&
    !/\b(?:students?|kids?|over|takes?|from|to|please|table|start|end|assessment|them|him|her|someone|another|starting|beginning|on|needs?|attention|physics?|phy|na|ol|online|additional|makeup)\b/i.test(
      value,
    );

  function parseSentence(text) {
    let body = withoutTime(clean(text).replace(/\s+please[.!]?$/i, ''))
      .replace(/\b(take|takes|taking|taken)-over\b/gi, '$1 over')
      .replace(/\bhand[\s-]+off\b/gi, 'handoff')
      .replace(/^please\s+/i, '')
      .replace(/\s+please$/i, '');
    let target = '',
      source = '',
      subjects = '',
      match;
    if (
      (match = body.match(
        /^(?:send|transfer|move|pass|give|hand\s*over|handover|handoff)\s+(?:over\s+)?to\s+(.+)$/i,
      ))
    ) {
      target = match[1];
    } else if (
      (match = body.match(
        /^(?:send|transfer|move|pass|give|hand\s*over|handover|handoff)\s+(.+?)\s+(?:over\s+)?to\s+(.+)$/i,
      ))
    ) {
      [, subjects, target] = match;
    } else if (
      (match = body.match(
        /^(.+?)\s+(?:goes?|going|should\s+go|will\s+go|moves?|transfers?)\s+to\s+(.+)$/i,
      ))
    ) {
      [, subjects, target] = match;
    } else if (
      (match = body.match(
        /^(.+?)\s+(?:is\s+|will\s+be\s+)?(?:taken\s+over|handed\s+over|transferred)\s+(?:by|to)\s+(.+)$/i,
      ))
    ) {
      [, subjects, target] = match;
    } else if (
      (match = body.match(
        /^(?:have\s+)?(.+?)\s+(?:will\s+|to\s+|is\s+)?(?:takes?|taking)\s*over\b\s*(.*)$/i,
      ))
    ) {
      target = match[1];
      subjects = match[2];
    } else if ((match = body.match(/^(?:take\s*over|hand\s*over|handover|handoff)\b\s*(.*)$/i))) {
      subjects = match[1];
      const destination = subjects.match(/(?:^|\s)(?:by|to)\s+(.+)$/i);
      if (destination) {
        target = destination[1];
        subjects = subjects.slice(0, destination.index);
      }
    } else return null;
    const from = subjects.match(/(?:^|\s)from\s+(.+)$/i);
    if (from) {
      source = from[1];
      subjects = subjects.slice(0, from.index).trim();
    }
    subjects = subjects.replace(/^(?:all(?: of)?|the rest of|the remaining)\s+/i, '');
    const group =
      subjects.match(/^(.+?)(?:'s|s')\s+(?:students?|kids?|group)(?:\s+over)?$/i) ||
      subjects.match(/^(.+?)\s+(?:students|kids)$/i);
    if (group) {
      source = group[1];
      subjects = '';
    }
    subjects = subjects
      .replace(/^(?:the\s+)?(?:students?|kids?)\s+/i, '')
      .replace(/^(?:the\s+)?(?:students|kids)$/i, '')
      .trim();
    target = target.replace(/^instructor\s+/i, '').trim();
    source = source.replace(/^instructor\s+/i, '').trim();
    return { target, source, subjects };
  }

  function collect({ sheet, bounds, heading, cell, text, address }) {
    const notes = [],
      noteCells = new Set();
    const merges = sheet['!merges'] || [];
    for (let r = heading + 1; r <= bounds.e.r; r++) {
      for (let c = bounds.s.c; c <= bounds.e.c; c++) {
        if (noteCells.has(address(r, c))) continue;
        const raw = text(cell(sheet, r, c));
        const splitAction =
          /^(?:take|takes|taking|hand)\s*$/i.test(raw) &&
          /^over\s*$/i.test(text(cell(sheet, r + 1, c)));
        if (!isNote(raw) && !splitAction) continue;
        const pieces = [raw];
        const evidenceCells = [address(r, c)];
        noteCells.add(address(r, c));
        const merge = merges.find((m) => m.s.r === r && m.s.c === c);
        let lastColumn = merge?.e.c ?? c;
        // Join explicitly connected horizontal fragments, not arbitrary nearby students.
        for (let cc = c + 1; cc <= Math.min(bounds.e.c, c + 2); cc++) {
          const value = text(cell(sheet, r, cc));
          if (
            !value ||
            typeof cell(sheet, r, cc)?.v === 'number' ||
            (/\b(?:to|from)\s*$/i.test(pieces.join(' ')) && !looksLikeName(withoutTime(value))) ||
            !(/\b(?:to|from)\s*$/i.test(pieces.join(' ')) || /^(?:to|from|at|@)\b/i.test(value))
          )
            break;
          pieces.push(value);
          evidenceCells.push(address(r, cc));
          noteCells.add(address(r, cc));
          lastColumn = cc;
        }
        for (let rr = r + 1; rr <= Math.min(bounds.e.r, r + 3); rr++) {
          const value = text(cell(sheet, rr, c));
          const needsObject = /^(?:take\s*over|hand\s*over|handover|handoff)\s*$/i.test(
            pieces.join(' '),
          );
          const needsDestination = /\b(?:to|from)\s*$/i.test(pieces.join(' '));
          const needsOver =
            /^(?:take|takes|taking|hand)\s*$/i.test(pieces.join(' ')) && /^over$/i.test(value);
          if (
            !value ||
            !(
              timeOnly.test(value) ||
              ((needsObject || needsDestination) && looksLikeName(value)) ||
              needsOver
            )
          )
            break;
          pieces.push(value);
          evidenceCells.push(address(rr, c));
          noteCells.add(address(rr, c));
        }
        // A time-only annotation immediately above belongs to this instruction.
        // Append it so the action remains at the start of the parsed sentence.
        const above = address(r - 1, c);
        const precedingTime = text(cell(sheet, r - 1, c));
        if (r - 1 > heading && !noteCells.has(above) && timeOnly.test(precedingTime)) {
          pieces.push(precedingTime);
          evidenceCells.push(above);
          noteCells.add(above);
        }
        notes.push({
          r,
          c,
          lastColumn,
          cells: evidenceCells,
          text: clean(pieces.join(' ')).slice(0, 2000),
        });
      }
    }
    return { notes, noteCells };
  }

  function incomingName(text) {
    const name = parseSentence(text)?.target || '';
    return looksLikeName(name) ? name : '';
  }

  function apply(appointments, context) {
    const {
      sheet,
      notes,
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
      unresolvedHandoffs = [],
      aiApplied = [],
      resolutions = [],
    } = context;
    const timeCache = new Map();
    function timesFor(c) {
      const tc = [...timeColumns].sort((a, b) => Math.abs(a - c) - Math.abs(b - c) || a - b)[0];
      if (!timeCache.has(tc)) {
        const values = [];
        for (let r = heading + 1; r <= bounds.e.r; r++) {
          const time = minutes(cell(sheet, r, tc)?.v);
          if (time !== null) values.push({ r, time });
        }
        timeCache.set(tc, values);
      }
      return timeCache.get(tc);
    }
    function noteTime(note) {
      const matches = [...note.text.matchAll(timePattern)];
      const times = timesFor(note.c),
        rowTime = times.find((t) => t.r === note.r)?.time;
      if (!matches.length) return /(?:\bat\b|@)\s*\d/i.test(note.text) ? null : (rowTime ?? null);
      if (matches.length > 1) return null;
      const hour = matches[0][1] ?? matches[0][4],
        minute = matches[0][2] ?? '00',
        suffix = matches[0][3] ?? matches[0][5];
      if (+minute > 59 || +hour > 23) return null;
      if (suffix) {
        const explicit = minutes(`${hour}:${minute} ${suffix.replaceAll('.', '')}`);
        const lane = appointments.filter(
          (appointment) =>
            appointment.sources.some(
              (source) => source.column >= note.c && source.column <= note.lastColumn,
            ) &&
            rowTime !== undefined &&
            minutes(appointment.start) <= rowTime &&
            rowTime < minutes(appointment.end),
        );
        if (
          lane.length &&
          !lane.some(
            (appointment) =>
              minutes(appointment.start) <= explicit && explicit < minutes(appointment.end),
          )
        ) {
          const target = parseSentence(note.text)?.target;
          const label = `Transfer${target ? ` to ${target}` : ''} at ${hour}:${minute} ${suffix.replaceAll('.', '').toUpperCase()} (May be incorrect input)`;
          // Bases have already been built; retain this visible note on the output too.
          for (const base of bases) {
            if (
              lane.some(
                (appointment) =>
                  key(appointment.student) === key(base.student) &&
                  base.sources.some(
                    (source) => source.column >= note.c && source.column <= note.lastColumn,
                  ),
              )
            ) {
              base.notes = [...new Set([...base.notes, label])];
              base.rationale = [
                ...(base.rationale || []),
                {
                  rule: 'Unresolved handoff',
                  detail: `${address(note.r, note.c)}: "${note.text}". The written time falls outside this student's appointment; no transfer was applied. ${label}.`,
                },
              ];
              // A contradictory written time must not silently fall back to a
              // color boundary elsewhere in this same appointment.
              events.set(
                base,
                events
                  .get(base)
                  .filter(
                    (event) =>
                      event.explicit ||
                      !lane.some(
                        (appointment) =>
                          minutes(appointment.start) <= event.time &&
                          event.time < minutes(appointment.end),
                      ),
                  ),
              );
            }
          }
          warnings.push(
            `${address(note.r, note.c)}: ${label}; transfer time is outside the student's appointment.`,
          );
          note.invalidTimeInput = true;
          return null;
        }
        return explicit;
      }
      if (+hour > 12 || +hour === 0) return +hour * 60 + +minute;
      const candidates = [(+hour % 12) * 60 + +minute, (+hour % 12) * 60 + +minute + 720];
      const reference = rowTime ?? times[0]?.time;
      if (reference === undefined) return null;
      candidates.sort((a, b) => Math.abs(a - reference) - Math.abs(b - reference));
      return Math.abs(candidates[0] - reference) === Math.abs(candidates[1] - reference)
        ? null
        : candidates[0];
    }
    const teacherRegistry = new Map();
    for (const name of [...instructors, ...appointments.map((a) => a.staff)])
      if (!teacherRegistry.has(key(name))) teacherRegistry.set(key(name), name);
    for (const note of notes) {
      const name = incomingName(note.text);
      if (!name) continue;
      const parts = key(name).split(' ');
      const alreadyRepresented = [...teacherRegistry.values()].some((known) => {
        const full = key(known).split(' ');
        return (
          full[0] === parts[0] &&
          full.length >= parts.length &&
          parts.slice(1).every((part, i) => full[i + 1].startsWith(part))
        );
      });
      if (!alreadyRepresented) teacherRegistry.set(key(name), name);
    }
    const teacherNames = [...teacherRegistry.values()];
    function resolveName(phrase, names, allowNew = false) {
      const wanted = key(phrase),
        parts = wanted.split(' ');
      const exact = names.filter((name) => key(name) === wanted);
      if (exact.length === 1) return exact[0];
      const matches = names.filter((name) => {
        const full = key(name).split(' ');
        return (
          full[0] === parts[0] &&
          full.length >= parts.length &&
          parts.slice(1).every((part, i) => full[i + 1].startsWith(part))
        );
      });
      const unique = [...new Set(matches)];
      if (unique.length === 1) return unique[0];
      return !unique.length && allowNew && looksLikeName(phrase) ? phrase.trim() : null;
    }
    // A repeated name at a color-change boundary is a continuation, not a
    // second appointment starting its own hour. Hourly repeats may span staff.
    const bases = [];
    for (const row of [...appointments].sort((a, b) => a.start.localeCompare(b.start))) {
      const prior = bases.find(
        (base) =>
          key(base.student) === key(row.student) &&
          base.staff !== row.staff &&
          row.start > base.start &&
          row.start < base.end &&
          base.sources.some((s) => s.column === row.sourceColumn),
      );
      const adjacent = bases.find(
        (base) =>
          key(base.student) === key(row.student) &&
          base.staff !== row.staff &&
          base.end === row.start &&
          base.durationMinutes === 60 &&
          row.durationMinutes === 60,
      );
      if (prior || adjacent) {
        const base = prior || adjacent;
        base.sources.push(...row.sources);
        base.namedEvents.push({
          time: minutes(row.start),
          target: row.staff,
          explicit: false,
          evidence: {
            rule: 'Handoff',
            detail: `Repeated student ${address(base.sources[0].row, base.sources[0].column)} -> ${address(row.sourceRow, row.sourceColumn)}: the new occurrence is assigned to ${row.staff} at ${row.start}. Treated as a continuation of the original appointment, not a new default hour.`,
          },
        });
        base.notes = [...new Set([...base.notes, ...row.notes])];
        base.online ||= row.online;
        if (adjacent && !prior) {
          base.end = row.end;
          base.durationMinutes = 120;
          base.bookedDurationMinutes = 120;
        }
      } else
        bases.push({
          ...row,
          notes: [...(row.notes || [])],
          sources: [...row.sources],
          namedEvents: [],
        });
    }
    const events = new Map(bases.map((base) => [base, [...base.namedEvents]]));
    for (const base of bases) {
      for (const source of base.sources) {
        let previous = base.staff;
        let previousCell = address(source.row, source.column);
        let previousColor = fill(cell(sheet, source.row, source.column));
        for (const { r, time } of timesFor(source.column)) {
          if (r < source.row || time < minutes(source.teachingStart) || time >= minutes(base.end))
            continue;
          if (time < minutes(base.start)) continue;
          // Another student's name ends this lane's evidence for this booking.
          if (source.stopRow !== null && r >= source.stopRow) break;
          const color = fill(cell(sheet, r, source.column));
          if (!color || assessmentFills.has(color)) continue;
          const assignment = instructorForColor(color, r, source.column);
          if (!assignment) continue;
          if (assignment.teacher === previous) {
            previousCell = address(r, source.column);
            previousColor = color;
            continue;
          }
          if (time > minutes(base.start))
            events.get(base).push({
              time,
              target: assignment.teacher,
              explicit: false,
              evidence: {
                rule: 'Handoff',
                detail:
                  color === previousColor
                    ? `Instructor labels for the shared color ${color} change the assignment from ${previous} at ${previousCell} to ${assignment.teacher} at ${address(r, source.column)}, time ${clock(time)}. This is a label-based change, not a change of fill color.`
                    : `Color changed from ${previousColor || 'unfilled'} at ${previousCell} (${previous}) to ${color} at ${address(r, source.column)} (${assignment.teacher}). The time label in ${address(r, [...timeColumns].sort((a, b) => Math.abs(a - source.column) - Math.abs(b - source.column) || a - b)[0])} supplies ${clock(time)}. Transfer inferred from that boundary.`,
              },
            });
          previous = assignment.teacher;
          previousCell = address(r, source.column);
          previousColor = color;
        }
      }
    }
    function acceptedEvents(base, reportConflicts = false) {
      const byTime = new Map();
      for (const event of events.get(base)) {
        if (!byTime.has(event.time)) byTime.set(event.time, []);
        byTime.get(event.time).push(event);
      }
      const accepted = [];
      let hasExplicit = false;
      for (const [time, list] of [...byTime].sort(([a], [b]) => a - b)) {
        const explicit = list.filter((event) => event.explicit);
        if (!explicit.length && hasExplicit) continue;
        const targets = [
          ...new Set((explicit.length ? explicit : list).map((event) => event.target)),
        ];
        if (targets.length > 1) {
          if (reportConflicts) {
            warnings.push(
              `Conflicting handoffs for ${base.student} at ${clock(time)}; that handoff was not applied.`,
            );
            base.rationale = [
              ...(base.rationale || []),
              {
                rule: 'Unresolved handoff',
                detail: `Conflicting destinations ${targets.join(', ')} at ${clock(time)}; this boundary was not applied.`,
              },
            ];
          }
          continue;
        }
        accepted.push({
          time,
          target: targets[0],
          evidence: (explicit.length ? explicit : list).find((event) => event.target === targets[0])
            ?.evidence,
        });
        if (explicit.length) hasExplicit = true;
      }
      return accepted;
    }
    function staffAt(base, time) {
      return (
        acceptedEvents(base)
          .filter((event) => event.time <= time)
          .at(-1)?.target || base.staff
      );
    }
    function contextualInstructor(note) {
      const ownColor = fill(cell(sheet, note.r, note.c));
      // A colored note is authoritative. Do not bypass an ambiguous shade or
      // assessment block by borrowing another instructor's nearby color.
      if (ownColor) return instructorForColor(ownColor, note.r, note.c)?.teacher;
      function neighbor(dr, dc) {
        for (let distance = 1; distance <= 4; distance++) {
          const r = note.r + dr * distance,
            c = note.c + dc * distance;
          if (
            r <= heading ||
            r > bounds.e.r ||
            c < bounds.s.c ||
            c > bounds.e.c ||
            timeColumns.includes(c)
          )
            return null;
          const color = fill(cell(sheet, r, c));
          // Stop at the first colored cell, including unknown/assessment fills.
          if (color) return instructorForColor(color, r, c)?.teacher || null;
        }
        return null;
      }
      const pairs = [
        [neighbor(0, -1), neighbor(0, 1)],
        [neighbor(-1, 0), neighbor(1, 0)],
      ];
      const enclosing = pairs.filter(([a, b]) => a && b && key(a) === key(b)).map(([a]) => a);
      // Both axes must agree when each provides an enclosing block. A single
      // nearby instructor, or two different bounding instructors, is not enough.
      if (pairs.some(([a, b]) => a && b && key(a) !== key(b))) return null;
      return enclosing.length && enclosing.every((name) => key(name) === key(enclosing[0]))
        ? enclosing[0]
        : null;
    }
    function pairedDestination(active, note) {
      const targets = active.map((base) => [
        ...new Set(
          events
            .get(base)
            .filter(
              (event) =>
                event.explicit &&
                event.namedTarget &&
                event.time === note.time &&
                event.noteCell !== address(note.r, note.c),
            )
            .map((event) => event.target),
        ),
      ]);
      if (!targets.length || targets.some((list) => list.length !== 1)) return null;
      return targets.every((list) => key(list[0]) === key(targets[0][0])) ? targets[0][0] : null;
    }
    function nearbySupports(note, teacher) {
      const ownColor = fill(cell(sheet, note.r, note.c));
      if (ownColor)
        return key(instructorForColor(ownColor, note.r, note.c)?.teacher || '') === key(teacher);
      for (const [dr, dc] of [
        [0, -1],
        [0, 1],
        [-1, 0],
        [1, 0],
      ]) {
        for (let distance = 1; distance <= 4; distance++) {
          const r = note.r + dr * distance,
            c = note.c + dc * distance;
          if (
            r <= heading ||
            r > bounds.e.r ||
            c < bounds.s.c ||
            c > bounds.e.c ||
            timeColumns.includes(c)
          )
            break;
          const color = fill(cell(sheet, r, c));
          if (!color) continue;
          if (key(instructorForColor(color, r, c)?.teacher || '') === key(teacher)) return true;
          break;
        }
      }
      return false;
    }
    const timedNotes = notes
      .map((note) => {
        const time = noteTime(note);
        return { ...note, time };
      })
      .sort(
        (a, b) =>
          (a.time ?? Infinity) - (b.time ?? Infinity) ||
          Number(Boolean(parseSentence(b.text)?.target)) -
            Number(Boolean(parseSentence(a.text)?.target)),
      );
    for (const note of timedNotes) {
      const warning = (reason) => {
        if (!note.invalidTimeInput) {
          const reference = note.time ?? timesFor(note.c).find((entry) => entry.r === note.r)?.time;
          for (const base of bases.filter(
            (base) =>
              reference !== undefined &&
              reference !== null &&
              minutes(base.start) <= reference &&
              reference < minutes(base.end) &&
              base.sources.some(
                (source) => source.column >= note.c && source.column <= note.lastColumn,
              ),
          ))
            base.rationale = [
              ...(base.rationale || []),
              {
                rule: 'Unresolved handoff',
                detail: `${address(note.r, note.c)}: "${note.text}". No transfer applied: ${reason}.`,
              },
            ];
        }
        if (!(note.invalidTimeInput && reason === 'time is missing or ambiguous'))
          warnings.push(
            `${address(note.r, note.c)}: handoff not applied (${reason}): "${note.text}".`,
          );
        if (unresolvedHandoffs.length >= 10) return;
        const nearby = [];
        for (let r = Math.max(heading + 1, note.r - 2); r <= Math.min(bounds.e.r, note.r + 2); r++)
          for (
            let c = Math.max(bounds.s.c, note.c - 4);
            c <= Math.min(bounds.e.c, note.c + 4);
            c++
          ) {
            const entry = cell(sheet, r, c),
              color = fill(entry);
            nearby.push({
              cell: address(r, c),
              text: String(entry?.v ?? '').slice(0, 200),
              color,
              instructor: instructorForColor(color, r, c)?.teacher || null,
              time: timesFor(c).find((t) => t.r === r)?.time ?? null,
            });
          }
        unresolvedHandoffs.push({
          cell: address(note.r, note.c),
          text: note.text,
          reason,
          time: note.time === null ? null : clock(note.time),
          instructors: teacherNames,
          appointments: bases
            .map((base, i) => ({
              id: `a${i}`,
              student: base.student,
              staff: staffAt(base, (note.time ?? minutes(base.start)) - 0.5),
              start: base.start,
              end: base.end,
            }))
            .filter(
              (base) =>
                note.time === null ||
                (minutes(base.start) <= note.time && note.time < minutes(base.end)),
            )
            .slice(0, 100),
          nearby,
        });
      };
      function useAiResolution() {
        const matches = resolutions.filter(
          (entry) => entry.cell === address(note.r, note.c) && entry.text === note.text,
        );
        if (matches.length !== 1) return false;
        const suggestion = matches[0],
          time = minutes(suggestion.time);
        const target = teacherNames.find((name) => name === suggestion.incomingInstructor);
        if (
          !target ||
          time === null ||
          (note.time !== null && time !== note.time) ||
          !Array.isArray(suggestion.studentIds) ||
          !suggestion.studentIds.length ||
          new Set(suggestion.studentIds).size !== suggestion.studentIds.length
        )
          return false;
        const selected = suggestion.studentIds.map((id) =>
          /^a\d+$/.test(id) ? bases[+id.slice(1)] : null,
        );
        if (
          selected.some((base) => !base || minutes(base.start) > time || time >= minutes(base.end))
        )
          return false;
        const parsed = parseSentence(note.text);
        const paired = pairedDestination(selected, note);
        if (paired && key(paired) !== key(target)) return false;
        if (
          parsed &&
          !parsed.subjects &&
          !parsed.source &&
          selected.some(
            (base) =>
              !base.sources.some(
                (source) => source.column >= note.c && source.column <= note.lastColumn,
              ),
          )
        )
          return false;
        if (parsed?.target && key(resolveName(parsed.target, teacherNames) || '') !== key(target))
          return false;
        if (parsed?.source) {
          const source = resolveName(parsed.source, teacherNames);
          if (!source || selected.some((base) => key(staffAt(base, time - 0.5)) !== key(source)))
            return false;
        }
        if (parsed?.subjects) {
          const phrases = parsed.subjects.split(/\s*(?:,|&|\band\b|\+)\s*/i).filter(Boolean);
          const active = bases.filter(
            (base) => minutes(base.start) <= time && time < minutes(base.end),
          );
          const resolved = phrases.map((phrase) =>
            resolveName(phrase, [...new Set(active.map((base) => base.student))]),
          );
          // When names are already unambiguous, AI cannot substitute other students.
          if (
            resolved.every(Boolean) &&
            (selected.length !== resolved.length ||
              selected.some((base) => !resolved.includes(base.student)))
          )
            return false;
        }
        for (const base of selected) {
          events.get(base).push({
            time,
            target,
            explicit: true,
            namedTarget: Boolean(parsed?.target),
            noteCell: address(note.r, note.c),
            evidence: {
              rule: 'AI-assisted handoff',
              detail: `AI-assisted resolution of "${note.text}" in ${(note.cells || [address(note.r, note.c)]).join(', ')} selected ${target} at ${clock(time)}. The proposed student, instructor and time passed validation against this appointment.`,
            },
          });
          base.notes.push('AI-assisted handoff');
        }
        aiApplied.push(
          `${address(note.r, note.c)}: AI-assisted handoff to ${target} at ${clock(time)}.`,
        );
        return true;
      }
      // Suggestions only get a chance at a note that the local rules cannot resolve.
      const unresolved = (reason) => {
        if (!useAiResolution()) warning(reason);
      };
      if (note.time === null) {
        unresolved('time is missing or ambiguous');
        continue;
      }
      const sentence = parseSentence(note.text);
      if (!sentence) {
        unresolved('sentence could not be resolved');
        continue;
      }
      let active = bases.filter(
        (base) => minutes(base.start) <= note.time && note.time < minutes(base.end),
      );
      if (!active.length) {
        unresolved('no matching appointment is active at that time');
        continue;
      }
      if (sentence.source) {
        const source = resolveName(sentence.source, teacherNames);
        if (!source) {
          unresolved('outgoing instructor is missing or ambiguous');
          continue;
        }
        active = active.filter((base) => key(staffAt(base, note.time - 0.5)) === key(source));
      }
      if (sentence.subjects) {
        const phrases = sentence.subjects.split(/\s*(?:,|&|\band\b|\+)\s*/i).filter(Boolean);
        const names = [...new Set(active.map((base) => base.student))];
        const resolved = phrases.map((phrase) => resolveName(phrase, names));
        if (resolved.some((name) => !name)) {
          unresolved('student reference is missing or ambiguous');
          continue;
        }
        active = active.filter((base) => resolved.includes(base.student));
      } else if (!sentence.source) {
        active = active.filter((base) =>
          base.sources.some(
            (source) => source.column >= note.c && source.column <= note.lastColumn,
          ),
        );
      }
      if (!active.length) {
        unresolved('no matching appointment is active at that time');
        continue;
      }
      const contextTeacher = contextualInstructor(note);
      const paired = !sentence.target && sentence.subjects ? pairedDestination(active, note) : null;
      if (paired && contextTeacher && key(paired) !== key(contextTeacher)) {
        warning('sending and receiving notes disagree about the incoming instructor');
        continue;
      }
      const target = sentence.target
        ? resolveName(sentence.target, teacherNames, true)
        : contextTeacher || (paired && nearbySupports(note, paired) ? paired : null);
      if (!target) {
        unresolved('incoming instructor is missing or ambiguous');
        continue;
      }
      for (const base of active) {
        if (
          !events
            .get(base)
            .some(
              (event) =>
                event.explicit && event.time === note.time && key(event.target) === key(target),
            )
        )
          events.get(base).push({
            time: note.time,
            target,
            explicit: true,
            namedTarget: Boolean(sentence.target),
            noteCell: address(note.r, note.c),
            evidence: {
              rule: 'Handoff',
              detail: `Written instruction in ${(note.cells || [address(note.r, note.c)]).join(', ')}: "${note.text}". Applied transfer from ${staffAt(base, note.time - 0.5)} to ${target} at ${clock(note.time)}${sentence.target ? '; destination named in the instruction' : contextTeacher ? '; destination inferred from surrounding instructor labels' : '; destination confirmed by the paired sending note and nearby instructor color'}${[...note.text.matchAll(timePattern)].length ? '; used the written time instead of the note row time' : '; used the schedule time beside the note'}.`,
            },
          });
      }
    }
    const output = [];
    for (const base of bases) {
      const accepted = acceptedEvents(base, true);
      let staff = base.staff,
        start = minutes(base.start),
        from = null;
      const parts = [];
      const appliedEvents = [];
      for (const event of accepted) {
        if (event.target === staff || event.time < start || event.time >= minutes(base.end))
          continue;
        appliedEvents.push(event);
        if (event.time > start)
          parts.push({ staff, start, end: event.time, from, to: event.target });
        from = staff;
        staff = event.target;
        start = event.time;
      }
      parts.push({ staff, start, end: minutes(base.end), from, to: null });
      for (const part of parts) {
        const notes = [...base.notes];
        if (part.from) notes.push(`Taking over from ${part.from} at ${clock(part.start)}`);
        if (part.to) notes.push(`Handoff to ${part.to} at ${clock(part.end)}`);
        if (parts.length > 1 && base.bookedDurationMinutes === 120)
          notes.push(`Part of 2-hour appointment ${base.start} - ${base.end}`);
        const { sources, sourceColumn, sourceRow, teachingStart, namedEvents, ...publicRow } = base;
        output.push({
          ...publicRow,
          staff: part.staff,
          table: part.staff === base.staff ? base.table : '',
          start: clock(part.start),
          end: clock(part.end),
          durationMinutes: part.end - part.start,
          notes,
          rationale: [
            ...(base.rationale || []),
            ...appliedEvents
              .filter((event) => event.time === part.start || event.time === part.end)
              .map((event) => event.evidence)
              .filter(Boolean),
          ],
          appointmentStart: base.start,
          appointmentEnd: base.end,
        });
      }
    }
    return output;
  }

  return { isNote, collect, incomingName, apply };
})();
