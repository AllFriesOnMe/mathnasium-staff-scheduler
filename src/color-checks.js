'use strict';

const ScheduleColorChecks = (() => {
  function create(context) {
    const {
      sheet,
      r,
      c,
      students,
      start,
      teachingStart,
      staff,
      color,
      bounds,
      heading,
      times,
      palette,
      assessmentFills,
      distance,
      cell,
      fill,
      address,
      clock,
      instructorForColor,
    } = context;
    const nearby = [];
    for (let rr = Math.max(heading + 1, r - 2); rr <= Math.min(bounds.e.r, r + 2); rr++)
      for (let cc = Math.max(bounds.s.c, c - 4); cc <= Math.min(bounds.e.c, c + 4); cc++) {
        const entry = cell(sheet, rr, cc),
          shade = fill(entry);
        nearby.push({
          cell: address(rr, cc),
          text: String(entry?.v ?? '').slice(0, 200),
          color: shade,
          instructor: instructorForColor(shade)?.teacher || null,
          time: times.get(rr) ?? null,
        });
      }
    const close = palette
      .map((entry) => ({ ...entry, distance: distance(color, entry.color) }))
      .filter((entry) => entry.distance <= 60)
      .sort((a, b) => a.distance - b.distance);
    if (!close.length) return null;
    // Never let a color-only check turn a possible assessment into teaching.
    if ([...assessmentFills].some((shade) => distance(color, shade) <= close[0].distance + 10))
      return null;
    const candidates = close
      .filter((entry) => nearby.some((neighbor) => neighbor.instructor === entry.instructor))
      .slice(0, 12)
      .map((entry) => ({ ...entry, distance: Math.round(entry.distance * 100) / 100 }));
    if (!candidates.length) return null;
    const instructors = [
      ...new Set([
        ...palette.map((entry) => entry.instructor),
        ...nearby.map((entry) => entry.instructor).filter(Boolean),
      ]),
    ];
    if (instructors.length > 100 || students.length > 100) return null;
    return {
      kind: 'color',
      cell: address(r, c),
      color,
      text: `Color ${color} for ${students.join(', ')}`.slice(0, 2000),
      reason:
        'Color is too different or too close to competing instructor shades for a reliable local match.',
      time: clock(teachingStart),
      instructors,
      candidates,
      nearby,
      appointments: students.map((student, i) => ({
        id: `a${i}`,
        student,
        staff,
        start: clock(start),
        end: clock(start + 60),
      })),
    };
  }

  function resolve(item, proposals) {
    const matches = Array.isArray(proposals)
      ? proposals.filter((entry) => entry?.cell === item.cell)
      : [];
    if (matches.length !== 1) return null;
    const entry = matches[0];
    if (
      entry.kind !== 'color' ||
      entry.text !== item.text ||
      entry.color !== item.color ||
      entry.time !== item.time ||
      entry.confident === false ||
      !item.candidates.some((candidate) => candidate.instructor === entry.incomingInstructor) ||
      !Array.isArray(entry.studentIds) ||
      entry.studentIds.length !== item.appointments.length ||
      new Set(entry.studentIds).size !== entry.studentIds.length ||
      !item.appointments.every((row) => entry.studentIds.includes(row.id))
    )
      return null;
    return { teacher: entry.incomingInstructor, table: '' };
  }
  return { create, resolve };
})();
