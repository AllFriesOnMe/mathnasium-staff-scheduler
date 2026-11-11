'use strict';

const ScheduleEdits = (() => {
  const minutes = (value) => SpreadsheetParser.minutes(value);
  const clock = (value) =>
    `${String(Math.floor(value / 60)).padStart(2, '0')}:${String(value % 60).padStart(2, '0')}`;
  const handoffNote = (value) =>
    /^(?:Handoff to |Taking over from |Part of 2-hour appointment |Table change to |\(switch to table |Continuing from Table )/i.test(
      value,
    );
  const editableLabel = (value) =>
    !handoffNote(value) &&
    !/^Needs attention$/i.test(value) &&
    !/^Transfer.*\(May be incorrect input\)$/i.test(value);
  const copy = (value) => JSON.parse(JSON.stringify(value));

  function create(input, roster = new Map()) {
    let rows = input;
    let sequence = 0;
    const history = [];
    const staff = new Map();
    const tables = new Map();
    const staffNotes = new Map();
    const bookings = new Map();
    for (const row of rows) {
      row._editId = `student-${++sequence}`;
      const key = JSON.stringify([
        row.student,
        row.appointmentStart || row.start,
        row.appointmentEnd || row.end,
      ]);
      if (!bookings.has(key)) bookings.set(key, `booking-${++sequence}`);
      row._bookingId = bookings.get(key);
      if (!staff.has(row.staff)) staff.set(row.staff, row.table || '');
      else if (!staff.get(row.staff) && row.table) staff.set(row.staff, row.table);
      if (!tables.has(row.staff)) tables.set(row.staff, new Set());
      if (row.table) tables.get(row.staff).add(String(row.table));
    }
    function checkpoint() {
      history.push({
        rows: copy(rows),
        tables: [...tables].map(([name, values]) => [name, [...values]]),
        staffNotes: copy([...staffNotes]),
      });
      if (history.length > 20) history.shift();
    }
    function tableValue(value, teacher) {
      if (value === undefined)
        return (
          (tables.get(teacher).has(staff.get(teacher))
            ? staff.get(teacher)
            : [...tables.get(teacher)][0]) || ''
        );
      const table = String(value).trim();
      if (table && !/^[1-9]\d{0,2}$/.test(table))
        throw new Error('Table number must be a whole number from 1 to 999, or blank.');
      return table;
    }
    const get = (id) => {
      const row = rows.find((entry) => entry._editId === id);
      if (!row) throw new Error('This appointment is no longer available.');
      return row;
    };
    const parts = (id) => {
      const booking = get(id)._bookingId;
      return rows
        .filter((row) => row._bookingId === booking)
        .sort((a, b) => a.start.localeCompare(b.start));
    };
    function saveDraft(draft, id) {
      const name = String(draft.student || '')
        .trim()
        .replace(/\s+/g, ' ');
      if (!name || name.length > 150)
        throw new Error('Enter a student name of up to 150 characters.');
      const start = minutes(draft.start),
        end = minutes(draft.end);
      if (start === null || end === null || end <= start)
        throw new Error('End time must be after start time on the same day.');
      if (!staff.has(draft.staff)) throw new Error('Choose an existing instructor.');
      let previous = start;
      let teacher = draft.staff;
      let table = tableValue(draft.table, teacher);
      const segments = [];
      for (const transfer of draft.handoffs || []) {
        const time = minutes(transfer.time);
        if (time === null || time <= previous || time >= end)
          throw new Error('Handoff times must be in order, strictly inside the appointment.');
        if (!staff.has(transfer.staff))
          throw new Error('Choose an existing instructor for each handoff.');
        const nextTable = tableValue(transfer.table, transfer.staff);
        if (transfer.staff === teacher && nextTable === table)
          throw new Error('Each handoff must change the instructor or table.');
        segments.push({ staff: teacher, table, start: clock(previous), end: clock(time) });
        previous = time;
        teacher = transfer.staff;
        table = nextTable;
      }
      segments.push({ staff: teacher, table, start: clock(previous), end: clock(end) });
      const old = id ? parts(id) : [];
      const template = old[0] || {};
      const booking = template._bookingId || `booking-${++sequence}`;
      const matchingName =
        old.length && SpreadsheetParser.key(template.student) === SpreadsheetParser.key(name);
      const binder = roster.get(SpreadsheetParser.key(name));
      const canonical = binder?.name || name;
      if (
        rows.some(
          (row) =>
            row._bookingId !== booking &&
            SpreadsheetParser.key(row.student) === SpreadsheetParser.key(canonical) &&
            minutes(row.start) < end &&
            start < minutes(row.end),
        )
      )
        throw new Error('This student already has an appointment during those times.');
      const labelText =
        draft.labels === undefined
          ? old
              .flatMap((row) => row.notes || [])
              .filter(editableLabel)
              .join('\n')
          : String(draft.labels);
      if (labelText.length > 2000)
        throw new Error('Additional labels must be 2,000 characters or fewer.');
      const notes = [
        ...new Set(
          labelText
            .split(/\r?\n|;/)
            .map((label) => label.trim())
            .filter(Boolean)
            .filter(editableLabel),
        ),
      ];
      if (!draft.attention) {
        for (let i = notes.length - 1; i >= 0; i--)
          if (/^Needs attention$/i.test(notes[i])) notes.splice(i, 1);
      } else if (!notes.some((note) => /^Needs attention$/i.test(note)))
        notes.push('Needs attention');
      const result = segments.map((segment, index) => ({
        ...template,
        ...segment,
        student: binder?.name || name,
        binder: binder?.location || (matchingName ? template.binder || '' : ''),
        online: Boolean(draft.online),
        table: segment.table,
        section: segment.staff,
        appointmentStart: clock(start),
        appointmentEnd: clock(end),
        sessionStart: clock(start),
        durationMinutes: minutes(segment.end) - minutes(segment.start),
        bookedDurationMinutes: end - start,
        manualRationale: [
          ...(template.manualRationale || []),
          {
            rule: 'Manual edit',
            detail: `${draft.manualAction || (old.length ? 'Edited appointment' : 'Added appointment')} by the user. This portion is assigned to ${segment.staff}, ${segment.table ? `Table ${segment.table}` : 'no table assigned'}, ${segment.start} - ${segment.end}. ${index ? `Incoming from ${segments[index - 1].staff} at ${segment.start}. ` : ''}${index < segments.length - 1 ? `Continues with ${segments[index + 1].staff} at ${segment.end}. ` : ''}Binder ${binder?.location || (matchingName ? template.binder || 'not found' : 'not found')}. These values replace automatic assignments for this portion.`,
          },
        ],
        notes: [
          ...notes,
          ...(index
            ? [
                segments[index - 1].staff === segment.staff
                  ? `Continuing from Table ${segments[index - 1].table || 'unassigned'} at ${segment.start}`
                  : `Taking over from ${segments[index - 1].staff} at ${segment.start}`,
              ]
            : []),
          ...(index < segments.length - 1
            ? [
                segments[index + 1].staff === segment.staff
                  ? `(switch to table ${segments[index + 1].table || 'unassigned'})`
                  : `Handoff to ${segments[index + 1].staff} at ${segment.end}`,
              ]
            : []),
          ...(segments.length > 1 && end - start === 120
            ? [`Part of 2-hour appointment ${clock(start)} - ${clock(end)}`]
            : []),
        ],
        _bookingId: booking,
        _editId: `student-${++sequence}`,
      }));
      checkpoint();
      for (const segment of result) if (segment.table) tables.get(segment.staff).add(segment.table);
      rows = [...rows.filter((row) => row._bookingId !== booking), ...result];
      return rows;
    }
    function draft(id) {
      const segments = parts(id),
        first = segments[0];
      const handoffs = segments
        .slice(1)
        .map((row) => ({ staff: row.staff, table: row.table || '', time: row.start }));
      if (!handoffs.length) {
        const unresolved = first.notes
          ?.map((note) =>
            note.match(/^Transfer to (.+) at (\d{1,2}:\d{2} [AP]M) \(May be incorrect input\)$/i),
          )
          .find(Boolean);
        if (unresolved) {
          const target = [...staff.keys()].find(
            (name) => SpreadsheetParser.key(name) === SpreadsheetParser.key(unresolved[1]),
          );
          const time = minutes(unresolved[2]);
          if (target && time !== null) handoffs.push({ staff: target, time: clock(time) });
        }
      }
      return {
        student: first.student,
        staff: first.staff,
        table: first.table || '',
        start: first.start,
        end: segments.at(-1).end,
        online: first.online,
        attention: first.notes?.some((note) => /^Needs attention$/i.test(note)),
        labels: [...new Set(segments.flatMap((row) => row.notes || []).filter(editableLabel))].join(
          '\n',
        ),
        handoffs,
      };
    }
    return {
      get rows() {
        return rows;
      },
      staff,
      staffNotes,
      saveStaffNote(teacher, text, id) {
        if (!staff.has(teacher)) throw new Error('Choose an existing instructor.');
        const value = String(text || '').trim();
        if (!value || value.length > 2000)
          throw new Error('Enter a note of up to 2,000 characters.');
        const notes = staffNotes.get(teacher) || [];
        if (id && !notes.some((note) => note.id === id))
          throw new Error('This note is no longer available.');
        checkpoint();
        staffNotes.set(
          teacher,
          id
            ? notes.map((note) => (note.id === id ? { ...note, text: value } : note))
            : [...notes, { id: `note-${++sequence}`, text: value }],
        );
      },
      removeStaffNote(teacher, id) {
        const notes = staffNotes.get(teacher) || [];
        if (!notes.some((note) => note.id === id))
          throw new Error('This note is no longer available.');
        checkpoint();
        staffNotes.set(
          teacher,
          notes.filter((note) => note.id !== id),
        );
      },
      tables,
      addTable(teacher, value) {
        if (!staff.has(teacher)) throw new Error('Choose an existing instructor.');
        const table = tableValue(value, teacher);
        if (!table) throw new Error('Enter a table number.');
        if (tables.get(teacher).has(table))
          throw new Error('This instructor already has that table.');
        checkpoint();
        tables.get(teacher).add(table);
      },
      renameTable(teacher, previous, value) {
        if (!tables.get(teacher)?.has(previous))
          throw new Error('This table is no longer available.');
        const table = tableValue(value, teacher);
        if (!table) throw new Error('Enter a table number.');
        if (table === previous) return;
        if (tables.get(teacher).has(table))
          throw new Error('This instructor already has that table.');
        checkpoint();
        tables.set(
          teacher,
          new Set([...tables.get(teacher)].map((number) => (number === previous ? table : number))),
        );
        const affected = new Set(
          rows
            .filter((row) => row.staff === teacher && row.table === previous)
            .map((row) => row._bookingId),
        );
        rows = rows.map((row) =>
          row.staff === teacher && row.table === previous
            ? {
                ...row,
                table,
                manualRationale: [
                  ...(row.manualRationale || []),
                  {
                    rule: 'Manual table change',
                    detail: `The user renamed Table ${previous} to Table ${table} for ${teacher}. This portion keeps its existing ${row.start} - ${row.end} time.`,
                  },
                ],
              }
            : row,
        );
        for (const booking of affected) {
          const segments = rows
            .filter((row) => row._bookingId === booking)
            .sort((a, b) => a.start.localeCompare(b.start));
          for (let index = 0; index < segments.length; index++) {
            const row = segments[index],
              before = segments[index - 1],
              after = segments[index + 1];
            row.notes = (row.notes || []).filter(
              (note) =>
                !/^(?:\(switch to table |Table change to |Continuing from Table )/i.test(note),
            );
            if (before?.staff === row.staff && before.table !== row.table)
              row.notes.push(
                `Continuing from Table ${before.table || 'unassigned'} at ${row.start}`,
              );
            if (after?.staff === row.staff && after.table !== row.table)
              row.notes.push(`(switch to table ${after.table || 'unassigned'})`);
          }
        }
      },
      roster,
      get,
      parts,
      draft,
      save: saveDraft,
      get canUndo() {
        return history.length > 0;
      },
      undo() {
        if (history.length) {
          const snapshot = history.pop();
          rows = snapshot.rows;
          tables.clear();
          for (const [name, values] of snapshot.tables) tables.set(name, new Set(values));
          staffNotes.clear();
          for (const [name, notes] of snapshot.staffNotes) staffNotes.set(name, notes);
        }
        return rows;
      },
      remove(id) {
        const booking = get(id)._bookingId;
        checkpoint();
        rows = rows.filter((row) => row._bookingId !== booking);
        return rows;
      },
      move(id, target, targetTable) {
        if (!staff.has(target)) throw new Error('Choose an existing instructor.');
        const selected = get(id),
          segments = parts(id),
          value = draft(id);
        const changed = segments.map((row) => ({
          ...row,
          staff: row === selected ? target : row.staff,
          table: row === selected ? tableValue(targetTable, target) : row.table,
        }));
        const merged = [];
        for (const row of changed) {
          if (merged.at(-1)?.staff === row.staff && merged.at(-1)?.table === row.table)
            merged.at(-1).end = row.end;
          else merged.push({ ...row });
        }
        value.staff = merged[0].staff;
        value.table = merged[0].table || '';
        value.manualAction = 'Moved by drag and drop';
        value.handoffs = merged
          .slice(1)
          .map((row) => ({ staff: row.staff, table: row.table || '', time: row.start }));
        return saveDraft(value, id);
      },
    };
  }
  return { create };
})();

const ScheduleEditor = (() => {
  let model = null,
    render = null,
    dragging = null,
    dialog = null;
  const element = (tag, text, className) => {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  };
  const button = (label, action, className = 'no-print') => {
    const node = element('button', label, className);
    node.type = 'button';
    node.addEventListener('click', action);
    return node;
  };
  function refresh(message) {
    render(model.rows);
    document.getElementById('editStatus').textContent = message;
    document.getElementById('resultSummary').textContent =
      `${model.staff.size} staff/sections, ${model.rows.length} appointment portions. Manual edits are included in print and CSV.`;
  }
  function open(staff, id) {
    dialog?.remove();
    dialog = element('dialog', undefined, 'schedule-dialog no-print');
    const form = element('form');
    form.noValidate = true;
    const value = id ? model.draft(id) : { student: '', staff, start: '', end: '', handoffs: [] };
    const heading = element('h3', id ? 'Edit student appointment' : `Add student to ${staff}`);
    heading.id = 'appointmentEditorTitle';
    dialog.setAttribute('aria-labelledby', heading.id);
    const hint = element('p', 'Changes apply to every instructor involved in this appointment.');
    const error = element('p', '', 'notice');
    error.setAttribute('role', 'alert');
    function input(label, type, initial = '') {
      const wrapper = element('label', label);
      const field = element('input');
      field.type = type;
      field.name = label.toLowerCase().replace(/\s+/g, '');
      field.value = initial;
      if (type === 'text') field.maxLength = 150;
      wrapper.append(field);
      form.append(wrapper);
      return field;
    }
    function instructors(selected) {
      const select = element('select');
      for (const name of model.staff.keys()) {
        const option = element('option', name);
        option.value = name;
        select.append(option);
      }
      select.value = selected;
      return select;
    }
    form.append(heading, hint);
    const name = input('Student name', 'text', value.student);
    name.required = true;
    const suggestions = element('datalist');
    suggestions.id = 'studentSuggestions';
    for (const student of model.roster.values()) {
      const option = element('option');
      option.value = student.name;
      suggestions.append(option);
    }
    name.setAttribute('list', suggestions.id);
    form.append(suggestions);
    const staffLabel = element('label', 'Starting instructor');
    const teacher = instructors(value.staff);
    teacher.name = 'staff';
    staffLabel.append(teacher);
    form.append(staffLabel);
    const table = input(
      'Table number',
      'number',
      value.table ?? [...model.tables.get(value.staff)][0] ?? '',
    );
    table.name = 'table';
    table.min = '1';
    table.max = '999';
    table.step = '1';
    teacher.addEventListener('change', () => {
      table.value = [...model.tables.get(teacher.value)][0] || '';
    });
    const start = input('Start time', 'time', value.start);
    const end = input('End time', 'time', value.end);
    start.addEventListener('change', () => {
      if (!end.value && start.value) {
        const total = SpreadsheetParser.minutes(start.value) + 60;
        if (total < 1440)
          end.value = `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
      }
    });
    const online = input('Online', 'checkbox');
    online.checked = Boolean(value.online);
    const attention = input('Needs attention', 'checkbox');
    attention.checked = Boolean(value.attention);
    const labelField = element('label', 'Additional labels');
    const labels = element('textarea');
    labels.name = 'additionallabels';
    labels.rows = 3;
    labels.maxLength = 2000;
    labels.value = value.labels || '';
    labels.placeholder = 'Physics; New binder';
    labelField.append(labels);
    form.append(
      labelField,
      element('p', 'Separate labels with a semicolon or a new line.', 'color-hint'),
    );
    const transfers = element('fieldset');
    transfers.append(element('legend', 'Handoffs / table changes'));
    const entries = [];
    function addTransfer(transfer = {}) {
      const row = element('div', undefined, 'handoff-fields');
      const label = element('label', 'Transfer to');
      const select = instructors(
        transfer.staff || [...model.staff.keys()].find((name) => name !== teacher.value),
      );
      select.name = 'handoffStaff';
      label.append(select);
      const timeLabel = element('label', 'At');
      const time = element('input');
      time.type = 'time';
      time.name = 'handoffTime';
      time.value = transfer.time || '';
      timeLabel.append(time);
      const tableLabel = element('label', 'Table number');
      const transferTable = element('input');
      transferTable.type = 'number';
      transferTable.name = 'handoffTable';
      transferTable.min = '1';
      transferTable.max = '999';
      transferTable.step = '1';
      transferTable.value = transfer.table ?? [...(model.tables.get(select.value) || [])][0] ?? '';
      select.addEventListener('change', () => {
        transferTable.value = [...model.tables.get(select.value)][0] || '';
      });
      tableLabel.append(transferTable);
      const entry = { row, select, time, table: transferTable };
      entries.push(entry);
      row.append(
        label,
        timeLabel,
        tableLabel,
        button('Remove handoff', () => {
          entries.splice(entries.indexOf(entry), 1);
          row.remove();
        }),
      );
      transfers.append(row);
    }
    value.handoffs.forEach(addTransfer);
    form.append(
      transfers,
      button('Add handoff', () => addTransfer()),
      error,
    );
    const actions = element('div', undefined, 'actions');
    const save = element('button', 'Save changes', 'primary');
    save.type = 'submit';
    actions.append(
      save,
      button('Cancel', () => dialog.close()),
    );
    form.append(actions);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      try {
        model.save(
          {
            student: name.value,
            staff: teacher.value,
            table: table.value,
            start: start.value,
            end: end.value,
            online: online.checked,
            attention: attention.checked,
            labels: labels.value,
            handoffs: entries.map(({ select, time, table }) => ({
              staff: select.value,
              table: table.value,
              time: time.value,
            })),
          },
          id,
        );
        dialog.close();
        refresh('Appointment saved.');
      } catch (problem) {
        error.textContent = problem.message;
      }
    });
    dialog.append(form);
    document.body.append(dialog);
    const currentDialog = dialog;
    dialog.addEventListener('close', () => currentDialog.remove(), { once: true });
    dialog.showModal();
    name.focus();
  }
  function addTable(staff, previous) {
    dialog?.remove();
    const current = element('dialog', undefined, 'schedule-dialog no-print');
    dialog = current;
    const form = element('form');
    form.noValidate = true;
    const heading = element('h3', `${previous ? 'Edit' : 'Add'} table for ${staff}`);
    heading.id = 'tableEditorTitle';
    current.setAttribute('aria-labelledby', heading.id);
    const label = element('label', 'Table number');
    const number = element('input');
    number.type = 'number';
    number.name = 'tablenumber';
    number.min = '1';
    number.max = '999';
    number.step = '1';
    number.value = previous || '';
    label.append(number);
    const error = element('p', '', 'notice');
    error.setAttribute('role', 'alert');
    const save = element('button', previous ? 'Save changes' : 'Add table', 'primary');
    save.type = 'submit';
    const actions = element('div', undefined, 'actions');
    actions.append(
      save,
      button('Cancel', () => current.close()),
    );
    form.append(heading, label, error, actions);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      try {
        if (previous) model.renameTable(staff, previous, number.value);
        else model.addTable(staff, number.value);
        current.close();
        refresh(
          previous
            ? `Table ${previous} changed to ${number.value} for ${staff}.`
            : `Table ${number.value} added for ${staff}.`,
        );
      } catch (problem) {
        error.textContent = problem.message;
      }
    });
    current.append(form);
    document.body.append(current);
    current.addEventListener('close', () => current.remove(), { once: true });
    current.showModal();
    number.focus();
  }
  function explain(row) {
    dialog?.remove();
    const current = element('dialog', undefined, 'schedule-dialog rationale-dialog no-print');
    dialog = current;
    const heading = element('h3', 'Why this assignment?');
    heading.id = 'assignmentExplanationTitle';
    current.setAttribute('aria-labelledby', heading.id);
    const format = (value) => displayTime(value);
    current.append(
      heading,
      element(
        'p',
        `${row.student}: ${format(row.start)} - ${format(row.end)} · ${row.staff} · ${row.table ? 'Table ' + row.table : 'No table assigned'}${row.binder ? ' · Binder ' + row.binder : ' · No binder location'}`,
      ),
    );
    function rules(title, items) {
      current.append(element('h4', title));
      const list = element('ul');
      const seen = new Set();
      for (const item of items) {
        const key = JSON.stringify(item);
        if (seen.has(key)) continue;
        seen.add(key);
        const entry = element('li');
        entry.append(
          element('strong', item.rule + ': '),
          document.createTextNode(displayNote(item.detail)),
        );
        list.append(entry);
      }
      current.append(list);
    }
    if (row.manualRationale?.length) rules('Manual edit history', row.manualRationale);
    if (row.rationale?.length)
      rules(
        row.manualRationale?.length
          ? 'Original import evidence (before manual edits)'
          : 'Rules and spreadsheet evidence',
        row.rationale,
      );
    else
      current.append(
        element('p', 'No spreadsheet rule evidence was recorded for this appointment.'),
      );
    if (!row.manualRationale?.length && !row.rationale?.some((item) => /handoff/i.test(item.rule)))
      current.append(element('p', 'No takeover was applied to this appointment portion.'));
    const actions = element('div', undefined, 'actions');
    actions.append(
      button('Edit appointment', () => open(row.staff, row._editId)),
      button('Close', () => current.close()),
    );
    current.append(actions);
    document.body.append(current);
    current.addEventListener('close', () => current.remove(), { once: true });
    current.showModal();
    actions.lastChild.focus();
  }
  function openStaffNote(staff, note) {
    dialog?.remove();
    const current = element('dialog', undefined, 'schedule-dialog no-print');
    dialog = current;
    const form = element('form');
    form.noValidate = true;
    const heading = element('h3', `${note ? 'Edit' : 'Add'} note for ${staff}`);
    heading.id = 'staffNoteTitle';
    current.setAttribute('aria-labelledby', heading.id);
    const label = element('label', 'Note');
    const text = element('textarea');
    text.name = 'staffnote';
    text.rows = 4;
    text.maxLength = 2000;
    text.value = note?.text || '';
    label.append(text);
    const error = element('p', '', 'notice');
    error.setAttribute('role', 'alert');
    const save = element('button', 'Save note', 'primary');
    save.type = 'submit';
    const actions = element('div', undefined, 'actions');
    actions.append(
      save,
      button('Cancel', () => current.close()),
    );
    form.append(heading, label, error, actions);
    form.addEventListener('submit', (event) => {
      event.preventDefault();
      try {
        model.saveStaffNote(staff, text.value, note?.id);
        current.close();
        refresh(`Note saved for ${staff}.`);
      } catch (problem) {
        error.textContent = problem.message;
      }
    });
    current.append(form);
    document.body.append(current);
    current.addEventListener('close', () => current.remove(), { once: true });
    current.showModal();
    text.focus();
  }
  function decorate(rows, renderer) {
    render = renderer;
    if (!model) return;
    document.getElementById('editToolbar')?.remove();
    const toolbar = element('div', undefined, 'no-print');
    toolbar.id = 'editToolbar';
    toolbar.append(
      element(
        'p',
        'Drag students between instructors, or use Edit to move them. Click a transfer note to adjust its time. Edits last until you generate again or close this page.',
      ),
    );
    const undo = button('Undo last edit', () => {
      model.undo();
      refresh('Last edit undone.');
    });
    undo.disabled = !model.canUndo;
    const status = element('p', '', 'notice');
    status.id = 'editStatus';
    status.setAttribute('role', 'status');
    toolbar.append(undo, status);
    const lists = document.getElementById('staffLists');
    lists.before(toolbar);
    // Keep empty instructors as drop targets and available in the form.
    for (const name of model.staff.keys()) {
      if (
        ![...lists.children].some(
          (block) =>
            block.querySelector('h3')?.textContent === name ||
            block.querySelector('h3')?.textContent.startsWith(name + ' (Table '),
        )
      ) {
        const block = element('section', undefined, 'staff-block empty-staff');
        block.append(element('h3', name));
        block.dataset.staff = name;
        for (const table of model.tables.get(name)) {
          const group = element('div', undefined, 'table-group');
          group.dataset.table = table;
          group.append(element('h4', `Table ${table}`), element('ul', undefined, 'staff-list'));
          block.append(group);
        }
        lists.append(block);
      }
    }
    for (const block of lists.children) {
      const heading = block.querySelector('h3');
      const staff = [...model.staff.keys()].find(
        (name) => heading.textContent === name || heading.textContent.startsWith(name + ' (Table '),
      );
      if (!staff) continue;
      const header = element('div', undefined, 'staff-heading');
      heading.before(header);
      header.append(
        heading,
        button('Add student', () => open(staff)),
        button('Add table', () => addTable(staff)),
      );
      for (const group of block.querySelectorAll('.table-group')) {
        if (!group.dataset.table) continue;
        const title = group.querySelector('h4');
        const tableHeader = element('div', undefined, 'table-heading');
        title.before(tableHeader);
        tableHeader.append(
          title,
          button('Edit table', () => addTable(staff, group.dataset.table)),
        );
      }
      const notes = model.staffNotes.get(staff) || [];
      const noteArea = element('div', undefined, 'instructor-notes');
      if (notes.length) {
        block.classList.remove('empty-staff');
        noteArea.append(element('h4', 'Notes'));
        for (const note of notes) {
          const entry = element('div', undefined, 'instructor-note');
          const controls = element('div', undefined, 'actions no-print');
          controls.append(
            button('Edit note', () => openStaffNote(staff, note)),
            button('Remove note', () => {
              model.removeStaffNote(staff, note.id);
              refresh(`Note removed from ${staff}'s schedule.`);
            }),
          );
          entry.append(element('p', note.text), controls);
          noteArea.append(entry);
        }
      }
      noteArea.append(button('Add a note', () => openStaffNote(staff)));
      block.append(noteArea);
      block.addEventListener('dragover', (event) => {
        if (dragging) {
          event.preventDefault();
          block.classList.add('drop-target');
        }
      });
      block.addEventListener('dragleave', (event) => {
        if (!block.contains(event.relatedTarget)) block.classList.remove('drop-target');
      });
      block.addEventListener('drop', (event) => {
        event.preventDefault();
        block.classList.remove('drop-target');
        if (!dragging) return;
        const id = dragging;
        dragging = null;
        const group = event.target.closest('.table-group');
        const table = group && block.contains(group) ? group.dataset.table : undefined;
        if (
          model.get(id).staff !== staff ||
          (table !== undefined && model.get(id).table !== table)
        ) {
          try {
            model.move(id, staff, table);
            refresh(`Student moved to ${staff}.`);
          } catch (problem) {
            document.getElementById('editStatus').textContent = problem.message;
          }
        }
      });
    }
    for (const node of lists.querySelectorAll('[data-edit-id]')) {
      const row = model.get(node.dataset.editId);
      node.draggable = true;
      node.addEventListener('dragstart', (event) => {
        dragging = row._editId;
        event.dataTransfer.setData('text/plain', row._editId);
        event.dataTransfer.effectAllowed = 'move';
      });
      node.addEventListener('dragend', () => {
        dragging = null;
        lists
          .querySelectorAll('.drop-target')
          .forEach((target) => target.classList.remove('drop-target'));
      });
      const controls = element('span', undefined, 'student-controls no-print');
      controls.append(button('Edit', () => open(row.staff, row._editId)));
      const remove = button('×', () => {
        model.remove(row._editId);
        refresh(
          `${row.student}'s appointment removed from all involved instructors. Use Undo to restore it.`,
        );
      });
      remove.title = `Remove ${row.student}'s entire appointment, including handoffs`;
      remove.setAttribute('aria-label', remove.title);
      controls.append(remove);
      node.append(controls);
    }
  }
  return {
    explain,
    load(rows, roster) {
      model = ScheduleEdits.create(rows, roster);
    },
    tablesFor(staff) {
      return model ? [...(model.tables.get(staff) || [])] : [];
    },
    staffNoteRows() {
      return model
        ? [...model.staffNotes].flatMap(([staff, notes]) =>
            notes.map((note) => [staff, '', '', '', '', note.text, '', '', '']),
          )
        : [];
    },
    clear() {
      model = null;
      dragging = null;
      dialog?.close();
      dialog?.remove();
      dialog = null;
    },
    decorate,
    note(row, text) {
      if (
        !/^(?:Handoff to |Taking over from |Table change to |\(switch to table |Continuing from Table |Transfer.*\(May be incorrect input\))/i.test(
          text,
        )
      )
        return document.createTextNode(text);
      const control = button(text, () => open(row.staff, row._editId), 'transfer-time');
      control.title = 'Edit appointment and handoff times';
      return control;
    },
  };
})();
