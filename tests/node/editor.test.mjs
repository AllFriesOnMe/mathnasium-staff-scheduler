import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { readFileSync } from 'node:fs';

function fixture() {
  const context = vm.createContext({});
  for (const file of ['src/spreadsheet.js', 'src/schedule-editor.js'])
    vm.runInContext(readFileSync(new URL(`../../${file}`, import.meta.url), 'utf8'), context);
  return vm.runInContext(
    `ScheduleEdits.create([
    {student:'Jane Doe',staff:'Alex',table:'1',start:'10:00',end:'10:30',appointmentStart:'10:00',appointmentEnd:'11:00',binder:'A1',online:true,notes:['Needs attention','Physics','Handoff to Jordan at 10:30']},
    {student:'Jane Doe',staff:'Jordan',table:'2',start:'10:30',end:'11:00',appointmentStart:'10:00',appointmentEnd:'11:00',binder:'A1',online:true,notes:['Needs attention','Physics','Taking over from Alex at 10:30']},
    {student:'Other Student',staff:'Morgan',table:'3',start:'12:00',end:'13:00',notes:[]}
  ], new Map([['jane doe',{name:'Jane Doe',location:'A1'}],['new student',{name:'New Student',location:'B2'}]]))`,
    context,
  );
}
const plain = (value) => JSON.parse(JSON.stringify(value));

test('Manual rationale is separate from original import evidence and follows undo', () => {
  const model = fixture(),
    row = model.rows[0],
    id = row._editId;
  row.rationale = [{ rule: 'Initial instructor', detail: 'Original spreadsheet rule' }];
  model.save({ ...model.draft(id), table: '7' }, id);
  const edited = model.rows.find((row) => row.student === 'Jane Doe');
  assert.ok(
    edited.manualRationale.some(
      (item) => item.detail.includes('user') && item.detail.includes('Table 7'),
    ),
  );
  assert.equal(edited.rationale[0].detail, 'Original spreadsheet rule');
  model.undo();
  assert.equal(model.rows[0].manualRationale, undefined);
});

test('Instructor notes can be added, edited, removed and undone independently of students', () => {
  const model = fixture(),
    before = plain(model.rows);
  model.saveStaffNote('Alex', 'Check binders\nReturn calculators');
  const note = model.staffNotes.get('Alex')[0];
  assert.deepEqual(plain(model.rows), before);
  assert.equal(model.staffNotes.has('Jordan'), false);
  model.saveStaffNote('Alex', 'Updated note', note.id);
  assert.equal(model.staffNotes.get('Alex')[0].text, 'Updated note');
  model.removeStaffNote('Alex', note.id);
  assert.equal(model.staffNotes.get('Alex').length, 0);
  model.undo();
  assert.equal(model.staffNotes.get('Alex')[0].text, 'Updated note');
  model.undo();
  assert.equal(model.staffNotes.get('Alex')[0].text, 'Check binders\nReturn calculators');
  model.undo();
  assert.equal(model.staffNotes.has('Alex'), false);
});
test('Student edits preserve instructor notes and include them in undo snapshots', () => {
  const model = fixture();
  model.saveStaffNote('Alex', 'Keep this note');
  model.remove(model.rows[0]._editId);
  assert.equal(model.staffNotes.get('Alex')[0].text, 'Keep this note');
  model.undo();
  assert.equal(model.staffNotes.get('Alex')[0].text, 'Keep this note');
});
test('Invalid instructor notes do not create history or change the schedule', () => {
  const model = fixture();
  for (const text of ['', '   ', 'x'.repeat(2001)])
    assert.throws(() => model.saveStaffNote('Alex', text));
  assert.throws(() => model.saveStaffNote('Unknown', 'A note'));
  assert.throws(() => model.saveStaffNote('Alex', 'A note', 'missing'));
  assert.equal(model.staffNotes.size, 0);
  assert.equal(model.canUndo, false);
});

test('Moving a handoff boundary updates both sides and preserves flags, binder and total duration', () => {
  const model = fixture(),
    id = model.rows[0]._editId;
  const draft = model.draft(id);
  draft.handoffs[0].time = '10:17';
  model.save(draft, id);
  const jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.deepEqual(plain(jane.map((row) => [row.start, row.end])), [
    ['10:00', '10:17'],
    ['10:17', '11:00'],
  ]);
  assert.equal(
    jane.reduce((total, row) => total + row.durationMinutes, 0),
    60,
  );
  assert.ok(
    jane.every(
      (row) =>
        row.binder === 'A1' &&
        row.online &&
        row.notes.includes('Physics') &&
        row.notes.includes('Needs attention'),
    ),
  );
  assert.ok(jane[0].notes.includes('Handoff to Jordan at 10:17'));
  assert.ok(jane[1].notes.includes('Taking over from Alex at 10:17'));
  assert.ok(!jane.some((row) => row.notes.some((note) => note.includes('10:30'))));
});
test('Drag move merges adjacent portions for the same instructor and updates table', () => {
  const model = fixture();
  model.move(model.rows[0]._editId, 'Jordan');
  const jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.equal(jane.length, 1);
  assert.equal(jane[0].table, '2');
  assert.equal(jane[0].durationMinutes, 60);
  assert.equal(jane[0].end, '11:00');
  assert.ok(!jane[0].notes.some((note) => /Handoff|Taking over/.test(note)));
});
test('Drag to a third instructor rewrites both handoff descriptions', () => {
  const model = fixture();
  model.move(model.rows[0]._editId, 'Morgan');
  const jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.equal(jane[0].staff, 'Morgan');
  assert.equal(jane[0].table, '3');
  assert.ok(jane[1].notes.includes('Taking over from Morgan at 10:30'));
});
test('Remove deletes the whole appointment, undo restores all its portions', () => {
  const model = fixture();
  const before = plain(model.rows);
  model.remove(model.rows[1]._editId);
  assert.equal(model.rows.length, 1);
  assert.equal(model.staff.size, 3);
  assert.ok(model.canUndo);
  model.undo();
  assert.deepEqual(plain(model.rows), before);
});
test('Add resolves binder names, supports multiple handoffs, and retains two-hour duration', () => {
  const model = fixture();
  model.save({
    student: 'New Student',
    staff: 'Alex',
    start: '14:00',
    end: '16:00',
    online: false,
    attention: true,
    handoffs: [
      { staff: 'Jordan', time: '14:23' },
      { staff: 'Morgan', time: '15:41' },
    ],
  });
  const added = model.rows.filter((row) => row.student === 'New Student');
  assert.equal(added.length, 3);
  assert.equal(
    added.reduce((n, row) => n + row.durationMinutes, 0),
    120,
  );
  assert.ok(
    added.every(
      (row) =>
        row.binder === 'B2' &&
        row.bookedDurationMinutes === 120 &&
        row.notes.includes('Needs attention'),
    ),
  );
});
for (const change of [
  { student: '' },
  { student: 'x'.repeat(151) },
  { start: '' },
  { end: '09:00' },
  { end: '10:00' },
  { start: '23:00', end: '01:00' },
  { staff: 'Unknown' },
  { handoffs: [{ staff: 'Jordan', time: '10:00' }] },
  { handoffs: [{ staff: 'Jordan', time: '11:00' }] },
  { handoffs: [{ staff: 'Alex', time: '10:20' }] },
  { handoffs: [{ staff: 'Unknown', time: '10:20' }] },
  {
    handoffs: [
      { staff: 'Jordan', time: '10:40' },
      { staff: 'Morgan', time: '10:20' },
    ],
  },
])
  test('Invalid edits are atomic: ' + JSON.stringify(change), () => {
    const model = fixture(),
      before = plain(model.rows),
      id = model.rows[0]._editId;
    assert.throws(() => model.save({ ...model.draft(id), ...change }, id));
    assert.deepEqual(plain(model.rows), before);
    assert.equal(model.canUndo, false);
  });
test('Renaming never retains another student binder; unknown students may be added', () => {
  const model = fixture(),
    id = model.rows[0]._editId;
  model.save({ ...model.draft(id), student: 'Unknown Student' }, id);
  assert.ok(
    model.rows.filter((row) => row.student === 'Unknown Student').every((row) => row.binder === ''),
  );
});

test('Adding an overlapping booking for the same student is rejected atomically', () => {
  const model = fixture(),
    before = plain(model.rows);
  assert.throws(
    () => model.save({ student: 'Jane Doe', staff: 'Morgan', start: '10:15', end: '11:15' }),
    /already has an appointment/,
  );
  assert.deepEqual(plain(model.rows), before);
});
test('An unresolved transfer note pre-fills the destination and written time for correction', () => {
  const model = fixture();
  const id = model.rows[0]._editId;
  model.save({ ...model.draft(id), handoffs: [] }, id);
  const jane = model.rows.find((row) => row.student === 'Jane Doe');
  jane.notes.push('Transfer to Jordan at 4:15 AM (May be incorrect input)');
  const draft = model.draft(jane._editId);
  assert.equal(draft.handoffs[0].staff, 'Jordan');
  assert.equal(draft.handoffs[0].time, '04:15');
  draft.handoffs[0].time = '10:15';
  model.save(draft, jane._editId);
  assert.ok(
    model.rows
      .filter((row) => row.student === 'Jane Doe')
      .every((row) => !row.notes.some((note) => note.includes('May be incorrect input'))),
  );
});

test('Additional labels can be edited, deduplicated and cleared across handoff portions', () => {
  const model = fixture();
  let id = model.rows[0]._editId;
  const draft = model.draft(id);
  assert.equal(draft.labels, 'Physics');
  model.save({ ...draft, labels: 'New binder; Homework\nNew binder\n  ' }, id);
  let jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.ok(
    jane.every(
      (row) =>
        row.notes.includes('New binder') &&
        row.notes.includes('Homework') &&
        !row.notes.includes('Physics'),
    ),
  );
  assert.ok(jane.every((row) => row.notes.filter((note) => note === 'New binder').length === 1));
  assert.ok(jane[0].notes.includes('Handoff to Jordan at 10:30'));
  assert.ok(jane.every((row) => row.notes.includes('Needs attention')));
  id = jane[0]._editId;
  model.save({ ...model.draft(id), labels: '' }, id);
  jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.ok(
    jane.every((row) => !row.notes.includes('New binder') && !row.notes.includes('Homework')),
  );
  assert.ok(jane[0].notes.includes('Handoff to Jordan at 10:30'));
});
test('Excessive additional labels are rejected without changing the schedule', () => {
  const model = fixture(),
    id = model.rows[0]._editId,
    before = plain(model.rows);
  assert.throws(() => model.save({ ...model.draft(id), labels: 'x'.repeat(2001) }, id), /2,000/);
  assert.deepEqual(plain(model.rows), before);
});

test('A student table can be edited without changing another handoff instructor table', () => {
  const model = fixture(),
    id = model.rows[0]._editId;
  const draft = model.draft(id);
  assert.equal(draft.table, '1');
  assert.equal(draft.handoffs[0].table, '2');
  model.save({ ...draft, table: '7' }, id);
  const jane = model.rows.filter((r) => r.student === 'Jane Doe');
  assert.equal(jane[0].table, '7');
  assert.equal(jane[1].table, '2');
  assert.ok(model.tables.get('Alex').has('7'));
  model.undo();
  assert.equal(model.rows[0].table, '1');
  assert.equal(model.tables.get('Alex').has('7'), false);
});
test('Same instructor table changes preserve booking duration and separate portions', () => {
  const model = fixture(),
    id = model.rows[0]._editId;
  model.save({ ...model.draft(id), handoffs: [{ staff: 'Alex', table: '4', time: '10:20' }] }, id);
  const jane = model.rows.filter((r) => r.student === 'Jane Doe');
  assert.equal(jane.length, 2);
  assert.equal(jane[0].table, '1');
  assert.equal(jane[1].table, '4');
  assert.ok(jane[0].notes.includes('(switch to table 4)'));
  assert.ok(jane[1].notes.includes('Continuing from Table 1 at 10:20'));
  assert.equal(
    jane.reduce((n, r) => n + r.durationMinutes, 0),
    60,
  );
  model.save(model.draft(jane[1]._editId), jane[1]._editId);
  assert.equal(model.rows.filter((r) => r.student === 'Jane Doe').length, 2);
});
test('Dragging to a specific table preserves separate tables for the same instructor', () => {
  const model = fixture();
  model.move(model.rows[0]._editId, 'Jordan', '5');
  const jane = model.rows.filter((r) => r.student === 'Jane Doe');
  assert.equal(jane.length, 2);
  assert.equal(jane[0].table, '5');
  assert.equal(jane[1].table, '2');
  assert.equal(jane[0].staff, 'Jordan');
  assert.equal(jane[1].staff, 'Jordan');
});
test('Add table creates an empty instructor table without changing any student and supports undo', () => {
  const model = fixture(),
    before = plain(model.rows);
  model.addTable('Alex', '8');
  assert.ok(model.tables.get('Alex').has('8'));
  assert.deepEqual(plain(model.rows), before);
  assert.throws(() => model.addTable('Alex', '8'), /already has/);
  model.undo();
  assert.equal(model.tables.get('Alex').has('8'), false);
  assert.deepEqual(plain(model.rows), before);
});

test('Rename table updates only its instructor, preserves assignments and supports undo', () => {
  const model = fixture(),
    before = plain(model.rows);
  model.renameTable('Alex', '1', '6');
  assert.equal(model.rows[0].table, '6');
  assert.equal(model.rows[1].table, '2');
  assert.ok(model.tables.get('Alex').has('6'));
  assert.ok(!model.tables.get('Alex').has('1'));
  model.save({ student: 'New Student', staff: 'Alex', start: '14:00', end: '15:00' });
  assert.equal(model.rows.find((row) => row.student === 'New Student').table, '6');
  model.undo();
  model.undo();
  assert.deepEqual(plain(model.rows), before);
  assert.ok(model.tables.get('Alex').has('1'));
});
test('Rename table rewrites switch and continuation notes on both portions', () => {
  const model = fixture(),
    id = model.rows[0]._editId;
  model.save({ ...model.draft(id), handoffs: [{ staff: 'Alex', table: '4', time: '10:20' }] }, id);
  model.renameTable('Alex', '4', '7');
  let jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.ok(jane[0].notes.includes('(switch to table 7)'));
  assert.equal(jane[1].table, '7');
  model.renameTable('Alex', '1', '5');
  jane = model.rows.filter((row) => row.student === 'Jane Doe');
  assert.ok(jane[1].notes.includes('Continuing from Table 5 at 10:20'));
  assert.equal(
    jane.reduce((n, row) => n + row.durationMinutes, 0),
    60,
  );
});
test('Duplicate or invalid table renames do not change rows or create undo entries', () => {
  const model = fixture(),
    before = plain(model.rows);
  for (const number of ['0', '-1', '1.5', '1000', ''])
    assert.throws(() => model.renameTable('Alex', '1', number));
  assert.deepEqual(plain(model.rows), before);
  assert.equal(model.canUndo, false);
  model.addTable('Alex', '8');
  assert.throws(() => model.renameTable('Alex', '1', '8'), /already has/);
  model.undo();
  assert.ok(!model.tables.get('Alex').has('8'));
});
for (const table of ['0', '-1', '1.5', '1000', '2, 3', 'abc'])
  test('Invalid student table is rejected atomically: ' + table, () => {
    const model = fixture(),
      before = plain(model.rows),
      id = model.rows[0]._editId;
    assert.throws(() => model.save({ ...model.draft(id), table }, id), /Table number/);
    assert.deepEqual(plain(model.rows), before);
  });
