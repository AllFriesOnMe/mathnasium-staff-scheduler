'use strict';
const passed = [],
  failures = [];
function test(name, run) {
  try {
    run();
    passed.push(name);
  } catch (error) {
    failures.push(name + ': ' + error.message);
  }
}
function assert(condition, message = 'Assertion failed') {
  if (!condition) throw new Error(message);
}
function rejects(fn) {
  let failed = false;
  try {
    fn();
  } catch {
    failed = true;
  }
  assert(failed);
}
const sheet = (rows) => XLSX.utils.aoa_to_sheet(rows);
const P = SpreadsheetParser;
const roster = P.readBinders(
  sheet([
    ['Center binder list'],
    [],
    ['Student', 'Binder'],
    ['Samara Adams', 'A1'],
    ['Ryan Clark', 'B2'],
    ['Veer Shah', 'E5'],
  ]),
);
