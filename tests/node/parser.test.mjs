import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
import assert from 'node:assert/strict';

test('Spreadsheet student and handoff regression matrices', (t) => {
  const context = vm.createContext({});
  for (const file of [
    'vendor/xlsx.full.min.js',
    'src/takeovers.js',
    'src/color-checks.js',
    'src/spreadsheet.js',
    'tests/helpers.js',
    'tests/parser.test.js',
    'tests/takeovers.test.js',
  ]) {
    vm.runInContext(readFileSync(new URL(`../../${file}`, import.meta.url), 'utf8'), context, {
      filename: file,
      timeout: 30000,
    });
  }
  const report = JSON.parse(
    vm.runInContext('JSON.stringify({passed: passed.length, failures})', context),
  );
  t.diagnostic(`${report.passed} student and handoff cases passed`);
  assert.equal(report.failures.length, 0, report.failures.join('\n'));
});
