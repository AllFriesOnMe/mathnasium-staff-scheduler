'use strict';
(async () => {
  const passed = [],
    failures = [];
  const assert = (ok, message = 'Assertion failed') => {
    if (!ok) throw new Error(message);
  };
  const check = async (name, run) => {
    try {
      await run();
      passed.push(name);
    } catch (error) {
      failures.push(`${name}: ${error.message}`);
    }
  };
  async function rejected(task, pattern) {
    let error;
    try {
      await task();
    } catch (caught) {
      error = caught;
    }
    assert(error && pattern.test(error.message), error?.message || 'Expected rejection');
  }
  const csv = (contents, name = 'students.csv') => new File([contents], name);
  const book = (...sheets) => {
    const b = XLSX.utils.book_new();
    sheets.forEach((s, i) =>
      XLSX.utils.book_append_sheet(b, XLSX.utils.aoa_to_sheet(s), `Sheet${i + 1}`),
    );
    return b;
  };
  const workbookFile = (b) =>
    new File([XLSX.write(b, { type: 'array', bookType: 'xlsx' })], 'test.xlsx');
  await check('Oversized file rejected before worker creation', async () => {
    const OriginalWorker = globalThis.Worker;
    globalThis.Worker = class {
      constructor() {
        throw new Error('Should not create worker');
      }
    };
    try {
      await rejected(
        () =>
          SpreadsheetJobs.run({ type: 'read', file: { size: SpreadsheetJobs.MAX_FILE_BYTES + 1 } }),
        /5 MB/,
      );
    } finally {
      globalThis.Worker = OriginalWorker;
    }
  });
  await check('Worker timeout terminates stalled work', async () => {
    const OriginalWorker = globalThis.Worker;
    let terminated = false;
    globalThis.Worker = class {
      postMessage() {}
      terminate() {
        terminated = true;
      }
    };
    try {
      await rejected(
        () => SpreadsheetJobs.run({ type: 'read', file: { size: 1 } }, { timeoutMs: 20 }),
        /stopped/,
      );
      assert(terminated);
    } finally {
      globalThis.Worker = OriginalWorker;
    }
  });
  await check('Cancellation terminates outdated work', async () => {
    const OriginalWorker = globalThis.Worker;
    let terminated = false;
    globalThis.Worker = class {
      postMessage() {}
      terminate() {
        terminated = true;
      }
    };
    try {
      const controller = new AbortController();
      const pending = SpreadsheetJobs.run(
        { type: 'read', file: { size: 1 } },
        { signal: controller.signal },
      );
      controller.abort();
      await rejected(() => pending, /Cancelled/);
      assert(terminated);
    } finally {
      globalThis.Worker = OriginalWorker;
    }
  });
  await check('Real worker imports CSV under strict CSP', async () => {
    const b = await SpreadsheetJobs.run({ type: 'read', file: csv('Student,Binder\nJane Doe,A1') });
    assert(b.SheetNames.length === 1);
  });
  await check('Renamed non-XLSX file rejected by signature', () =>
    rejected(
      () => SpreadsheetJobs.run({ type: 'read', file: csv('not a workbook', 'bad.xlsx') }),
      /not a valid/,
    ),
  );
  await check('Empty file rejected', () =>
    rejected(() => SpreadsheetJobs.run({ type: 'read', file: csv('') }), /nonempty/),
  );
  await check('Workbook with over 20 sheets rejected', async () => {
    const b = book(
      ...Array.from({ length: 21 }, () => [
        ['Student', 'Binder'],
        ['Jane', 'A1'],
      ]),
    );
    await rejected(() => SpreadsheetJobs.run({ type: 'read', file: workbookFile(b) }), /20 sheets/);
  });
  await check('Sheet beyond allowed rows rejected even when read is truncated', async () => {
    const b = book([
      ['Student', 'Binder'],
      ['Jane', 'A1'],
    ]);
    b.Sheets.Sheet1['!ref'] = 'A1:B10002';
    b.Sheets.Sheet1.A10002 = { t: 's', v: 'Too far' };
    await rejected(
      () => SpreadsheetJobs.run({ type: 'read', file: workbookFile(b) }),
      /10,000 rows/,
    );
  });
  await check('Sparse oversized sheet area rejected', async () => {
    const b = book([
      ['Student', 'Binder'],
      ['Jane', 'A1'],
    ]);
    b.Sheets.Sheet1['!ref'] = 'A1:CV3000';
    b.Sheets.Sheet1.CV3000 = { t: 's', v: 'Too far' };
    await rejected(
      () => SpreadsheetJobs.run({ type: 'read', file: workbookFile(b) }),
      /too many cells/,
    );
  });
  await check('Worker generates schedules and matches binders', async () => {
    const binders = book([
      ['Student', 'Binder'],
      ['Jane Doe', 'A1'],
    ]).Sheets.Sheet1;
    const schedule = book([
      ['Time', 'IC 1'],
      ['10:00', 'Jane D'],
    ]).Sheets.Sheet1;
    const result = await SpreadsheetJobs.run({ type: 'generate', binders, schedule });
    assert(result.rows[0].binder === 'A1' && result.rows[0].end === '11:00');
  });
  await check('CSP blocks outbound fetch', async () => {
    await rejected(() => fetch('/src/scheduler.js'), /fetch|network|load/i);
  });
  document.getElementById('testResults').textContent =
    `${passed.length} passed; ${failures.length} failed\n` +
    [...passed.map((n) => `PASS ${n}`), ...failures.map((n) => `FAIL ${n}`)].join('\n');
  document.title = failures.length ? 'FAIL' : 'PASS';
})();
