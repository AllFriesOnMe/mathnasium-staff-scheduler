import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { build, publicFiles, root } from '../../scripts/build.mjs';

test('Build copies only approved files and removes stale student data', () => {
  const output = build();
  writeFileSync(join(output, 'private-students.csv'), 'Sensitive fixture');
  build();
  const files = readdirSync(output, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) =>
      join(entry.parentPath, entry.name)
        .slice(output.length + 1)
        .replaceAll('\\', '/'),
    )
    .sort();
  assert.deepEqual(files, [...publicFiles].sort());
  assert.equal(existsSync(join(output, 'private-students.csv')), false);
});
test('Deployment allows only same-origin connections, scripts and workers', () => {
  const headers = readFileSync(join(root, '_headers'), 'utf8');
  for (const directive of [
    "script-src 'self'",
    "worker-src 'self'",
    "connect-src 'self'",
    "frame-ancestors 'none'",
    "form-action 'none'",
  ])
    assert.ok(headers.includes(directive));
  assert.ok(!headers.includes('unsafe-inline') && !headers.includes('unsafe-eval'));
  const html = readFileSync(join(root, 'index.html'), 'utf8');
  assert.ok(!/<style\b|<script\s*>|\son\w+\s*=/i.test(html));
  for (const file of [
    'src/scheduler.js',
    'src/spreadsheet.js',
    'src/takeovers.js',
    'src/spreadsheet-jobs.js',
    'src/spreadsheet-worker.js',
  ])
    assert.ok(
      !/\.innerHTML\s*=|\.outerHTML\s*=|\beval\(/.test(readFileSync(join(root, file), 'utf8')),
    );
});
test('Git ignore covers workbooks, exports, credentials, and build output', () => {
  const ignored = readFileSync(join(root, '.gitignore'), 'utf8');
  for (const pattern of ['*.xlsx', '*.csv', '*.pdf', '.env', '.env.*', '.local-qa/', 'dist/'])
    assert.ok(ignored.split(/\r?\n/).includes(pattern));
  assert.ok(!ignored.split(/\r?\n/).includes('vendor/'));
});
test('Production preview applies headers and rejects private paths', async () => {
  const child = spawn(process.execPath, ['scripts/serve.mjs', '--dist'], {
    cwd: root,
    env: { ...process.env, PORT: '0' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  try {
    const [chunk] = await once(child.stdout, 'data');
    const url = String(chunk).match(/http:\/\/127\.0\.0\.1:\d+/)?.[0];
    assert.ok(url);
    const page = await fetch(url);
    assert.equal(page.status, 200);
    assert.equal(page.headers.get('x-frame-options'), 'DENY');
    assert.ok(page.headers.get('content-security-policy').includes("connect-src 'self'"));
    assert.equal(page.headers.get('x-content-type-options'), 'nosniff');
    for (const path of [
      '/tests/parser.html',
      '/.env',
      '/.git/config',
      '/_headers',
      '/test-app.html',
      '/student.xlsx',
      '/scripts/build.mjs',
    ])
      assert.equal((await fetch(url + path)).status, 404);
    assert.equal((await fetch(url, { method: 'POST', body: 'student data' })).status, 405);
  } finally {
    child.kill();
  }
});
