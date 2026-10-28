import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';

// Optional dependency-free browser runner. Supply EDGE_PATH for another local
// Chromium browser; the local --test server must already be running.
const executable =
  process.env.EDGE_PATH || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
const profile = mkdtempSync(join(tmpdir(), 'mathnasium-browser-'));
const browser = spawn(
  executable,
  [
    '--headless',
    '--disable-gpu',
    '--no-first-run',
    '--no-default-browser-check',
    '--remote-debugging-port=0',
    `--user-data-dir=${profile}`,
    'about:blank',
  ],
  { windowsHide: true, stdio: 'ignore' },
);
let socket;
try {
  let port;
  for (let tries = 0; tries < 100; tries++) {
    try {
      port = readFileSync(join(profile, 'DevToolsActivePort'), 'utf8').split('\n')[0];
      break;
    } catch {
      await delay(100);
    }
  }
  if (!port) throw new Error('Browser did not start.');
  const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  socket = new WebSocket(targets.find((target) => target.type === 'page').webSocketDebuggerUrl);
  await new Promise((resolve) => socket.addEventListener('open', resolve, { once: true }));
  let id = 0;
  const pending = new Map();
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (pending.has(message.id)) {
      pending.get(message.id)(message);
      pending.delete(message.id);
    }
  });
  const command = (method, params = {}) =>
    new Promise((resolve) => {
      const key = ++id;
      pending.set(key, resolve);
      socket.send(JSON.stringify({ id: key, method, params }));
    });
  for (const file of ['tests/parser.html', 'tests/security.html']) {
    await command('Page.navigate', { url: `http://127.0.0.1:8788/${file}` });
    let report = '';
    for (let tries = 0; tries < 400; tries++) {
      await delay(100);
      const result = await command('Runtime.evaluate', {
        expression: "document.getElementById('testResults')?.textContent",
      });
      report = result.result?.result?.value || '';
      if (/\d+ passed; \d+ failed/.test(report)) break;
    }
    console.log(`${file}: ${report}`);
    if (!/\d+ passed; 0 failed/.test(report)) process.exitCode = 1;
  }
  await command('Page.navigate', { url: 'http://127.0.0.1:8788/' });
  let ready = false;
  for (let tries = 0; tries < 100; tries++) {
    await delay(100);
    const result = await command('Runtime.evaluate', {
      expression: "typeof generate === 'function' && document.readyState === 'complete'",
    });
    if (result.result?.result?.value) {
      ready = true;
      break;
    }
  }
  if (!ready) throw new Error('Production app page did not load.');
  const generated = await command('Runtime.evaluate', {
    awaitPromise: true,
    expression: `(async () => {
    for (const [kind, data] of [['binders', 'Student,Binder\\nJane Doe,A1'], ['schedule', 'Time,HS1\\n10:00,Jane D']]) {
      const list = new DataTransfer(); list.items.add(new File([data], kind + '.csv'));
      document.getElementById(kind).files = list.files;
      await importFile(kind);
    }
    await generate();
    return !document.getElementById('results').hidden && document.getElementById('staffLists').textContent.includes('Jane Doe (A1)');
  })()`,
  });
  if (generated.result?.result?.value !== true)
    throw new Error('Production app import/generation failed.');
  console.log('PASS Production page imports and generates with unchanged security headers');
  const downloads = join(profile, 'downloads');
  mkdirSync(downloads);
  await command('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: downloads });
  await command('Runtime.evaluate', { expression: "document.getElementById('export').click()" });
  const csvPath = join(downloads, 'staff-schedules.csv');
  for (let tries = 0; tries < 100 && !existsSync(csvPath); tries++) await delay(100);
  if (!existsSync(csvPath) || !readFileSync(csvPath, 'utf8').includes('Jane Doe'))
    throw new Error('Production CSV download failed.');
  console.log('PASS Real CSV download under strict security headers');
  const pdf = await command('Page.printToPDF', { printBackground: true });
  if (
    !pdf.result?.data ||
    !Buffer.from(pdf.result.data, 'base64').subarray(0, 4).equals(Buffer.from('%PDF'))
  )
    throw new Error('Print/PDF smoke check failed.');
  console.log('PASS Browser print/PDF generation');
} finally {
  socket?.close();
  browser.kill();
  // The directory was freshly created under the OS temp directory above.
  await delay(500);
  try {
    rmSync(profile, { recursive: true, force: true });
  } catch {
    /* Browser shutdown can briefly retain its temporary profile. */
  }
}
