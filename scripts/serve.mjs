import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { publicFiles, root } from './build.mjs';

// Development/QA only. Production serves dist through Cloudflare static assets.
const productionPreview = process.argv.includes('--dist');
const testMode = process.argv.includes('--test');
const directory = productionPreview ? join(root, 'dist') : root;
const allowed = new Set(publicFiles.filter((file) => file !== '_headers'));
if (testMode)
  for (const file of [
    'tests/parser.html',
    'tests/parser.test.js',
    'tests/color-preferences.test.js',
    'tests/takeovers.test.js',
    'tests/helpers.js',
    'tests/app.test.js',
    'tests/security.html',
    'tests/security.test.js',
  ])
    allowed.add(file);
const headers = Object.fromEntries(
  readFileSync(join(directory, '_headers'), 'utf8')
    .split(/\r?\n/)
    .filter((line) => /^\s+\S/.test(line))
    .map((line) => {
      const colon = line.indexOf(':');
      return [line.slice(0, colon).trim(), line.slice(colon + 1).trim()];
    }),
);
const types = {
  html: 'text/html; charset=utf-8',
  js: 'text/javascript; charset=utf-8',
  css: 'text/css; charset=utf-8',
  txt: 'text/plain; charset=utf-8',
};
const server = createServer((request, response) => {
  if (!['GET', 'HEAD'].includes(request.method)) {
    response.writeHead(405, { Allow: 'GET, HEAD' });
    response.end();
    return;
  }
  let file;
  try {
    file =
      decodeURIComponent(new URL(request.url, 'http://localhost').pathname).slice(1) ||
      'index.html';
  } catch {
    response.writeHead(400);
    response.end();
    return;
  }
  const testApp = testMode && file === 'test-app.html';
  if (testApp) file = 'index.html';
  if (!allowed.has(file)) {
    response.writeHead(404, headers);
    response.end('Not found');
    return;
  }
  const responseHeaders = { ...headers };
  // Test-only routes let the browser harness embed the app. Normal and --dist
  // responses keep the production DENY/frame-ancestors 'none' policy intact.
  if (testApp) {
    responseHeaders['Content-Security-Policy'] = headers['Content-Security-Policy'].replace(
      "frame-ancestors 'none'",
      "frame-ancestors 'self'",
    );
    responseHeaders['X-Frame-Options'] = 'SAMEORIGIN';
  }
  if (testMode && (file === 'tests/parser.html' || file === 'tests/security.html'))
    responseHeaders['Content-Security-Policy'] += "; frame-src 'self'";
  try {
    const content = readFileSync(join(directory, file));
    response.writeHead(200, {
      ...responseHeaders,
      'Content-Type': types[file.split('.').pop()] || 'application/octet-stream',
      'Cache-Control': 'no-store',
    });
    response.end(request.method === 'HEAD' ? undefined : content);
  } catch {
    response.writeHead(404, headers);
    response.end('Not found');
  }
});
const port = Number(process.env.PORT || 8788);
server.listen(port, '127.0.0.1', () =>
  console.log(
    `Local: http://127.0.0.1:${server.address().port}${testMode ? '/tests/parser.html' : '/'}`,
  ),
);
