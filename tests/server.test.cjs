const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const { createRequestHandler } = require('../server.cjs');

async function withServer(run) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'ng-addin-server-'));
  await fs.writeFile(path.join(root, 'index.html'), '<h1>NG test</h1>');
  await fs.mkdir(path.join(root, '.git'));
  await fs.writeFile(path.join(root, '.git', 'config'), 'private repo data');
  const server = http.createServer(createRequestHandler(root));
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  try {
    await run(`http://127.0.0.1:${address.port}`);
  } finally {
    await new Promise(resolve => server.close(resolve));
    await fs.rm(root, { recursive: true, force: true });
  }
}

test('serves an existing file with its content type', async () => {
  await withServer(async base => {
    const response = await fetch(`${base}/index.html`);
    assert.equal(response.status, 200);
    assert.match(response.headers.get('content-type'), /text\/html/);
    assert.equal(await response.text(), '<h1>NG test</h1>');
  });
});

test('returns 404 for a file that does not exist', async () => {
  await withServer(async base => {
    const response = await fetch(`${base}/missing.js`);
    assert.equal(response.status, 404);
  });
});

test('rejects path traversal outside the served root', async () => {
  await withServer(async base => {
    const response = await fetch(`${base}/..%2fsecret.txt`);
    assert.equal(response.status, 403);
  });
});

test('does not serve hidden repository metadata', async () => {
  await withServer(async base => {
    const response = await fetch(`${base}/.git/config`);
    assert.equal(response.status, 403);
  });
});
