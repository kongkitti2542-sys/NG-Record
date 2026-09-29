const fs = require('node:fs');
const path = require('node:path');
const https = require('node:https');

const MIME_TYPES = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.manifest': 'text/cache-manifest; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.xml': 'application/xml; charset=utf-8',
};

function isInsideRoot(root, filePath) {
  const relativePath = path.relative(root, filePath);
  return relativePath === '' || (!relativePath.startsWith(`..${path.sep}`) && relativePath !== '..' && !path.isAbsolute(relativePath));
}

function createRequestHandler(root) {
  const rootPath = path.resolve(root);

  return async (request, response) => {
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      response.writeHead(405, { Allow: 'GET, HEAD' });
      response.end('Method not allowed');
      return;
    }

    let pathname;
    try {
      pathname = decodeURIComponent(new URL(request.url, 'https://localhost').pathname);
    } catch {
      response.writeHead(400);
      response.end('Invalid URL');
      return;
    }

    if (pathname.includes('\0')) {
      response.writeHead(400);
      response.end('Invalid path');
      return;
    }

    const requestedPath = pathname === '/' ? '/index.html' : pathname;
    const pathSegments = requestedPath.split(/[\\/]+/).filter(Boolean);
    if (pathSegments.some(segment => segment.startsWith('.') || ['docs', 'node_modules', 'tests'].includes(segment.toLowerCase()))) {
      response.writeHead(403);
      response.end('Forbidden');
      return;
    }
    let filePath = path.resolve(rootPath, `.${requestedPath}`);
    if (!isInsideRoot(rootPath, filePath)) {
      response.writeHead(403);
      response.end('Forbidden');
      return;
    }
    if (!Object.hasOwn(MIME_TYPES, path.extname(filePath).toLowerCase())) {
      response.writeHead(404);
      response.end('Not found');
      return;
    }

    try {
      let fileInfo = await fs.promises.stat(filePath);
      if (fileInfo.isDirectory()) {
        filePath = path.join(filePath, 'index.html');
        if (!isInsideRoot(rootPath, filePath)) {
          response.writeHead(403);
          response.end('Forbidden');
          return;
        }
        fileInfo = await fs.promises.stat(filePath);
      }
      if (!fileInfo.isFile()) throw Object.assign(new Error('Not a regular file'), { code: 'ENOENT' });

      response.writeHead(200, {
        'Content-Length': fileInfo.size,
        'Content-Type': MIME_TYPES[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
        'X-Content-Type-Options': 'nosniff',
      });
      if (request.method === 'HEAD') {
        response.end();
        return;
      }
      fs.createReadStream(filePath).pipe(response);
    } catch (error) {
      response.writeHead(error.code === 'ENOENT' || error.code === 'ENOTDIR' ? 404 : 500);
      response.end(error.code === 'ENOENT' || error.code === 'ENOTDIR' ? 'Not found' : 'Internal server error');
    }
  };
}

function createHttpsServer(options = {}) {
  const root = options.root || __dirname;
  const pfxPath = options.pfxPath || path.join(__dirname, 'localhost.pfx');
  const passphrase = options.passphrase ?? process.env.NG_ADDIN_CERT_PASSWORD;
  if (!fs.existsSync(pfxPath)) {
    throw new Error(`Development certificate not found: ${pfxPath}. Run . .\\setup-addin.ps1 first.`);
  }
  return https.createServer({ pfx: fs.readFileSync(pfxPath), passphrase }, createRequestHandler(root));
}

module.exports = { createHttpsServer, createRequestHandler };

if (require.main === module) {
  try {
    const server = createHttpsServer();
    server.listen(3000, 'localhost', () => console.log('NG Record Add-in: https://localhost:3000/index.html'));
    server.on('error', error => {
      console.error(`Could not start the NG Record Add-in server: ${error.message}`);
      process.exitCode = 1;
    });
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
