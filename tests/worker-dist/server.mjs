import { createReadStream, existsSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { extname, join, resolve, sep } from 'node:path';

const root = resolve(new URL('../..', import.meta.url).pathname);
const port = Number(process.env.WORKER_DIST_PORT ?? 6012);
const mime = new Map([
  ['.html', 'text/html; charset=utf-8'],
  ['.js', 'text/javascript; charset=utf-8'],
  ['.mjs', 'text/javascript; charset=utf-8'],
  ['.wasm', 'application/wasm'],
  ['.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
  ['.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
  ['.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
]);

createServer((request, response) => {
  const pathname = new URL(request.url ?? '/', 'http://127.0.0.1').pathname;
  const consumerRoot = join(tmpdir(), 'ooxml-worker-consumer-dist');
  const fromConsumer = pathname.startsWith('/consumer/');
  const servingRoot = fromConsumer ? consumerRoot : root;
  const relative = fromConsumer
    ? pathname.slice('/consumer/'.length)
    : pathname === '/' ? 'tests/worker-dist/fixture.html' : pathname.slice(1);
  const filePath = resolve(servingRoot, relative || 'index.html');
  if (!filePath.startsWith(`${servingRoot}${sep}`) || !existsSync(filePath) || !statSync(filePath).isFile()) {
    response.writeHead(404).end('Not found');
    return;
  }
  response.setHeader('Content-Type', mime.get(extname(filePath)) ?? 'application/octet-stream');
  response.setHeader('Cache-Control', 'no-store');
  createReadStream(filePath).pipe(response);
}).listen(port, '127.0.0.1', () => {
  console.log(`Production worker fixture: http://127.0.0.1:${port}/`);
});
