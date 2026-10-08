// Lokal visning af vinduet (demo-data):  node tools/server.mjs  ->  http://localhost:8091/kalender.html
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
const rod = fileURLToPath(new URL('..', import.meta.url));
const typer = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.xml': 'application/xml' };
createServer(async (req, res) => {
  const sti = normalize(decodeURIComponent(new URL(req.url, 'http://x').pathname)).replace(/^([\/])+/, '') || 'kalender.html';
  if (sti.includes('..')) { res.writeHead(400).end(); return; }
  try {
    const data = await readFile(join(rod, sti));
    res.writeHead(200, { 'content-type': typer[extname(sti)] || 'application/octet-stream', 'cache-control': 'no-store' }).end(data);
  } catch { res.writeHead(404).end('Ikke fundet'); }
}).listen(8091, () => console.log('http://localhost:8091/kalender.html'));
