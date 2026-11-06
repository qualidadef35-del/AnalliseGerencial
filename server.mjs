/**
 * HEMU | hospedagem alternativa ao Netlify (Render / Node.js).
 * Serve as paginas estaticas e reutiliza a mesma funcao de leitura do Excel.
 * Sem dependencias externas e sem chaves de acesso.
 */
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import dados from './netlify/functions/dados.mjs';

const ROOT = dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 10000;

const staticRoutes = new Map([
  ['/', 'index.html'],
  ['/index.html', 'index.html'],
  ['/evolucao.html', 'evolucao.html'],
]);

async function handle(req, res) {
  const pathname = new URL(req.url || '/', 'http://localhost').pathname;
  try {
    if (pathname === '/health') {
      res.writeHead(200, { 'content-type':'application/json; charset=utf-8', 'cache-control':'no-store' });
      res.end(JSON.stringify({ ok: true, service: 'HEMU Analise Gerencial' }));
      return;
    }
    if (pathname === '/.netlify/functions/dados') {
      // A rota e mantida igual para que o dashboard funcione no Render e no Netlify.
      const resposta = await dados(new Request('http://localhost' + (req.url || pathname), { method: req.method }));
      const headers = Object.fromEntries(resposta.headers.entries());
      res.writeHead(resposta.status, headers);
      res.end(Buffer.from(await resposta.arrayBuffer()));
      return;
    }
    const file = staticRoutes.get(pathname);
    if (file) {
      const body = await readFile(join(ROOT, file));
      res.writeHead(200, {
        'content-type':'text/html; charset=utf-8',
        'cache-control':'no-cache, max-age=0, must-revalidate',
        'x-content-type-options':'nosniff',
      });
      res.end(body);
      return;
    }
    if (pathname === '/favicon.ico') { res.writeHead(204); res.end(); return; }
    res.writeHead(404, { 'content-type':'text/plain; charset=utf-8' });
    res.end('Pagina nao encontrada');
  } catch (error) {
    console.error('HEMU: erro de requisicao', error);
    res.writeHead(500, { 'content-type':'application/json; charset=utf-8', 'cache-control':'no-store' });
    res.end(JSON.stringify({ok:false,error:'Erro interno do servidor'}));
  }
}

createServer(handle).listen(PORT, '0.0.0.0', () => {
  console.log('HEMU - servidor iniciado na porta ' + PORT);
});
