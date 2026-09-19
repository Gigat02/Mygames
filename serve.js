/* Piccolo server statico per provare MyGames in locale.
   Serve perché il file Excel viene letto con fetch, che da file:// è bloccato.

       node serve.js      →  http://localhost:4180
*/

const http = require('http');
const fs = require('fs');
const path = require('path');

const PORTA = process.env.PORT || 4180;
const RADICE = __dirname;

const TIPI = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.webmanifest': 'application/manifest+json'
};

http.createServer((richiesta, risposta) => {
  const indirizzo = decodeURIComponent(richiesta.url.split('?')[0]);
  let file = path.join(RADICE, indirizzo === '/' ? 'index.html' : indirizzo);

  if (!path.resolve(file).startsWith(RADICE)) {
    risposta.writeHead(403).end('Accesso negato');
    return;
  }

  fs.stat(file, (errore, info) => {
    if (errore) {
      risposta.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Non trovato');
      return;
    }
    if (info.isDirectory()) file = path.join(file, 'index.html');

    fs.readFile(file, (erroreLettura, contenuto) => {
      if (erroreLettura) {
        risposta.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('Non trovato');
        return;
      }
      risposta.writeHead(200, {
        'Content-Type': TIPI[path.extname(file).toLowerCase()] || 'application/octet-stream',
        'Cache-Control': 'no-store'
      }).end(contenuto);
    });
  });
}).listen(PORTA, () => {
  console.log(`MyGames in ascolto su http://localhost:${PORTA}`);
});
