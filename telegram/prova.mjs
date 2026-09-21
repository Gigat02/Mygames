/* Verifiche del relay, senza toccare Telegram né Cloudflare:
   la chiamata a sendMessage viene intercettata e ispezionata.

       node telegram/prova.mjs
*/
import worker from './worker.js';

const SITO = 'https://gigat02.github.io';
const ambiente = { TOKEN_BOT: 'finto:token', CHAT_ID: '123456789' };

let ultimaChiamata = null;
const fetchVero = globalThis.fetch;
globalThis.fetch = async (url, opzioni) => {
  if (String(url).includes('api.telegram.org')) {
    ultimaChiamata = { url: String(url), corpo: JSON.parse(opzioni.body) };
    return new Response(JSON.stringify({ ok: true }), { status: 200 });
  }
  return fetchVero(url, opzioni);
};

function posta(dati, origine = SITO, env = ambiente) {
  return worker.fetch(new Request('https://relay.example/', {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=UTF-8', Origin: origine },
    body: JSON.stringify(dati)
  }), env);
}

const buona = {
  nome: 'Marco',
  giorno: '2026-09-25',
  nota: 'dopo cena, da me',
  giochi: ['7 Wonders', 'Splendor', 'Skull'],
  inviatoIl: new Date().toISOString()
};

let falliti = 0;
function verifica(titolo, condizione, extra = '') {
  console.log(`${condizione ? 'OK  ' : 'KO  '} ${titolo}${extra ? ' — ' + extra : ''}`);
  if (!condizione) falliti++;
}

// 1. controllo di salute
const salute = await worker.fetch(new Request('https://relay.example/', { method: 'GET' }), ambiente);
verifica('GET risponde 200 e configurato:true',
  salute.status === 200 && (await salute.clone().json()).configurato === true);

// 2. preflight
const pre = await worker.fetch(new Request('https://relay.example/', {
  method: 'OPTIONS', headers: { Origin: SITO }
}), ambiente);
verifica('OPTIONS risponde 204 con CORS',
  pre.status === 204 && pre.headers.get('Access-Control-Allow-Origin') === SITO);

// 3. validazioni
const casi = [
  ['nome mancante', { ...buona, nome: '   ' }, 'manca il nome'],
  ['giorno non valido', { ...buona, giorno: '25/09/2026' }, 'giorno non valido'],
  ['nessun gioco', { ...buona, giochi: [] }, 'nessun gioco scelto'],
  ['troppi giochi', { ...buona, giochi: Array(21).fill('X') }, 'troppi giochi in una sola richiesta']
];
for (const [titolo, dati, atteso] of casi) {
  const r = await posta(dati);
  const corpo = await r.json();
  verifica(`rifiuta: ${titolo}`, r.status === 400 && corpo.errore === atteso, corpo.errore);
}

// 4. origine estranea
const estranea = await posta(buona, 'https://sito-a-caso.example');
verifica('rifiuta un\'origine estranea', estranea.status === 403);

// 5. relay non configurato
const scarico = await posta(buona, SITO, {});
verifica('segnala i segreti mancanti', scarico.status === 500);

// 6. invio riuscito
ultimaChiamata = null;
const esito = await posta(buona);
verifica('accetta una richiesta valida', esito.status === 200 && (await esito.json()).esito === 'inviato');
verifica('chiama sendMessage sulla chat giusta',
  ultimaChiamata && ultimaChiamata.corpo.chat_id === '123456789'
  && ultimaChiamata.url.includes('/sendMessage'));

// 7. niente HTML iniettabile
const cattivo = await posta({ ...buona, nome: '<b>Ivan</b>', nota: '<a href="x">clic</a>' });
verifica('neutralizza l\'HTML nei campi liberi',
  cattivo.status === 200 && !ultimaChiamata.corpo.text.includes('<b>Ivan</b>')
  && ultimaChiamata.corpo.text.includes('&lt;b&gt;Ivan&lt;/b&gt;'));

// 8. orologio sballato
await posta({ ...buona, inviatoIl: '1999-01-01T00:00:00Z' });
verifica('ignora un orologio assurdo',
  !ultimaChiamata.corpo.text.includes('1999'), ultimaChiamata.corpo.text.split('\n').pop());

console.log('\n--- messaggio che arriverebbe su Telegram ---');
await posta(buona);
console.log(ultimaChiamata.corpo.text);
console.log('--------------------------------------------\n');

console.log(falliti === 0 ? 'Tutte le verifiche superate.' : `${falliti} verifiche fallite.`);
process.exit(falliti === 0 ? 0 : 1);
