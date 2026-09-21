/* ============================================================
   MyGames — relay verso Telegram
   ------------------------------------------------------------
   Micro-servizio da pubblicare su Cloudflare Workers. Riceve dal
   sito una richiesta di serata e la inoltra a MyGamesBot, che la
   recapita alla chat di chi possiede la collezione.

   Esiste per un motivo solo: il token del bot è una credenziale e
   non può stare nel JavaScript del sito, che chiunque può leggere.
   Qui vive come segreto, sui server di Cloudflare.

   Segreti attesi (vedi LEGGIMI.md):
     TOKEN_BOT   il token che restituisce BotFather
     CHAT_ID     l'identificativo numerico della chat di destinazione

   Variabile facoltativa:
     ORIGINI_CONSENTITE   elenco separato da virgole dei siti
                          autorizzati a chiamare il relay
   ============================================================ */

const ORIGINI_PREDEFINITE = 'https://gigat02.github.io';

const LIMITI = {
  nome: 60,
  nota: 200,
  giochi: 20,
  titoloGioco: 120
};

export default {
  async fetch(richiesta, ambiente) {
    const consentite = (ambiente.ORIGINI_CONSENTITE || ORIGINI_PREDEFINITE)
      .split(',').map((o) => o.trim()).filter(Boolean);
    const origine = richiesta.headers.get('Origin') || '';
    const origineOk = consentite.includes(origine);

    const intestazioni = {
      'Access-Control-Allow-Origin': origineOk ? origine : consentite[0],
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      'Cache-Control': 'no-store',
      'Content-Type': 'application/json; charset=utf-8'
    };

    const risposta = (stato, corpo) =>
      new Response(JSON.stringify(corpo), { status: stato, headers: intestazioni });

    if (richiesta.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: intestazioni });
    }

    // Aprendo l'indirizzo nel browser si controlla che il relay sia in piedi.
    if (richiesta.method === 'GET') {
      return risposta(200, {
        servizio: 'MyGames — relay Telegram',
        configurato: Boolean(ambiente.TOKEN_BOT && ambiente.CHAT_ID)
      });
    }

    if (richiesta.method !== 'POST') {
      return risposta(405, { errore: 'metodo non ammesso' });
    }

    if (origine && !origineOk) {
      return risposta(403, { errore: 'origine non autorizzata' });
    }

    if (!ambiente.TOKEN_BOT || !ambiente.CHAT_ID) {
      return risposta(500, { errore: 'il relay non ha ancora TOKEN_BOT e CHAT_ID' });
    }

    let dati;
    try {
      const grezzo = await richiesta.text();
      if (grezzo.length > 8000) return risposta(413, { errore: 'richiesta troppo lunga' });
      dati = JSON.parse(grezzo);
    } catch (e) {
      return risposta(400, { errore: 'corpo della richiesta illeggibile' });
    }

    const problema = valida(dati);
    if (problema) return risposta(400, { errore: problema });

    const testo = componiMessaggio(dati);

    const esito = await fetch(`https://api.telegram.org/bot${ambiente.TOKEN_BOT}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chat_id: ambiente.CHAT_ID,
        text: testo,
        parse_mode: 'HTML',
        disable_web_page_preview: true
      })
    });

    if (!esito.ok) {
      const dettaglio = await esito.text();
      console.log('Telegram ha rifiutato il messaggio:', esito.status, dettaglio);
      return risposta(502, { errore: 'Telegram non ha accettato il messaggio' });
    }

    return risposta(200, { esito: 'inviato' });
  }
};

/* ------------------------------------------------------------
   Validazione: il sito è pubblico, quindi non ci si fida di nulla
   ------------------------------------------------------------ */

function testoPulito(valore, massimo) {
  if (typeof valore !== 'string') return '';
  // via i caratteri di controllo, che in un messaggio non servono
  return valore.replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, massimo);
}

function valida(dati) {
  if (!dati || typeof dati !== 'object') return 'dati mancanti';
  if (!testoPulito(dati.nome, LIMITI.nome)) return 'manca il nome';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dati.giorno || '')) return 'giorno non valido';
  if (!Array.isArray(dati.giochi) || dati.giochi.length === 0) return 'nessun gioco scelto';
  if (dati.giochi.length > LIMITI.giochi) return 'troppi giochi in una sola richiesta';
  if (dati.giochi.some((g) => !testoPulito(g, LIMITI.titoloGioco))) return 'titolo di gioco vuoto';
  return null;
}

/* ------------------------------------------------------------
   Composizione del messaggio
   ------------------------------------------------------------ */

function proteggi(testo) {
  return testo.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function giornoInLettere(iso) {
  const data = new Date(`${iso}T12:00:00Z`);
  if (isNaN(data)) return iso;
  const formattato = new Intl.DateTimeFormat('it-IT', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/Rome'
  }).format(data);
  return formattato.charAt(0).toUpperCase() + formattato.slice(1);
}

function istante(iso) {
  // L'orologio di chi invia potrebbe essere sballato: se la data è assurda
  // o illeggibile si ripiega su quella del relay.
  let quando = new Date(iso);
  if (isNaN(quando) || Math.abs(Date.now() - quando.getTime()) > 2 * 24 * 60 * 60 * 1000) {
    quando = new Date();
  }
  return new Intl.DateTimeFormat('it-IT', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', timeZone: 'Europe/Rome'
  }).format(quando).replace(',', ' alle');
}

function componiMessaggio(dati) {
  const nome = proteggi(testoPulito(dati.nome, LIMITI.nome));
  const nota = proteggi(testoPulito(dati.nota, LIMITI.nota));
  const giochi = dati.giochi
    .map((g) => proteggi(testoPulito(g, LIMITI.titoloGioco)))
    .filter(Boolean);

  const righe = [
    '\u{1F3B2} <b>Nuova richiesta da MyGames</b>',
    '',
    `\u{1F464} <b>${nome}</b> vorrebbe giocare`,
    `\u{1F4C5} ${proteggi(giornoInLettere(dati.giorno))}`,
    '',
    giochi.length === 1 ? '\u{1F3AF} <b>Gioco scelto</b>' : '\u{1F3AF} <b>Giochi scelti</b>',
    ...giochi.map((g) => `• ${g}`)
  ];

  if (nota) righe.push('', `\u{1F4DD} <i>${nota}</i>`);
  righe.push('', `\u{1F552} <i>Richiesta inviata il ${istante(dati.inviatoIl)}</i>`);

  return righe.join('\n');
}
