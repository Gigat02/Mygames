# Le richieste di serata su Telegram

Quando un amico sceglie dei giochi nel sito e preme **Proponi una serata**, il sito manda i dati
a un piccolo servizio (il *relay*) che li gira a **MyGamesBot**, il quale scrive nella tua chat
personale di Telegram.

```
amico sul sito  ──POST──▶  relay su Cloudflare  ──API──▶  MyGamesBot  ──▶  la tua chat
                           (custodisce il token)
```

Il relay serve perché il token del bot è una credenziale: se stesse nel JavaScript del sito
sarebbe leggibile da chiunque apra la pagina, e chiunque potrebbe usare il bot per scriverti.
Su Cloudflare invece resta un segreto lato server.

Tutto quello che segue si fa una volta sola. Il piano gratuito di Cloudflare basta e avanza:
100.000 richieste al giorno, nessuna carta di credito.

---

## 1. Creare il bot (2 minuti)

1. Apri Telegram e cerca **@BotFather** (quello con la spunta blu).
2. Premi *Avvia*, poi scrivi `/newbot`.
3. Alla domanda sul **nome** rispondi `MyGames` — è l'etichetta che vedrai in cima alla chat.
4. Alla domanda sullo **username** rispondi `MyGamesBot`. Deve finire per `bot` ed essere
   libero: se è già preso, prova `MyGamesGigatBot` o `Gigat02MyGamesBot`.
5. BotFather risponde con un messaggio che contiene il **token**, qualcosa come
   `8123456789:AAH...`.

> **Il token è una password.** Non incollarlo in chat, non metterlo nel repository, non
> mandarlo a nessuno: serve solo nel punto 4, dentro Cloudflare. Se per sbaglio finisce in
> giro — anche solo in un messaggio a qualcuno — scrivi `/revoke` a BotFather, scegli il bot
> e ne ottieni uno nuovo: il vecchio smette di funzionare all'istante. Il `CHAT_ID` invece
> non cambia mai, nemmeno dopo una revoca.

Facoltativo, per rifinire: `/setdescription` per la descrizione e `/setuserpic` per usare
`assets/icona-512.png` come immagine del bot.

## 2. Far sì che il bot possa scriverti

Un bot non può scrivere per primo a qualcuno che non gli ha mai parlato.

1. Cerca il tuo bot su Telegram (`@MyGamesBot`) e apri la chat.
2. Premi **Avvia**.

## 3. Trovare il tuo CHAT_ID

Su Telegram cerca **@userinfobot**, aprilo e premi *Avvia*: risponde subito con il tuo
identificativo numerico, del tipo `Id: 123456789`. Per una chat personale il `CHAT_ID`
coincide con l'identificativo dell'utente, quindi quel numero è quello che serve.

Non richiede il token e non lo espone da nessuna parte.

<details>
<summary>In alternativa, con il token</summary>

Apri nel browser `https://api.telegram.org/bot<TOKEN>/getUpdates` e cerca nella risposta
`"chat":{"id":123456789`. Se vedi `"result":[]`, torna sulla chat del bot, scrivigli `ciao`
e ricarica la pagina.

Funziona, ma mette il token nella barra degli indirizzi, dove finisce nella cronologia: da
preferire il metodo qui sopra.

</details>

## 4. Pubblicare il relay su Cloudflare

1. Vai su <https://dash.cloudflare.com/sign-up> e crea un account gratuito (email e password,
   niente carta).
2. Nel menu a sinistra apri **Workers & Pages**, poi **Create** → **Start with Hello World!**
   → **Get started**.
3. Dai al worker il nome `mygames-telegram` e premi **Deploy**.
4. Premi **Edit code**, cancella tutto quello che c'è nell'editor e incolla il contenuto di
   [`worker.js`](worker.js). Poi **Deploy**.
5. Torna alla pagina del worker e apri **Settings** → **Variables and Secrets**
   (in alcune versioni: *Settings* → *Variables*). Aggiungi due voci di tipo **Secret**:

   | Nome        | Valore                          |
   |-------------|---------------------------------|
   | `TOKEN_BOT` | il token del punto 1            |
   | `CHAT_ID`   | il numero del punto 3           |

   Scegli **Secret** e non *Text*: così il valore resta nascosto anche nel pannello.
6. Salva e fai di nuovo **Deploy**.
7. In cima alla pagina trovi l'indirizzo del worker, del tipo
   `https://mygames-telegram.<tuo-sottodominio>.workers.dev`. **Quello puoi condividerlo
   tranquillamente**: non è un segreto.

Per verificare che sia vivo, aprilo nel browser: deve rispondere
`{"servizio":"MyGames — relay Telegram","configurato":true}`. Se `configurato` è `false`,
i due segreti non sono stati salvati.

## 5. Collegare il sito al relay

Apri `js/app.js`, trova la riga

```js
var RELAY_TELEGRAM = '';
```

e mettici dentro l'indirizzo del punto 4:

```js
var RELAY_TELEGRAM = 'https://mygames-telegram.tuo-sottodominio.workers.dev';
```

Poi `git commit` e `git push`: dopo un paio di minuti GitHub Pages si allinea e il pulsante
funziona. Finché la riga resta vuota, il modulo lo dice apertamente invece di fallire in
silenzio.

---

## Come arriva il messaggio

```
🎲 Nuova richiesta da MyGames

👤 Marco vorrebbe giocare
📅 Venerdì 25 settembre 2026

🎯 Giochi scelti
• 7 Wonders
• Splendor
• Skull

📝 dopo cena, da me
🕒 Richiesta inviata il 21/09/2026 alle 18:42
```

Data e ora sono quelle del clic, lette dall'orologio di chi invia; se quell'orologio è
palesemente sbagliato (più di due giorni di scarto) il relay usa la propria.

## Se qualcosa non va

| Sintomo | Causa probabile |
|---|---|
| «il relay non ha ancora TOKEN_BOT e CHAT_ID» | i segreti non sono stati salvati, o manca il Deploy dopo averli aggiunti |
| «origine non autorizzata» | stai provando da un indirizzo diverso da `https://gigat02.github.io` (vedi sotto) |
| Il pulsante dice che manca l'indirizzo del relay | il punto 5 non è stato fatto |

Quando Telegram rifiuta il messaggio, il relay riporta la sua spiegazione nel campo
`dettaglio`. Le due che capitano in fase di configurazione:

| `dettaglio` | Che cosa è successo |
|---|---|
| `Unauthorized` | il `TOKEN_BOT` non corrisponde: di solito è stato incollato con uno spazio davanti o dietro, o manca un pezzo |
| `chat not found` | il `CHAT_ID` è sbagliato, oppure non hai mai premuto *Avvia* sulla chat del bot |

I segreti su Cloudflare non si possono rileggere: se uno dei due è sospetto, si cancella e
si riscrive da capo.

I log delle chiamate si leggono nella scheda **Logs** del worker su Cloudflare.

### Provare in locale

Il relay accetta solo chiamate che arrivano dal sito pubblicato. Per provarlo anche da
`node serve.js`, aggiungi su Cloudflare una variabile di tipo **Text** chiamata
`ORIGINI_CONSENTITE` con valore:

```
https://gigat02.github.io,http://localhost:4180
```

## Sicurezza, in breve

- Il token vive solo fra i segreti di Cloudflare. Il sito non lo vede mai.
- L'indirizzo del relay è pubblico: chi lo scopre può mandarti richieste finte. Il relay si
  difende accettando solo chiamate provenienti dal sito, rifiutando corpi oltre gli 8 KB,
  al massimo 20 giochi per richiesta e ripulendo ogni testo prima di inoltrarlo.
- Il messaggio lo compone il relay a partire da campi verificati: il sito non può far
  scrivere al bot testo arbitrario.
- Se un giorno ricevi richieste indesiderate, cambia l'indirizzo del worker (basta
  rinominarlo) e aggiorna `RELAY_TELEGRAM`.
