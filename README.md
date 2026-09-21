# MyGames

Collezione personale di giochi da tavolo: elenco filtrabile per nome, numero di giocatori,
durata e tag, con le espansioni possedute annidate sotto ciascun gioco.

Sito: **https://gigat02.github.io/Mygames/**

Salvato sulla schermata Home del telefono si apre a schermo intero con il suo logo.

## Come funziona

- I dati stanno in un unico file Excel, `dati/giochi.xlsx`, con tre fogli:
  - **Giochi** — `Nome`, `Giocatori min`, `Giocatori max`, `Durata min`, `Durata max`, `Tag`,
    `Descrizione`, `Note`
  - **Espansioni** — `Gioco base`, `Espansione`, `Note`
  - **Config** — `password_hash` (SHA-256 della password dell'area riservata) e `versione`
- La pagina legge il file con [SheetJS](https://sheetjs.com/) direttamente nel browser: nessun
  server, nessuna build, nessuna dipendenza da installare.
- L'area **Gestione** permette di aggiungere, modificare ed eliminare giochi ed espansioni.
  Ogni modifica viene riscritta nello stesso file Excel.

## Interfaccia

- L'impaginazione parte dallo schermo del telefono: i filtri diventano fogli che salgono dal
  basso, l'elenco diventa una scheda per gioco. Da 620 px in su i filtri tornano popover e da
  860 px l'elenco torna una tabella.
- Il tema **scuro è quello predefinito**; il pulsante nella testata passa al chiaro e la scelta
  resta memorizzata nel browser.
- Salvando il sito sulla schermata Home (`manifest.webmanifest` + `apple-touch-icon`) si
  apre in modalità `standalone` con il marchio dell'app come icona; le icone PNG si
  rigenerano con `python tools/genera_icone.py`.
- Ogni gioco si apre con un tocco e mostra un riassunto di due righe più le espansioni
  possedute. I riassunti stanno in `tools/descrizioni.py` e finiscono nella colonna
  `Descrizione`; la ricerca testuale li considera insieme a nome, tag ed espansioni.
- I filtri sono quattro menu a tendina — giocatori, durata, tag e ordinamento — più
  l'interruttore per i soli giochi con espansioni. La durata offre cinque intervalli
  preimpostati e due campi `da`/`a` per scriverne uno qualsiasi. I filtri attivi compaiono come
  etichette sotto la barra e si tolgono con un tocco.

## Modificare la collezione

1. Apri il sito e premi **Gestione**.
2. Inserisci la password (quella iniziale è `mygames`: cambiala dal pulsante *Cambia password*).
3. Aggiungi o correggi le voci. Le modifiche restano nel browser finché non le pubblichi.
4. Premi **Salva su GitHub**: il file `dati/giochi.xlsx` viene aggiornato nel repository e il
   sito si riallinea in un paio di minuti.

### Il token GitHub

Per scrivere sul repository serve un [fine-grained personal access token](https://github.com/settings/personal-access-tokens/new)
limitato a questo repository, con permesso **Contents: Read and write**. Va incollato una volta
sola nell'area Gestione: resta nel `localStorage` del browser e non finisce mai nel sito.

Senza token il pulsante **Scarica giochi.xlsx** produce il file aggiornato, da sostituire a mano
in `dati/giochi.xlsx`.

> La password protegge l'interfaccia, non il repository: è comodità, non sicurezza. Ciò che
> impedisce davvero le modifiche altrui è il token, che solo tu possiedi.

## Sviluppo in locale

```bash
node serve.js
```

Poi apri <http://localhost:4180>. Il server statico serve perché `fetch` non legge i file da
`file://`.

Per rigenerare il file Excel di partenza (operazione una tantum, già eseguita):

```bash
python tools/seed_excel.py
```
