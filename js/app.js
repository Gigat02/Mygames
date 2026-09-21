/* ============================================================
   MyGames — logica dell'applicazione
   Il file dati/giochi.xlsx è l'unica fonte dei dati: viene letto
   all'avvio e riscritto (via API GitHub) a ogni modifica.
   ============================================================ */

(function () {
  'use strict';

  var REPO = 'Gigat02/Mygames';
  var RAMO = 'main';
  var PERCORSO_EXCEL = 'dati/giochi.xlsx';

  var CHIAVE_TEMA = 'mygames:tema';
  var CHIAVE_TOKEN = 'mygames:token';
  var CHIAVE_CACHE = 'mygames:bozza';
  var CHIAVE_SBLOCCO = 'mygames:sbloccato';
  var CHIAVE_SCELTI = 'mygames:scelti';
  var CHIAVE_NOME = 'mygames:nome';

  /* Indirizzo del relay che inoltra le richieste a MyGamesBot su Telegram.
     Il token del bot vive come segreto dentro il relay, mai in questa pagina.
     Istruzioni per crearlo: telegram/LEGGIMI.md. Finché resta vuoto, il
     modulo di richiesta lo dice apertamente invece di fallire in silenzio. */
  var RELAY_TELEGRAM = '';

  var PRESET_DURATA = [
    { id: 'lampo', etichetta: 'Fino a 15 minuti', min: null, max: 15 },
    { id: 'breve', etichetta: 'Da 16 a 30 minuti', min: 16, max: 30 },
    { id: 'media', etichetta: 'Da 31 a 60 minuti', min: 31, max: 60 },
    { id: 'lunga', etichetta: 'Da 61 a 120 minuti', min: 61, max: 120 },
    { id: 'epica', etichetta: 'Oltre 120 minuti', min: 121, max: null }
  ];

  var ORDINI = [
    { id: 'nome', etichetta: 'Nome A → Z' },
    { id: 'nome-desc', etichetta: 'Nome Z → A' },
    { id: 'durata', etichetta: 'Durata crescente' },
    { id: 'durata-desc', etichetta: 'Durata decrescente' },
    { id: 'giocatori-desc', etichetta: 'Più giocatori' },
    { id: 'giocatori', etichetta: 'Meno giocatori' }
  ];

  var stato = {
    giochi: [],
    espansioni: [],
    passwordHash: '',
    versione: 1,
    versioneFile: 1,
    bozza: false,
    aperti: {},
    scelti: []
  };

  var filtri = {
    testo: '',
    giocatori: null,
    durataMin: null,
    durataMax: null,
    tag: [],
    tuttiITag: false,
    soloEspansioni: false,
    ordine: 'nome'
  };

  var el = {};
  var tendine = [];
  var tendinaAperta = null;
  var cercaTag = '';

  /* ---------------------------------------------------------
     Utilità
     --------------------------------------------------------- */

  function $(id) { return document.getElementById(id); }

  function chiave(testo) {
    return String(testo == null ? '' : testo).trim().toLowerCase();
  }

  function numero(valore) {
    if (valore === null || valore === undefined || valore === '') return null;
    var n = parseInt(valore, 10);
    return isNaN(n) ? null : n;
  }

  function separaTag(testo) {
    if (!testo) return [];
    return String(testo).split(',').map(function (t) { return t.trim(); })
      .filter(function (t) { return t.length > 0; });
  }

  function senzaAccenti(testo) {
    return chiave(testo).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function creaElemento(tag, classe, testo) {
    var nodo = document.createElement(tag);
    if (classe) nodo.className = classe;
    if (testo !== undefined && testo !== null) nodo.textContent = testo;
    return nodo;
  }

  function avvisa(testo) {
    el.avviso.textContent = testo;
    el.avviso.hidden = false;
    clearTimeout(avvisa._t);
    avvisa._t = setTimeout(function () { el.avviso.hidden = true; }, 3200);
  }

  function leggiLocale(chiaveStorage) {
    try { return localStorage.getItem(chiaveStorage); } catch (e) { return null; }
  }

  function scriviLocale(chiaveStorage, valore) {
    try {
      if (valore === null) localStorage.removeItem(chiaveStorage);
      else localStorage.setItem(chiaveStorage, valore);
    } catch (e) { /* spazio non disponibile: si prosegue senza cache */ }
  }

  function sha256(testo) {
    if (!window.crypto || !window.crypto.subtle) {
      return Promise.reject(new Error('Le funzioni di sicurezza del browser non sono disponibili su questa pagina.'));
    }
    var dati = new TextEncoder().encode(testo);
    return window.crypto.subtle.digest('SHA-256', dati).then(function (buffer) {
      return Array.prototype.map.call(new Uint8Array(buffer), function (b) {
        return ('00' + b.toString(16)).slice(-2);
      }).join('');
    });
  }

  /* ---------------------------------------------------------
     Lettura e scrittura del file Excel
     --------------------------------------------------------- */

  function daWorkbook(wb) {
    var giochi = [];
    var espansioni = [];
    var configurazione = {};

    var foglioGiochi = wb.Sheets[wb.SheetNames.indexOf('Giochi') >= 0 ? 'Giochi' : wb.SheetNames[0]];
    XLSX.utils.sheet_to_json(foglioGiochi, { defval: null }).forEach(function (riga) {
      var nome = riga['Nome'];
      if (!nome || !String(nome).trim()) return;
      giochi.push({
        nome: String(nome).trim(),
        min: numero(riga['Giocatori min']),
        max: numero(riga['Giocatori max']),
        dmin: numero(riga['Durata min']),
        dmax: numero(riga['Durata max']),
        tag: separaTag(riga['Tag']),
        descrizione: riga['Descrizione'] ? String(riga['Descrizione']).trim() : '',
        note: riga['Note'] ? String(riga['Note']).trim() : ''
      });
    });

    if (wb.Sheets['Espansioni']) {
      XLSX.utils.sheet_to_json(wb.Sheets['Espansioni'], { defval: null }).forEach(function (riga) {
        var nome = riga['Espansione'];
        var base = riga['Gioco base'];
        if (!nome || !base) return;
        espansioni.push({
          base: String(base).trim(),
          nome: String(nome).trim(),
          note: riga['Note'] ? String(riga['Note']).trim() : ''
        });
      });
    }

    if (wb.Sheets['Config']) {
      XLSX.utils.sheet_to_json(wb.Sheets['Config'], { defval: null }).forEach(function (riga) {
        if (riga['Chiave']) configurazione[String(riga['Chiave']).trim()] = riga['Valore'];
      });
    }

    return { giochi: giochi, espansioni: espansioni, configurazione: configurazione };
  }

  function ordinaPerNome(elenco) {
    return elenco.slice().sort(function (a, b) { return a.nome.localeCompare(b.nome, 'it'); });
  }

  function aWorkbook() {
    var wb = XLSX.utils.book_new();

    var righeGiochi = [['Nome', 'Giocatori min', 'Giocatori max', 'Durata min', 'Durata max',
      'Tag', 'Descrizione', 'Note']];
    ordinaPerNome(stato.giochi).forEach(function (g) {
      righeGiochi.push([g.nome, g.min, g.max, g.dmin, g.dmax, g.tag.join(', '),
        g.descrizione || null, g.note || null]);
    });
    var fg = XLSX.utils.aoa_to_sheet(righeGiochi);
    fg['!cols'] = [{ wch: 42 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
      { wch: 62 }, { wch: 90 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, fg, 'Giochi');

    var righeEspansioni = [['Gioco base', 'Espansione', 'Note']];
    stato.espansioni.slice()
      .sort(function (a, b) { return a.base.localeCompare(b.base, 'it') || a.nome.localeCompare(b.nome, 'it'); })
      .forEach(function (e) { righeEspansioni.push([e.base, e.nome, e.note || null]); });
    var fe = XLSX.utils.aoa_to_sheet(righeEspansioni);
    fe['!cols'] = [{ wch: 42 }, { wch: 42 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, fe, 'Espansioni');

    var fc = XLSX.utils.aoa_to_sheet([
      ['Chiave', 'Valore'],
      ['password_hash', stato.passwordHash],
      ['versione', stato.versione]
    ]);
    fc['!cols'] = [{ wch: 24 }, { wch: 72 }];
    XLSX.utils.book_append_sheet(wb, fc, 'Config');

    return wb;
  }

  function caricaDati() {
    return fetch(PERCORSO_EXCEL + '?t=' + Date.now(), { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('Impossibile leggere ' + PERCORSO_EXCEL + ' (' + r.status + ')');
        return r.arrayBuffer();
      })
      .then(function (buffer) {
        var lettura = daWorkbook(XLSX.read(new Uint8Array(buffer), { type: 'array' }));
        stato.giochi = lettura.giochi;
        stato.espansioni = lettura.espansioni;
        stato.passwordHash = String(lettura.configurazione.password_hash || '');
        stato.versioneFile = numero(lettura.configurazione.versione) || 1;
        stato.versione = stato.versioneFile;
        applicaBozza();
      });
  }

  function applicaBozza() {
    var grezzo = leggiLocale(CHIAVE_CACHE);
    if (!grezzo) return;
    var bozza;
    try { bozza = JSON.parse(grezzo); } catch (e) { scriviLocale(CHIAVE_CACHE, null); return; }
    if (!bozza || !bozza.versione || bozza.versione <= stato.versioneFile) {
      scriviLocale(CHIAVE_CACHE, null);
      return;
    }
    stato.giochi = bozza.giochi || stato.giochi;
    stato.espansioni = bozza.espansioni || stato.espansioni;
    stato.passwordHash = bozza.passwordHash || stato.passwordHash;
    stato.versione = bozza.versione;
    stato.bozza = true;
  }

  function segnaModifica() {
    stato.versione = Math.max(stato.versione, stato.versioneFile) + 1;
    stato.bozza = true;
    scriviLocale(CHIAVE_CACHE, JSON.stringify({
      versione: stato.versione,
      giochi: stato.giochi,
      espansioni: stato.espansioni,
      passwordHash: stato.passwordHash
    }));
    disegnaTutto();
  }

  /* ---------------------------------------------------------
     Filtri
     --------------------------------------------------------- */

  function espansioniDi(gioco) {
    var k = chiave(gioco.nome);
    return stato.espansioni.filter(function (e) { return chiave(e.base) === k; });
  }

  function conteggioTag() {
    var visti = {};
    stato.giochi.forEach(function (g) {
      g.tag.forEach(function (t) { visti[t] = (visti[t] || 0) + 1; });
    });
    return visti;
  }

  function elencoTag() {
    var visti = conteggioTag();
    return Object.keys(visti).sort(function (a, b) {
      return visti[b] - visti[a] || a.localeCompare(b, 'it');
    });
  }

  function presetAttivo() {
    for (var i = 0; i < PRESET_DURATA.length; i++) {
      var p = PRESET_DURATA[i];
      if (p.min === filtri.durataMin && p.max === filtri.durataMax) return p;
    }
    return null;
  }

  function filtra() {
    var testo = senzaAccenti(filtri.testo);

    var elenco = stato.giochi.filter(function (g) {
      var esp = espansioniDi(g);

      if (testo) {
        var corrisponde = senzaAccenti(g.nome).indexOf(testo) >= 0
          || g.tag.some(function (t) { return senzaAccenti(t).indexOf(testo) >= 0; })
          || senzaAccenti(g.descrizione).indexOf(testo) >= 0
          || esp.some(function (e) { return senzaAccenti(e.nome).indexOf(testo) >= 0; });
        if (!corrisponde) return false;
      }

      if (filtri.giocatori) {
        var min = g.min || 1;
        var max = g.max || 99;
        if (filtri.giocatori < min || filtri.giocatori > max) return false;
      }

      if (filtri.durataMin !== null || filtri.durataMax !== null) {
        var da = g.dmin || g.dmax;
        var a = g.dmax || g.dmin;
        if (!da && !a) return false;
        if (filtri.durataMin !== null && a < filtri.durataMin) return false;
        if (filtri.durataMax !== null && da > filtri.durataMax) return false;
      }

      if (filtri.tag.length) {
        var presenti = g.tag.map(chiave);
        var verifica = filtri.tag.map(chiave);
        var esito = filtri.tuttiITag
          ? verifica.every(function (t) { return presenti.indexOf(t) >= 0; })
          : verifica.some(function (t) { return presenti.indexOf(t) >= 0; });
        if (!esito) return false;
      }

      if (filtri.soloEspansioni && esp.length === 0) return false;

      return true;
    });

    var perDurata = function (g) { return g.dmin || g.dmax || 0; };

    switch (filtri.ordine) {
      case 'nome-desc':
        elenco.sort(function (a, b) { return b.nome.localeCompare(a.nome, 'it'); });
        break;
      case 'durata':
        elenco.sort(function (a, b) { return perDurata(a) - perDurata(b) || a.nome.localeCompare(b.nome, 'it'); });
        break;
      case 'durata-desc':
        elenco.sort(function (a, b) { return perDurata(b) - perDurata(a) || a.nome.localeCompare(b.nome, 'it'); });
        break;
      case 'giocatori':
        elenco.sort(function (a, b) { return (a.max || 99) - (b.max || 99) || a.nome.localeCompare(b.nome, 'it'); });
        break;
      case 'giocatori-desc':
        elenco.sort(function (a, b) { return (b.max || 99) - (a.max || 99) || a.nome.localeCompare(b.nome, 'it'); });
        break;
      default:
        elenco.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'it'); });
    }

    return elenco;
  }

  function etichettaDurata() {
    if (filtri.durataMin === null && filtri.durataMax === null) return 'Qualsiasi';
    var preset = presetAttivo();
    if (preset) {
      if (preset.min === null) return 'Fino a ' + preset.max + '′';
      if (preset.max === null) return 'Oltre ' + (preset.min - 1) + '′';
      return preset.min + '–' + preset.max + '′';
    }
    if (filtri.durataMin === null) return 'Fino a ' + filtri.durataMax + '′';
    if (filtri.durataMax === null) return 'Da ' + filtri.durataMin + '′';
    return filtri.durataMin + '–' + filtri.durataMax + '′';
  }

  function etichettaOrdine() {
    var trovato = ORDINI.filter(function (o) { return o.id === filtri.ordine; })[0];
    return trovato ? trovato.etichetta : ORDINI[0].etichetta;
  }

  function filtriAttivi() {
    return !!(filtri.testo || filtri.giocatori || filtri.durataMin !== null
      || filtri.durataMax !== null || filtri.tag.length || filtri.soloEspansioni);
  }

  function azzeraFiltri() {
    filtri.testo = '';
    filtri.giocatori = null;
    filtri.durataMin = null;
    filtri.durataMax = null;
    filtri.tag = [];
    filtri.soloEspansioni = false;
    el.testo.value = '';
    disegnaFiltri();
    disegnaLista();
  }

  /* ---------------------------------------------------------
     Menu a tendina
     --------------------------------------------------------- */

  function chiudiTendine() {
    tendine.forEach(function (t) {
      t.pannello.hidden = true;
      t.bottone.setAttribute('aria-expanded', 'false');
    });
    tendinaAperta = null;
    el.velo.hidden = true;
    document.body.style.overflow = '';
  }

  function apriTendina(voce) {
    var giaAperta = tendinaAperta === voce;
    chiudiTendine();
    if (giaAperta) return;
    voce.pannello.hidden = false;
    voce.bottone.setAttribute('aria-expanded', 'true');
    tendinaAperta = voce;
    el.velo.hidden = false;
    if (window.matchMedia('(max-width: 619px)').matches) document.body.style.overflow = 'hidden';
  }

  function preparaTendine() {
    tendine = Array.prototype.map.call(document.querySelectorAll('.tendina'), function (nodo) {
      var voce = {
        nome: nodo.getAttribute('data-tendina'),
        nodo: nodo,
        bottone: nodo.querySelector('.tendina__btn'),
        pannello: nodo.querySelector('.tendina__pannello')
      };
      voce.bottone.addEventListener('click', function (ev) {
        ev.stopPropagation();
        apriTendina(voce);
      });
      voce.pannello.addEventListener('click', function (ev) { ev.stopPropagation(); });
      return voce;
    });

    el.velo.addEventListener('click', chiudiTendine);
    document.addEventListener('click', chiudiTendine);
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape' && tendinaAperta) chiudiTendine();
    });

    Array.prototype.forEach.call(document.querySelectorAll('[data-chiudi]'), function (b) {
      b.addEventListener('click', chiudiTendine);
    });
  }

  /* ---------------------------------------------------------
     Presentazione dei valori
     --------------------------------------------------------- */

  function mostraGiocatori(g) {
    if (!g.min && !g.max) return '—';
    var min = g.min || g.max;
    var max = g.max || g.min;
    if (max >= 99) return min + '+';
    return min === max ? String(min) : min + '–' + max;
  }

  function mostraDurata(g) {
    if (!g.dmin && !g.dmax) return '—';
    var da = g.dmin || g.dmax;
    var a = g.dmax || g.dmin;
    return (da === a ? String(da) : da + '–' + a) + '′';
  }

  /* ---------------------------------------------------------
     Disegno dei filtri
     --------------------------------------------------------- */

  function disegnaNumeri() {
    el.numeriGiocatori.textContent = '';
    for (var n = 1; n <= 12; n++) {
      (function (valore) {
        var b = creaElemento('button', 'numero', valore === 12 ? '12+' : String(valore));
        b.type = 'button';
        b.setAttribute('aria-pressed', filtri.giocatori === valore ? 'true' : 'false');
        b.addEventListener('click', function () {
          filtri.giocatori = filtri.giocatori === valore ? null : valore;
          disegnaFiltri();
          disegnaLista();
        });
        el.numeriGiocatori.appendChild(b);
      })(n);
    }
  }

  function disegnaPresetDurata(saltaIntervallo) {
    var attivo = presetAttivo();
    el.presetDurata.textContent = '';
    PRESET_DURATA.forEach(function (p) {
      var b = creaElemento('button', 'opzione', p.etichetta);
      b.type = 'button';
      b.setAttribute('aria-pressed', attivo && attivo.id === p.id ? 'true' : 'false');
      b.addEventListener('click', function () {
        if (attivo && attivo.id === p.id) {
          filtri.durataMin = null;
          filtri.durataMax = null;
        } else {
          filtri.durataMin = p.min;
          filtri.durataMax = p.max;
        }
        disegnaFiltri();
        disegnaLista();
      });
      el.presetDurata.appendChild(b);
    });

    // mentre l'utente scrive nei due campi non li riscriviamo sotto le sue dita
    if (saltaIntervallo) return;
    el.durataDa.value = filtri.durataMin === null ? '' : filtri.durataMin;
    el.durataA.value = filtri.durataMax === null ? '' : filtri.durataMax;
  }

  function disegnaListaTag() {
    var conteggi = conteggioTag();
    var ricerca = senzaAccenti(cercaTag);
    el.listaTag.textContent = '';

    var tutti = elencoTag();
    var scelti = tutti.filter(function (t) { return filtri.tag.indexOf(t) >= 0; });
    var restanti = tutti.filter(function (t) {
      return filtri.tag.indexOf(t) < 0 && (!ricerca || senzaAccenti(t).indexOf(ricerca) >= 0);
    });

    scelti.concat(restanti).forEach(function (t) {
      var attivo = filtri.tag.indexOf(t) >= 0;
      var b = creaElemento('button', 'scelta');
      b.type = 'button';
      b.setAttribute('aria-pressed', attivo ? 'true' : 'false');
      b.appendChild(creaElemento('span', null, t));
      b.appendChild(creaElemento('span', 'scelta__conteggio', String(conteggi[t] || 0)));
      b.addEventListener('click', function () {
        var i = filtri.tag.indexOf(t);
        if (i >= 0) filtri.tag.splice(i, 1); else filtri.tag.push(t);
        disegnaFiltri();
        disegnaLista();
      });
      el.listaTag.appendChild(b);
    });

    if (!el.listaTag.childNodes.length) {
      el.listaTag.appendChild(creaElemento('p', 'nota', 'Nessun tag corrisponde.'));
    }
  }

  function disegnaListaOrdine() {
    el.listaOrdine.textContent = '';
    ORDINI.forEach(function (o) {
      var b = creaElemento('button', 'scelta', o.etichetta);
      b.type = 'button';
      b.setAttribute('aria-pressed', filtri.ordine === o.id ? 'true' : 'false');
      b.addEventListener('click', function () {
        filtri.ordine = o.id;
        chiudiTendine();
        disegnaFiltri();
        disegnaLista();
      });
      el.listaOrdine.appendChild(b);
    });
  }

  function aggiungiChip(testo, alRimuovere) {
    var b = creaElemento('button', 'attivo');
    b.type = 'button';
    b.appendChild(creaElemento('span', null, testo));
    b.appendChild(creaElemento('span', 'attivo__x', '×'));
    b.setAttribute('aria-label', 'Rimuovi il filtro ' + testo);
    b.addEventListener('click', function () {
      alRimuovere();
      disegnaFiltri();
      disegnaLista();
    });
    el.attivi.appendChild(b);
  }

  function disegnaChipAttivi() {
    el.attivi.textContent = '';

    if (filtri.giocatori) {
      aggiungiChip(filtri.giocatori + (filtri.giocatori === 12 ? '+ giocatori' : ' giocatori'), function () {
        filtri.giocatori = null;
      });
    }

    if (filtri.durataMin !== null || filtri.durataMax !== null) {
      aggiungiChip(etichettaDurata(), function () {
        filtri.durataMin = null;
        filtri.durataMax = null;
      });
    }

    filtri.tag.forEach(function (t) {
      aggiungiChip(t, function () {
        var i = filtri.tag.indexOf(t);
        if (i >= 0) filtri.tag.splice(i, 1);
      });
    });

    if (filtri.soloEspansioni) {
      aggiungiChip('Con espansioni', function () { filtri.soloEspansioni = false; });
    }

    if (filtriAttivi()) {
      var azzera = creaElemento('button', 'attivo attivo--azzera', 'Azzera tutto');
      azzera.type = 'button';
      azzera.addEventListener('click', azzeraFiltri);
      el.attivi.appendChild(azzera);
    }
  }

  function segnaTendina(nome, attivo) {
    var voce = tendine.filter(function (t) { return t.nome === nome; })[0];
    if (voce) voce.bottone.classList.toggle('tendina__btn--attivo', !!attivo);
  }

  function disegnaFiltri(saltaIntervallo) {
    el.valGiocatori.textContent = filtri.giocatori
      ? (filtri.giocatori === 12 ? '12 o più' : filtri.giocatori + ' giocatori')
      : 'Tutti';
    el.valDurata.textContent = etichettaDurata();
    el.valTag.textContent = filtri.tag.length === 0 ? 'Tutti'
      : (filtri.tag.length === 1 ? filtri.tag[0] : filtri.tag.length + ' tag');
    el.valOrdine.textContent = etichettaOrdine();

    segnaTendina('giocatori', filtri.giocatori);
    segnaTendina('durata', filtri.durataMin !== null || filtri.durataMax !== null);
    segnaTendina('tag', filtri.tag.length);

    el.btnSoloEspansioni.setAttribute('aria-pressed', filtri.soloEspansioni ? 'true' : 'false');

    Array.prototype.forEach.call(document.querySelectorAll('[data-modo-tag]'), function (b) {
      var attivo = (b.getAttribute('data-modo-tag') === 'tutti') === filtri.tuttiITag;
      b.setAttribute('aria-pressed', attivo ? 'true' : 'false');
    });

    disegnaNumeri();
    disegnaPresetDurata(saltaIntervallo);
    disegnaListaTag();
    disegnaListaOrdine();
    disegnaChipAttivi();
  }

  /* ---------------------------------------------------------
     Disegno della collezione
     --------------------------------------------------------- */

  function disegnaLista() {
    var elenco = filtra();
    el.righe.textContent = '';

    elenco.forEach(function (g) {
      var esp = espansioniDi(g);
      var apribile = !!(g.descrizione || esp.length);
      var aperto = apribile && !!stato.aperti[chiave(g.nome)];

      var riga = creaElemento('div', 'riga' + (apribile ? ' riga--cliccabile' : '') + (aperto ? ' riga--aperta' : ''));
      riga.setAttribute('role', 'row');

      if (apribile) {
        riga.tabIndex = 0;
        riga.setAttribute('aria-expanded', aperto ? 'true' : 'false');
        riga.appendChild(creaElemento('span', 'riga__freccia', '▶'));
      }

      var nome = creaElemento('span', 'riga__nome');
      nome.setAttribute('role', 'cell');
      nome.appendChild(document.createTextNode(g.nome));
      if (esp.length) nome.appendChild(creaElemento('span', 'riga__espansioni', esp.length + ' esp.'));
      riga.appendChild(nome);

      var dati = creaElemento('span', 'riga__dati');
      var giocatori = creaElemento('span', 'riga__dato');
      giocatori.setAttribute('role', 'cell');
      giocatori.appendChild(creaElemento('strong', null, mostraGiocatori(g)));
      giocatori.appendChild(document.createTextNode(' giocatori'));
      dati.appendChild(giocatori);

      var durata = creaElemento('span', 'riga__dato');
      durata.setAttribute('role', 'cell');
      durata.appendChild(creaElemento('strong', null, mostraDurata(g)));
      dati.appendChild(durata);
      riga.appendChild(dati);

      var contenitoreTag = creaElemento('span', 'riga__tag');
      contenitoreTag.setAttribute('role', 'cell');
      g.tag.forEach(function (t) { contenitoreTag.appendChild(creaElemento('span', 'tag', t)); });
      riga.appendChild(contenitoreTag);

      var scelto = stato.scelti.indexOf(g.nome) >= 0;
      var scegli = creaElemento('button', 'scegli', scelto ? '✓' : '+');
      scegli.type = 'button';
      scegli.setAttribute('aria-pressed', scelto ? 'true' : 'false');
      scegli.setAttribute('title', scelto ? 'Togli dalla richiesta' : 'Aggiungi alla richiesta');
      scegli.setAttribute('aria-label', (scelto ? 'Togli ' : 'Aggiungi ') + g.nome
        + (scelto ? ' dalla richiesta' : ' alla richiesta'));
      scegli.addEventListener('click', function (ev) {
        ev.stopPropagation();
        alternaScelta(g.nome);
      });
      riga.appendChild(scegli);

      if (g.note) riga.appendChild(creaElemento('p', 'riga__note', g.note));

      el.righe.appendChild(riga);

      if (apribile) {
        var apri = function () {
          stato.aperti[chiave(g.nome)] = !stato.aperti[chiave(g.nome)];
          disegnaLista();
        };
        riga.addEventListener('click', apri);
        riga.addEventListener('keydown', function (ev) {
          if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); apri(); }
        });

        if (aperto) {
          var blocco = creaElemento('div', 'annidate');
          if (g.descrizione) blocco.appendChild(creaElemento('p', 'dettaglio', g.descrizione));
          esp.forEach(function (e) {
            var voce = creaElemento('div', 'annidata');
            voce.appendChild(creaElemento('span', 'annidata__etichetta', 'Espansione'));
            voce.appendChild(creaElemento('span', null, e.nome));
            if (e.note) voce.appendChild(creaElemento('span', 'annidata__note', '— ' + e.note));
            blocco.appendChild(voce);
          });
          el.righe.appendChild(blocco);
        }
      }
    });

    el.vuoto.hidden = elenco.length > 0;
    el.lista.hidden = elenco.length === 0;
    el.risultati.textContent = elenco.length === stato.giochi.length
      ? 'Tutti i ' + stato.giochi.length + ' giochi'
      : elenco.length + (elenco.length === 1 ? ' gioco su ' : ' giochi su ') + stato.giochi.length;
  }

  function disegnaConteggio() {
    el.conteggio.textContent = stato.giochi.length + ' giochi · ' + stato.espansioni.length + ' espansioni'
      + (stato.bozza ? ' · modifiche non pubblicate' : '');
  }

  function disegnaTagModulo() {
    var tutti = elencoTag();
    el.elencoTag.textContent = '';
    el.tagSuggeriti.textContent = '';
    tutti.forEach(function (t) {
      var opzione = document.createElement('option');
      opzione.value = t;
      el.elencoTag.appendChild(opzione);

      var pillola = creaElemento('button', 'pillola', t);
      pillola.type = 'button';
      pillola.addEventListener('click', function () {
        var correnti = separaTag(el.gTag.value);
        if (correnti.map(chiave).indexOf(chiave(t)) < 0) correnti.push(t);
        el.gTag.value = correnti.join(', ');
      });
      el.tagSuggeriti.appendChild(pillola);
    });
  }

  function disegnaTutto() {
    disegnaConteggio();
    disegnaFiltri();
    disegnaLista();
    disegnaBarraScelta();
    disegnaTagModulo();
    disegnaGestione();
    disegnaStatoSalvataggio();
  }

  /* ---------------------------------------------------------
     Richiesta di una serata (inoltrata a Telegram dal relay)
     --------------------------------------------------------- */

  function leggiScelti() {
    try {
      var grezzo = sessionStorage.getItem(CHIAVE_SCELTI);
      var elenco = grezzo ? JSON.parse(grezzo) : [];
      return Array.isArray(elenco) ? elenco : [];
    } catch (e) { return []; }
  }

  function scriviScelti() {
    try { sessionStorage.setItem(CHIAVE_SCELTI, JSON.stringify(stato.scelti)); } catch (e) { /* ignora */ }
  }

  function alternaScelta(nome) {
    var i = stato.scelti.indexOf(nome);
    if (i >= 0) stato.scelti.splice(i, 1); else stato.scelti.push(nome);
    scriviScelti();
    disegnaLista();
    disegnaBarraScelta();
    if (!el.modale.hidden) {
      if (stato.scelti.length) disegnaScelti(); else chiudiRichiesta();
    }
  }

  function disegnaBarraScelta() {
    var quanti = stato.scelti.length;
    var visibile = quanti > 0 && !el.vistaCollezione.hidden;
    el.barraScelta.hidden = !visibile;
    document.body.classList.toggle('con-barra', visibile);
    el.sceltaTesto.textContent = quanti === 1 ? '1 gioco scelto' : quanti + ' giochi scelti';
  }

  function disegnaScelti() {
    el.elencoScelti.textContent = '';
    stato.scelti.forEach(function (nome) {
      var voce = creaElemento('button', 'attivo');
      voce.type = 'button';
      voce.appendChild(creaElemento('span', null, nome));
      voce.appendChild(creaElemento('span', 'attivo__x', '×'));
      voce.setAttribute('aria-label', 'Togli ' + nome + ' dalla richiesta');
      voce.addEventListener('click', function () { alternaScelta(nome); });
      el.elencoScelti.appendChild(voce);
    });
  }

  function oggiISO() {
    var adesso = new Date();
    return new Date(adesso.getTime() - adesso.getTimezoneOffset() * 60000)
      .toISOString().slice(0, 10);
  }

  function notaRichiesta(testo, errore) {
    el.messaggioRichiesta.hidden = false;
    el.messaggioRichiesta.className = 'messaggio' + (errore ? ' messaggio--errore' : '');
    el.messaggioRichiesta.textContent = testo;
  }

  function apriRichiesta() {
    if (!stato.scelti.length) return;
    chiudiTendine();
    el.messaggioRichiesta.hidden = true;
    el.rNome.value = leggiLocale(CHIAVE_NOME) || '';
    el.rGiorno.min = oggiISO();
    if (!el.rGiorno.value) el.rGiorno.value = oggiISO();
    disegnaScelti();
    el.modale.hidden = false;
    document.body.style.overflow = 'hidden';
    setTimeout(function () { (el.rNome.value ? el.rGiorno : el.rNome).focus(); }, 50);
  }

  function chiudiRichiesta() {
    el.modale.hidden = true;
    document.body.style.overflow = '';
  }

  function inviaRichiesta(ev) {
    ev.preventDefault();
    var nome = el.rNome.value.trim();
    var giorno = el.rGiorno.value;
    if (!nome || !giorno || !stato.scelti.length) return;

    if (!RELAY_TELEGRAM) {
      notaRichiesta('L\'invio su Telegram non è ancora configurato: manca l\'indirizzo '
        + 'del relay in js/app.js. Vedi telegram/LEGGIMI.md.', true);
      return;
    }

    scriviLocale(CHIAVE_NOME, nome);
    el.btnInviaRichiesta.disabled = true;
    notaRichiesta('Invio in corso…');

    /* text/plain evita la richiesta di preflight: al relay arriva comunque JSON. */
    fetch(RELAY_TELEGRAM, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
      body: JSON.stringify({
        nome: nome,
        giorno: giorno,
        nota: el.rNota.value.trim(),
        giochi: stato.scelti.slice(),
        inviatoIl: new Date().toISOString()
      })
    })
      .then(function (r) {
        return r.json().catch(function () { return {}; }).then(function (corpo) {
          if (!r.ok) throw new Error(corpo.errore || ('il relay ha risposto ' + r.status));
          return corpo;
        });
      })
      .then(function () {
        chiudiRichiesta();
        stato.scelti = [];
        scriviScelti();
        el.rNota.value = '';
        disegnaLista();
        disegnaBarraScelta();
        avvisa('Richiesta inviata su Telegram.');
      })
      .catch(function (errore) {
        notaRichiesta('Non sono riuscito a inviare la richiesta: ' + errore.message
          + '. Riprova fra poco.', true);
      })
      .then(function () { el.btnInviaRichiesta.disabled = false; });
  }

  /* ---------------------------------------------------------
     Area di gestione
     --------------------------------------------------------- */

  function disegnaGestione() {
    if (!el.gestioneLista) return;

    var ricerca = senzaAccenti(el.gCerca.value || '');
    el.gestioneLista.textContent = '';

    var selezionato = el.eBase.value;
    el.eBase.textContent = '';
    var vuota = document.createElement('option');
    vuota.value = '';
    vuota.textContent = '— scegli un gioco —';
    el.eBase.appendChild(vuota);
    ordinaPerNome(stato.giochi).forEach(function (g) {
      var opzione = document.createElement('option');
      opzione.value = g.nome;
      opzione.textContent = g.nome;
      el.eBase.appendChild(opzione);
    });
    el.eBase.value = selezionato;

    ordinaPerNome(stato.giochi).forEach(function (g) {
      var esp = espansioniDi(g);
      if (ricerca) {
        var trovato = senzaAccenti(g.nome).indexOf(ricerca) >= 0
          || esp.some(function (e) { return senzaAccenti(e.nome).indexOf(ricerca) >= 0; });
        if (!trovato) return;
      }

      var voce = creaElemento('div', 'voce');
      var testo = creaElemento('div', 'voce__testo');
      testo.appendChild(creaElemento('div', 'voce__nome', g.nome));
      var meta = [mostraGiocatori(g) + ' giocatori', mostraDurata(g)];
      if (g.tag.length) meta.push(g.tag.join(', '));
      if (g.note) meta.push(g.note);
      testo.appendChild(creaElemento('div', 'voce__meta', meta.join(' · ')));
      voce.appendChild(testo);

      var azioni = creaElemento('div', 'voce__azioni');
      var modifica = creaElemento('button', 'voce__azione', 'Modifica');
      modifica.type = 'button';
      modifica.addEventListener('click', function () { caricaGiocoNelModulo(g); });
      var elimina = creaElemento('button', 'voce__azione voce__azione--pericolo', 'Elimina');
      elimina.type = 'button';
      elimina.addEventListener('click', function () { eliminaGioco(g); });
      azioni.appendChild(modifica);
      azioni.appendChild(elimina);
      voce.appendChild(azioni);
      el.gestioneLista.appendChild(voce);

      esp.forEach(function (e) {
        var vocEsp = creaElemento('div', 'voce voce--espansione');
        var testoEsp = creaElemento('div', 'voce__testo');
        testoEsp.appendChild(creaElemento('div', 'voce__nome', '↳ ' + e.nome));
        if (e.note) testoEsp.appendChild(creaElemento('div', 'voce__meta', e.note));
        vocEsp.appendChild(testoEsp);

        var azioniEsp = creaElemento('div', 'voce__azioni');
        var modificaEsp = creaElemento('button', 'voce__azione', 'Modifica');
        modificaEsp.type = 'button';
        modificaEsp.addEventListener('click', function () { caricaEspansioneNelModulo(e); });
        var eliminaEsp = creaElemento('button', 'voce__azione voce__azione--pericolo', 'Elimina');
        eliminaEsp.type = 'button';
        eliminaEsp.addEventListener('click', function () { eliminaEspansione(e); });
        azioniEsp.appendChild(modificaEsp);
        azioniEsp.appendChild(eliminaEsp);
        vocEsp.appendChild(azioniEsp);
        el.gestioneLista.appendChild(vocEsp);
      });
    });

    if (!el.gestioneLista.childNodes.length) {
      el.gestioneLista.appendChild(creaElemento('div', 'voce', 'Nessuna voce trovata.'));
    }
  }

  function caricaGiocoNelModulo(g) {
    el.gOriginale.value = g.nome;
    el.gNome.value = g.nome;
    el.gMin.value = g.min || '';
    el.gMax.value = g.max === 99 ? '' : (g.max || '');
    el.gDmin.value = g.dmin || '';
    el.gDmax.value = g.dmax || '';
    el.gTag.value = g.tag.join(', ');
    el.gDescrizione.value = g.descrizione || '';
    el.gNote.value = g.note || '';
    el.btnSalvaGioco.textContent = 'Salva modifiche';
    el.btnAnnullaGioco.hidden = false;
    el.gNome.focus();
    el.gNome.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function azzeraModuloGioco() {
    el.formGioco.reset();
    el.gOriginale.value = '';
    el.btnSalvaGioco.textContent = 'Aggiungi gioco';
    el.btnAnnullaGioco.hidden = true;
  }

  function caricaEspansioneNelModulo(e) {
    el.eOriginaleBase.value = e.base;
    el.eOriginaleNome.value = e.nome;
    el.eBase.value = e.base;
    el.eNome.value = e.nome;
    el.eNote.value = e.note || '';
    el.btnSalvaEspansione.textContent = 'Salva modifiche';
    el.btnAnnullaEspansione.hidden = false;
    el.eNome.focus();
    el.eNome.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function azzeraModuloEspansione() {
    el.formEspansione.reset();
    el.eOriginaleBase.value = '';
    el.eOriginaleNome.value = '';
    el.btnSalvaEspansione.textContent = 'Aggiungi espansione';
    el.btnAnnullaEspansione.hidden = true;
  }

  function salvaGioco(ev) {
    ev.preventDefault();
    var nome = el.gNome.value.trim();
    if (!nome) return;

    var originale = el.gOriginale.value;
    var duplicato = stato.giochi.some(function (g) {
      return chiave(g.nome) === chiave(nome) && chiave(g.nome) !== chiave(originale);
    });
    if (duplicato) {
      avvisa('Esiste già un gioco con questo nome.');
      return;
    }

    var dati = {
      nome: nome,
      min: numero(el.gMin.value),
      max: numero(el.gMax.value),
      dmin: numero(el.gDmin.value),
      dmax: numero(el.gDmax.value),
      tag: separaTag(el.gTag.value),
      descrizione: el.gDescrizione.value.trim(),
      note: el.gNote.value.trim()
    };
    if (dati.dmin && !dati.dmax) dati.dmax = dati.dmin;
    if (dati.dmax && !dati.dmin) dati.dmin = dati.dmax;

    if (originale) {
      var indice = stato.giochi.findIndex(function (g) { return chiave(g.nome) === chiave(originale); });
      if (indice >= 0) stato.giochi[indice] = dati;
      if (chiave(originale) !== chiave(nome)) {
        stato.espansioni.forEach(function (e) {
          if (chiave(e.base) === chiave(originale)) e.base = nome;
        });
      }
      avvisa('Gioco aggiornato.');
    } else {
      stato.giochi.push(dati);
      avvisa('Gioco aggiunto.');
    }

    azzeraModuloGioco();
    segnaModifica();
  }

  function eliminaGioco(g) {
    var esp = espansioniDi(g);
    var messaggio = 'Eliminare "' + g.nome + '"?'
      + (esp.length ? '\nVerranno eliminate anche ' + esp.length + ' espansioni.' : '');
    if (!window.confirm(messaggio)) return;
    stato.giochi = stato.giochi.filter(function (x) { return chiave(x.nome) !== chiave(g.nome); });
    stato.espansioni = stato.espansioni.filter(function (e) { return chiave(e.base) !== chiave(g.nome); });
    avvisa('Gioco eliminato.');
    segnaModifica();
  }

  function salvaEspansione(ev) {
    ev.preventDefault();
    var base = el.eBase.value;
    var nome = el.eNome.value.trim();
    if (!base || !nome) return;

    var originaleBase = el.eOriginaleBase.value;
    var originaleNome = el.eOriginaleNome.value;

    var duplicato = stato.espansioni.some(function (e) {
      return chiave(e.base) === chiave(base) && chiave(e.nome) === chiave(nome)
        && !(chiave(e.base) === chiave(originaleBase) && chiave(e.nome) === chiave(originaleNome));
    });
    if (duplicato) {
      avvisa('Questa espansione è già presente.');
      return;
    }

    var dati = { base: base, nome: nome, note: el.eNote.value.trim() };

    if (originaleNome) {
      var indice = stato.espansioni.findIndex(function (e) {
        return chiave(e.base) === chiave(originaleBase) && chiave(e.nome) === chiave(originaleNome);
      });
      if (indice >= 0) stato.espansioni[indice] = dati;
      avvisa('Espansione aggiornata.');
    } else {
      stato.espansioni.push(dati);
      avvisa('Espansione aggiunta.');
    }

    stato.aperti[chiave(base)] = true;
    azzeraModuloEspansione();
    segnaModifica();
  }

  function eliminaEspansione(e) {
    if (!window.confirm('Eliminare l\'espansione "' + e.nome + '"?')) return;
    stato.espansioni = stato.espansioni.filter(function (x) {
      return !(chiave(x.base) === chiave(e.base) && chiave(x.nome) === chiave(e.nome));
    });
    avvisa('Espansione eliminata.');
    segnaModifica();
  }

  /* ---------------------------------------------------------
     Salvataggio del file
     --------------------------------------------------------- */

  function disegnaStatoSalvataggio() {
    if (!el.statoModifiche) return;
    if (stato.bozza) {
      el.statoModifiche.className = 'stato stato--modificato';
      el.statoModifiche.textContent = 'Ci sono modifiche salvate solo su questo browser: pubblicale per aggiornare il file Excel online.';
    } else {
      el.statoModifiche.className = 'stato';
      el.statoModifiche.textContent = 'Nessuna modifica in sospeso: il file online è aggiornato.';
    }
  }

  function messaggio(testo, errore) {
    el.messaggioSalvataggio.hidden = false;
    el.messaggioSalvataggio.className = 'messaggio' + (errore ? ' messaggio--errore' : '');
    el.messaggioSalvataggio.textContent = testo;
  }

  function intestazioniGitHub(token) {
    return {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json'
    };
  }

  function salvaSuGitHub() {
    var token = el.ghToken.value.trim();
    if (!token) {
      messaggio('Inserisci un token GitHub con permesso di scrittura, oppure scarica il file e caricalo a mano.', true);
      return;
    }
    scriviLocale(CHIAVE_TOKEN, token);

    var indirizzo = 'https://api.github.com/repos/' + REPO + '/contents/' + PERCORSO_EXCEL;
    var intestazioni = intestazioniGitHub(token);
    el.btnPubblica.disabled = true;
    messaggio('Salvataggio in corso…');

    fetch(indirizzo + '?ref=' + RAMO, { headers: intestazioni, cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (meta) {
        var corpo = {
          message: 'Aggiorna la collezione (' + stato.giochi.length + ' giochi, ' + stato.espansioni.length + ' espansioni)',
          content: XLSX.write(aWorkbook(), { type: 'base64', bookType: 'xlsx' }),
          branch: RAMO
        };
        if (meta && meta.sha) corpo.sha = meta.sha;
        return fetch(indirizzo, { method: 'PUT', headers: intestazioni, body: JSON.stringify(corpo) });
      })
      .then(function (r) {
        return r.json().then(function (corpo) {
          if (!r.ok) throw new Error(corpo && corpo.message ? corpo.message : 'Errore ' + r.status);
          return corpo;
        });
      })
      .then(function () {
        stato.versioneFile = stato.versione;
        messaggio('Salvato. Il file Excel online è aggiornato; il sito pubblico si allinea entro un paio di minuti.');
        avvisa('Modifiche pubblicate.');
        disegnaTutto();
      })
      .catch(function (errore) {
        messaggio('Salvataggio non riuscito: ' + errore.message, true);
      })
      .then(function () { el.btnPubblica.disabled = false; });
  }

  function scaricaExcel() {
    XLSX.writeFile(aWorkbook(), 'giochi.xlsx');
    messaggio('File scaricato: sostituisci dati/giochi.xlsx nel repository per pubblicarlo.');
  }

  function cambiaPassword() {
    var nuova = window.prompt('Nuova password (minimo 4 caratteri):');
    if (nuova === null) return;
    nuova = nuova.trim();
    if (nuova.length < 4) { avvisa('Password troppo corta.'); return; }
    sha256(nuova).then(function (hash) {
      stato.passwordHash = hash;
      segnaModifica();
      avvisa('Password aggiornata: ricordati di pubblicare le modifiche.');
    }).catch(function (errore) { avvisa(errore.message); });
  }

  /* ---------------------------------------------------------
     Navigazione fra le viste
     --------------------------------------------------------- */

  function mostraVista(quale) {
    var gestione = quale === 'gestione';
    chiudiTendine();
    el.vistaCollezione.hidden = gestione;
    el.vistaGestione.hidden = !gestione;
    el.btnGestione.textContent = gestione ? 'Collezione' : 'Gestione';
    disegnaBarraScelta();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function sbloccato() {
    try { return sessionStorage.getItem(CHIAVE_SBLOCCO) === '1'; } catch (e) { return false; }
  }

  function aggiornaSblocco() {
    var ok = sbloccato();
    el.pannelloSblocco.hidden = ok;
    el.pannelloGestione.hidden = !ok;
    if (ok) {
      el.ghToken.value = leggiLocale(CHIAVE_TOKEN) || '';
      disegnaGestione();
      disegnaStatoSalvataggio();
    }
  }

  function tentaSblocco(ev) {
    ev.preventDefault();
    var inserita = el.campoPassword.value;
    sha256(inserita).then(function (hash) {
      if (stato.passwordHash && hash === stato.passwordHash) {
        try { sessionStorage.setItem(CHIAVE_SBLOCCO, '1'); } catch (e) { /* ignora */ }
        el.erroreSblocco.hidden = true;
        el.campoPassword.value = '';
        aggiornaSblocco();
      } else {
        el.erroreSblocco.hidden = false;
        el.erroreSblocco.textContent = 'Password errata.';
      }
    }).catch(function (errore) {
      el.erroreSblocco.hidden = false;
      el.erroreSblocco.textContent = errore.message;
    });
  }

  /* ---------------------------------------------------------
     Tema — scuro se non è stato scelto altro
     --------------------------------------------------------- */

  var ICONA_SOLE = '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round">'
    + '<circle cx="12" cy="12" r="4.2"/><path d="M12 2.6v2.2M12 19.2v2.2M2.6 12h2.2M19.2 12h2.2M5.4 5.4l1.6 1.6M17 17l1.6 1.6M18.6 5.4L17 7M7 17l-1.6 1.6"/></svg>';
  var ICONA_LUNA = '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">'
    + '<path d="M20 14.4A8.4 8.4 0 0 1 9.6 4a8.4 8.4 0 1 0 10.4 10.4z"/></svg>';

  function applicaTema(tema) {
    var scuro = tema !== 'light';
    document.documentElement.setAttribute('data-tema', scuro ? 'dark' : 'light');
    el.iconaTema.innerHTML = scuro ? ICONA_SOLE : ICONA_LUNA;
    el.btnTema.setAttribute('aria-label', scuro ? 'Passa al tema chiaro' : 'Passa al tema scuro');
    var colore = document.querySelector('meta[name="theme-color"]');
    if (colore) colore.setAttribute('content', scuro ? '#111113' : '#f6f6f4');
  }

  function alternaTema() {
    var nuovo = document.documentElement.getAttribute('data-tema') === 'light' ? 'dark' : 'light';
    scriviLocale(CHIAVE_TEMA, nuovo);
    applicaTema(nuovo);
  }

  /* ---------------------------------------------------------
     Avvio
     --------------------------------------------------------- */

  function preparaFiltri() {
    el.testo.addEventListener('input', function () {
      filtri.testo = el.testo.value;
      disegnaChipAttivi();
      disegnaLista();
    });

    // I due campi restano come li scrive l'utente: se li inverte, il limite
    // inferiore e quello superiore vengono messi in ordine solo internamente.
    var leggiIntervallo = function () {
      var da = numero(el.durataDa.value);
      var a = numero(el.durataA.value);
      if (da !== null && a !== null && da > a) { var scambio = da; da = a; a = scambio; }
      filtri.durataMin = da;
      filtri.durataMax = a;
      disegnaFiltri(true);
      disegnaLista();
    };
    [el.durataDa, el.durataA].forEach(function (campo) {
      campo.addEventListener('input', leggiIntervallo);
      campo.addEventListener('change', leggiIntervallo);
    });

    el.cercaTag.addEventListener('input', function () {
      cercaTag = el.cercaTag.value;
      disegnaListaTag();
    });

    Array.prototype.forEach.call(document.querySelectorAll('[data-modo-tag]'), function (b) {
      b.addEventListener('click', function () {
        filtri.tuttiITag = b.getAttribute('data-modo-tag') === 'tutti';
        disegnaFiltri();
        disegnaLista();
      });
    });

    el.btnSoloEspansioni.addEventListener('click', function () {
      filtri.soloEspansioni = !filtri.soloEspansioni;
      disegnaFiltri();
      disegnaLista();
    });

    Array.prototype.forEach.call(document.querySelectorAll('[data-azzera]'), function (b) {
      b.addEventListener('click', function () {
        var quale = b.getAttribute('data-azzera');
        if (quale === 'giocatori') filtri.giocatori = null;
        if (quale === 'durata') { filtri.durataMin = null; filtri.durataMax = null; }
        if (quale === 'tag') filtri.tag = [];
        disegnaFiltri();
        disegnaLista();
      });
    });
  }

  function avvia() {
    el = {
      conteggio: $('conteggio-testata'),
      avviso: $('avviso'),
      velo: $('velo'),
      btnTema: $('btn-tema'),
      iconaTema: $('icona-tema'),
      btnGestione: $('btn-gestione'),
      vistaCollezione: $('vista-collezione'),
      vistaGestione: $('vista-gestione'),
      testo: $('f-testo'),
      valGiocatori: $('val-giocatori'),
      valDurata: $('val-durata'),
      valTag: $('val-tag'),
      valOrdine: $('val-ordine'),
      numeriGiocatori: $('numeri-giocatori'),
      presetDurata: $('preset-durata'),
      durataDa: $('d-da'),
      durataA: $('d-a'),
      cercaTag: $('cerca-tag'),
      listaTag: $('lista-tag'),
      listaOrdine: $('lista-ordine'),
      btnSoloEspansioni: $('btn-solo-espansioni'),
      attivi: $('attivi'),
      risultati: $('risultati'),
      lista: $('lista'),
      righe: $('righe'),
      vuoto: $('vuoto'),
      barraScelta: $('barra-scelta'),
      sceltaTesto: $('scelta-testo'),
      modale: $('modale-richiesta'),
      elencoScelti: $('elenco-scelti'),
      rNome: $('r-nome'),
      rGiorno: $('r-giorno'),
      rNota: $('r-nota'),
      btnInviaRichiesta: $('btn-invia-richiesta'),
      messaggioRichiesta: $('messaggio-richiesta'),
      pannelloSblocco: $('pannello-sblocco'),
      pannelloGestione: $('pannello-gestione'),
      formSblocco: $('form-sblocco'),
      campoPassword: $('campo-password'),
      erroreSblocco: $('errore-sblocco'),
      formGioco: $('form-gioco'),
      gOriginale: $('g-originale'),
      gNome: $('g-nome'),
      gMin: $('g-min'),
      gMax: $('g-max'),
      gDmin: $('g-dmin'),
      gDmax: $('g-dmax'),
      gTag: $('g-tag'),
      gDescrizione: $('g-descrizione'),
      gNote: $('g-note'),
      elencoTag: $('elenco-tag'),
      tagSuggeriti: $('tag-suggeriti'),
      btnSalvaGioco: $('btn-salva-gioco'),
      btnAnnullaGioco: $('btn-annulla-gioco'),
      formEspansione: $('form-espansione'),
      eOriginaleBase: $('e-originale-base'),
      eOriginaleNome: $('e-originale-nome'),
      eBase: $('e-base'),
      eNome: $('e-nome'),
      eNote: $('e-note'),
      btnSalvaEspansione: $('btn-salva-espansione'),
      btnAnnullaEspansione: $('btn-annulla-espansione'),
      gCerca: $('g-cerca'),
      gestioneLista: $('gestione-lista'),
      statoModifiche: $('stato-modifiche'),
      ghToken: $('gh-token'),
      btnPubblica: $('btn-pubblica'),
      messaggioSalvataggio: $('messaggio-salvataggio')
    };

    applicaTema(leggiLocale(CHIAVE_TEMA) === 'light' ? 'light' : 'dark');
    el.btnTema.addEventListener('click', alternaTema);

    el.btnGestione.addEventListener('click', function () {
      mostraVista(el.vistaGestione.hidden ? 'gestione' : 'collezione');
      if (!el.vistaGestione.hidden) aggiornaSblocco();
    });

    el.formSblocco.addEventListener('submit', tentaSblocco);
    el.formGioco.addEventListener('submit', salvaGioco);
    el.formEspansione.addEventListener('submit', salvaEspansione);
    el.btnAnnullaGioco.addEventListener('click', azzeraModuloGioco);
    el.btnAnnullaEspansione.addEventListener('click', azzeraModuloEspansione);
    el.gCerca.addEventListener('input', disegnaGestione);
    el.btnPubblica.addEventListener('click', salvaSuGitHub);
    $('btn-scarica').addEventListener('click', scaricaExcel);
    $('btn-cambia-password').addEventListener('click', cambiaPassword);
    $('btn-dimentica-token').addEventListener('click', function () {
      el.ghToken.value = '';
      scriviLocale(CHIAVE_TOKEN, null);
      avvisa('Token rimosso da questo dispositivo.');
    });

    stato.scelti = leggiScelti();

    $('btn-apri-richiesta').addEventListener('click', apriRichiesta);
    $('btn-chiudi-richiesta').addEventListener('click', chiudiRichiesta);
    $('btn-annulla-richiesta').addEventListener('click', chiudiRichiesta);
    $('form-richiesta').addEventListener('submit', inviaRichiesta);
    $('btn-svuota-scelta').addEventListener('click', function () {
      stato.scelti = [];
      scriviScelti();
      disegnaLista();
      disegnaBarraScelta();
    });
    el.modale.addEventListener('click', function (ev) {
      if (ev.target === el.modale) chiudiRichiesta();
    });
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape' && !el.modale.hidden) chiudiRichiesta();
    });

    preparaTendine();
    preparaFiltri();

    caricaDati()
      .then(function () {
        disegnaTutto();
        if (stato.bozza) avvisa('Ci sono modifiche non ancora pubblicate.');
      })
      .catch(function (errore) {
        el.vuoto.hidden = false;
        el.vuoto.textContent = errore.message;
        el.lista.hidden = true;
      });
  }

  document.addEventListener('DOMContentLoaded', avvia);
})();
