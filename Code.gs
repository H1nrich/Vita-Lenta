/**
 * ============================
 *  CONFIGURAZIONE & COSTANTI
 * ============================
 */

const SHEET_NAME = 'Voti';

const AUTHORIZED = [
  "Gaetano Orefice","Daniela Parascandolo","Aurora Orefice","Enrico Scolastico",
  "Gianpiero Fasulo","Federica Devastato","Eleonora Galdi","Fabio Vassallo",
  "Fabio Sivero","Luigi De Simone","Chiara Vassallo"
];

const DEFAULT_Q_COUNT = 28; // Cambia qui se vuoi 29 domande di default


/**
 * ============================
 *  FUNZIONI DI UTILITÀ
 * ============================
 */

/**
 * Risposta JSON standardizzata.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * ============================
 *  FUNZIONI PRINCIPALI (API)
 * ============================
 */

/**
 * Gestisce le richieste POST (ricezione voto).
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok:false, error: 'Nessun payload ricevuto' });
    }

    const payload = JSON.parse(e.postData.contents);
    const voter = (payload.voter || '').toString().trim();
    const votes = payload.votes || {};

    // Autorizzazione voter
    if (!voter || AUTHORIZED.map(n=>n.toLowerCase()).indexOf(voter.toLowerCase()) === -1) {
      return jsonResponse({ ok:false, error: 'Voter non autorizzato' });
    }

    // Controllo auto-voto
    for (const k in votes) {
      const val = (votes[k] || '').toString().trim();
      if (!val) continue;
      if (val.toLowerCase() === voter.toLowerCase()) {
        return jsonResponse({ ok:false, error: 'Auto-voto non consentito' });
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Header dinamico
    const header = ['timestamp','voter'].concat(Object.keys(votes));
    if (sheet.getLastRow() === 0) sheet.appendRow(header);

    // Riga dati
    const row = [new Date(), voter];
    for (const k of Object.keys(votes)) row.push(votes[k] || '');
    sheet.appendRow(row);

    return jsonResponse({ ok:true });

  } catch (err) {
    return jsonResponse({ ok:false, error: err.message });
  }
}


/**
 * Gestisce le richieste GET.
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;

    if (action === 'results') {
      return jsonResponse(getAggregatedResults());
    }

    return jsonResponse({ ok:true, message:'Web App attivo' });

  } catch (err) {
    return jsonResponse({ ok:false, error: err.message });
  }
}


/**
 * ============================
 *  LOGICA DI CALCOLO RISULTATI
 * ============================
 */

/**
 * Aggrega i risultati dal foglio.
 */
function getAggregatedResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Foglio inesistente → risultati vuoti
  if (!sheet) {
    const res = {};
    for (let i = 0; i < DEFAULT_Q_COUNT; i++) {
      res['q' + i] = AUTHORIZED.map(n => ({ name: n, count: 0, percent: 0 }));
    }
    return res;
  }

  const data = sheet.getDataRange().getValues();

  // Nessun dato
  if (!data || data.length === 0) {
    const res = {};
    for (let i = 0; i < DEFAULT_Q_COUNT; i++) {
      res['q' + i] = AUTHORIZED.map(n => ({ name: n, count: 0, percent: 0 }));
    }
    return res;
  }

  // Solo header → nessun voto
  if (data.length <= 1) {
    const header = data[0] || [];
    const qCount = Math.max((header.length - 2), 0);
    const res = {};
    for (let i = 0; i < qCount; i++) {
      res['q' + i] = AUTHORIZED.map(n => ({ name: n, count: 0, percent: 0 }));
    }
    return res;
  }

  const header = data[0];
  const rows = data.slice(1);
  const qCount = Math.max(header.length - 2, 0);
  const results = {};

  for (let qi = 0; qi < qCount; qi++) {
    const counts = {};
    AUTHORIZED.forEach(n => counts[n] = 0);

    rows.forEach(r => {
      const val = r[2 + qi];
      if (val && AUTHORIZED.includes(val)) counts[val]++;
    });

    const total = Object.values(counts).reduce((a, b) => a + b, 0);

    const arr = AUTHORIZED.map(name => ({
      name,
      count: counts[name] || 0,
      percent: total > 0 ? (counts[name] / total * 100) : 0
    }));

    arr.sort((a, b) => b.count - a.count);
    results['q' + qi] = arr;
  }

  return results;
}
