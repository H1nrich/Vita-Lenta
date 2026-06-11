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

const DEFAULT_Q_COUNT = 28;


/**
 * ============================
 *  FUNZIONI DI UTILITÀ
 * ============================
 */

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Restituisce la lista dei voter che hanno già votato.
 */
function getVotersWhoVoted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // salta header
  const voterCol = 1;         // colonna "voter" (indice 1)

  const voted = new Set();
  rows.forEach(r => {
    const v = (r[voterCol] || '').toString().trim();
    if (v) voted.add(v);
  });

  return Array.from(voted);
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
      return jsonResponse({ ok: false, error: 'Nessun payload ricevuto' });
    }

    const payload = JSON.parse(e.postData.contents);
    const voter = (payload.voter || '').toString().trim();
    const votes = payload.votes || {};

    // Autorizzazione voter
    if (!voter || !AUTHORIZED.some(n => n.toLowerCase() === voter.toLowerCase())) {
      return jsonResponse({ ok: false, error: 'Voter non autorizzato' });
    }

    // Controllo: ha già votato?
    const alreadyVoted = getVotersWhoVoted();
    if (alreadyVoted.some(v => v.toLowerCase() === voter.toLowerCase())) {
      return jsonResponse({ ok: false, error: 'Hai già votato' });
    }

    // Controllo auto-voto
    for (const k in votes) {
      const val = (votes[k] || '').toString().trim();
      if (!val) continue;
      if (val.toLowerCase() === voter.toLowerCase()) {
        return jsonResponse({ ok: false, error: 'Auto-voto non consentito' });
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    // Header dinamico
    const header = ['timestamp', 'voter'].concat(Object.keys(votes));
    if (sheet.getLastRow() === 0) sheet.appendRow(header);

    // Riga dati
    const row = [new Date(), voter];
    for (const k of Object.keys(votes)) row.push(votes[k] || '');
    sheet.appendRow(row);

    return jsonResponse({ ok: true });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
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

    // Restituisce anche la lista di chi ha già votato
    if (action === 'voted') {
      return jsonResponse({ ok: true, voted: getVotersWhoVoted() });
    }

    return jsonResponse({ ok: true, message: 'Web App attivo' });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}


/**
 * ============================
 *  RESET VOTI
 *  → Esegui manualmente da Apps Script
 * ============================
 */

/**
 * Cancella tutti i voti dal foglio (lascia l'header intatto).
 * Per eseguire: apri Apps Script → seleziona questa funzione → clicca ▶ Esegui
 */
function resetVoti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('Foglio "' + SHEET_NAME + '" non trovato. Niente da resettare.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log('Nessun voto da cancellare.');
    return;
  }

  // Cancella tutte le righe dati (dalla 2 in poi), mantiene l'header
  sheet.deleteRows(2, lastRow - 1);
  Logger.log('Reset completato: cancellate ' + (lastRow - 1) + ' righe.');
}


/**
 * ============================
 *  LOGICA DI CALCOLO RISULTATI
 * ============================
 */

function getAggregatedResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const emptyResult = (count) => {
    const res = {};
    for (let i = 0; i < count; i++) {
      res['q' + i] = AUTHORIZED.map(n => ({ name: n, count: 0, percent: 0 }));
    }
    return res;
  };

  if (!sheet) return emptyResult(DEFAULT_Q_COUNT);

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) return emptyResult(DEFAULT_Q_COUNT);
  if (data.length <= 1) {
    const qCount = Math.max((data[0] || []).length - 2, 0);
    return emptyResult(qCount || DEFAULT_Q_COUNT);
  }

  const header = data[0];
  const rows = data.slice(1);
  const qCount = Math.max(header.length - 2, 0);
  const results = {};

  for (let qi = 0; qi < qCount; qi++) {
    const counts = {};
    AUTHORIZED.forEach(n => counts[n] = 0);

    rows.forEach(r => {
      const raw = r[2 + qi];
      const val = (raw === undefined || raw === null) ? '' : raw.toString().trim();
      if (!val) return;
      const matched = AUTHORIZED.find(a => a.toLowerCase() === val.toLowerCase());
      if (matched) counts[matched] = (counts[matched] || 0) + 1;
    });

    const totalCount = Object.values(counts).reduce((a, b) => a + b, 0);
    const arr = AUTHORIZED.map(name => ({
      name,
      count: counts[name] || 0,
      percent: totalCount > 0 ? (counts[name] / totalCount * 100) : 0
    }));

    arr.sort((a, b) => b.count - a.count);
    results['q' + qi] = arr;
  }

  return results;
}
