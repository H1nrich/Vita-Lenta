const SHEET_NAME = 'Voti';
const AUTHORIZED = [
  "Gaetano Orefice","Daniela Parascandolo","Aurora Orefice","Enrico Scolastico",
  "Gianpiero Fasulo","Federica Devastato","Eleonora Galdi","Fabio Vassallo",
  "Fabio Sivero","Luigi De Simone","Chiara Vassallo"
];

function renderResults(data){
  const resultsDiv = document.getElementById('results');
  resultsDiv.innerHTML = '';
  // Filtra le chiavi che eccedono il numero di domande nel frontend
  const maxIndex = questions.length - 1; // se questions ha 28 elementi, maxIndex = 27
  Object.keys(data).filter(k=>{
    const qi = parseInt(k.replace('q',''));
    return !isNaN(qi) && qi <= maxIndex;
  }).forEach(qk=>{
    const qi = parseInt(qk.replace('q',''));
    // ... resto del codice esistente per creare la card ...
  });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const voter = (payload.voter || '').toString().trim();
    const votes = payload.votes || {};
    if (!voter || AUTHORIZED.map(n=>n.toLowerCase()).indexOf(voter.toLowerCase()) === -1) {
      return jsonResponse({ ok:false, error: 'Voter non autorizzato' });
    }
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
    const header = ['timestamp','voter'].concat(Object.keys(votes));
    if (sheet.getLastRow() === 0) sheet.appendRow(header);
    const row = [new Date(), voter];
    for (const k of Object.keys(votes)) row.push(votes[k] || '');
    sheet.appendRow(row);
    return jsonResponse({ ok:true });
  } catch (err) {
    return jsonResponse({ ok:false, error: err.message });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'results') {
    return jsonResponse(getAggregatedResults());
  }
  return jsonResponse({ ok:true, message:'Web App attivo' });
}

function getAggregatedResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  // Se non esiste il foglio, assumiamo un numero di domande di default (modifica se vuoi)
  const DEFAULT_Q_COUNT = 28; // metti 29 se vuoi 29 domande
  if (!sheet) {
    const res = {};
    const qCount = DEFAULT_Q_COUNT;
    for (let i=0;i<qCount;i++) res['q'+i] = AUTHORIZED.map(n=>({ name:n, count:0, percent:0 }));
    return res;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    const header = data[0] || [];
    const qCount = Math.max((header.length - 2), 0);
    const res = {};
    for (let i=0;i<qCount;i++) res['q'+i] = AUTHORIZED.map(n=>({ name:n, count:0, percent:0 }));
    return res;
  }

  const header = data[0];
  const rows = data.slice(1);
  const qCount = Math.max(header.length - 2, 0); // numero di colonne domanda dedotto dall'header
  const results = {};
  for (let qi=0; qi<qCount; qi++) {
    const counts = {};
    AUTHORIZED.forEach(n=>counts[n]=0);
    rows.forEach(r=>{
      const val = r[2+qi];
      if (val && AUTHORIZED.includes(val)) counts[val]++;
    });
    const total = Object.values(counts).reduce((a,b)=>a+b,0);
    const arr = AUTHORIZED.map(name=>({
      name,
      count: counts[name] || 0,
      percent: total>0 ? (counts[name]/total*100) : 0
    }));
    arr.sort((a,b)=>b.count - a.count);
    results['q'+qi] = arr;
  }
  return results;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
