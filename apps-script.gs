// Google Apps Script für das Bewertungs-Sheet.
// Einrichtung: siehe README.md. Diesen Code komplett in Code.gs einfügen.

const SHEET_NAME = 'Bewertungen';
const HEADER = ['id', 'zeit', 'geschmack', 'schaerfe', 'bestellen', 'vollstaendig'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Neue Bewertung vom iPad. Doppelte ids (erneuter Versuch nach Funkloch) werden ignoriert.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const r = JSON.parse(e.postData.contents);
    if (!r.id) return json_({ ok: false, error: 'id fehlt' });
    const sh = sheet_();
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (x) { return x[0]; }) : [];
    if (ids.indexOf(r.id) === -1) {
      sh.appendRow([
        r.id,
        r.ts ? new Date(r.ts) : new Date(),
        r.taste || '',
        r.spice || '',
        r.order || '',
        r.complete ? 'ja' : 'nein'
      ]);
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Alle Bewertungen für die Auswertung auf dem iPad.
function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, HEADER.length).getValues() : [];
  return json_({
    ok: true,
    rows: rows.map(function (x) {
      return {
        id: x[0],
        ts: x[1],
        taste: x[2] === '' ? null : Number(x[2]),
        spice: x[3] || null,
        order: x[4] || null,
        complete: x[5] === 'ja'
      };
    })
  });
}
