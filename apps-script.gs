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

// Einmal von Hand ausführen (oben Funktion „einrichten“ wählen → Ausführen):
// färbt das Blatt „Bewertungen“ ein und legt das Blatt „Auswertung“ mit Formeln an.
// Kann gefahrlos erneut ausgeführt werden, die Daten bleiben unberührt.
function einrichten() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet_();
  const GREEN = '#D5EFD8', LGREEN = '#E6F4D7', YELLOW = '#FBEFC9', ORANGE = '#FCDFC7', RED = '#F8D3D1';

  // --- Blatt „Bewertungen“: Kopf, Spaltenbreiten, Farben ---
  sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#F1F2F5');
  sh.getRange('B2:B').setNumberFormat('dd.MM. HH:mm');
  sh.getRange('C2:F').setHorizontalAlignment('center');
  sh.setColumnWidth(1, 170);
  [2, 3, 4, 5, 6].forEach(function (c) { sh.setColumnWidth(c, 120); });

  function txt(range, value, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(value).setBackground(color).setRanges([sh.getRange(range)]).build();
  }
  function num(range, value, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(value).setBackground(color).setRanges([sh.getRange(range)]).build();
  }
  sh.setConditionalFormatRules([
    num('C2:C', 5, GREEN), num('C2:C', 4, LGREEN), num('C2:C', 3, YELLOW), num('C2:C', 2, ORANGE), num('C2:C', 1, RED),
    txt('D2:D', 'genau richtig', GREEN), txt('D2:D', 'zu mild', YELLOW), txt('D2:D', 'zu scharf', RED),
    txt('E2:E', 'ja', GREEN), txt('E2:E', 'vielleicht', YELLOW), txt('E2:E', 'nein', RED),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('nein').setFontColor('#9AA0AC').setRanges([sh.getRange('F2:F')]).build()
  ]);

  // --- Blatt „Auswertung“ ---
  let a = ss.getSheetByName('Auswertung');
  if (a) a.clear(); else a = ss.insertSheet('Auswertung');
  a.setConditionalFormatRules([]);
  const B = 'Bewertungen!';
  // Deutsche (und die meisten anderen) Sheets trennen Formel-Argumente mit „;“, englische mit „,“.
  const S = /^en/.test(ss.getSpreadsheetLocale()) ? ',' : ';';

  a.getRange('A1').setValue('Çiğ Köfte · Auswertung').setFontSize(16).setFontWeight('bold');
  a.getRange('A3:B5').setValues([
    ['Antworten', ''],
    ['Ø Geschmack (1 bis 5)', ''],
    ['Würden bestellen (Ja)', '']
  ]);
  a.getRange('B3').setFormula('=COUNTA(' + B + 'A2:A)');
  a.getRange('B4').setFormula('=IFERROR(ROUND(AVERAGE(' + B + 'C2:C)' + S + '1)' + S + '"–")');
  a.getRange('B5').setFormula('=IFERROR(COUNTIF(' + B + 'E2:E' + S + '"ja")/COUNTA(' + B + 'E2:E)' + S + '"–")').setNumberFormat('0%');
  a.getRange('A3:A5').setFontWeight('bold');
  a.getRange('B3:B5').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('left');

  function block(row, title, col, items) {
    a.getRange(row, 1, 1, 4).setValues([[title, 'Anzahl', 'Anteil', '']]).setFontWeight('bold').setBackground('#F1F2F5');
    const first = row + 1, last = row + items.length;
    items.forEach(function (it, i) {
      const r = first + i;
      const crit = typeof it[1] === 'number' ? it[1] : '"' + it[1] + '"';
      a.getRange(r, 1).setValue(it[0]).setBackground(it[2]);
      a.getRange(r, 2).setFormula('=COUNTIF(' + B + col + '2:' + col + S + crit + ')');
      a.getRange(r, 3).setFormula('=IFERROR(B' + r + '/SUM(B$' + first + ':B$' + last + ')' + S + '0)').setNumberFormat('0%');
      // Balken aus Blockzeichen: 20 Zeichen = 100 %
      a.getRange(r, 4).setFormula('=REPT("█"' + S + 'ROUND(C' + r + '*20' + S + '0))').setFontColor(it[3]);
    });
  }
  block(7, 'Geschmack', 'C', [
    ['Hammer (5)', 5, GREEN, '#55B25E'],
    ['Lecker (4)', 4, LGREEN, '#8DC351'],
    ['Okay (3)', 3, YELLOW, '#EDB13A'],
    ['Geht so (2)', 2, ORANGE, '#EE8A3F'],
    ['Gar nicht (1)', 1, RED, '#E0574F']
  ]);
  block(14, 'Schärfe', 'D', [
    ['Genau richtig', 'genau richtig', GREEN, '#55B25E'],
    ['Zu mild', 'zu mild', YELLOW, '#EDB13A'],
    ['Zu scharf', 'zu scharf', RED, '#E0574F']
  ]);
  block(19, 'Würde bestellen', 'E', [
    ['Ja', 'ja', GREEN, '#55B25E'],
    ['Vielleicht', 'vielleicht', YELLOW, '#EDB13A'],
    ['Nein', 'nein', RED, '#E0574F']
  ]);
  a.setColumnWidth(1, 190);
  a.setColumnWidths(2, 2, 80);
  a.setColumnWidth(4, 260);
  a.setHiddenGridlines(true);

  ss.setActiveSheet(a);
  ss.moveActiveSheet(1);
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
