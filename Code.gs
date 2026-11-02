const MAX_SPOTS = 100;      // nombre de places
const SHEET_NAME = 'Reservations';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['Place', 'Prénom', 'Nom', 'CIN', 'Date']); }
  return sh;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safe_(t) { t = String(t || '').trim().slice(0, 40); return /^[=+\-@]/.test(t) ? "'" + t : t; }

// Le site demande juste le nombre de places prises (aucune donnée personnelle)
function doGet() { return json_({ taken: Math.max(0, sheet_().getLastRow() - 1), max: MAX_SPOTS }); }

// Le site envoie une réservation
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fn = safe_(d.fn), ln = safe_(d.ln);
    const cin = String(d.cin || '').toUpperCase().replace(/\s+/g, '');
    if (!fn || !ln || !/^[A-Z]{1,2}[0-9]{3,8}$/.test(cin)) return json_({ ok: false, error: 'invalid' });
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n >= MAX_SPOTS) return json_({ ok: false, error: 'full' });
    if (n > 0) {
      const cins = sh.getRange(2, 4, n, 1).getValues().flat().map(String);
      if (cins.indexOf(cin) !== -1) return json_({ ok: false, error: 'duplicate' });
    }
    const place = n + 1;
    sh.appendRow([place, fn, ln, cin, new Date()]);
    return json_({ ok: true, place: place, taken: place });
  } finally {
    lock.releaseLock();
  }
}
