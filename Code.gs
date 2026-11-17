/**
 * Bump & Binge — the tiny backend that stores movie recs in this Google Sheet.
 *
 * Paste this whole file into Extensions → Apps Script of your Google Sheet,
 * then Deploy → New deployment → Web app:
 *   Execute as: Me
 *   Who has access: Anyone
 * Copy the Web app URL into SCRIPT_URL at the top of index.html.
 *
 * Rows live on a tab called "Recs". You can edit or delete rows by hand;
 * the page picks the change up on its next refresh.
 */

const SHEET_NAME = 'Recs';
const HEADERS = ['id', 'added', 'title', 'year', 'vibe', 'why', 'from', 'watched'];
const VIBES = ['cozy', 'laugh', 'cry', 'romance', 'family', 'edge', 'classic'];
const MAX_ROWS = 2000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('D:D').setNumberFormat('@'); // keep years as text
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Trim, cap length, and stop anything that would run as a spreadsheet formula.
function clean_(value, max) {
  let s = String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doGet() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  values.shift(); // headers
  const recs = values
    .filter(r => r[0] && r[2])
    .map(r => ({
      id: String(r[0]),
      added: r[1] instanceof Date ? r[1].getTime() : Number(r[1]) || 0,
      title: String(r[2]),
      year: String(r[3] || ''),
      vibe: String(r[4] || ''),
      why: String(r[5] || ''),
      from: String(r[6] || ''),
      watched: r[7] === true || String(r[7]).toUpperCase() === 'TRUE'
    }));
  return json_({ ok: true, recs: recs });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    const sh = getSheet_();

    if (body.action === 'add') {
      if (body.website) return json_({ ok: true }); // spam-bot honeypot
      const title = clean_(body.title, 120);
      if (!title) return json_({ ok: false, error: 'missing_title' });
      if (sh.getLastRow() > MAX_ROWS) return json_({ ok: false, error: 'full' });
      const year = /^\d{4}$/.test(String(body.year || '')) ? String(body.year) : '';
      const vibe = VIBES.indexOf(body.vibe) >= 0 ? body.vibe : '';
      const id = Utilities.getUuid();
      sh.appendRow([id, new Date(), title, year, vibe, clean_(body.why, 400), clean_(body.from, 60), false]);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'watched') {
      const last = sh.getLastRow();
      if (last < 2) return json_({ ok: false, error: 'not_found' });
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(body.id)) {
          sh.getRange(i + 2, 8).setValue(body.watched === true);
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    return json_({ ok: false, error: 'bad_action' });
  } finally {
    lock.releaseLock();
  }
}
