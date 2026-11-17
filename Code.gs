/**
 * Bump to Binge — the tiny backend that stores recommendations in this Google Sheet.
 *
 * Paste this whole file into Extensions → Apps Script of your Google Sheet.
 * First time: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone),
 *   then copy the Web app URL into SCRIPT_URL at the top of index.html.
 * Updating later: Deploy → Manage deployments → ✏️ → Version: New version → Deploy
 *   (this keeps the same URL).
 *
 * Rows live on a tab called "Recs". Columns are found by their header names, so you can
 * reorder them, and any missing ones are added automatically. Edit or delete rows by hand;
 * the page picks the change up on its next refresh.
 */

const SHEET_NAME = 'Recs';
const HEADERS = ['id', 'added', 'type', 'title', 'genre', 'subgenre', 'vibe', 'why', 'from', 'watched'];
const TYPES = ['movie', 'tv'];
const VIBES = ['cozy', 'laugh', 'cry', 'romance', 'binge', 'edge', 'classic'];
const MAX_ROWS = 2000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return sh;
  }
  // Add any header this version needs that an older sheet doesn't have yet.
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = HEADERS.filter(h => have.indexOf(h) < 0);
  if (missing.length) {
    const start = have.filter(String).length ? lastCol + 1 : 1;
    sh.getRange(1, start, 1, missing.length).setValues([missing]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// header name → 0-based column index
function columns_(sh) {
  const row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach((h, i) => { if (h !== '' && !(h in map)) map[String(h)] = i; });
  return map;
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
  const col = columns_(sh);
  const values = sh.getDataRange().getValues();
  values.shift(); // headers
  const get = (r, name) => (name in col ? r[col[name]] : '');
  const recs = values
    .filter(r => get(r, 'id') && get(r, 'title'))
    .map(r => {
      const added = get(r, 'added');
      const watched = get(r, 'watched');
      return {
        id: String(get(r, 'id')),
        added: added instanceof Date ? added.getTime() : Number(added) || 0,
        type: String(get(r, 'type') || 'movie'),
        title: String(get(r, 'title')),
        genre: String(get(r, 'genre') || ''),
        subgenre: String(get(r, 'subgenre') || ''),
        vibe: String(get(r, 'vibe') || ''),
        why: String(get(r, 'why') || ''),
        from: String(get(r, 'from') || ''),
        watched: watched === true || String(watched).toUpperCase() === 'TRUE'
      };
    });
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
    const col = columns_(sh);

    if (body.action === 'add') {
      if (body.website) return json_({ ok: true }); // spam-bot honeypot
      const title = clean_(body.title, 120);
      if (!title) return json_({ ok: false, error: 'missing_title' });
      if (sh.getLastRow() > MAX_ROWS) return json_({ ok: false, error: 'full' });
      const id = Utilities.getUuid();
      const fields = {
        id: id,
        added: new Date(),
        type: TYPES.indexOf(body.type) >= 0 ? body.type : 'movie',
        title: title,
        genre: clean_(body.genre, 40),
        subgenre: clean_(body.subgenre, 40),
        vibe: VIBES.indexOf(body.vibe) >= 0 ? body.vibe : '',
        why: clean_(body.why, 400),
        from: clean_(body.from, 60),
        watched: false
      };
      const row = new Array(sh.getLastColumn()).fill('');
      Object.keys(fields).forEach(k => { if (k in col) row[col[k]] = fields[k]; });
      sh.appendRow(row);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'watched') {
      const last = sh.getLastRow();
      if (last < 2) return json_({ ok: false, error: 'not_found' });
      const ids = sh.getRange(2, col.id + 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(body.id)) {
          sh.getRange(i + 2, col.watched + 1).setValue(body.watched === true);
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
