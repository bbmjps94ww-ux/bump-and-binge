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

var SHEET_NAME = 'Recs';
var HEADERS = ['id', 'added', 'type', 'title', 'genre', 'subgenre', 'vibe', 'why', 'from', 'watched'];
var TYPES = ['movie', 'tv'];
var VIBES = ['cozy', 'laugh', 'cry', 'romance', 'binge', 'edge', 'classic'];
var MAX_ROWS = 2000;

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return sh;
  }
  // Add any header this version needs that an older sheet doesn't have yet.
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var have = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var missing = HEADERS.filter(function (h) { return have.indexOf(h) < 0; });
  if (missing.length) {
    var start = have.filter(String).length ? lastCol + 1 : 1;
    sh.getRange(1, start, 1, missing.length).setValues([missing]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// header name → 0-based column index
function columns_(sh) {
  var row = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var map = {};
  row.forEach(function (h, i) { if (h !== '' && !(String(h) in map)) map[String(h)] = i; });
  return map;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Trim, cap length, and stop anything that would run as a spreadsheet formula.
function clean_(value, max) {
  var s = String(value == null ? '' : value).replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function doGet() {
  var sh = getSheet_();
  var col = columns_(sh);
  var values = sh.getDataRange().getValues();
  values.shift(); // headers
  function get(r, name) { return name in col ? r[col[name]] : ''; }
  var recs = values
    .filter(function (r) { return get(r, 'id') && get(r, 'title'); })
    .map(function (r) {
      var added = get(r, 'added');
      var watched = get(r, 'watched');
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
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    var sh = getSheet_();
    var col = columns_(sh);

    if (body.action === 'add') {
      if (body.website) return json_({ ok: true }); // spam-bot honeypot
      var title = clean_(body.title, 120);
      if (!title) return json_({ ok: false, error: 'missing_title' });
      if (sh.getLastRow() > MAX_ROWS) return json_({ ok: false, error: 'full' });
      var id = Utilities.getUuid();
      var fields = {
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
      var row = [];
      for (var c = 0; c < sh.getLastColumn(); c++) row.push('');
      Object.keys(fields).forEach(function (k) { if (k in col) row[col[k]] = fields[k]; });
      sh.appendRow(row);
      return json_({ ok: true, id: id });
    }

    if (body.action === 'watched') {
      var last = sh.getLastRow();
      if (last < 2) return json_({ ok: false, error: 'not_found' });
      var ids = sh.getRange(2, col.id + 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
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
