// Beanly — Google Sheet sync backend (v1.2.0)
//
// Setup (first time):
// 1. Create a new Google Sheet, then Extensions > Apps Script.
// 2. Delete any starter code and paste this whole file in.
// 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
// 4. Copy the Web app URL (ends in /exec) into Beanly: hold the version number > Google Sheet sync > Save.
//
// Updating from an older version of this script:
// Paste this over the old code, then Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
// The URL stays the same, so nothing changes in the app.
//
// The "Beans", "Logs" and "Removed" tabs are created automatically.

var HEADERS = {
  Beans:   ['Bean', 'LastUsed', 'LastUsedDisplay', 'GrindSize', 'PullTime', 'Comments', 'Updated'],
  Logs:    ['Id', 'Name', 'Date', 'DateDisplay', 'GrindSize', 'PullTime', 'Comments'],
  Removed: ['Bean', 'Timestamp']
};
// Columns (1-based) kept as plain text so Sheets never turns them into dates or numbers.
var TEXT_COLS = { Beans: [1, 2, 3, 6], Logs: [1, 2, 3, 4, 7], Removed: [1] };

function getSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function orNull_(v) {
  return (v === '' || v === undefined || v === null) ? null : v;
}
function isoStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return (v === null || v === undefined) ? '' : String(v);
}
function dispStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'd MMM yyyy');
  return (v === null || v === undefined) ? '' : String(v);
}

function readState_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var beans = {}, logs = [], removed = {};

  var b = getSheet_(ss, 'Beans').getDataRange().getValues();
  for (var i = 1; i < b.length; i++) {
    var r = b[i];
    if (!r[0]) continue;
    beans[String(r[0])] = {
      lastUsed: isoStr_(r[1], tz),
      lastUsedDisplay: dispStr_(r[2], tz),
      grindSize: orNull_(r[3]),
      pullTime: orNull_(r[4]),
      comments: r[5] === null || r[5] === undefined ? '' : String(r[5]),
      updated: Number(r[6]) || 0
    };
  }

  var l = getSheet_(ss, 'Logs').getDataRange().getValues();
  for (var j = 1; j < l.length; j++) {
    var q = l[j];
    if (!q[0]) continue;
    logs.push({
      id: String(q[0]),
      name: String(q[1]),
      date: isoStr_(q[2], tz),
      dateDisplay: dispStr_(q[3], tz),
      grindSize: orNull_(q[4]),
      pullTime: orNull_(q[5]),
      comments: q[6] === null || q[6] === undefined ? '' : String(q[6])
    });
  }

  var d = getSheet_(ss, 'Removed').getDataRange().getValues();
  for (var k = 1; k < d.length; k++) {
    if (!d[k][0]) continue;
    removed[String(d[k][0])] = Number(d[k][1]) || 0;
  }
  return { beans: beans, logs: logs, removed: removed };
}

function writeSheet_(ss, name, rows) {
  var sh = getSheet_(ss, name);
  sh.clearContents();
  var all = [HEADERS[name]].concat(rows);
  var height = Math.max(all.length, sh.getMaxRows());
  TEXT_COLS[name].forEach(function (c) {
    sh.getRange(1, c, height, 1).setNumberFormat('@');
  });
  sh.getRange(1, 1, all.length, HEADERS[name].length).setValues(all);
}

function writeState_(state) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var beanRows = Object.keys(state.beans).map(function (n) {
    var x = state.beans[n];
    return [n, x.lastUsed, x.lastUsedDisplay,
      x.grindSize === null ? '' : x.grindSize,
      x.pullTime === null ? '' : x.pullTime,
      x.comments || '', x.updated || 0];
  });
  var logRows = state.logs.map(function (e) {
    return [e.id, e.name, e.date, e.dateDisplay,
      e.grindSize === null ? '' : e.grindSize,
      e.pullTime === null ? '' : e.pullTime,
      e.comments || ''];
  });
  var removedRows = Object.keys(state.removed).map(function (n) {
    return [n, state.removed[n]];
  });
  writeSheet_(ss, 'Beans', beanRows);
  writeSheet_(ss, 'Logs', logRows);
  writeSheet_(ss, 'Removed', removedRows);
}

// MERGE-START
function normBean_(bn) {
  var c = {};
  for (var k in bn) c[k] = bn[k];
  c.updated = Number(bn.updated) || Date.parse(bn.lastUsed) || 0;
  return c;
}
function merge_(a, b) {
  a = a || {}; b = b || {};
  var logMap = {};
  (a.logs || []).concat(b.logs || []).forEach(function (l) {
    if (l && l.id && !logMap[l.id]) logMap[l.id] = l;
  });
  var rem = {};
  [a.removed || {}, b.removed || {}].forEach(function (r) {
    Object.keys(r).forEach(function (n) {
      rem[n] = Math.max(Number(rem[n]) || 0, Number(r[n]) || 0);
    });
  });
  var names = {};
  Object.keys(a.beans || {}).concat(Object.keys(b.beans || {})).forEach(function (n) { names[n] = true; });
  var out = {};
  Object.keys(names).forEach(function (n) {
    var x = a.beans && a.beans[n] ? normBean_(a.beans[n]) : null;
    var y = b.beans && b.beans[n] ? normBean_(b.beans[n]) : null;
    var win = (x && y) ? (y.updated > x.updated ? y : x) : (x || y);
    var t = Number(rem[n]) || 0;
    if (!t || win.updated > t) out[n] = win;
  });
  return {
    beans: out,
    logs: Object.keys(logMap).map(function (k) { return logMap[k]; }),
    removed: rem
  };
}
// MERGE-END

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_(readState_());
}

// Each phone sends its full state; the sheet MERGES it in (never overwrites), under a lock,
// so two devices saving at the same moment can't lose each other's edits.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var incoming = JSON.parse(e.postData.contents);
    var merged = merge_(readState_(), incoming);
    writeState_(merged);
  } finally {
    lock.releaseLock();
  }
  return json_({ status: 'ok' });
}
