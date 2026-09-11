// Pool Water Test Records — JSON API for the Dosing Lookup app
// Deploy this as a Web App (Anyone can access) and paste the resulting
// URL into the dosing app's config.

var POOL_SHEETS = ['WWW','TPY','DR','Tengah','HS','Lentor','CC','HB','BK','JEIP1','JEIP2','JEOP1','JEOP2'];
var DAYS_BACK = 90; // how much history to serve

function testDebug() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('WWW');
  var raw = sheet.getRange(1, 1, 6, 25).getValues();
  raw.forEach(function(row, i) {
    var desc = row.map(function(cell) {
      return (cell instanceof Date ? 'DATE(' + cell + ')' : typeof cell + ':' + cell);
    });
    Logger.log('Row ' + i + ': ' + desc.join(' | '));
  });
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Debug mode: ?debug=1&sheet=WWW  -> dumps raw header + first data rows
  if (e && e.parameter && e.parameter.debug) {
    var dbgName = e.parameter.sheet || 'WWW';
    var dbgSheet = ss.getSheetByName(dbgName);
    if (!dbgSheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'sheet not found', name: dbgName, allSheets: ss.getSheets().map(function(s){return s.getName();}) }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var raw = dbgSheet.getRange(1, 1, 6, 10).getValues();
    var typed = raw.map(function(row){
      return row.map(function(cell){ return { value: cell, type: typeof cell, isDate: cell instanceof Date }; });
    });
    return ContentService.createTextOutput(JSON.stringify({ sheet: dbgName, rows: typed }, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ?pool=WWW  -> only scan that one sheet instead of all 13 (much faster).
  // Omit ?pool to get the old behaviour (every pool in one response).
  var requestedPool = e && e.parameter && e.parameter.pool;
  var sheetsToProcess = requestedPool ? [requestedPool] : POOL_SHEETS;

  var cache = CacheService.getScriptCache();
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - DAYS_BACK);
  var out = {};

  sheetsToProcess.forEach(function(name) {
    var cacheKey = 'pool_v1_' + name;
    var cached = cache.get(cacheKey);
    if (cached) {
      out[name] = JSON.parse(cached);
      return;
    }

    var sheet = ss.getSheetByName(name);
    if (!sheet) { out[name] = []; return; }
    var data = sheet.getDataRange().getValues();
    if (data.length < 3) { out[name] = []; return; }

    var groupRow = data[0];
    var subRow = data[1];
    var cols = findColumns(groupRow, subRow);

    var records = [];
    for (var r = 2; r < data.length; r++) {
      var row = data[r];
      var dateVal = parseDate(row[cols.date]);
      if (!dateVal) continue;
      if (dateVal < cutoff) continue;

      records.push({
        date: Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
        fc: { morning: numOrNull(row[cols.fcM]), afternoon: numOrNull(row[cols.fcA]), night: numOrNull(row[cols.fcN]) },
        ph: { morning: numOrNull(row[cols.phM]), afternoon: numOrNull(row[cols.phA]), night: numOrNull(row[cols.phN]) },
        temp: { morning: numOrNull(row[cols.tempM]), afternoon: numOrNull(row[cols.tempA]), night: numOrNull(row[cols.tempN]) },
        dosing: { chlorine: strOrNull(row[cols.dosCl]), acid: strOrNull(row[cols.dosAcid]) },
        attended: numOrNull(row[cols.attended]),
        attended8pm: numOrNull(row[cols.attended]),
        expected7am: numOrNull(row[cols.expected])
      });
    }
    records.sort(function(a, b) { return a.date < b.date ? 1 : -1; });
    out[name] = records;

    // Cache for 5 minutes so the next visitor gets an instant response.
    // (Wrapped in try/catch: CacheService rejects values over 100KB, in
    // which case we just skip caching that pool rather than failing.)
    try { cache.put(cacheKey, JSON.stringify(records), 300); } catch (err) {}
  });

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'string' && v.trim() !== '') {
    var d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function findColumns(groupRow, subRow) {
  var cols = { date: 0 };
  var currentGroup = '';
  var attendedSet = false;
  var expectedSet = false;
  for (var c = 0; c < groupRow.length; c++) {
    if (groupRow[c]) currentGroup = String(groupRow[c]).trim();
    var sub = subRow[c] ? String(subRow[c]).trim() : '';

    if (currentGroup === 'FC') {
      if (sub === 'Morning') cols.fcM = c;
      if (sub === 'Afternoon') cols.fcA = c;
      if (sub === 'Night') cols.fcN = c;
    }
    if (currentGroup === 'PH') {
      if (sub === 'Morning') cols.phM = c;
      if (sub === 'Afternoon') cols.phA = c;
      if (sub === 'Night') cols.phN = c;
    }
    if (currentGroup === 'Pool Temp') {
      if (sub === 'Morning') cols.tempM = c;
      if (sub === 'Afternoon') cols.tempA = c;
      if (sub === 'Night') cols.tempN = c;
    }
    if (currentGroup === 'Dosing') {
      if (sub === 'Chlorine') cols.dosCl = c;
      if (sub === 'Acid') cols.dosAcid = c;
    }
    // Matched on the subheader text alone (not the group row above it) —
    // the group cells for these two got auto-converted to time values
    // (8:00 PM / 7:00 AM) by Sheets, so they can't be matched as text.
    if (sub === 'Attended' && !attendedSet) {
      cols.attended = c;
      attendedSet = true;
    }
    if (sub === 'Expected' && !expectedSet) {
      cols.expected = c;
      expectedSet = true;
    }
  }
  return cols;
}

function numOrNull(v) {
  return (typeof v === 'number') ? v : null;
}
function strOrNull(v) {
  return (v === null || v === undefined || v === '') ? null : String(v);
}
