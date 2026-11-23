/**
 * Store.gs — reads and writes the Google Sheet that acts as the database.
 * Each tab is a table. Row 1 = headers. Columns ending in "Json" hold JSON text.
 * Everything else in the project talks to the data only through these functions.
 */
var TABLES = {
  POs: ['id','po','item','wh','desc','supplier','eta','clockFrom','cntr','plt','cs','lbs','load','alloc','prodQty','salesQty','so','qc','qcAt','relAt','sap','instrSap','extra','tplInstr','source','addedBy','unplannedJson','planOk','nudged','updatedAt'],
  Submissions: ['subId','id','seq','superseded','by','at','recv','plt','cs','lbs','firstJson','diffNote','lot','exp','cond','damage','dmgText','comment','filesJson'],
  Issues: ['id','type','owner','ownerManual','status','opened','urgent','urgentManual','openManual','decision','decisionNote','updatedAt'],
  Notes: ['id','who','text','at','ext'],
  Delays: ['delayId','id','reason','note','newDate','by','at','date','closedAt','filesJson'],
  Recounts: ['id','at','from','active'],
  Docs: ['id','cat','name','url','by','at'],
  SupplierQty: ['id','kind','plt','cs','lbs','by','at'],
  Users: ['token','wh','name','email','active','user','salt','passHash','lastLogin'],
  Sessions: ['sid','token','expires','at'],
  Config: ['key','value'],
  Log: ['at','who','action','id','detail']
};

var _cache = {};

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing tab "' + name + '". Run setup() first.');
  return sh;
}

/** Read a whole tab as an array of objects. JSON columns are parsed. */
function readAll(name) {
  if (_cache[name]) return _cache[name];
  var sh = sheet_(name), cols = TABLES[name];
  var last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = rows.filter(function (r) { return r.join('') !== ''; }).map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = normalize_(c, r[j]); });
    return o;
  });
  _cache[name] = out;
  return out;
}

function normalize_(col, v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'America/Toronto', /At$|^at$/.test(col) ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  if (/Json$/.test(col)) {
    if (v === '' || v == null) return null;
    try { return JSON.parse(v); } catch (e) { return null; }
  }
  if (v === true || v === false) return v;
  if (v === 'TRUE') return true;
  if (v === 'FALSE') return false;
  return v === null ? '' : v;
}

function toRow_(name, o) {
  return TABLES[name].map(function (c) {
    var v = o[c];
    if (/Json$/.test(c)) return v == null ? '' : JSON.stringify(v);
    if (v === undefined || v === null) return '';
    // keep PO numbers, lots, tokens and dates as text so Sheets doesn't reformat them
    // a leading apostrophe tells Sheets "keep as text" (PO 41499-1, lot 0042, dates stay exactly as typed)
    if (typeof v === 'string' && v !== '' && /^(id|po|item|lot|token|user|salt|passHash|sid|subId|delayId|eta|clockFrom|recv|exp|qcAt|relAt|sap|newDate|date|opened|nudged|from|cntr|so|key|value)$/.test(c)) return "'" + v;
    return v;
  });
}

/** Append one object as a new row. */
function append(name, o) {
  var sh = sheet_(name);
  sh.appendRow(toRow_(name, o));
  delete _cache[name];
}

/** Update the row with the same key, or append it. */
function upsert(name, key, o) {
  var rows = readAll(name);
  var found = rows.filter(function (r) { return String(r[key]) === String(o[key]); })[0];
  if (found) {
    var merged = {};
    TABLES[name].forEach(function (c) { merged[c] = (c in o) ? o[c] : found[c]; });
    sheet_(name).getRange(found._row, 1, 1, TABLES[name].length).setValues([toRow_(name, merged)]);
  } else {
    append(name, o);
  }
  delete _cache[name];
}

/** Update the first row where match(row) is true, or append. */
function upsertWhere(name, match, o) {
  var found = readAll(name).filter(match)[0];
  if (found) sheet_(name).getRange(found._row, 1, 1, TABLES[name].length).setValues([toRow_(name, o)]);
  else append(name, o);
  delete _cache[name];
}

/** Delete every row where match(row) is true (bottom first, so row numbers stay right). */
function deleteWhere(name, match) {
  var rows = readAll(name).filter(match).map(function (r) { return r._row; }).sort(function (a, b) { return b - a; });
  var sh = sheet_(name);
  rows.forEach(function (n) { sh.deleteRow(n); });
  delete _cache[name];
}

function clearCache_() { _cache = {}; }

/** Config tab as a simple key → value map. */
function config_() {
  var m = {};
  readAll('Config').forEach(function (r) { m[r.key] = r.value; });
  return m;
}

/** Save an uploaded file (base64) into Drive: <root>/<PO id>/<name>. Returns {name,url,fileId}. */
function saveFile_(id, f) {
  var root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID') || config_().ROOT_FOLDER_ID);
  var folderName = String(id).replace(/\|/g, ' · ');
  var it = root.getFoldersByName(folderName);
  var folder = it.hasNext() ? it.next() : root.createFolder(folderName);
  var bytes = Utilities.base64Decode(String(f.data).split(',').pop());
  var blob = Utilities.newBlob(bytes, f.type || 'application/octet-stream', f.name || 'file');
  var file = folder.createFile(blob);
  return { name: file.getName(), url: file.getUrl(), fileId: file.getId() };
}

function logAction_(who, action, id, detail) {
  append('Log', { at: nowStamp_(), who: who, action: action, id: id || '', detail: detail || '' });
}

function nowStamp_() { return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd'); }
