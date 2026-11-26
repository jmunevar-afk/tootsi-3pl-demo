/**
 * Store.gs — reads and writes the Google Sheet that acts as the database.
 * Each tab is a table. Row 1 = headers. Columns ending in "Json" hold JSON text.
 * Everything else in the project talks to the data only through these functions.
 */
var TABLES = {
  POs: ['id','po','item','wh','desc','supplier','eta','clockFrom','cntr','plt','cs','lbs','load','alloc','prodQty','salesQty','so','qc','qcAt','relAt','sap','instrSap','extra','tplInstr','source','addedBy','unplannedJson','planOk','nudged','updatedAt','closed','manual'],
  Submissions: ['subId','id','seq','superseded','by','at','recv','plt','cs','lbs','firstJson','diffNote','lot','exp','cond','damage','dmgText','comment','filesJson','ship','grpPlt','lotsJson'],
  Issues: ['id','type','owner','ownerManual','status','opened','urgent','urgentManual','openManual','decision','decisionNote','updatedAt','resolvedManual'],
  Notes: ['id','who','text','at','ext'],
  Delays: ['delayId','id','reason','note','newDate','by','at','date','closedAt','filesJson'],
  Recounts: ['id','at','from','active','base'],
  Docs: ['id','cat','name','url','by','at'],
  SupplierQty: ['id','kind','plt','cs','lbs','by','at'],
  Users: ['token','wh','name','email','active','user','salt','passHash','lastLogin'],
  Sessions: ['sid','token','expires','at'],
  Checks: ['key','at','by','off','countJson'],
  Slips: ['po','status','url','fileId','source','action','checked'],
  MasterSync: ['key','at','po','item','wh','field','value','status','row'],
  MasterSeen: ['id','recv','row','checked'],
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
function readAll_(name) {
  if (_cache[name]) return _cache[name];
  var sh = sheet_(name), cols = TABLES[name];
  var last = sh.getLastRow();
  var rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, cols.length).getValues();
  // number the rows BEFORE dropping blank ones, so _row always points at the real sheet row
  var out = rows.map(function (r, i) {
    var o = { _row: i + 2, _blank: r.join('') === '' };
    cols.forEach(function (c, j) { o[c] = normalize_(c, r[j]); });
    return o;
  }).filter(function (o) { return !o._blank; });
  out.forEach(function (o) { delete o._blank; });
  _cache[name] = out;
  return out;
}

function normalize_(col, v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, dbTz_(), /At$|^at$/.test(col) ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
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

var TEXT_COLS_ = /^(field|row|status|source|action|checked|fileId|url|id|po|item|lot|token|user|salt|passHash|sid|subId|delayId|ship|eta|clockFrom|recv|exp|qcAt|relAt|sap|newDate|date|opened|nudged|from|cntr|so|key|value|closed|manual)$/;
/**
 * Turn an object into a sheet row.
 * Every piece of text gets a leading apostrophe: Sheets then stores it exactly as typed.
 * Without it, "=..." would become a formula, "-2 cs" an error and "12/11" a date.
 * A number typed as text ("21") becomes a real number, except in text columns (PO 0042 stays 0042).
 */
function toRow_(name, o) {
  return TABLES[name].map(function (c) {
    var v = o[c];
    if (/Json$/.test(c)) return v == null ? '' : "'" + JSON.stringify(v);
    if (v === undefined || v === null) return '';
    if (typeof v !== 'string') return v;
    if (v === '') return '';
    v = v.slice(0, 5000);
    if (!TEXT_COLS_.test(c) && /^-?\d+(\.\d+)?$/.test(v)) return Number(v);
    return "'" + v;
  });
}

/** Append one object as a new row. */
function append_(name, o) {
  var sh = sheet_(name);
  sh.appendRow(toRow_(name, o));
  delete _cache[name];
}

/** Update the row with the same key, or append_ it. */
function upsert_(name, key, o) {
  var rows = readAll_(name);
  var found = rows.filter(function (r) { return String(r[key]) === String(o[key]); })[0];
  if (found) {
    var merged = {};
    TABLES[name].forEach(function (c) { merged[c] = (c in o) ? o[c] : found[c]; });
    sheet_(name).getRange(found._row, 1, 1, TABLES[name].length).setValues([toRow_(name, merged)]);
  } else {
    append_(name, o);
  }
  delete _cache[name];
}

/** Update the first row where match(row) is true, or append_. */
function upsertWhere_(name, match, o) {
  var found = readAll_(name).filter(match)[0];
  if (found) sheet_(name).getRange(found._row, 1, 1, TABLES[name].length).setValues([toRow_(name, o)]);
  else append_(name, o);
  delete _cache[name];
}

/** Delete every row where match(row) is true (bottom first, so row numbers stay right). */
function deleteWhere_(name, match) {
  var rows = readAll_(name).filter(match).map(function (r) { return r._row; }).sort(function (a, b) { return b - a; });
  var sh = sheet_(name);
  rows.forEach(function (n) { sh.deleteRow(n); });
  delete _cache[name];
}

var _tz = null;
/** Time zone of the database Sheet, so dates typed in it read back as the same day. */
function dbTz_() { if (!_tz) _tz = tzOr_(function () { return ss_().getSpreadsheetTimeZone(); }, 'America/Toronto'); return _tz; }

/** Google sometimes returns no usable time zone for a sheet. Use the fallback then. */
function tzOr_(get, fallback) {
  try { var z = get(); if (z && typeof z === 'string' || z instanceof String) return String(z); } catch (e) {}
  return fallback;
}

function clearCache_() { _cache = {}; }

/** Config tab as a simple key → value map. */
function config_() {
  var m = {};
  readAll_('Config').forEach(function (r) { m[r.key] = r.value; });
  return m;
}

/** Save an uploaded file (base64) into Drive: <root>/<PO id>/<name>. Returns {name,url,fileId}. */
/** Staff open photos from the hub: view access for anyone at Tootsi with the link. Never public. */
function shareDomain_(file) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {} }
function rootFolder_() { return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID') || config_().ROOT_FOLDER_ID); }
/** One folder per PO (all lines of a PO share it): "PO Hub files/43725". */
function poFolder_(id) {
  var root = rootFolder_(), name = String(id).split('|')[0];
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
/** Server-side checks on any uploaded file: photos and PDFs only, 10 MB max. */
function checkFile_(f) {
  if (!f || !f.data) throw new Error('Empty file.');
  var type = String(f.type || ''), data = String(f.data).split(',').pop();
  if (!/^image\/(jpeg|png|webp|heic|heif|gif)$/.test(type) && type !== 'application/pdf') throw new Error('Only photos and PDF files can be added.');
  if (data.length > 14000000) throw new Error('File too large (10 MB max).');
  // the first bytes must really be a photo or a PDF (a renamed web page or program is refused)
  var head = Utilities.base64Decode(data.slice(0, 24)).slice(0, 12).map(function (b) { return b & 255; });
  var hex = head.map(function (b) { return ('0' + b.toString(16)).slice(-2); }).join('');
  var okMagic = /^ffd8ff/.test(hex) || /^89504e47/.test(hex) || /^25504446/.test(hex) || /^47494638/.test(hex) || (/^52494646/.test(hex) && hex.slice(16, 24) === '57454250') || hex.slice(8, 16) === '66747970';
  if (!okMagic) throw new Error('Only photos and PDF files can be added.');
  return { name: String(f.name || 'file').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120), type: type, data: data };
}
function saveFile_(id, f) {
  var bytes = Utilities.base64Decode(String(f.data).split(',').pop());
  var file = poFolder_(id).createFile(Utilities.newBlob(bytes, f.type || 'application/octet-stream', f.name || 'file'));
  shareDomain_(file);
  return { name: file.getName(), url: file.getUrl(), fileId: file.getId() };
}
/** Uploads land first in "PO Hub files/_incoming", named "<warehouse>__<name>". The browser only gets the file id.
 *  The folder is created by setup() and its id kept, so two uploads at the same time never create two folders. */
function incomingFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('INCOMING_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  return withLock_(function () {
    var root = rootFolder_(), it = root.getFoldersByName('_incoming'), f = it.hasNext() ? it.next() : root.createFolder('_incoming');
    props.setProperty('INCOMING_ID', f.getId()); return f;
  });
}
function saveIncoming_(who, f) {
  var file = incomingFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.type, who + '__' + f.name));
  return { fid: file.getId(), name: f.name };
}
/** Before saving a receipt: the file must exist and must have been uploaded by this same warehouse. Nothing moves yet. */
function verifyFile_(fid, who) {
  var file;
  try { file = DriveApp.getFileById(String(fid)); } catch (e) { throw new Error('A photo was not uploaded. Please add it again.'); }
  var pre = who + '__', name = file.getName(), inc = incomingFolder_().getId(), ok = false, ps = file.getParents();
  while (ps.hasNext()) if (ps.next().getId() === inc) ok = true;
  if (!ok || file.isTrashed() || name.indexOf(pre) !== 0) throw new Error('A photo was not uploaded. Please add it again.');
  return { name: name.slice(pre.length), url: file.getUrl() };
}
/** After the receipt is saved (outside the lock): move each photo into its PO folder, drop the prefix, share with Tootsi.
 *  The saved link stays valid even if a move fails; the file then simply stays in _incoming. */
function finishMoves_(moves) {
  (moves || []).forEach(function (m) {
    try {
      var file = DriveApp.getFileById(m.fid), pre = m.who + '__';
      file.moveTo(poFolder_(m.poId));
      if (file.getName().indexOf(pre) === 0) file.setName(file.getName().slice(pre.length));
      shareDomain_(file);
    } catch (e) { Logger.log('Could not move ' + m.fid + ': ' + e.message); }
  });
}
/** Daily: uploads never used in a receipt (older than 3 days) go to the trash. */
/** A packing slip PDF for the 3PL portal (the 3PL has no Google account, so the server sends the file itself). */
function slipFile_(fileId) {
  var blob = DriveApp.getFileById(String(fileId)).getBlob();
  if (blob.getBytes().length > 4000000) throw new Error('Packing slip too large to open here. Ask Tootsi to email it.');
  return { name: blob.getName() || 'packing-slip.pdf', type: blob.getContentType() || 'application/pdf', data: Utilities.base64Encode(blob.getBytes()) };
}
function cleanIncoming_() {
  var linked = {};
  readAll_('Submissions').forEach(function (s) { (s.filesJson || []).forEach(function (f) { if (f.fileId) linked[f.fileId] = s.id; }); });
  var it = incomingFolder_().getFiles(), old = new Date(Date.now() - 3 * 864e5), n = 0, retry = [];
  while (it.hasNext()) {
    var f = it.next(), id = f.getId();
    if (linked[id]) { retry.push({ fid: id, who: f.getName().split('__')[0], poId: linked[id] }); continue; } // part of a receipt: move it, never trash it
    if (f.getDateCreated() < old) { f.setTrashed(true); n++; }
  }
  finishMoves_(retry);
  return n;
}
function logAction_(who, action, id, detail) {
  append_('Log', { at: nowStamp_(), who: who, action: action, id: id || '', detail: detail || '' });
}

function nowStamp_() { return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd'); }
