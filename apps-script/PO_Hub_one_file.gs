/**
 * PO Hub · Tootsi Impex · Google Workspace pilot
 * ONE-FILE INSTALL: paste this whole file into Extensions › Apps Script (replace everything in Code.gs), Save, then Run ▶ setup.
 * Source and docs: github.com/jmunevar-afk/tootsi-3pl-demo/tree/main/apps-script
 */

// ==================== Code.gs ====================
/**
 * Code.gs — web app entry point.
 *
 * Two deployments of this same script:
 *  1) INTERNAL hub  → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone within Tootsi
 *  2) 3PL portal    → Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 *     3PLs open the portal link and sign in with the username and password from addThreePL().
 *     Tootsi staff opening any link go straight to the hub (Google account check). Add ?view=3pl to test the 3PL screen.
 *
 * The script always runs as you. The Sheet and the Drive folder are never shared with the 3PLs:
 * they only ever see what portalLoad() returns (see portalRow() in Logic.gs).
 */
function doGet(e) {
  var view = e && e.parameter && e.parameter.view;
  // Tootsi staff: Google has already signed them in. The domain check is the login.
  if (view !== '3pl') {
    try {
      hubUser_();
      return tpl_('Hub').evaluate().setTitle('Tootsi · PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) { /* not a Tootsi Google account: show the 3PL sign-in page */ }
  }
  // 3PL partners: the page opens on the sign-in screen (username + password from the Users tab).
  var t = tpl_('Portal');
  t.lang = 'fr'; t.hubUrl = String(config_().HUB_URL || '');
  return t.evaluate().setTitle('Tootsi · 3PL Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function page_(name, vars) {
  var t = tpl_(name);
  Object.keys(vars || {}).forEach(function (k) { t[k] = vars[k]; });
  return t.evaluate().setTitle('Tootsi PO Hub').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Used inside HTML templates: <?!= include('Styles') ?> */
/** The receiving form code, sent as data (fallback if the page builder damaged the inline copy). Public page code only. */
function clientCode(name) {
  if (['Receive'].indexOf(name) < 0) throw new Error('Not available.');
  var h = include(name), m = h.match(/<script[^>]*>([\s\S]*)<\/script>/);
  return m ? m[1] : '';
}
function include(name) {
  return typeof HTML_FILES !== 'undefined' ? HTML_FILES[name] : HtmlService.createHtmlOutputFromFile(name).getContent();
}
/** The one-file install (PO_Hub.gs) carries the pages inside HTML_FILES. The multi-file install uses the .html files. */
function tpl_(name) {
  return typeof HTML_FILES !== 'undefined' ? HtmlService.createTemplate(HTML_FILES[name]) : HtmlService.createTemplateFromFile(name);
}

// ==================== Setup.gs ====================
/**
 * Setup.gs — run these once from the Apps Script editor (Run ▶).
 *
 *  setup()            creates every tab with headers, the Drive folder, and default settings.
 *  addThreePL('TP','Trans-Plus','receiving@…')  creates a private portal link for a 3PL. Run once per 3PL.
 *  loadSampleData()   optional: fills a few sample POs so you can try it right away.
 *  installTriggers()  turns on the daily jobs (import from the upcoming sheet, reminders, digest email).
 */
/**
 * Admin functions must be run by the owner from the editor (or by a time trigger).
 * Google lets a web page call any function whose name has no trailing "_", so each admin function checks first.
 * A time trigger passes an event with its triggerUid. The id must belong to one of this project's triggers.
 */
function requireOwner_(e) {
  if (e && e.triggerUid) {
    var ok = ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === String(e.triggerUid); });
    if (ok) return;
  }
  var a = String(Session.getActiveUser().getEmail() || '').toLowerCase(), o = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!a || a !== o) throw new Error('Only the owner can run this from the Apps Script editor.');
}

function setup(e) {
  requireOwner_(e);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the PO Hub Google Sheet (Extensions > Apps Script).');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var cols = TABLES[name];
    if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#16181a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  var s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  var props = PropertiesService.getScriptProperties();
  var cfgFolder = config_().ROOT_FOLDER_ID;          // folder created ahead of time (Config tab)
  if (cfgFolder && !props.getProperty('ROOT_FOLDER_ID')) props.setProperty('ROOT_FOLDER_ID', String(cfgFolder));
  if (!props.getProperty('ROOT_FOLDER_ID')) {
    var parent = DriveApp.getFileById(ss.getId()).getParents();
    var folder = (parent.hasNext() ? parent.next() : DriveApp.getRootFolder()).createFolder('PO Hub files');
    props.setProperty('ROOT_FOLDER_ID', folder.getId());
  }
  // the upload folder exists before anyone uploads (two first uploads at once would otherwise create two folders)
  try { incomingFolder_(); } catch (err) { Logger.log('Could not create _incoming: ' + err.message); }
  // Tootsi staff must be able to open photos and documents: view access for the whole domain (never "anyone")
  try { rootFolder_().setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) { Logger.log('Could not share the files folder with the domain: ' + err.message); }
  var me = Session.getActiveUser().getEmail();
  var defaults = {
    DOMAIN: (me.split('@')[1] || 'tootsi.com'),
    SLA_NUDGE_H: 9, SLA_ESC_H: 18, LB_TOLERANCE: 0.01, QC_DAYS: 2,
    HOLIDAYS: '2026-12-25,2026-12-28,2027-01-01,2027-04-02,2027-05-24,2027-06-24,2027-07-01,2027-09-06,2027-10-11',
    EMAIL_ON: 'no', PILOT_ONLY_ME: 'yes', FALLBACK_EMAIL: me, CC_ALWAYS: '', EMAIL_3PL: 'no',
    TEAM_Purchasing: '', TEAM_Logistics: '', TEAM_QC: '', TEAM_Receiving: '', TEAM_InventoryControl: me, TEAM_Fulfillment: '',
    LANG_TP: 'fr', LANG_DL: 'fr', SESSION_HOURS: 12, GO_LIVE: today_(), COMPARE_LB: 'no',
    SLIP_LOG_ID: '' /* Arash's "Packing Slip Extractor Log" spreadsheet id */, SLIP_TAB: 'PO Status', SLIP_TO_3PL: 'after' /* after = the 3PL sees the slip once its count is sent (keeps the count blind); no; yes */,
    MASTER_SHEET_ID: '' /* third_party_whse master id */, MASTER_TAB: '3PL', MASTER_MODE: 'preview', HUB_URL: '', PORTAL_URL: '',
    UPCOMING_SHEET_ID: '' /* paste the ID of the UPCOMING FTL FCL sheet */, UPCOMING_TAB: 'UPCOMING FTL FCL', DIGEST_TO: me
  };
  var have = config_();
  Object.keys(defaults).forEach(function (k) { if (!(k in have)) append_('Config', { key: k, value: defaults[k] }); });
  SpreadsheetApp.getUi && toastSafe_('Setup done. Next: addThreePL, then Deploy (two web apps).');
}

/**
 * Create a 3PL login. Example: addThreePL('TP', 'Trans-Plus', 'receiving@trans-plus.com', 'transplus')
 * The temporary password shows in the execution log. Give it to the 3PL by phone, not by email.
 */
function addThreePL(wh, name, email, user) {
  requireOwner_(null);
  wh = wh || 'TP'; name = name || (wh === 'TP' ? 'Trans-Plus' : 'Deslauriers');
  user = String(user || (wh === 'TP' ? 'transplus' : 'deslauriers')).toLowerCase();
  if (readAll_('Users').some(function (r) { return String(r.user).toLowerCase() === user; })) throw new Error('Username already used: ' + user);
  var token = Utilities.getUuid().replace(/-/g, '');
  var pass = newPassword_(), salt = Utilities.getUuid();
  append_('Users', { token: token, wh: wh, name: name, email: email || '', active: true, user: user, salt: salt, passHash: hash_(salt, pass) });
  Logger.log('3PL login for ' + name + ' → link: ' + (config_().PORTAL_URL || '<PORTAL_URL>') + ' · username: ' + user + ' · password: ' + pass);
  return { user: user, password: pass };
}

/** New password for a 3PL login (forgotten password, or a receiver left). Old sessions are signed out. */
function resetThreePLPassword(user) {
  requireOwner_(null);
  var u = readAll_('Users').filter(function (r) { return String(r.user).toLowerCase() === String(user).toLowerCase(); })[0];
  if (!u) throw new Error('No such username: ' + user);
  var pass = newPassword_(); u.salt = Utilities.getUuid(); u.passHash = hash_(u.salt, pass); upsert_('Users', 'token', u);
  deleteWhere_('Sessions', function (r) { return r.token === u.token; });
  Logger.log('New password for ' + u.user + ': ' + pass);
  return pass;
}

/** Turn off a 3PL login. It also signs out any open session. */
function disableThreePL(user) {
  requireOwner_(null);
  var u = readAll_('Users').filter(function (r) { return String(r.user).toLowerCase() === String(user).toLowerCase(); })[0];
  if (!u) return;
  u.active = false; upsert_('Users', 'token', u);
  deleteWhere_('Sessions', function (r) { return r.token === u.token; });
}

function newPassword_() {
  var ch = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789', p = '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now());
  for (var i = 0; i < 12; i++) p += ch[(bytes[i] & 255) % ch.length];
  return p;
}

function installTriggers(e) {
  requireOwner_(e);
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(7).nearMinute(40).everyDays(1).inTimezone('America/Toronto').create();
  ScriptApp.newTrigger('syncSources').timeBased().everyHours(1).create();
}

function toastSafe_(m) { try { SpreadsheetApp.getActiveSpreadsheet().toast(m, 'PO Hub', 8); } catch (e) {} }

/** Optional: a few sample POs to try the pilot. Safe to delete the rows afterwards. */
function loadSampleData(e) {
  requireOwner_(e);
  var t = today_(), d = function (n) { return addDays(t, n); };
  [
    { id: '90101|10201', po: '90101', item: '10201', wh: 'TP', desc: 'RAW WHOLE CASHEWS (W-320)', supplier: 'Cashew exporter · Vietnam', eta: d(1), cntr: 'TCLU4826103', plt: 28, cs: 1100, lbs: 55000, load: "FCL 40' · Palletized", alloc: 'Production', qc: 'Not received', instrSap: 'FOR PRODUCTION ONLY. No double stacking.', tplInstr: 'Do not accept double-stacked pallets. Photograph any torn bag before unloading.', source: 'sample' },
    { id: '90106|10220', po: '90106', item: '10220', wh: 'TP', desc: 'CASHEW PIECES (SSP2)', supplier: 'Cashew processor · Ghana', eta: d(-4), cntr: 'UETU3091547', plt: 18, cs: 700, lbs: 35000, load: "FCL 20' · Palletized", alloc: 'Sales', so: 'SO 118441', qc: 'Not received', source: 'sample' },
    { id: '90203|14015', po: '90203', item: '14015', wh: 'DL', desc: 'SULTANA RAISINS (TURKEY)', supplier: 'Raisin packer · Türkiye', eta: d(0), cntr: 'MSBU1188204', plt: 20, cs: 1600, lbs: 44096, load: "FCL 20' · Floor loaded", alloc: 'Production', qc: 'Not received', tplInstr: 'Floor-loaded. Photograph the open container before unloading.', source: 'sample' },
    { id: '90301|10501', po: '90301', item: '10501', wh: 'BM', desc: 'FANCY JUMBO PECAN HALVES', supplier: 'Pecan sheller · Mexico', eta: d(0), cntr: '', plt: 22, cs: 1680, lbs: 41580, load: 'FTL · Palletized', alloc: 'Sales', so: 'SO 118455', qc: 'Not received', source: 'sample' }
  ].forEach(function (p) { p.updatedAt = nowStamp_(); upsert_('POs', 'id', p); });
}

// ==================== Store.gs ====================
/**
 * Store.gs — reads and writes the Google Sheet that acts as the database.
 * Each tab is a table. Row 1 = headers. Columns ending in "Json" hold JSON text.
 * Everything else in the project talks to the data only through these functions.
 */
var TABLES = {
  POs: ['id','po','item','wh','desc','supplier','eta','clockFrom','cntr','plt','cs','lbs','load','alloc','prodQty','salesQty','so','qc','qcAt','relAt','sap','instrSap','extra','tplInstr','source','addedBy','unplannedJson','planOk','nudged','updatedAt','closed','manual'],
  Submissions: ['subId','id','seq','superseded','by','at','recv','plt','cs','lbs','firstJson','diffNote','lot','exp','cond','damage','dmgText','comment','filesJson','ship','grpPlt'],
  Issues: ['id','type','owner','ownerManual','status','opened','urgent','urgentManual','openManual','decision','decisionNote','updatedAt'],
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

// ==================== Logic.gs ====================
/**
 * Logic.gs — business rules. Pure functions: no Sheets, no Drive.
 * They take a "state" object (all tables) and a "ctx" (today, current hour, config).
 */
var TPL_WH = ['TP', 'DL'], OWN_WH = ['BM', 'LY', 'HB'];
var DELAY_OWNER = { r_notarr: 'Logistics', r_customs: 'Logistics', r_docs: 'Purchasing', r_seal: 'QC', r_wrong: 'Purchasing', r_dmg: 'QC', r_dock: 'Inventory Control', r_other: 'Inventory Control' };
var REASON_EN = { r_notarr: 'Truck or container not arrived', r_customs: 'Held at port or customs', r_docs: 'Missing documents (BOL, packing list)', r_seal: 'Seal broken or missing', r_wrong: 'Wrong product or labels', r_dmg: 'Damaged load, waiting for a decision', r_dock: 'Dock full or no staff', r_other: 'Other' };
var SLA_EN = { up: 'Upcoming', tom: 'Tomorrow', due: 'Expected today', nudge: '1 business day late', esc: '2+ business days late', paused: 'Delay reported' };

function isTPL(w) { return TPL_WH.indexOf(w) >= 0; }
/** A shipment = what arrives together: one container, or one PO when there is no container number. */
function shipKey(p) { return (p.wh || '?') + '|' + (p.cntr ? 'C:' + p.cntr : 'P:' + p.po); }
/** True while nothing of the shipment has been received: pallets are counted once, on that first receiving. */
/** Nothing of the shipment was ever received (a recount does not make it fresh again). */
function shipFresh(p, st) { return !shipLines(p, st).some(function (x) { return isDone(x, st) || (st.hist[x.id] || []).length > 0; }); }
/** Pallets are counted on the first receiving of a shipment, and again when the line that holds the pallet count is recounted.
 *  ONE rule, used by the server checks AND by the 3PL / hub screens. */
function shipAskP(p, st) {
  if (shipFresh(p, st)) return true;
  return shipLines(p, st).some(function (x) { return x.wh === p.wh && !isDone(x, st) && (st.hist[x.id] || []).some(function (s) { return s.plt !== '' && s.plt != null; }); });
}
/** The supplier packing slip shows the expected cartons, so a 3PL sees it only AFTER sending its own count (SLIP_TO_3PL = after),
 *  never (no), or always (yes: breaks the blind count). */
function slipFor3pl(p, st, ctx) {
  var mode = String(ctx.cfg.SLIP_TO_3PL || 'after'), s = st.slip[p.po];
  if (!s || s.status !== 'FOUND' || !s.fileId || mode === 'no') return false;
  return mode === 'yes' || !!st.sub[p.id];
}
function isDone(p, st) { return !!st.sub[p.id] || !!p.closed; }
/** Lines of the same shipment (same warehouse and container, or same PO). */
function shipLines(p, st) { var k = shipKey(p); return st.pos.filter(function (x) { return shipKey(x) === k && !(x.unplannedJson && !truthy(x.planOk)); }); }
/** Expected pallets for a whole shipment. The Upcoming sheet repeats the container total on every line, so take it once. */
function shipPallets(lines) {
  var v = lines.map(function (x) { return num(x.plt); }).filter(function (n) { return n != null; });
  if (!v.length) return null;
  if (lines.length === 1) return v[0];
  var fromSheet = lines.every(function (x) { return x.source === 'sheet'; });
  return fromSheet ? Math.max.apply(null, v) : v.reduce(function (a, b) { return a + b; }, 0);
}
function isFloor(p) { return /floor/i.test(p.load || ''); }
function num(v) { if (v === '' || v == null) return null; var n = +v; return isFinite(n) ? n : null; }
function fmtN(n) { return n == null || n === '' ? '—' : Number(n).toLocaleString('en-CA'); }
function addDays(d, n) { var x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + n); var p = function (v) { return (v < 10 ? '0' : '') + v; }; return x.getFullYear() + '-' + p(x.getMonth() + 1) + '-' + p(x.getDate()); }
function daysBetween(a, b) { return Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 864e5); }

/** Business hours (Mon–Fri, 8:00–17:00) from startDate 08:00 until now. Holidays from config (comma list of yyyy-mm-dd). */
function bizHours(startDate, ctx) {
  var h = 0, s = new Date(startDate + 'T08:00:00'), e = new Date(ctx.today + 'T' + (ctx.nowH < 10 ? '0' : '') + ctx.nowH + ':00:00');
  var hol = String(ctx.cfg.HOLIDAYS || '').split(',').map(function (x) { return x.trim(); });
  for (var x = new Date(s); x < e; x.setHours(x.getHours() + 1)) {
    var d = x.getDay(), hr = x.getHours();
    var ds = addDays(x.getFullYear() + '-' + ('0' + (x.getMonth() + 1)).slice(-2) + '-' + ('0' + x.getDate()).slice(-2), 0);
    if (d > 0 && d < 6 && hr >= 8 && hr < 17 && hol.indexOf(ds) < 0) h++;
  }
  return h;
}

/** Index the raw tables so lookups by PO id are fast. */
function indexState(raw) {
  var st = { slip: {}, mseen: {}, pos: raw.pos, sub: {}, hist: {}, issue: {}, notes: {}, delay: {}, delayLog: {}, req: {}, docs: {}, pl: {}, lr: {} };
  raw.subs.slice().sort(function (a, b) { return a.seq - b.seq; }).forEach(function (s) {
    if (s.superseded === true || s.superseded === 'TRUE') (st.hist[s.id] = st.hist[s.id] || []).push(s); else st.sub[s.id] = s;
  });
  raw.issues.forEach(function (i) { if (!st.issue[i.id]) st.issue[i.id] = i; }); // first row wins (same row the save updates)
  raw.notes.forEach(function (n) { (st.notes[n.id] = st.notes[n.id] || []).push(n); });
  raw.delays.forEach(function (d) { if (d.closedAt) (st.delayLog[d.id] = st.delayLog[d.id] || []).push(d); else st.delay[d.id] = d; });
  raw.recounts.forEach(function (r) { if (r.active === true || r.active === 'TRUE') st.req[r.id] = r; });
  raw.docs.forEach(function (d) { (st.docs[d.id] = st.docs[d.id] || []).push(d); });
  (raw.slips || []).forEach(function (s) { st.slip[s.po] = s; });
  (raw.mseen || []).forEach(function (m) { st.mseen[m.id] = m; });
  raw.supq.forEach(function (q) { if (q.kind === 'pl') st.pl[q.id] = q; else st.lr[q.id] = q; });
  return st;
}

function diffs(p, got, ctx) {
  var d = []; if (!p || !got) return d;
  var tol = Number(ctx.cfg.LB_TOLERANCE || 0.01);
  var ep = got.grpPlt !== undefined && got.grpPlt !== '' && got.grpPlt != null ? num(got.grpPlt) : num(p.plt), ec = num(p.cs), el = num(p.lbs), gp = num(got.plt), gc = num(got.cs), gl = num(got.lbs);
  if (!isFloor(p) && ep != null && gp != null && gp !== ep) d.push('pallets');
  if (ec != null && gc != null && gc !== ec) d.push('cases');
  // the Upcoming sheet "Line Weight" may be gross: compare lb only when COMPARE_LB = yes
  if (String(ctx.cfg.COMPARE_LB || 'no') === 'yes' && el && gl != null && Math.abs(gl - el) / el > tol) d.push('lbs');
  return d;
}

function gapText(p, got, ctx) {
  if (!got) return '';
  var tol = Number(ctx.cfg.LB_TOLERANCE || 0.01), parts = [];
  var ep = got.grpPlt !== undefined && got.grpPlt !== '' && got.grpPlt != null ? num(got.grpPlt) : num(p.plt), ec = num(p.cs), el = num(p.lbs), gp = num(got.plt), gc = num(got.cs), gl = num(got.lbs);
  if (ec != null && gc != null && gc !== ec) parts.push((gc < ec ? 'Short ' : 'Over ') + fmtN(Math.abs(gc - ec)) + ' cs');
  if (!isFloor(p) && ep != null && gp != null && gp !== ep) parts.push((gp < ep ? 'short ' : 'over ') + Math.abs(gp - ep) + ' plt');
  if (String(ctx.cfg.COMPARE_LB || 'no') === 'yes' && el && gl != null && Math.abs(gl - el) / el > tol) parts.push((gl < el ? 'light ' : 'heavy ') + fmtN(Math.abs(gl - el)) + ' lb');
  var s = parts.join(', '); return s.charAt(0).toUpperCase() + s.slice(1);
}

/** Receiving clock level for a PO. */
function sla(p, st, ctx) {
  if (isDone(p, st)) return null;
  if (!p.wh) return null;
  if (st.delay[p.id]) return { lvl: 'paused', h: 0 };
  var from = st.req[p.id] ? st.req[p.id].from : (p.clockFrom && (!p.eta || p.clockFrom > p.eta) ? p.clockFrom : p.eta);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(from || ''))) return null;
  // lines that were due before the hub went live start their clock on the go-live day
  var live = String(ctx.cfg.GO_LIVE || '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(live) && from < live) from = live;
  if (from > ctx.today) return { lvl: from === addDays(ctx.today, 1) ? 'tom' : 'up', h: 0 };
  var h = bizHours(from, ctx), nud = Number(ctx.cfg.SLA_NUDGE_H || 9), esc = Number(ctx.cfg.SLA_ESC_H || 18);
  return { lvl: h >= esc ? 'esc' : h >= nud ? 'nudge' : 'due', h: h };
}
function isLate(p, st, ctx) { var s = sla(p, st, ctx); return !!s && (s.lvl === 'nudge' || s.lvl === 'esc'); }

/** What the issue for a PO should be right now. */
function issueState(p, st, ctx) {
  if (p.closed) return { type: '', owner: 'Purchasing', urgent: false, opened: ctx.today };
  var s = st.sub[p.id], pl = st.pl[p.id], sl = sla(p, st, ctx), types = [];
  var owner = 'Purchasing', urgent = false, opened = ctx.today;
  var unpl = p.unplannedJson && !p.planOk;
  if (unpl) { types.push('Arrived at ' + p.wh + ', not on the upcoming list'); owner = 'Logistics'; urgent = true; opened = p.eta; }
  if (pl && !s && diffs(p, pl, ctx).length) types.push('Packing list ≠ PO (' + gapText(p, pl, ctx) + ')');
  if (s && diffs(p, s, ctx).length) types.push(gapText(p, s, ctx));
  var lastH = (st.hist[p.id] || []).slice(-1)[0];
  if (!s && st.req[p.id] && lastH && diffs(p, lastH, ctx).length) types.push(gapText(p, lastH, ctx) + ' (recount asked)');
  else if (s && num(s.cs) === 0 && num(p.cs) == null) types.push('Item did not arrive (0 cs counted)');
  var hist = st.hist[p.id] || [];
  if ((s && s.damage) || hist.some(function (h) { return h.damage; })) { types.push('damage'); urgent = true; }
  var dl = st.delay[p.id];
  if (dl) { types.push('Delay: ' + REASON_EN[dl.reason]); owner = DELAY_OWNER[dl.reason]; opened = dl.date || ctx.today; }
  else if (sl && (sl.lvl === 'nudge' || sl.lvl === 'esc')) {
    types.push('Not received at ' + p.wh); if (!unpl) owner = 'Inventory Control'; if (sl.lvl === 'esc') urgent = true;
    opened = st.req[p.id] ? st.req[p.id].from : (p.clockFrom || p.eta);
    var live = String(ctx.cfg.GO_LIVE || ''); if (/^\d{4}-\d{2}-\d{2}$/.test(live) && opened < live) opened = live;
  }
  return { type: types.join(' + '), owner: owner, urgent: urgent, opened: opened };
}

/** Returns the list of issue rows that must be written back (created, changed, auto-closed) + system notes. */
function syncIssues(st, ctx) {
  var writes = [], notes = [];
  st.pos.forEach(function (p) {
    var want = issueState(p, st, ctx), is = st.issue[p.id];
    if (want.type && !is) {
      is = { id: p.id, type: want.type, owner: want.owner, ownerManual: false, status: 'Open', opened: want.opened, urgent: want.urgent, urgentManual: false, openManual: false, decision: (st.sub[p.id] && st.sub[p.id].damage) ? 'Pending' : '', decisionNote: '', updatedAt: ctx.stamp, _new: true };
      st.issue[p.id] = is; writes.push(is); return;
    }
    if (!is) return;
    var changed = false;
    if (want.type && want.type !== is.type) {
      if (is.status !== 'Open') { is.status = 'Open'; is.opened = ctx.today; notes.push({ id: p.id, who: 'System', text: 'Issue reopened: ' + want.type, at: ctx.stamp, ext: false }); }
      is.type = want.type; if (!truthy(is.ownerManual)) is.owner = want.owner; is._chg = true;
      if (st.sub[p.id] && st.sub[p.id].damage && !is.decision) is.decision = 'Pending';
      changed = true;
    }
    if (want.type && !truthy(is.urgentManual) && !truthy(is.urgent) && want.urgent) { is.urgent = true; is._esc = true; changed = true; }
    if (!want.type && is.status === 'Open' && !truthy(is.openManual)) {
      is.status = 'Resolved'; changed = true;
      notes.push({ id: p.id, who: 'System', text: 'Issue closed automatically: nothing left to fix (' + is.type + ').', at: ctx.stamp, ext: false });
    }
    if (changed) { is.updatedAt = ctx.stamp; writes.push(is); }
  });
  return { writes: writes, notes: notes };
}
function truthy(v) { return v === true || v === 'TRUE' || v === 'true'; }

function openIssue(p, st) { var i = st.issue[p.id]; return i && i.status === 'Open' ? i : null; }
function issueOverdue(i, ctx) { if (!i || i.status !== 'Open') return false; return bizHours(i.opened, ctx) >= (truthy(i.urgent) ? Number(ctx.cfg.SLA_NUDGE_H || 9) : Number(ctx.cfg.SLA_ESC_H || 18)); }
function notInSap(p, st, ctx) { var s = st.sub[p.id]; return !!s && !p.sap && daysBetween(s.recv, ctx.today) >= 2; }

function availSales(p, st, ctx) {
  if (p.alloc !== 'Sales' && p.alloc !== 'Split') return null;
  if (p.alloc === 'Split' && (p.salesQty === '' || p.salesQty == null)) return null;
  var s = st.sub[p.id], qd = Number(ctx.cfg.QC_DAYS || 2);
  var base = p.alloc === 'Sales' ? num(s ? s.cs : p.cs) : num(p.salesQty);
  if (base == null) return null;
  var qty = p.alloc === 'Split' && s && num(p.cs) ? Math.round(base * num(s.cs) / num(p.cs)) : base;
  if (p.qc === 'Released') return { qty: qty, date: p.relAt || ctx.today, est: false };
  var d = s ? addDays(p.qcAt || s.recv, qd) : (p.eta ? addDays(p.eta, qd) : '');
  return { qty: qty, date: d, est: true, qcLate: !!s && d < ctx.today };
}

function countsList(p, st) {
  var out = [], h = st.hist[p.id] || [], s = st.sub[p.id];
  h.forEach(function (x, i) {
    if (x.firstJson) out.push({ l: 'Submission ' + (i + 1) + ': first count', v: x.firstJson, by: x.by });
    out.push({ l: 'Submission ' + (i + 1) + ': ' + (x.firstJson ? 'recount' : 'count'), v: { plt: x.plt, cs: x.cs, lbs: x.lbs }, by: x.by });
  });
  if (s) {
    var pre = h.length ? 'Submission ' + (h.length + 1) + ': ' : '';
    if (s.firstJson) out.push({ l: pre + 'First count', v: s.firstJson, by: s.by });
    out.push({ l: pre + (s.firstJson ? 'Recount (final)' : 'Count (final)'), v: { plt: s.plt, cs: s.cs, lbs: s.lbs }, by: s.by });
  }
  return out;
}

function needsAction(p, st, ctx) {
  if (p.closed) return [];
  var is = openIssue(p, st), s = st.sub[p.id], sl = sla(p, st, ctx), a = [], dl = st.delay[p.id];
  if (p.unplannedJson && !truthy(p.planOk)) a.push({ t: 'Arrived at ' + p.wh + ', not on list · check PO', u: true });
  if (!p.wh) a.push({ t: 'Set destination', u: false });
  if (is && is.decision === 'Pending') a.push({ t: 'Dock decision needed', u: true });
  if (dl) a.push({ t: 'Delay (' + DELAY_OWNER[dl.reason] + '): ' + REASON_EN[dl.reason] });
  else if (sl && sl.lvl === 'esc') a.push({ t: 'Not received at ' + p.wh + ' · escalated', u: true });
  else if (sl && sl.lvl === 'nudge') a.push({ t: 'Not received at ' + p.wh + ' · nudge' });
  if (is && issueOverdue(is, ctx)) a.push({ t: 'Overdue: ' + is.owner, u: true });
  else if (is && !dl && !(sl && (sl.lvl === 'nudge' || sl.lvl === 'esc')) && !(p.unplannedJson && !truthy(p.planOk))) a.push({ t: 'Issue · ' + is.owner });
  if (notInSap(p, st, ctx)) a.push({ t: 'Enter in SAP' });
  if (!s && st.mseen[p.id]) a.push({ t: 'Master sheet: received ' + st.mseen[p.id].recv + ' (units entered) · confirm', u: true });
  var sl = st.slip[p.po];
  if (!s && (!sl || sl.status !== 'FOUND') && p.eta && daysBetween(ctx.today, p.eta) <= 7 && !p.unplannedJson) a.push({ t: 'Packing slip missing' });
  if (!s && p.wh && num(p.cs) == null && !p.unplannedJson) a.push({ t: 'Add expected cases (not compared yet)' });
  if (!p.alloc && !p.unplannedJson && p.eta && daysBetween(ctx.today, p.eta) <= 3) a.push({ t: 'Set allocation' });
  return a.sort(function (x, y) { return (y.u ? 1 : 0) - (x.u ? 1 : 0); });
}

function qtyStatus(p, st, ctx) {
  if (p.closed) return { cls: 'p-grey', t: 'Received before the hub', sub: String(p.closed) };
  var s = st.sub[p.id], pl = st.pl[p.id];
  if (s && diffs(p, s, ctx).length) return { cls: 'p-dmg', t: gapText(p, s, ctx) };
  if (s && num(s.cs) === 0 && num(p.cs) == null) return { cls: 'p-dmg', t: 'Did not arrive (0 cs)' };
  if (s && num(p.cs) == null) return { cls: 'p-grey', t: 'Received · cases not compared', sub: 'No expected cases on the PO line' };
  if (s) return { cls: 'p-done', t: 'Matches PO', sub: pl && diffs(p, pl, ctx).length ? 'Packing list differed' : '' };
  if (pl && diffs(p, pl, ctx).length) return { cls: 'p-dmg', t: 'Packing list ≠ PO' };
  if (pl) return { cls: 'p-done', t: 'Packing list OK' };
  return { cls: 'p-grey', t: num(p.cs) == null ? 'No expected cases' : 'Not checked' };
}

function filesOf(p, st) {
  var out = [];
  var add = function (arr, old) { (arr || []).forEach(function (f) { out.push({ cat: f.cat, name: f.name, url: f.url, old: !!old }); }); };
  if (p.unplannedJson) add(p.unplannedJson.files, false);
  (st.delayLog[p.id] || []).concat(st.delay[p.id] ? [st.delay[p.id]] : []).forEach(function (d) { add(d.filesJson, false); });
  (st.hist[p.id] || []).forEach(function (h) { add(h.filesJson, true); });
  if (st.sub[p.id]) add(st.sub[p.id].filesJson, false);
  return out;
}

function timeline(p, st, ctx) {
  var s = st.sub[p.id], pl = st.pl[p.id], sl = sla(p, st, ctx), unpl = !!p.unplannedJson;
  var late = !s && sl && (sl.lvl === 'nudge' || sl.lvl === 'esc');
  var dd = s && p.eta ? daysBetween(p.eta, s.recv) : 0;
  var steps = [
    ['On the upcoming list', unpl ? 'No: arrived without notice' : (p.source === 'manual' ? 'Added by ' + p.addedBy : 'From the shipment sheet'), !unpl, unpl],
    ['Packing list received', pl ? 'Entered by ' + pl.by : (st.slip[p.po] && st.slip[p.po].status === 'FOUND' ? 'Packing slip on file' : ''), !!pl || !!(st.slip[p.po] && st.slip[p.po].status === 'FOUND'), false],
    ['Arrived at ' + (p.wh || '?'), p.closed ? 'Before the hub (' + p.closed + ')' : s ? s.recv + (dd ? ' (' + Math.abs(dd) + ' d ' + (dd < 0 ? 'early' : 'late') + ')' : '') : st.delay[p.id] ? 'Delay: ' + REASON_EN[st.delay[p.id].reason] : unpl ? 'Reported ' + p.unplannedJson.at : (p.eta ? 'ETA ' + p.eta : ''), !!s || unpl || !!p.closed, late],
    ['Counted and documented', s ? s.by : '', !!s, false],
    ['Quantity verified', s ? (diffs(p, s, ctx).length ? gapText(p, s, ctx) : 'Matches PO') : '', !!s && !diffs(p, s, ctx).length, !!s && diffs(p, s, ctx).length > 0],
    ['Entered in SAP / WMS', p.sap ? p.sap : notInSap(p, st, ctx) ? 'Missing ' + daysBetween(s.recv, ctx.today) + ' d' : '', !!p.sap, notInSap(p, st, ctx)],
    ['QC released', p.qc === 'Released' ? p.relAt : p.qc === 'Locked' ? 'Locked' + (p.qcAt ? ' ' + daysBetween(p.qcAt, ctx.today) + ' d' : '') : '', p.qc === 'Released', false]
  ];
  return steps.map(function (z) { return { t: z[0], when: z[1], done: !!z[2] && !z[3], bad: !!z[3] }; });
}

function claimText(p, st, ctx) {
  var s = st.sub[p.id], pl = st.pl[p.id], lr = st.lr[p.id], c = countsList(p, st), is = st.issue[p.id];
  var ev = filesOf(p, st).concat((st.docs[p.id] || []).map(function (d) { return { cat: d.cat, name: d.name }; }));
  var q = function (z) { return z.plt + ' plt / ' + z.cs + ' cs / ' + (z.lbs === '' || z.lbs == null ? '— lb (from photos / SAP)' : z.lbs + ' lb'); };
  var sl = shipLines(p, st), holder = s && s.grpPlt !== '' && s.grpPlt != null;
  var o = ['CLAIM / NCR SUMMARY · PO ' + p.po + (p.item ? ' · item ' + p.item : ''), 'Supplier: ' + p.supplier,
    sl.length > 1 ? 'Shipment: ' + (p.cntr ? 'container ' + p.cntr : 'PO ' + p.po) + ' · ' + sl.length + ' items' + (holder ? ' · the pallet count below is for the WHOLE shipment (' + s.grpPlt + ' expected)' : '') : '', 'Item: ' + p.item + ' ' + p.desc + (s && s.lot ? ' · Lot ' + s.lot : '') + (s && s.exp ? ' · Exp ' + s.exp : ''),
    'Received at: ' + p.wh + ' on ' + (s ? s.recv : '(not yet)') + (p.cntr ? ' · Container ' + p.cntr : '') + ' · ' + p.load, '', 'Quantities', '  PO:              ' + q(p)];
  if (pl) o.push('  Packing list:    ' + q(pl));
  if (lr) o.push('  Loading report:  ' + q(lr));
  c.forEach(function (z) { o.push('  ' + z.l + ': ' + q(z.v) + ' (' + z.by + ')'); });
  if (s && diffs(p, s, ctx).length) o.push('  Result: ' + gapText(p, s, ctx) + ' vs PO');
  o.push('  Physical counts: ' + c.length + ' of 3 required by Purchasing (first count and recount by the 3PL count as 2)' + (c.length < 3 ? ' (ask for another count before sending)' : ''), '');
  (st.hist[p.id] || []).forEach(function (h, i) { o.push('Submission ' + (i + 1) + ' (' + h.by + ', ' + h.recv + '): condition ' + h.cond + (h.damage ? ' · ' + h.dmgText : '')); });
  if (s) { o.push('Load condition on arrival: ' + s.cond); if (s.damage) o.push('Damage: ' + s.dmgText); if (s.diffNote) o.push('Receiver note: ' + s.diffNote); if (s.comment) o.push('Receiver comment: ' + s.comment); }
  if (is && is.decision && is.decision !== 'Pending') o.push('Dock decision: ' + is.decision + (is.decisionNote ? ' · ' + is.decisionNote : ''));
  o.push('', 'Conditions on the PO (from SAP): ' + (p.instrSap || '—')); if (p.extra) o.push('Additional packaging conditions: ' + p.extra);
  o.push('', 'Evidence (' + ev.length + ' files):'); ev.forEach(function (f) { o.push('  - ' + f.cat + (f.old ? ' (earlier submission)' : '') + ': ' + f.name + (f.url ? ' ' + f.url : '')); });
  return o.join('\n');
}

/** Everything the internal hub needs for one PO. Internal users only. */
function hubRow(p, st, ctx) {
  var s = st.sub[p.id], sl = sla(p, st, ctx), is = st.issue[p.id];
  return {
    id: p.id, po: p.po, item: p.item, wh: p.wh, desc: p.desc, supplier: p.supplier, eta: p.eta, clockFrom: p.clockFrom, cntr: p.cntr, plt: p.plt, cs: p.cs, lbs: p.lbs, load: p.load,
    alloc: p.alloc, prodQty: p.prodQty, salesQty: p.salesQty, so: p.so, qc: p.qc, qcAt: p.qcAt, relAt: p.relAt, sap: p.sap, instrSap: p.instrSap, extra: p.extra, tplInstr: p.tplInstr,
    source: p.source, addedBy: p.addedBy, unplanned: p.unplannedJson && !truthy(p.planOk) ? p.unplannedJson : null, nudged: p.nudged, floor: isFloor(p), isTPL: isTPL(p.wh),
    sla: sl, slaText: sl ? SLA_EN[sl.lvl] : '', sub: s ? subView(s) : null, hasHist: (st.hist[p.id] || []).length > 0,
    delay: st.delay[p.id] ? { reason: st.delay[p.id].reason, reasonText: REASON_EN[st.delay[p.id].reason], owner: DELAY_OWNER[st.delay[p.id].reason], note: st.delay[p.id].note, newDate: st.delay[p.id].newDate, by: st.delay[p.id].by, at: st.delay[p.id].at } : null,
    req: st.req[p.id] ? { at: st.req[p.id].at } : null,
    issue: is ? { type: is.type, owner: is.owner, status: is.status, opened: is.opened, urgent: truthy(is.urgent), decision: is.decision, decisionNote: is.decisionNote, overdue: issueOverdue(is, ctx) } : null,
    needs: needsAction(p, st, ctx), qty: qtyStatus(p, st, ctx), avail: availSales(p, st, ctx), notInSap: notInSap(p, st, ctx),
    pl: st.pl[p.id] || null, lr: st.lr[p.id] || null,
    slip: st.slip[p.po] ? { status: st.slip[p.po].status, url: st.slip[p.po].url, source: st.slip[p.po].source, action: st.slip[p.po].action, has: !!st.slip[p.po].fileId } : null,
    masterRecv: st.mseen[p.id] ? st.mseen[p.id].recv : '',
    closed: p.closed || '', ship: shipKey(p), shipN: shipLines(p, st).length, manual: p.manual || ''
  };
}
function subView(s) { return { recv: s.recv, plt: s.plt, cs: s.cs, lbs: s.lbs, by: s.by, at: s.at, first: s.firstJson, diffNote: s.diffNote, lot: s.lot, exp: s.exp, cond: s.cond, damage: truthy(s.damage), dmgText: s.dmgText, comment: s.comment }; }

function hubDetail(p, st, ctx) {
  var r = hubRow(p, st, ctx);
  r.counts = countsList(p, st); r.timeline = timeline(p, st, ctx); r.files = filesOf(p, st);
  // documents and notes belong to the PO, so show those of every line of the same PO
  var same = st.pos.filter(function (x) { return x.po === p.po; });
  r.docs = []; r.notes = [];
  same.forEach(function (x) { r.docs = r.docs.concat(st.docs[x.id] || []); r.notes = r.notes.concat(st.notes[x.id] || []); });
  r.notes.sort(function (a, b) { return String(a.at) < String(b.at) ? -1 : 1; });
  r.lines = shipLines(p, st).map(function (x) { return { id: x.id, po: x.po, item: x.item, desc: x.desc, cs: x.cs, lbs: x.lbs, done: isDone(x, st), sub: st.sub[x.id] ? { cs: st.sub[x.id].cs, lot: st.sub[x.id].lot } : null }; });
  r.shipPlt = shipPallets(shipLines(p, st)); r.askP = shipAskP(p, st); r.fresh = shipFresh(p, st);
  r.claim = (st.sub[p.id] || (st.hist[p.id] || []).length) ? claimText(p, st, ctx) : '';
  return r;
}

/** The ONLY data a 3PL ever receives. Allow-list. No expected cases or lb, no allocation, no notes, no other 3PL. */
function portalRow(p, st, ctx) {
  var s = st.sub[p.id], sl = sla(p, st, ctx), dl = st.delay[p.id];
  return {
    id: p.id, ship: shipKey(p), askP: shipAskP(p, st), fresh: shipFresh(p, st), shipPlt: shipPallets(shipLines(p, st)), po: p.po, item: p.item, desc: p.desc, eta: p.eta, cntr: p.cntr, plt: p.plt, load: p.load, floor: isFloor(p), instr: p.tplInstr || '',
    sla: sl ? { lvl: sl.lvl } : null, req: !!st.req[p.id], slip: slipFor3pl(p, st, ctx),
    delay: dl ? { reason: dl.reason, newDate: dl.newDate } : null,
    sent: s ? { recv: s.recv, at: s.at, plt: s.plt, cs: s.cs, lbs: s.lbs, first: s.firstJson, diffNote: s.diffNote, lot: s.lot, exp: s.exp, cond: s.cond, damage: truthy(s.damage), dmgText: s.dmgText, comment: s.comment,
      files: (s.filesJson || []).map(function (f) { return { cat: f.cat, name: f.name }; }) } : null
  };
}

// ==================== Api.gs ====================
/**
 * Api.gs — the only functions the web pages can call (google.script.run).
 * Every call checks who is asking. 3PL calls need their private link token and only ever
 * touch POs for their own warehouse. Hub calls need a Google account in the Tootsi domain.
 */

// ---------- who is asking ----------
function hubUser_() {
  var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  var dom = String(config_().DOMAIN || 'tootsi.com').toLowerCase();
  if (!email || email.split('@')[1] !== dom) throw new Error('Access denied. Sign in with your @' + dom + ' Google account.');
  return email;
}
/**
 * 3PL access. A 3PL signs in with a username and password (Users tab).
 * The server gives back a session code that lasts SESSION_HOURS (default 12 h, one shift).
 * Every portal call sends that code. The server checks it each time, so nobody can ask for another 3PL's data.
 */
function tplUser_(sid) {
  var t = String(sid || '');
  if (t.length < 20) throw new Error('SESSION');
  var now = Date.now();
  var s = readAll_('Sessions').filter(function (r) { return r.sid === t && Number(r.expires) > now; })[0];
  if (!s) throw new Error('SESSION');
  var u = readAll_('Users').filter(function (r) { return r.token === s.token && truthy(r.active); })[0];
  if (!u) throw new Error('SESSION');
  return u; // {wh, name, email}
}

function hash_(salt, pass) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/** Sign in. Returns a session code. 10 wrong passwords lock the username for 10 minutes. */
function portalLogin(user, pass) {
  user = String(user || '').toLowerCase().trim().slice(0, 64);
  var key = 'fail_' + user.replace(/[^a-z0-9._-]/g, ''), cache = CacheService.getScriptCache();
  var fails = Number(cache.get(key) || 0);
  if (fails >= 10) throw new Error('LOCKED');
  var u = readAll_('Users').filter(function (r) { return truthy(r.active) && r.user && String(r.user).toLowerCase() === String(user || '').toLowerCase().trim(); })[0];
  if (!u || !u.passHash || hash_(u.salt, String(pass || '')) !== u.passHash) {
    cache.put(key, String(fails + 1), 600);
    throw new Error('BADLOGIN');
  }
  cache.remove(key);
  var sid = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var hours = Number(config_().SESSION_HOURS || 12);
  return withLock_(function () {
    // drop expired sessions so the tab stays small
    var now = Date.now();
    deleteWhere_('Sessions', function (r) { return Number(r.expires) < now; });
    append_('Sessions', { sid: sid, token: u.token, expires: now + hours * 3600000, at: nowStamp_() });
    u.lastLogin = nowStamp_(); upsert_('Users', 'token', u);
    logAction_(u.name, 'login', '', '');
    return { sid: sid, wh: u.wh, name: u.name, lang: String(config_()['LANG_' + u.wh] || 'fr') };
  });
}

/** The page reports where its code was damaged (written to the Log tab, at most 5 times an hour). */
function portalDiag(token, info) {
  var u = tplUser_(token); diagLog_(u.name, info); return true;
}
function hubDiag(info) {
  var me = hubUser_(); diagLog_(me, info); return true;
}
function diagLog_(who, info) {
  var c = CacheService.getScriptCache(), n = Number(c.get('diag') || 0); if (n >= 5) return; c.put('diag', String(n + 1), 3600);
  logAction_(who, 'page-diag', '', String(info || '').slice(0, 1500));
}
function portalLogout(sid) {
  try { tplUser_(sid); } catch (e) { return true; } // only a valid session can be signed out
  withLock_(function () { deleteWhere_('Sessions', function (r) { return r.sid === sid; }); });
  return true;
}

// ---------- state ----------
function ctx_() {
  var cfg = config_(), d = new Date();
  return { cfg: cfg, today: today_(), nowH: Number(Utilities.formatDate(d, 'America/Toronto', 'H')), stamp: nowStamp_() };
}
function loadState_() {
  clearCache_();
  return indexState({ slips: readAll_('Slips'), mseen: readAll_('MasterSeen'), pos: readAll_('POs'), subs: readAll_('Submissions'), issues: readAll_('Issues'), notes: readAll_('Notes'), delays: readAll_('Delays'), recounts: readAll_('Recounts'), docs: readAll_('Docs'), supq: readAll_('SupplierQty') });
}
/** Recompute issues and save what changed. Also emails new urgent issues. */
function sync_(st, ctx) {
  var r = syncIssues(st, ctx), mail = [];
  r.writes.forEach(function (is) {
    var isNew = is._new, esc = is._esc, chg = is._chg; delete is._new; delete is._esc; delete is._chg;
    upsert_('Issues', 'id', is);
    var quiet = /^Not received at/.test(is.type) && !truthy(is.urgent); // 1 day late: the hub shows it; email only once escalated
    if ((isNew || esc || chg) && !quiet) mail.push(is); // email only when an issue is created or becomes urgent
  });
  notifyIssues_(mail, ctx);
  var seenN = {};
  r.notes.forEach(function (n) { var k = String(n.id).split('|')[0] + '|' + n.text.replace(/\(.*\)/, ''); if (seenN[k]) return; seenN[k] = true; append_('Notes', n); });
  return st;
}
var _locked = false;
/** Run fn while holding the script lock, so two people never write at the same time. Safe to call inside itself. */
function withLock_(fn) {
  if (_locked) return fn();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('The hub is busy saving someone else\'s change. Please try again in a few seconds.');
  _locked = true;
  try { return fn(); } finally { _locked = false; lock.releaseLock(); }
}
/** Page loads: recompute issues; take the lock only if something must be saved. */
function syncSafe_(ctx) {
  var st = loadState_(), r = syncIssues(st, ctx);
  if (!r.writes.length && !r.notes.length) return loadState_();
  return withLock_(function () { return sync_(loadState_(), ctx); });
}
function findPO_(st, id) { var p = st.pos.filter(function (x) { return String(x.id) === String(id); })[0]; if (!p) throw new Error('PO not found.'); return p; }
function savePO_(p, ctx) { p.updatedAt = ctx.stamp; upsert_('POs', 'id', p); }
function note_(id, who, text, ctx, ext) { append_('Notes', { id: id, who: who, text: text, at: ctx.stamp, ext: !!ext }); }

// =====================================================================
// 3PL PORTAL
// =====================================================================
function portalLoad(token) {
  var u = tplUser_(token), ctx = ctx_();
  var st = syncSafe_(ctx);
  var rows = st.pos.filter(function (p) { return p.wh === u.wh && !p.closed && (!p.unplannedJson || truthy(p.planOk) || st.sub[p.id]); })
    .map(function (p) { return portalRow(p, st, ctx); });
  return { wh: u.wh, name: u.name, today: ctx.today, rows: rows };
}
/** A PO line of this 3PL, or "PO not found" (same message either way, so nobody can probe other warehouses). */
function myPO_(st, id, u) { var p = st.pos.filter(function (x) { return String(x.id) === String(id); })[0]; if (!p || p.wh !== u.wh || p.closed || (p.unplannedJson && !truthy(p.planOk))) throw new Error('PO not found.'); return p; }

/** Step 1 of any upload: one file per call, so big containers never hit Google's size limit. Returns a file id, never a link. */
function portalUpload(token, f) { var u = tplUser_(token); uploadQuota_(u.wh); return saveIncoming_(u.wh, checkFile_(f)); }

/** Blind check for a whole shipment. Answers which lines (and whether the pallet total) are off, never the expected numbers.
 *  At most 2 wrong checks per receiving round, recorded on the server. */
function portalCheckShip(token, id, f) {
  var u = tplUser_(token);
  return withLock_(function () { var ctx = ctx_(), st = loadState_(); return checkShipCall_(myPO_(st, id, u), st, ctx, f || {}, u.name + ' receiving'); });
}
function portalSubmitShip(token, id, f) {
  var u = tplUser_(token);
  var out = withLock_(function () { var ctx = ctx_(), st = loadState_(); return submitShip_(myPO_(st, id, u), st, ctx, f || {}, u.name + ' receiving', true, u.wh); });
  finishMoves_(out._moves); delete out._moves;
  return out;
}

function portalDelay(token, id, d) {
  var u = tplUser_(token); d = d || {};
  var ctx0 = ctx_(); myPO_(loadState_(), id, u);
  if (!DELAY_OWNER[d.reason] || String(d.note || '').trim().length < 5) throw new Error('Reason and details are required.');
  var newDate = String(d.newDate || '');
  if (newDate && (!/^\d{4}-\d{2}-\d{2}$/.test(newDate) || newDate < ctx0.today || newDate > addDays(ctx0.today, 21))) throw new Error('The new date must be within the next 3 weeks.');
  // photos are saved before taking the lock (Drive is slow)
  var files = (d.files || []).slice(0, 10).map(function (f) { var s = saveFile_(id, checkFile_(f)); s.cat = 'Delay photo'; return s; });
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = myPO_(st, id, u);
    var lines = shipLines(p, st).filter(function (x) { return x.wh === u.wh && !isDone(x, st) && !st.delay[x.id]; });
    var note = String(d.note).trim().slice(0, 1000);
    lines.forEach(function (x, i) {
      append_('Delays', { delayId: Utilities.getUuid(), id: x.id, reason: d.reason, note: note, newDate: newDate, by: u.name + ' receiving', at: ctx.stamp, date: ctx.today, closedAt: '', filesJson: i ? [] : files });
      var is = st.issue[x.id]; if (is && is.status !== 'Open') { is.status = 'Open'; is.opened = ctx.today; upsert_('Issues', 'id', is); }
    });
    notePOs_(lines, u.name + ' receiving', 'Delay reported: ' + REASON_EN[d.reason] + '. ' + note + (newDate ? ' New date: ' + newDate : ''), ctx, true);
    logAction_(u.name, 'delay', p.id, d.reason);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

function portalSolveDelay(token, id) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = myPO_(st, id, u);
    var lines = shipLines(p, st).filter(function (x) { return x.wh === u.wh && st.delay[x.id]; });
    if (!lines.length) return portalLoad(token);
    var reason = st.delay[lines[0].id].reason;
    lines.forEach(function (x) {
      var dl = st.delay[x.id], from = dl.newDate && dl.newDate > ctx.today ? dl.newDate : ctx.today;
      dl.closedAt = ctx.stamp; upsert_('Delays', 'delayId', dl);
      x.clockFrom = from; savePO_(x, ctx);
    });
    notePOs_(lines, u.name + ' receiving', 'Delay solved (' + REASON_EN[reason] + '). Receiving clock restarts ' + lines[0].clockFrom + '.', ctx, true);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

function portalUnplanned(token, d) {
  var u = tplUser_(token); d = d || {};
  if (!(d.files && d.files.length) || !String(d.desc || '').trim()) throw new Error('Paperwork photo and product are required.');
  var po = String(d.po || '').replace(/[^\d-]/g, '').slice(0, 15);
  var mine = po ? loadState_().pos.filter(function (x) { return x.po === po && x.wh === u.wh && !x.closed; }) : [];
  if (mine.length) throw new Error('This PO is already in your list. Look for it under "To receive".');
  // photos saved before the lock; Tootsi (Logistics) confirms the PO in the hub, which links it to the planned lines
  var folder = 'unplanned', files = d.files.slice(0, 10).map(function (f) { var s = saveFile_(folder, checkFile_(f)); s.cat = 'Paperwork (arrived not on list)'; return s; })
    .concat((d.loadFiles || []).slice(0, 10).map(function (f) { var s = saveFile_(folder, checkFile_(f)); s.cat = 'Load photo (arrived not on list)'; return s; }));
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_();
    var n = st.pos.filter(function (p) { return /^NEW-/.test(p.id); }).length + 1, id = 'NEW-' + n;
    var p = { id: id, po: po || id, item: '', wh: u.wh, desc: String(d.desc).trim().toUpperCase().slice(0, 200), supplier: String(d.sup || '').trim().slice(0, 200) || 'Unknown', eta: ctx.today, clockFrom: '', cntr: String(d.cntr || '').trim().toUpperCase().slice(0, 30),
      plt: num(d.plt) != null && num(d.plt) >= 0 && num(d.plt) <= 200 ? num(d.plt) : '', cs: '', lbs: '', load: 'Unknown', alloc: '', qc: 'Not received', source: 'unplanned', addedBy: u.name + ' receiving',
      unplannedJson: { by: u.name + ' receiving', at: ctx.stamp, files: files }, planOk: false };
    savePO_(p, ctx);
    note_(id, u.name + ' receiving', 'Arrived at ' + u.wh + ' but not on the upcoming list. Paperwork attached.', ctx, true);
    logAction_(u.name, 'unplanned', id, d.desc);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

// ---------- receiving a shipment (3PL portal and Tootsi receiving at BM/LY/HB) ----------
/** Lines to receive now: every line of the shipment not yet received. */
function openLines_(p, st) { return shipLines(p, st).filter(function (x) { return x.wh === p.wh && !isDone(x, st); }); }
/** Pallets are counted once per shipment, and only on the first receiving (not on a recount of one line). */
/** Pallets are counted on the first receiving of a shipment, and again when the line that holds the pallet count is recounted. */
function palletsAsked_(p, st) { return shipAskP(p, st); }
function palletsAskedOld_(p, st) {
  if (shipFresh(p, st)) return true;
  return openLines_(p, st).some(function (x) { return (st.hist[x.id] || []).some(function (s) { return s.plt !== '' && s.plt != null; }); });
}
/** A receiving round starts at the last recount request, and never looks back more than 7 days. */
/** A receiving round: the checks made after the last recount request (each request stores how many checks existed then),
 *  and never more than 7 days back. */
function roundStart_(lines, st, ctx) {
  var t = { at: addDays(ctx.today, -7) + ' 00:00', base: 0 };
  lines.forEach(function (x) { var q = st.req[x.id]; if (q && num(q.base) != null && num(q.base) > t.base) t.base = num(q.base); });
  return t;
}
function wrongChecks_(key, since) {
  return readAll_('Checks').filter(function (c) { return c.key === key; }).slice(since.base || 0)
    .filter(function (c) { return c.off !== '' && String(c.at) >= since.at; });
}

/** Compare a shipment count to what is expected. Returns the off lines and whether the pallet total is off. */
function shipDiff_(p, st, ctx, f) {
  var lines = openLines_(p, st), byId = {};
  (f.lines || []).forEach(function (L) { byId[L.id] = L; });
  var off = [], askP = palletsAsked_(p, st), grp = shipPallets(shipLines(p, st));
  var pltOff = askP && !isFloor(p) && grp != null && num(f.plt) !== grp;
  lines.forEach(function (x) {
    var L = byId[x.id] || {};
    if (diffs(x, { plt: null, cs: num(L.cs), lbs: num(L.lbs) }, ctx).length) off.push(x.id);
  });
  return { lines: lines, off: off, plt: pltOff, askP: askP, grp: grp };
}
function checkShipCall_(p, st, ctx, f, by) {
  var r = shipDiff_(p, st, ctx, f), key = shipKey(p), since = roundStart_(r.lines, st, ctx);
  var wrong = wrongChecks_(key, since);
  if (wrong.length >= 2) { var last = wrong[wrong.length - 1]; return { match: false, limit: true, off: String(last.off).split(',').filter(String) }; }
  var offList = r.off.slice(); if (r.plt) offList.push('pallets');
  if (offList.length) append_('Checks', { key: key, at: ctx.stamp, by: by, off: offList.join(','), countJson: { plt: num(f.plt), lines: (f.lines || []).map(function (L) { return { id: L.id, cs: num(L.cs), lbs: num(L.lbs) }; }) } });
  return { match: !offList.length, off: offList, tries: wrong.length + (offList.length ? 1 : 0) };
}

/** The line that carries the pallet count of a shipment: the one that carried it before (recount), else the first open line. */
function palletHolder_(lines, st) {
  var h = lines.filter(function (x) { return (st.hist[x.id] || []).some(function (s) { return s.plt !== '' && s.plt != null; }); })[0];
  return (h || lines[0] || {}).id;
}

/**
 * Save a shipment receipt. All checks first, then all rows; photos are moved into the PO folder AFTER the lock
 * (see finishMoves_), so a big container never blocks other users and a failure never leaves half a receipt.
 */
function submitShip_(p, st, ctx, f, by, ext, wh) {
  var lines = openLines_(p, st), key = shipKey(p);
  if (!lines.length) throw new Error('Already received. Contact Tootsi Inventory Control to change it.');
  var byId = {}; (f.lines || []).forEach(function (L) { byId[L.id] = L; });
  if (lines.some(function (x) { return !byId[x.id]; })) throw new Error('Enter every item of this shipment (0 cases if an item did not come).');
  if (!f.recv || !/^\d{4}-\d{2}-\d{2}$/.test(f.recv) || f.recv > ctx.today || f.recv < addDays(ctx.today, -60)) throw new Error('Check the date received (not in the future, not more than 60 days ago).');
  var askP = palletsAsked_(p, st);
  if (askP && (num(f.plt) == null || num(f.plt) < 0 || num(f.plt) > 200)) throw new Error('Enter the number of pallets.');
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs), lb = num(L.lbs);
    if (cs == null || cs < 0 || cs > 1000000 || (lb != null && (lb < 0 || lb > 200000))) throw new Error('Enter cases (0 or more) for item ' + x.item + '.');
    if (cs > 0 && (!String(L.lot || '').trim() || !/^\d{4}-\d{2}-\d{2}$/.test(String(L.exp || '')) || L.exp <= f.recv || L.exp > addDays(f.recv, 3660))) throw new Error('Lot and expiry are required for item ' + x.item + '. Expiry must be after the date received.');
    if (cs > 0 && !(L.label || []).length) throw new Error('Label photo missing for item ' + x.item + '.');
  });
  if (!f.cond) throw new Error('Choose the load condition.');
  var bad = f.cond !== 'Good';
  if (bad && String(f.dmgText || '').trim().length < 10) throw new Error('Describe the damage.');
  var fresh = shipFresh(p, st); // the "before unloading" photo only makes sense on the first receiving
  if ((fresh && !(f.load || []).length) || !(f.doc || []).length || (bad && !(f.dmg || []).length)) throw new Error('Photos and documents are missing.');
  // 1. every photo must be a real upload of this warehouse (checked BEFORE the count, so a fake submit learns nothing)
  var files = {}, verify = function (ids) { return (ids || []).slice(0, 20).map(function (fid) { if (!files[fid]) files[fid] = verifyFile_(fid, wh); return files[fid]; }); };
  var sharedRaw = [[f.load, 'Load photo (before unloading)'], [f.doc, 'Receiving document'], [f.dmg, 'Damage photo']];
  sharedRaw.forEach(function (z) { verify(z[0]); });
  lines.forEach(function (x) { verify(byId[x.id].label); });
  var used = CacheService.getScriptCache();
  Object.keys(files).forEach(function (fid) { if (used.get('used_' + fid)) throw new Error('A photo was already used in another receipt. Please add it again.'); });
  // 2. the count: a submit with a different count is recorded like a check (same 2-try limit)
  var r = shipDiff_(p, st, ctx, f), offAll = r.off.length || r.plt;
  var wrong = wrongChecks_(key, roundStart_(lines, st, ctx));
  var record = function () { var offList = r.off.slice(); if (r.plt) offList.push('pallets'); append_('Checks', { key: key, at: ctx.stamp, by: by, off: offList.join(','), countJson: { plt: num(f.plt), lines: (f.lines || []).map(function (L) { return { id: L.id, cs: num(L.cs), lbs: num(L.lbs) }; }) } }); };
  if (offAll && !wrong.length) { record(); throw new Error('RECOUNT'); }
  if (offAll && String(f.diffNote || '').trim().length < 10) { record(); throw new Error('Explain the difference.'); }
  // after 2 wrong counts the answer is the same whatever the count: a note is required, so retrying teaches nothing
  if (wrong.length >= 2 && String(f.diffNote || '').trim().length < 10) throw new Error('Explain the difference.');
  // first count = the first wrong check of this round, from the server (never from the browser)
  var first = wrong.length ? (wrong[0].countJson || {}) : null, firstLine = {};
  if (first) (first.lines || []).forEach(function (L) { firstLine[L.id] = L; });
  // 3. write all rows
  var moves = [], lead = lines[0].id, holder = askP ? palletHolder_(lines, st) : null;
  var asFiles = function (ids, cat, poId) { return (ids || []).slice(0, 20).map(function (fid) { var v = files[fid]; moves.push({ fid: fid, who: wh, poId: poId }); return { cat: cat, name: v.name, url: v.url, fileId: fid }; }); };
  var shared = [];
  sharedRaw.forEach(function (z) { shared = shared.concat(asFiles(z[0], z[1], lead)); });
  var multi = shipLines(p, st).length > 1;
  var comment = String(f.comment || '').trim().slice(0, 500), note = String(f.diffNote || '').trim().slice(0, 1000);
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs), lb = num(L.lbs), isHolder = x.id === holder;
    var fl = firstLine[x.id], lineOff = r.off.indexOf(x.id) >= 0, pOff = isHolder && r.plt;
    var hist = st.hist[x.id] || [];
    append_('Submissions', { subId: Utilities.getUuid(), id: x.id, seq: hist.length + 1, superseded: false, by: by, at: ctx.stamp, recv: f.recv,
      // the pallet count of the shipment sits on ONE line only, so a pallet gap is one issue and one claim, not one per item
      plt: isHolder ? num(f.plt) : '', cs: cs, lbs: lb == null ? '' : lb,
      firstJson: first && (lineOff || pOff) ? { plt: pOff ? first.plt : null, cs: fl ? fl.cs : null, lbs: fl ? fl.lbs : null } : null,
      diffNote: (lineOff || pOff) ? note : '', lot: String(L.lot || '').trim().slice(0, 60), exp: L.exp || '', cond: f.cond, damage: bad, dmgText: bad ? String(f.dmgText).trim().slice(0, 1000) : '', comment: comment,
      filesJson: shared.concat(asFiles(L.label, 'Label photo', x.id)), ship: key, grpPlt: isHolder && multi && r.grp != null ? r.grp : '' });
    if (st.req[x.id]) { st.req[x.id].active = false; upsert_('Recounts', 'id', st.req[x.id]); }
    if (st.delay[x.id]) { var dl = st.delay[x.id]; dl.closedAt = ctx.stamp; upsert_('Delays', 'delayId', dl); }
    if (!x.qc || x.qc === 'Not received') { x.qc = 'Locked'; x.qcAt = f.recv; savePO_(x, ctx); }
  });
  if (comment) notePOs_(lines, by, comment, ctx, ext);
  logAction_(by, 'receipt', key, lines.length + ' lines' + (askP ? ' / ' + num(f.plt) + ' plt' : '') + (bad ? ' / ' + f.cond : ''));
  Object.keys(files).forEach(function (fid) { used.put('used_' + fid, '1', 21600); });
  queueMaster_(lines, byId, f, wh, ctx);
  sync_(loadState_(), ctx);
  return { ok: true, flagged: !!offAll || bad, n: lines.length, _moves: moves };
}

/** After a 3PL receipt: queue the master-sheet fills (only empty cells will ever be filled; see syncMaster_ in Jobs). */
function queueMaster_(lines, byId, f, wh, ctx) {
  if (!isTPL(wh) || !config_().MASTER_SHEET_ID || String(config_().MASTER_MODE || 'preview') === 'off') return;
  var docs = ''; try { docs = poFolder_(lines[0].id).getUrl(); } catch (e) {}
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs); if (!(cs > 0)) return;
    var vals = { item: x.item, desc: x.desc, recv: f.recv, exp: L.exp, units: cs, docs: docs };
    Object.keys(vals).forEach(function (k) {
      if (vals[k] === '' || vals[k] == null) return;
      upsertWhere_('MasterSync', function (r) { return r.key === x.id + '|' + k; }, { key: x.id + '|' + k, at: ctx.stamp, po: x.po, item: x.item, wh: wh, field: k, value: String(vals[k]), status: 'pending', row: '' });
    });
  });
}

/** Supplier packing slip of a PO (from Arash's extractor). The 3PL gets it only for its own POs, and only if SLIP_TO_3PL = yes. */
function portalSlip(token, id) {
  var u = tplUser_(token), st = loadState_();
  // a sent line is no longer "open", so look it up directly (still only this 3PL's own, planned PO)
  var p = st.pos.filter(function (x) { return String(x.id) === String(id); })[0];
  if (!p || p.wh !== u.wh || (p.unplannedJson && !truthy(p.planOk))) throw new Error('PO not found.');
  if (!slipFor3pl(p, st, ctx_())) throw new Error('No packing slip yet.');
  return slipFile_(st.slip[p.po].fileId);
}
function hubSlip(id) { hubUser_(); var st = loadState_(), p = findPO_(st, id), s = st.slip[p.po]; if (!s || !s.fileId) throw new Error('No packing slip yet.'); return slipFile_(s.fileId); }

/** One note per PO (notes are shown per PO), even when a shipment has many lines. */
function notePOs_(lines, who, text, ctx, ext) { var seen = {}; lines.forEach(function (x) { if (seen[x.po]) return; seen[x.po] = true; note_(x.id, who, text, ctx, ext); }); }

// =====================================================================
// INTERNAL HUB (Tootsi domain only)
// =====================================================================
function hubLoad() {
  var me = hubUser_(), ctx = ctx_();
  var st = syncSafe_(ctx);
  return { me: me, today: ctx.today, rows: st.pos.map(function (p) { return hubRow(p, st, ctx); }), cfg: { nudgeH: ctx.cfg.SLA_NUDGE_H || 9, escH: ctx.cfg.SLA_ESC_H || 18, portalUrl: String(ctx.cfg.PORTAL_URL || ''), goLive: String(ctx.cfg.GO_LIVE || '') } };
}
/** Staff uploads (BM/LY/HB receiving photos). One file per call. */
function hubUpload(f) { var me = hubUser_(); uploadQuota_('HUB' + me); return saveIncoming_('HUB', checkFile_(f)); }
/** At most 200 uploads per hour per warehouse, so nobody can fill the Drive. */
function uploadQuota_(who) {
  var c = CacheService.getScriptCache(), k = 'upq_' + who + '_' + Utilities.formatDate(new Date(), 'America/Toronto', 'yyyyMMddHH'), n = Number(c.get(k) || 0);
  if (n >= 200) throw new Error('Too many photos this hour. Please wait a few minutes.');
  c.put(k, String(n + 1), 3700);
}
function hubDetailLoad(id) {
  hubUser_(); var ctx = ctx_(), st = loadState_();
  return hubDetail(findPO_(st, id), st, ctx);
}
/** One entry point for all internal edits. action = what to do, d = data. */
function hubAction(id, action, d) {
  var me = hubUser_();
  if (action === 'syncNow') { // staff: pull the sheets now. Outside the hub lock, at most once every 2 minutes.
    var cc = CacheService.getScriptCache(); if (cc.get('syncNow')) throw new Error('A sync ran less than 2 minutes ago. Please wait.');
    cc.put('syncNow', me, 120); logAction_(me, 'syncNow', '', '');
    return { ok: true, sync: syncSources_() };
  }
  var preDoc = action === 'doc' ? saveFile_(id, checkFile_((d || {}).file)) : null; // Drive save outside the lock
  var out = withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = id ? findPO_(st, id) : null, is = id ? st.issue[id] : null;
    d = d || {};
    switch (action) {
      case 'who': // allocation, SO, QC, extra conditions, 3PL instructions
        if (d.alloc === 'Split') { var a = num(d.prodQty), b = num(d.salesQty); if (a == null || b == null || a < 0 || b < 0 || a + b !== num(p.cs)) throw new Error('Production + sales must equal ' + p.cs + ' cs.'); p.prodQty = a; p.salesQty = b; }
        else { p.prodQty = ''; p.salesQty = ''; }
        var oldQc = p.qc; p.alloc = d.alloc || ''; p.so = String(d.so || '').trim(); p.qc = d.qc || p.qc; p.extra = String(d.extra || '').trim(); if ('tplInstr' in d) p.tplInstr = String(d.tplInstr || '').trim();
        if (p.qc === 'Locked' && oldQc !== 'Locked') p.qcAt = ctx.today; if (p.qc === 'Released' && oldQc !== 'Released') p.relAt = ctx.today;
        savePO_(p, ctx); note_(id, me, 'Saved: ' + (p.alloc || 'no allocation') + (p.alloc === 'Split' ? ' ' + p.prodQty + '/' + p.salesQty : '') + (p.so ? ', ' + p.so : '') + ', QC ' + p.qc, ctx); break;
      case 'expected': // expected qty, destination, ETA (for POs imported without cases)
        // fields changed by hand are marked "manual": the Upcoming sheet import will not overwrite them
        var man = String(p.manual || '').split(',').filter(String), mark = function (k) { if (man.indexOf(k) < 0) man.push(k); }, ch = [];
        ['plt', 'cs', 'lbs'].forEach(function (k) { if (d[k] !== undefined && d[k] !== '') { var v = num(d[k]); if (v == null || v < 0) throw new Error('Quantities must be 0 or more.'); if (String(v) !== String(p[k])) { ch.push(k + ' ' + v); p[k] = v; if (k !== 'cs') mark(k); } } });
        if (d.wh && d.wh !== p.wh && st.sub[p.id]) throw new Error('Already received at ' + p.wh + '. The destination can no longer change.');
        if (d.wh && OWN_WH.concat(TPL_WH).indexOf(d.wh) < 0) throw new Error('Unknown warehouse.');
        if (d.eta && !/^\d{4}-\d{2}-\d{2}$/.test(d.eta)) throw new Error('Bad date.');
        // ETA, destination and pallets belong to the whole shipment: change every open line of it, not just this one
        var sib = openLines_(p, st).filter(function (x) { return x.id !== p.id; });
        var whCh = d.wh && d.wh !== p.wh;
        [p].concat(sib).forEach(function (x) {
          var mx = String(x.manual || '').split(',').filter(String), mk = function (k) { if (mx.indexOf(k) < 0) mx.push(k); };
          if (d.eta && d.eta !== x.eta) { x.eta = d.eta; mk('eta'); }
          if (d.wh && d.wh !== x.wh) { x.wh = d.wh; mk('wh'); }
          if (x !== p && d.plt !== undefined && d.plt !== '' && num(d.plt) != null && String(num(d.plt)) !== String(x.plt)) { x.plt = num(d.plt); mk('plt'); }
          if (x === p) { man.forEach(mk); }
          x.manual = mx.join(','); savePO_(x, ctx);
        });
        if (whCh) ch.push('destination ' + d.wh);
        if (d.eta) ch.push('ETA ' + d.eta);
        note_(id, me, 'Expected data updated' + (ch.length ? ': ' + ch.join(', ') : '') + (sib.length && (d.eta || d.wh) ? ' (whole shipment, ' + (sib.length + 1) + ' lines)' : ''), ctx); break;

      case 'expectedBulk': // expected cases for several lines of the same PO in one save
        var setN = 0;
        Object.keys(d.cs || {}).forEach(function (lid) {
          var x = findPO_(st, lid); if (x.po !== p.po || isDone(x, st)) return;
          var v = num(d.cs[lid]); if (v == null || v < 0) throw new Error('Cases must be 0 or more.');
          if (String(v) !== String(x.cs)) { x.cs = v; savePO_(x, ctx); setN++; }
        });
        note_(id, me, 'Expected cases set on ' + setN + ' line(s) of PO ' + p.po, ctx); break;
      case 'prehub': // the shipment physically arrived before the hub existed
        var pl2 = openLines_(p, st); if (!pl2.length) throw new Error('Nothing left to close.');
        pl2.forEach(function (x) { x.closed = ctx.today + ' · ' + me; savePO_(x, ctx); });
        notePOs_(pl2, me, 'Marked as received before the hub (' + pl2.length + ' line' + (pl2.length > 1 ? 's' : '') + '). No 3PL receiving expected.', ctx); break;
      case 'reopen': if (!p.closed) throw new Error('Not closed.'); p.closed = ''; savePO_(p, ctx); note_(id, me, 'Reopened: receiving expected again', ctx); break;
      case 'supq': // packing list / loading report quantities
        var v = { plt: num(d.plt), cs: num(d.cs), lbs: num(d.lbs) };
        if (v.cs == null || v.cs < 0 || (v.plt != null && v.plt < 0) || (v.lbs != null && v.lbs < 0)) throw new Error('Enter the cases (pallets and lb are optional).');
        if (v.plt == null) v.plt = ''; if (v.lbs == null) v.lbs = '';
        upsertSupQ_(id, d.kind === 'lr' ? 'lr' : 'pl', v, me, ctx); note_(id, me, (d.kind === 'lr' ? 'Loading report' : 'Packing list') + ' quantities: ' + v.plt + ' plt / ' + v.cs + ' cs / ' + v.lbs + ' lb', ctx); break;
      case 'issue':
        if (!is) throw new Error('No issue on this PO.');
        if (d.owner) { is.owner = d.owner; is.ownerManual = true; }
        if (d.urgent !== undefined) { is.urgent = !!d.urgent; is.urgentManual = true; }
        if (d.status) { is.status = d.status; if (d.status === 'Open') { is.opened = ctx.today; is.openManual = true; } else is.openManual = false; }
        if (d.decision !== undefined) { is.decision = d.decision; is.decisionNote = String(d.decisionNote || '').trim(); }
        is.updatedAt = ctx.stamp; upsert_('Issues', 'id', is); note_(id, me, 'Issue updated: ' + JSON.stringify(d), ctx);
        if (d.owner) notifyIssue_(is, ctx); break;
      case 'recount':
        var s = st.sub[id]; if (!s || !isTPL(p.wh)) throw new Error('Nothing to recount.');
        s.superseded = true; upsert_('Submissions', 'subId', s);
        upsert_('Recounts', 'id', { id: id, at: ctx.stamp, from: ctx.today, active: true, base: readAll_('Checks').filter(function (c) { return c.key === shipKey(p); }).length }); note_(id, me, 'Asked ' + p.wh + ' for another count', ctx); notify3pl_(p, 'recount', ctx); break;
      case 'nudge':
        if (p.nudged === ctx.today) throw new Error('Already nudged today.');
        p.nudged = ctx.today; savePO_(p, ctx); note_(id, me, 'Nudge sent to ' + p.wh, ctx); notify3pl_(p, 'nudge', ctx); break;
      case 'sap': p.sap = ctx.today; savePO_(p, ctx); note_(id, me, 'Entered in SAP / WMS', ctx); break;
      case 'note': if (String(d.text || '').trim()) note_(id, me, String(d.text).trim(), ctx); break;
      case 'doc': var f = preDoc; append_('Docs', { id: id, cat: d.cat || 'Other', name: f.name, url: f.url, by: me, at: ctx.stamp }); break;
      case 'receive': // Tootsi receiving at BM / LY / HB, a whole shipment at once
        if (isTPL(p.wh)) throw new Error('3PL POs are received by the 3PL.');
        return submitShip_(p, st, ctx, d, 'Tootsi receiving (' + p.wh + ') · ' + me, false, 'HUB');
      case 'check': if (isTPL(p.wh)) throw new Error('3PL POs are received by the 3PL.'); return checkShipCall_(p, st, ctx, d, me);
      case 'confirmUnplanned':
        if (!p.unplannedJson || truthy(p.planOk)) throw new Error('This is not an unplanned arrival.');
        var np = String(d.po || '').trim(); if (!/^\d{3,10}(-\d+)?$/.test(np)) throw new Error('Enter the PO number from SAP.');
        var cs = num(d.cs), lb = num(d.lbs); if (cs == null || cs < 0 || lb == null || lb < 0) throw new Error('Enter expected cases and lb.');
        // the PO may already be in the hub (no destination, or planned for another warehouse): move those lines here, no duplicate
        var planned = st.pos.filter(function (x) { return x.po === np && x.id !== p.id && !isDone(x, st); });
        var elsewhere = planned.filter(function (x) { return x.wh && x.wh !== p.wh; });
        if (elsewhere.length && !d.move) throw new Error('MOVE?' + elsewhere[0].wh + '?' + elsewhere.length);
        if (planned.length) {
          planned.forEach(function (x) { var m = String(x.manual || '').split(',').filter(String); if (m.indexOf('wh') < 0) m.push('wh'); x.manual = m.join(','); x.wh = p.wh; x.clockFrom = ctx.today; savePO_(x, ctx); });
          (p.unplannedJson.files || []).forEach(function (f) { append_('Docs', { id: planned[0].id, cat: f.cat, name: f.name, url: f.url, by: p.unplannedJson.by, at: p.unplannedJson.at }); });
          p.planOk = true; p.closed = ctx.today + ' · merged into PO ' + np; savePO_(p, ctx);
          notePOs_(planned, me, 'Arrived at ' + p.wh + ' without notice; confirmed by ' + me + '. ' + planned.length + ' planned line(s) moved to ' + p.wh + '. Paperwork from ' + p.unplannedJson.by + ' attached.', ctx);
          break;
        }
        p.po = np; p.item = String(d.item || '').trim(); p.cs = cs; p.lbs = lb; p.planOk = true; p.clockFrom = ctx.today; savePO_(p, ctx);
        note_(id, me, 'PO confirmed as ' + np + ' and added to the plan. ' + p.wh + ' can now receive it.', ctx); break;
      case 'addPO':
        var po = String(d.po || '').trim(), item = String(d.item || '').trim();
        if (!/^\d{3,10}(-\d+)?$/.test(po)) throw new Error('Enter the PO number (digits, suffix like -1 allowed).');
        if (!item || !d.desc || !d.supplier || !d.eta || OWN_WH.concat(TPL_WH).indexOf(d.wh) < 0) throw new Error('Fill in all required fields.');
        var nid = po + '|' + item; if (st.pos.some(function (x) { return x.id === nid; })) throw new Error('PO ' + po + ' item ' + item + ' is already in the hub.');
        ['plt', 'cs', 'lbs'].forEach(function (k) { var n = num(d[k]); if (n == null || n < 0) throw new Error('Quantities must be 0 or more.'); });
        savePO_({ id: nid, po: po, item: item, wh: d.wh, desc: String(d.desc).toUpperCase(), supplier: d.supplier, eta: d.eta, cntr: String(d.cntr || '').toUpperCase(), plt: num(d.plt), cs: num(d.cs), lbs: num(d.lbs), load: d.load || '',
          alloc: d.alloc || '', so: d.so || '', qc: 'Not received', instrSap: d.instr || '', source: 'manual', addedBy: me }, ctx);
        note_(nid, me, 'PO added by hand', ctx); break;
      default: throw new Error('Unknown action.');
    }
    logAction_(me, action, id, '');
    sync_(loadState_(), ctx);
    return { ok: true };
  });
  if (out && out._moves) { finishMoves_(out._moves); delete out._moves; }
  return out;
}
function upsertSupQ_(id, kind, v, me, ctx) {
  upsertWhere_('SupplierQty', function (r) { return r.id === id && r.kind === kind; }, { id: id, kind: kind, plt: v.plt, cs: v.cs, lbs: v.lbs, by: me, at: ctx.stamp });
}

// ---------- email (MailApp is included with Google Workspace) ----------
function notifyIssue_(is, ctx) { notifyIssues_([is], ctx); }
/** One email per PO and owner, listing its lines (a 17-line container sends 1 email, not 17). */
function notifyIssues_(list, ctx) {
  var cfg = ctx.cfg; if (!list.length || String(cfg.EMAIL_ON || 'no') !== 'yes') return;
  var groups = {};
  list.forEach(function (is) { var k = String(is.id).split('|')[0] + '#' + is.owner; (groups[k] = groups[k] || []).push(is); });
  Object.keys(groups).forEach(function (k) {
    var g = groups[k], is = g[0], po = k.split('#')[0];
    var to = String(cfg['TEAM_' + String(is.owner).replace(/\s/g, '')] || cfg.FALLBACK_EMAIL || '').trim();
    if (String(cfg.PILOT_ONLY_ME || 'yes') === 'yes') to = cfg.FALLBACK_EMAIL;
    if (!to) return;
    var cc = String(cfg.PILOT_ONLY_ME || 'yes') === 'yes' ? '' : (cfg.CC_ALWAYS || '');
    var urgent = g.some(function (x) { return truthy(x.urgent); });
    MailApp.sendEmail({ to: to, cc: cc, subject: (urgent ? '[URGENT] ' : '') + 'PO Hub · PO ' + po + ' · ' + is.type + (g.length > 1 ? ' (+' + (g.length - 1) + ' more lines)' : ''),
      body: 'Owner: ' + is.owner + '\n\n' + g.map(function (x) { return '• ' + String(x.id).replace('|', ' item ') + ': ' + x.type + ' (since ' + x.opened + ')'; }).join('\n') + '\n\nOpen the PO Hub: ' + (cfg.HUB_URL || '') });
  });
}
function notify3pl_(p, kind, ctx) {
  var cfg = ctx.cfg; if (String(cfg.EMAIL_3PL || 'no') !== 'yes') return;
  var u = readAll_('Users').filter(function (r) { return r.wh === p.wh && truthy(r.active) && r.email; })[0]; if (!u) return;
  var fr = String(cfg['LANG_' + p.wh] || 'fr') === 'fr';
  var subj = kind === 'recount' ? (fr ? 'Tootsi demande un recompte · PO ' : 'Tootsi asks for a recount · PO ') : (fr ? 'Rappel de réception · PO ' : 'Receiving reminder · PO ');
  MailApp.sendEmail({ to: u.email, subject: subj + p.po, body: (fr ? 'Ouvrez votre portail Tootsi pour recevoir ou signaler un retard : ' : 'Open your Tootsi portal to receive or report a delay: ') + (cfg.PORTAL_URL || '') });
}

// ==================== Jobs.gs ====================
/**
 * Jobs.gs — scheduled jobs (installTriggers() turns them on).
 *
 * importUpcoming(): every 2 hours, copies new lines from the "Upcoming FTL FCL" sheet into the POs tab.
 *   - One row per PO line (PO + item). Never overwrites a line that already exists (people may have edited it).
 *   - That sheet has no expected cases, so the hub flags "Enter expected cases" (the blind count needs it).
 *   - WHS "-" means not decided yet: the line is imported without a destination and flagged "Set destination".
 * dailyJob(): 7:40 every morning: recompute issues, remind 3PLs about late POs (if EMAIL_3PL = yes),
 *   and email a short digest to DIGEST_TO.
 */
function importUpcoming(e) {
  requireOwner_(e);
  return importUpcoming_();
}
function importUpcoming_() {
  var cfg = config_(); if (!cfg.UPCOMING_SHEET_ID) return 0;
  var src = SpreadsheetApp.openById(cfg.UPCOMING_SHEET_ID);
  var sh = src.getSheetByName(cfg.UPCOMING_TAB || 'UPCOMING FTL FCL');
  if (!sh) throw new Error('Tab "' + (cfg.UPCOMING_TAB || 'UPCOMING FTL FCL') + '" not found in the upcoming sheet.');
  // read dates in the source sheet's own time zone (that sheet is on GMT; reading it in Toronto time moved every ETA one day back)
  var srcTz = tzOr_(function () { return src.getSpreadsheetTimeZone(); }, 'Etc/GMT');
  var v = sh.getDataRange().getValues(); if (v.length < 2) return 0;
  var h = v[0].map(function (x) { return String(x).replace(/\s+/g, ' ').trim().toLowerCase(); });
  var col = function (name) { return h.indexOf(name.replace(/\s+/g, ' ').toLowerCase()); };
  var c = { dec: col('WH DECISION (MAJID)'), po: col('PO#'), ven: col('Vendor Name'), skid: col('#Skid'), size: col('Ship Size'), csize: col('CSize'), pltzd: col('PLTZD'), cntr: col('CNTR/TrackNum'), whs: col('WHS'), eta: col('ETA Destination'), item: col('Item No.'), desc: col('Item Description'), wt: col('Line Weight') };
  if (c.po < 0 || c.item < 0) throw new Error('Could not find the PO# / Item No. columns in the upcoming sheet.');
  var lines = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], po = String(r[c.po]).trim(), item = String(r[c.item]).trim();
    if (!po || !item) continue;
    var eta = r[c.eta] instanceof Date ? Utilities.formatDate(r[c.eta], srcTz, 'yyyy-MM-dd') : (/^\d{4}-\d{2}-\d{2}/.test(String(r[c.eta])) ? String(r[c.eta]).slice(0, 10) : '');
    var whs = String(r[c.whs] || '').trim().toUpperCase();
    // WHS "-" = not decided in WHS, but Majid may already have decided in "WH DECISION (MAJID)"
    if (TPL_WH.concat(OWN_WH).indexOf(whs) < 0 && c.dec >= 0) { var dec = String(r[c.dec] || '').trim().toUpperCase().replace(/[^A-Z]/g, ''); if (TPL_WH.concat(OWN_WH).indexOf(dec) >= 0) whs = dec; }
    var size = String(r[c.size] || '').trim(), csize = String(r[c.csize] || '').trim(), pal = String(r[c.pltzd] || '').trim() === '1';
    var load = (size === 'FCL' ? 'FCL ' + (csize && csize !== '-' ? csize : '') : size) + ' · ' + (size === 'FCL' && !pal ? 'Floor loaded' : 'Palletized');
    lines.push({ id: po + '|' + item, po: po, item: item, wh: (TPL_WH.concat(OWN_WH).indexOf(whs) >= 0 ? whs : ''), desc: String(r[c.desc]).trim(), supplier: String(r[c.ven]).trim(), eta: eta,
      cntr: String(r[c.cntr] || '').trim().toUpperCase(), plt: num(r[c.skid]), lbs: num(r[c.wt]), load: load.replace(/\s+/g, ' ').trim() });
  }
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), have = {}, added = 0, changed = 0, notesByPO = {};
    st.pos.forEach(function (p) { have[p.id] = p; });
    lines.forEach(function (L) {
      var p = have[L.id];
      if (!p) {
        L.cs = ''; L.qc = 'Not received'; L.source = 'sheet'; L.addedBy = 'Import'; L.updatedAt = ctx.stamp;
        if (L.plt == null) L.plt = ''; if (L.lbs == null) L.lbs = '';
        append_('POs', L); have[L.id] = L; added++; return;
      }
      // keep the hub in step with the sheet: ETA, destination, container, pallets, weight, load type
      // never after receiving, and never for a field someone changed by hand in the hub
      if (p.source !== 'sheet' || st.sub[p.id] || p.closed) return;
      var manual = String(p.manual || '').split(','), diff = [];
      ['eta', 'wh', 'cntr', 'plt', 'lbs', 'load'].forEach(function (k) {
        var nv = L[k] == null ? '' : L[k], ov = p[k] == null ? '' : p[k];
        if (manual.indexOf(k) >= 0 || nv === '' || String(nv) === String(ov)) return;
        diff.push(k + ' ' + (ov === '' ? '(blank)' : ov) + ' → ' + nv); p[k] = nv;
      });
      if (diff.length) { savePO_(p, ctx); (notesByPO[p.po] = notesByPO[p.po] || { id: p.id, d: {} }); diff.forEach(function (t) { notesByPO[p.po].d[t] = true; }); changed++; }
    });
    // one note per PO (a 17-line container gets one note, not 17)
    Object.keys(notesByPO).forEach(function (po) { note_(notesByPO[po].id, 'Upcoming sheet', 'Updated from the Upcoming sheet: ' + Object.keys(notesByPO[po].d).join(', '), ctx); });
    if (added || changed) logAction_('Import', 'importUpcoming', '', added + ' new PO lines, ' + changed + ' updated');
    return added;
  });
}

/** Every hour (trigger) and on "Sync now": Upcoming sheet → hub, Arash's packing slip log → hub, hub ↔ master sheet. */
function syncSources(e) { requireOwner_(e); return syncSources_(); }
function syncSources_() {
  var out = {};
  try { out.upcoming = importUpcoming_(); } catch (err) { out.upcoming = 'error: ' + err.message; logAction_('Sync', 'error', 'upcoming', err.message); }
  try { out.slips = importSlips_(); } catch (err) { out.slips = 'error: ' + err.message; logAction_('Sync', 'error', 'slips', err.message); }
  try { out.master = syncMaster_(); } catch (err) { out.master = 'error: ' + err.message; logAction_('Sync', 'error', 'master', err.message); }
  return out;
}

/** Packing slips: Arash's extractor keeps a "PO Status" tab (PO, FOUND/MISSING, file link, action). Copy it into the Slips tab. */
function importSlips_() {
  var cfg = config_(); if (!cfg.SLIP_LOG_ID) return 0;
  var sh = SpreadsheetApp.openById(cfg.SLIP_LOG_ID).getSheetByName(cfg.SLIP_TAB || 'PO Status'); if (!sh) throw new Error('Packing slip tab not found.');
  var v = sh.getDataRange().getDisplayValues(); if (v.length < 2) return 0;
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase(); }), c = function (n) { return h.indexOf(n); };
  var ci = { po: c('po'), st: c('packing slip'), file: c('packing slip file'), src: c('source / why it is missing'), act: c('action needed') };
  if (ci.po < 0 || ci.st < 0) throw new Error('Packing slip log: columns PO / Packing slip not found.');
  var rows = v.slice(1).filter(function (r) { return /^\d{4,6}(-\d+)?$/.test(String(r[ci.po]).trim()); }).map(function (r) {
    var url = ci.file >= 0 ? String(r[ci.file]).trim() : '', m = url.match(/\/d\/([\w-]{20,})/);
    return { po: String(r[ci.po]).trim(), status: String(r[ci.st]).trim().toUpperCase(), url: url, fileId: m ? m[1] : '', source: ci.src >= 0 ? String(r[ci.src]).slice(0, 300) : '', action: ci.act >= 0 ? String(r[ci.act]).slice(0, 300) : '', checked: today_() };
  });
  return withLock_(function () {
    var have = {}; readAll_('Slips').forEach(function (s) { have[s.po] = s; });
    var n = 0;
    rows.forEach(function (r) { var o = have[r.po]; if (!o || o.status !== r.status || o.url !== r.url) { upsert_('Slips', 'po', r); n++; } });
    return n;
  });
}

/**
 * Hub ↔ master sheet (third_party_whse, tab "3PL").
 * READ: if the master shows a plausible receiving date for an open TP/DL line, the hub shows it (one click closes the line).
 * WRITE: after a 3PL receipt, the hub fills EMPTY master cells only (Product, Description, Receiving Date, Expiry Date,
 *   Units shipped to 3PL, Receiving Docs) on the row with the same PO + warehouse (+ item). Never overwrites.
 *   MASTER_MODE = preview (default): nothing is written, the MasterSync tab shows what would be written. on: writes. off: nothing.
 */
function syncMaster_() {
  var cfg = config_(), mode = String(cfg.MASTER_MODE || 'preview'); if (!cfg.MASTER_SHEET_ID || mode === 'off') return 'off';
  var sh = SpreadsheetApp.openById(cfg.MASTER_SHEET_ID).getSheetByName(cfg.MASTER_TAB || '3PL'); if (!sh) throw new Error('Master tab not found.');
  var top = sh.getRange(1, 1, 10, Math.min(sh.getLastColumn(), 60)).getDisplayValues(), hr = -1;
  top.forEach(function (r, i) { if (hr < 0 && r.some(function (x) { return String(x).trim() === 'PO Number'; })) hr = i + 1; });
  if (hr < 0) throw new Error('Master: header row with "PO Number" not found.');
  var head = top[hr - 1].map(function (x) { return String(x).replace(/\s+/g, ' ').trim(); }), col = function (n) { return head.indexOf(n) + 1; };
  var C = { po: col('PO Number'), wh: col('Warehouse'), item: col('Product'), desc: col('Description'), recv: col('Receiving Date'), exp: col('Expiry Date'), units: col('Units shipped to 3PL'), docs: col('Receiving Docs'), status: col('Report Alert Status') };
  if (!C.po || !C.wh) throw new Error('Master: PO Number / Warehouse columns not found.');
  var last = sh.getLastRow(); if (last <= hr) return 0;
  var n = last - hr;
  // one read per column, everything else is compared in memory (no slow cell-by-cell reads)
  var disp = function (c) { return c ? sh.getRange(hr + 1, c, n, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : []; };
  var raw = function (c) { return c ? sh.getRange(hr + 1, c, n, 1).getValues().map(function (r) { return r[0]; }) : []; };
  var M = { po: disp(C.po), wh: disp(C.wh), item: disp(C.item), desc: disp(C.desc), recv: raw(C.recv), recvT: disp(C.recv), exp: disp(C.exp), expRaw: raw(C.exp), units: disp(C.units), docs: disp(C.docs), status: disp(C.status) };
  var normPO = function (x) { return String(x).replace(/^PO\s*#?/i, '').trim(); };
  var rowsFor = function (po, wh) { var out = []; for (var i = 0; i < n; i++) if (normPO(M.po[i]) === po && M.wh[i].toUpperCase() === wh) out.push(i); return out; };
  var ctx = ctx_(), seen = 0, written = 0, previewed = 0;
  withLock_(function () {
    var st = loadState_();
    // READ: the master's "Receiving Date" often holds a PLANNED date, so trust a row only when it really looks received:
    // units filled, a real date between Jan 2026 and today, and a status that does not say "pending"
    var have = {}; readAll_('MasterSeen').forEach(function (s) { have[s.id] = s; });
    st.pos.forEach(function (p) {
      if (!isTPL(p.wh) || isDone(p, st)) return;
      var hit = rowsFor(p.po, p.wh).filter(function (i) { return !C.item || M.item[i] === p.item; })[0];
      var ds = '';
      if (hit != null) {
        var d = M.recv[hit], units = num(String(M.units[hit]).replace(/[^\d.]/g, ''));
        if (d instanceof Date && !isNaN(d) && units > 0 && !/pending/i.test(M.status[hit])) {
          var t = Utilities.formatDate(d, 'America/Toronto', 'yyyy-MM-dd');
          if (t >= '2026-01-01' && t <= ctx.today) ds = t;
        }
      }
      if (ds && (!have[p.id] || have[p.id].recv !== ds)) { upsert_('MasterSeen', 'id', { id: p.id, recv: ds, row: String(hr + 1 + hit), checked: ctx.today }); seen++; }
      if (!ds && have[p.id]) deleteWhere_('MasterSeen', function (r) { return r.id === p.id; });
    });
    // WRITE: only NEW items each hour (pending), plus the previewed ones when MASTER_MODE is switched to on
    var todo = readAll_('MasterSync').filter(function (q) {
      return q.status === 'pending' || (mode === 'on' && /^preview/.test(q.status)) || (q.status === 'no master row' && String(q.at) >= addDays(ctx.today, -7));
    });
    todo.forEach(function (q) {
      var save = function () { upsertWhere_('MasterSync', function (r) { return r.key === q.key; }, q); };
      var cands = rowsFor(q.po, q.wh), withItem = cands.filter(function (i) { return M.item[i] === q.item; }), blank = cands.filter(function (i) { return M.item[i] === ''; });
      var row = withItem.length === 1 ? withItem[0] : (!withItem.length && blank.length === 1 ? blank[0] : null);
      if (row == null) { q.status = withItem.length > 1 || blank.length > 1 ? 'several master rows: fill by hand' : 'no master row'; save(); return; }
      var c = C[q.field]; if (!c) { q.status = 'column not in master'; save(); return; }
      var cur = String((q.field === 'recv' ? M.recvT : M[q.field] || [])[row] || '').trim();
      q.row = String(hr + 1 + row);
      if (cur !== '') {
        // dates are compared as dates (the master shows 23/Nov/26, the hub writes 2026-11-23)
        var rawD = q.field === 'recv' ? M.recv[row] : q.field === 'exp' ? M.expRaw[row] : null;
        var curD = rawD instanceof Date && !isNaN(rawD) ? Utilities.formatDate(rawD, 'America/Toronto', 'yyyy-MM-dd') : '';
        var same = cur === String(q.value) || (curD && curD === String(q.value)) || (q.field === 'units' && num(cur.replace(/[^\d.]/g, '')) === num(q.value));
        q.status = same ? 'already the same' : 'differs: master has ' + cur.slice(0, 40); save(); return;
      }
      if (mode !== 'on') { q.status = 'preview: would write'; previewed++; save(); return; }
      // right before writing, make sure the row was not moved (someone may sort or insert rows in the master)
      if (normPO(sh.getRange(hr + 1 + row, C.po).getDisplayValue()) !== q.po) { q.status = 'master rows moved: will retry'; save(); return; }
      var v = String(q.value), m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      var val = (q.field === 'recv' || q.field === 'exp') && m ? new Date(+m[1], +m[2] - 1, +m[3])
        : q.field === 'units' ? num(v)
        : q.field === 'docs' && /^https:\/\/drive\.google\.com\//.test(v) ? '=HYPERLINK("' + v.replace(/"/g, '') + '","PO Hub files")'
        : "'" + v; // text always as text: never a formula in Majid's sheet
      sh.getRange(hr + 1 + row, c).setValue(val);
      if (q.field === 'recv') { M.recvT[row] = v; M.recv[row] = val; } else { if (M[q.field]) M[q.field][row] = v; if (q.field === 'exp') M.expRaw[row] = val; }
      q.status = 'written'; written++; save();
    });
  });
  return { seen: seen, written: written, preview: previewed };
}

function dailyJob(e) {
  requireOwner_(e);
  try { cleanIncoming_(); } catch (err) { logAction_('Daily', 'error', 'incoming', err.message); }
  try { importSlips_(); syncMaster_(); } catch (err) { logAction_('Daily', 'error', 'sync', err.message); }
  try { importUpcoming_(); } catch (e) { logAction_('Import', 'error', '', e.message); }
  var ctx = ctx_(), cfg = ctx.cfg, st, late;
  withLock_(function () {
    st = sync_(loadState_(), ctx);
    late = st.pos.filter(function (p) { return isLate(p, st, ctx); });
    // one reminder per shipment (container or PO), not one per item line
    var sent = {};
    late.forEach(function (p) {
      if (!isTPL(p.wh) || p.nudged === ctx.today || sla(p, st, ctx).lvl !== 'nudge') return;
      var k = shipKey(p); p.nudged = ctx.today; savePO_(p, ctx);
      if (!sent[k]) { sent[k] = true; notify3pl_(p, 'nudge', ctx); }
    });
  });
  var open = st.pos.filter(function (p) { return openIssue(p, st); });
  var sap = st.pos.filter(function (p) { return notInSap(p, st, ctx); });
  if (!cfg.DIGEST_TO || String(cfg.EMAIL_ON || 'no') !== 'yes') return;
  var line = function (p) { var i = st.issue[p.id]; return '• PO ' + p.po + (p.item ? ' / ' + p.item : '') + ' · ' + p.wh + ' · ' + p.desc + (i && i.status === 'Open' ? ' — ' + i.type + ' (owner ' + i.owner + ')' : ''); };
  // one line per shipment / PO, not per item line
  var once = function (arr, k) { var s = {}; return arr.filter(function (p) { var x = k(p); if (s[x]) return false; s[x] = true; return true; }); };
  late = once(late, shipKey); open = once(open, function (p) { return p.po; }); sap = once(sap, shipKey);
  var body = 'PO Hub · ' + ctx.today + '\n\nLate at the warehouse (' + late.length + ' shipments):\n' + late.map(line).join('\n') +
    '\n\nOpen issues (' + open.length + '):\n' + open.map(line).join('\n') + '\n\nArrived 2+ days, not in SAP (' + sap.length + '):\n' + sap.map(line).join('\n') + '\n\n' + (cfg.HUB_URL || '');
  MailApp.sendEmail({ to: cfg.DIGEST_TO, subject: 'PO Hub daily · ' + late.length + ' late · ' + open.length + ' issues', body: body });
}

// ==================== pages (HTML) ====================
var HTML_FILES = {"Hub": "<!DOCTYPE html>\n<html lang=\"en\">\n<head><base target=\"_top\"><meta charset=\"utf-8\"><?!= include('Styles') ?></head>\n<body>\n<header>\n  <?!= include('Logo') ?>\n  <div><h1>PO <span>Hub</span></h1><div class=\"hsub\">Every PO, every document, every department · pilot</div></div>\n  <div class=\"who\"><a class=\"btn ghost\" id=\"tplview\" href=\"#\" target=\"_blank\" rel=\"noopener\" style=\"display:none;min-height:36px;padding:4px 12px;font-size:13px;text-decoration:none\">👁 See the 3PL view</a><div class=\"nm\"><div id=\"uname\" style=\"font-weight:500\"></div><div class=\"hsub\">Tootsi Impex</div></div></div>\n</header>\n<main id=\"main\"><div class=\"loading\">Loading POs…</div></main>\n<div class=\"scrim\" id=\"scrim\"></div>\n<aside class=\"drawer\" id=\"drawer\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"dtitle\" tabindex=\"-1\"></aside>\n<div class=\"toast\" id=\"toast\" role=\"status\" aria-live=\"polite\"></div>\n<div class=\"busy\" id=\"busy\" aria-live=\"polite\">Saving…</div>\n<p id=\"rc\" class=\"sr\" aria-live=\"polite\"></p>\n<?!= include('Client') ?>\n<?!= include('Lang') ?>\n<?!= include('Receive') ?>\n<script>\nlet lang=\"en\";const t=k=>(L.en[k])||k;const loadTxt=s=>s;\nconst CONDS=[[\"Good\",\"✅\",\"c_Good\"],[\"Shifted load\",\"↔️\",\"c_Shifted\"],[\"Collapsed pallets\",\"⬇️\",\"c_Collapsed\"],[\"Double-stacked\",\"⏫\",\"c_Double\"],[\"Torn / leaking bags\",\"🩹\",\"c_Torn\"],[\"Wet / dirty\",\"💧\",\"c_Wet\"],[\"Other\",\"❓\",\"c_Other\"]];\nconst WHS=[\"TP\",\"DL\",\"BM\",\"LY\",\"HB\"],isTPL=w=>w===\"TP\"||w===\"DL\";\nconst OWNERS=[\"Purchasing\",\"Logistics\",\"QC\",\"Receiving\",\"Inventory Control\",\"Fulfillment\"],DECISIONS=[\"Pending\",\"Refuse\",\"Accept with reservation\",\"Repalletize and accept\"];\nconst CATS=[\"Supplier packing list\",\"Supplier loading report\",\"Bill of lading\",\"Certificate of analysis\",\"QC report\",\"NCR / claim\",\"Photos\",\"Other\"];\nconst ICON={up:\"⚪\",tom:\"🔵\",due:\"🟡\",nudge:\"🟠\",esc:\"🔴\",paused:\"⏸\"};\nlet DATA=null,whF=\"\",kpiF=\"\",ownF=\"\",qRaw=\"\",view=\"table\",welcome=true,openId=null;\ntry{welcome=localStorage.getItem(\"pohub_welcome\")!==\"0\"}catch(e){}\nlet HUBLINES=[];\nconst RCV={title:\"Tootsi receiving\",nextLabel:\"Back to the PO\",lines:()=>HUBLINES,askP:r=>!!r.askP,fresh:r=>!!r.fresh,upload:f=>call(\"hubUpload\",f),\n check:(r,c)=>call(\"hubAction\",r.id,\"check\",c),submit:(r,pl)=>call(\"hubAction\",r.id,\"receive\",pl),after:async(r)=>{await load();openPO(r.id)}};\n\nasync function load(){try{ensureReceive(i=>call(\"hubDiag\",i));DATA=await call(\"hubLoad\");$(\"uname\").textContent=DATA.me;const tv=$(\"tplview\");if(DATA.cfg&&DATA.cfg.portalUrl){tv.href=DATA.cfg.portalUrl+(DATA.cfg.portalUrl.indexOf(\"?\")<0?\"?\":\"&\")+\"view=3pl\";tv.style.display=\"\"}render()}catch(e){$(\"main\").innerHTML=`<div class=\"err\">${esc(e.message)}</div>`}}\nconst slaPill=r=>r.sla?`<span class=\"sla sla-${r.sla.lvl}\" ${r.sla.h?`title=\"${r.sla.h} business hours\"`:\"\"}><span aria-hidden=\"true\">${ICON[r.sla.lvl]}</span> ${esc(r.slaText)}</span>`:\"\";\nconst late=r=>r.sla&&(r.sla.lvl===\"nudge\"||r.sla.lvl===\"esc\");\nconst allocPill=r=>r.alloc?`<span class=\"pill ${r.alloc===\"Production\"?\"a-prod\":r.alloc===\"Sales\"?\"p-done\":\"p-todo\"}\">${esc(r.alloc)}${r.alloc===\"Split\"?(r.prodQty!==\"\"?` ${fmt(r.prodQty)}/${fmt(r.salesQty)}`:\" (set qty)\"):\"\"}</span>`:`<span class=\"pill p-dmg\">Not set</span>`;\nconst qcText=r=>r.qc===\"Locked\"?`QC locked${r.qcAt?\" \"+daysBetween(r.qcAt,DATA.today)+\" d\":\"\"}`:r.qc===\"Released\"?\"QC released\":\"QC: waiting\";\n// the cards count shipments (a container or a PO), not item lines: a 17-line container counts once\nconst KPIS=[\n {id:\"week\",l:\"Shipments arriving this week\",f:r=>!r.sub&&!r.closed&&r.eta>=DATA.today&&daysBetween(DATA.today,r.eta)<=6,main:1,by:\"ship\"},\n {id:\"rcv\",l:\"Shipments late at warehouse\",f:r=>late(r)||!!r.unplanned,alert:1,main:1,by:\"ship\"},\n {id:\"sap\",l:\"Shipments arrived 2+ days, not in SAP\",f:r=>r.notInSap,alert:1,main:1,by:\"ship\"},\n {id:\"iss\",l:\"POs with an open issue\",f:r=>r.issue&&r.issue.status===\"Open\",alert:1,main:1,by:\"po\"},\n {id:\"over\",l:\"Issues overdue\",f:r=>r.issue&&r.issue.overdue},{id:\"dly\",l:\"Delays reported by 3PL\",f:r=>!!r.delay},\n {id:\"cases\",l:\"No expected cases yet\",f:r=>!r.sub&&!r.closed&&(r.cs===\"\"||r.cs==null)},{id:\"closed\",l:\"Received before the hub\",f:r=>!!r.closed},{id:\"dest\",l:\"No destination\",f:r=>!r.wh},\n {id:\"alloc\",l:\"No allocation set\",f:r=>!r.alloc},{id:\"qc\",l:\"QC locked\",f:r=>r.qc===\"Locked\"}];\nfunction statusCell(r){let m;const s=r.sub;\n if(r.closed)m=`<span class=\"pill p-grey\">Received before the hub</span>`;\n else if(r.req)m=`<span class=\"pill p-warn\">Recount asked</span>`;\n else if(s){const dd=r.eta?daysBetween(r.eta,s.recv):0;m=`<span class=\"pill ${s.damage?\"p-dmg\":\"p-done\"}\">Arrived ${fds(s.recv)}</span>${dd?`<div class=\"small\">${Math.abs(dd)} d ${dd<0?\"early\":\"late\"} vs ETA ${fds(r.eta)}</div>`:\"\"}`}\n else if(r.unplanned)m=`<span class=\"pill p-dmg\">Arrived · not on list</span><div class=\"small\">Reported ${esc(r.unplanned.at)}</div>`;\n else if(r.sla&&[\"due\",\"nudge\",\"esc\",\"paused\"].includes(r.sla.lvl))m=slaPill(r)+(r.delay?`<div class=\"small\">${esc(r.delay.reasonText)}${r.delay.newDate?\" · new date \"+fds(r.delay.newDate):\"\"}</div>`:`<div class=\"small\">ETA ${fds(r.eta)}</div>`);\n else m=`<span class=\"pill p-todo\">${r.eta?\"ETA \"+fds(r.eta):\"No ETA\"}</span>`;\n return m+`<div class=\"small\">${qcText(r)}</div>${r.notInSap?`<div class=\"small\">Not in SAP</div>`:\"\"}`}\nfunction baseRows(){const q=qRaw.trim().toLowerCase();return DATA.rows.filter(r=>(!whF||r.wh===whF)&&(!ownF||(r.issue&&r.issue.status===\"Open\"&&r.issue.owner===ownF))&&(!q||[r.po,r.item,r.desc,r.supplier,r.so,r.sub&&r.sub.lot,r.cntr].join(\" \").toLowerCase().includes(q)))}\nfunction weekBoard(rows){const mon=DATA.today,cols=[{k:\"late\",h:\"Late / arrived earlier\"}];for(let i=0;i<5;i++){const d=addDays(mon,i);cols.push({k:d,h:fd(d)})}cols.push({k:\"next\",h:\"Later\"});\n const colOf=r=>{const d=r.sub?r.sub.recv:(r.clockFrom||r.eta);if(!d)return\"next\";if(d<mon)return\"late\";if(d>addDays(mon,4))return\"next\";return d};\n return `<div class=\"week\">${cols.map(c=>{const it=rows.filter(r=>colOf(r)===c.k);return `<div class=\"wcol ${c.k===DATA.today?\"today\":\"\"}\"><h4>${esc(c.h)}</h4>${it.length?it.map(r=>`<button type=\"button\" class=\"wcard ${r.sub?\"done\":\"lvl-\"+(r.sla?r.sla.lvl:\"none\")}\" data-open=\"${esc(r.id)}\"><b>${esc(r.po)}</b> · ${esc(r.wh||\"?\")}<br>${esc(String(r.desc).slice(0,26))}<br><span class=\"small\" style=\"margin:0\">${r.sub?\"Arrived \"+fds(r.sub.recv):r.unplanned?\"Not on list\":esc(r.slaText||\"\")}</span></button>`).join(\"\"):`<div class=\"small\">—</div>`}</div>`}).join(\"\")}</div>`}\nconst EXP=new Set(),LV=[\"esc\",\"nudge\",\"due\",\"paused\",\"tom\",\"up\"],rk=r=>r.sla?LV.indexOf(r.sla.lvl):99;\n/** Lines of one shipment collapse to one row (the most urgent) unless the group is opened. */\nfunction groupRows(rows){const m=new Map();rows.forEach(r=>{if(!m.has(r.ship))m.set(r.ship,[]);m.get(r.ship).push(r)});const out=[];\n m.forEach((ls,k)=>{const open=EXP.has(k)||!!qRaw.trim();const sorted=ls.slice().sort((a,b)=>rk(a)-rk(b));(open?sorted:[sorted[0]]).forEach((r,i)=>out.push({r,n:i===0?ls.length:1,open}))});return out}\nfunction render(){\n const base=baseRows(),kf=KPIS.find(k=>k.id===kpiF),rows=base.filter(r=>!kf||kf.f(r));\n $(\"main\").innerHTML=`${welcome?`<div class=\"welcome\"><div style=\"display:flex;justify-content:space-between;gap:10px\"><h3>PO Hub pilot</h3><button type=\"button\" class=\"btn ghost\" id=\"wx\" style=\"min-height:34px;padding:4px 10px\">Hide</button></div>\n  <div style=\"font-size:13.5px\">Every PO from the upcoming sheet, with the 3PL receiving, photos, counts and issues in one place. Click a PO number to see everything. Lines of the same container are grouped: click \"items\" to open them. The Upcoming sheet has no cases, so the 3PL count is only compared on pallets until you add <b>expected cases</b> (open a PO, then \"Expected cases for all lines\"). Shipments that arrived before the hub: open them and click <b>Received before the hub</b>.</div></div>`:\"\"}\n <div style=\"display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px\"><h2 style=\"margin:0\" tabindex=\"-1\">All POs <span class=\"ichip\">Internal only</span></h2>\n  <div style=\"display:flex;gap:8px;flex-wrap:wrap\"><div class=\"view\"><button type=\"button\" data-v=\"table\" aria-pressed=\"${view===\"table\"}\" class=\"${view===\"table\"?\"on\":\"\"}\">Table</button><button type=\"button\" data-v=\"week\" aria-pressed=\"${view===\"week\"}\" class=\"${view===\"week\"?\"on\":\"\"}\">This week</button></div>\n  <button type=\"button\" class=\"btn ghost\" id=\"refresh\">↻ Refresh</button><button type=\"button\" class=\"btn ghost\" id=\"syncnow\" title=\"Pull the Upcoming sheet, Arash's packing slips and the master sheet now (also runs every hour)\">⇄ Sync sheets</button><button type=\"button\" class=\"btn\" id=\"addpo\">＋ Add upcoming PO</button></div></div>\n <div class=\"kpis\">${KPIS.filter(k=>k.main).map(k=>{const n=new Set(base.filter(k.f).map(r=>k.by===\"po\"?r.po:r.ship)).size;return `<button type=\"button\" class=\"kpi ${kpiF===k.id?\"on\":\"\"} ${k.alert&&n?\"alert\":\"\"}\" data-k=\"${k.id}\" aria-pressed=\"${kpiF===k.id}\"><div class=\"n\">${n}</div><div class=\"l\">${k.l}</div></button>`}).join(\"\")}</div>\n <div class=\"bar\"><label for=\"q\" class=\"sr\">Search</label><input id=\"q\" type=\"search\" placeholder=\"Search PO, item, product, supplier, lot, container or sales order…\" value=\"${esc(qRaw)}\">\n  <div class=\"chips\">${[\"\",...WHS].map(w=>`<button type=\"button\" data-w=\"${w}\" class=\"${whF===w?\"on\":\"\"}\" aria-pressed=\"${whF===w}\">${w||\"All\"}${w?`<small>${isTPL(w)?\"3PL\":\"Tootsi\"}</small>`:\"\"}</button>`).join(\"\")}</div>\n  <select id=\"more\" aria-label=\"More filters\"><option value=\"\">More filters…</option>${KPIS.filter(k=>!k.main).map(k=>`<option value=\"${k.id}\" ${kpiF===k.id?\"selected\":\"\"}>${k.l} (${base.filter(k.f).length})</option>`).join(\"\")}</select>\n  <select id=\"ownF\" aria-label=\"Issue owner\"><option value=\"\">Issues: any owner</option>${OWNERS.map(o=>`<option ${ownF===o?\"selected\":\"\"}>${o}</option>`).join(\"\")}</select>\n  ${kpiF||whF||ownF||qRaw?`<button type=\"button\" class=\"btn ghost\" id=\"clr\">Clear filters</button>`:\"\"}</div>\n <details class=\"legend\"><summary>What do the colors mean?</summary><div><span class=\"sla sla-up\">⚪ Upcoming</span><span class=\"sla sla-due\">🟡 Expected today</span><span class=\"sla sla-nudge\">🟠 1 business day late: nudge</span><span class=\"sla sla-esc\">🔴 2+ business days late: escalated</span><span class=\"sla sla-paused\">⏸ Delay reported by 3PL</span><span>Business hours Mon–Fri 8:00–17:00.</span></div></details>\n ${view===\"week\"?weekBoard(rows):`<div class=\"tw\"><table><caption class=\"sr\">Purchase orders</caption><thead><tr><th scope=\"col\">Needs action</th><th scope=\"col\">PO</th><th scope=\"col\">Dest.</th><th scope=\"col\">Item</th><th scope=\"col\">Status</th><th scope=\"col\">Who it's for</th><th scope=\"col\">Qty check</th></tr></thead><tbody>\n ${rows.length?groupRows(rows).map(({r,n,open})=>`${n>1?`<tr class=\"grp\"><td colspan=\"7\"><button type=\"button\" class=\"grpbtn\" data-grp=\"${esc(r.ship)}\" aria-expanded=\"${open}\">${open?\"▾\":\"▸\"} ${r.cntr?\"Container \"+esc(r.cntr):\"PO \"+esc(r.po)} · <b>${n} items</b> ${open?\"\":\"(showing the most urgent)\"}</button></td></tr>`:\"\"}<tr data-po=\"${esc(r.id)}\" class=\"lvl-${r.sla?r.sla.lvl:\"none\"}\">\n  <td data-l=\"Needs action\" class=\"wide ${r.needs.length?\"\":\"nona\"}\">${r.needs.length?r.needs.map(a=>`<div class=\"need ${a.u?\"u\":\"\"}\">${esc(a.t)}</div>`).join(\"\"):`<span class=\"small\">—</span>`}</td>\n  <td data-l=\"PO\"><button type=\"button\" class=\"pobtn\" data-open=\"${esc(r.id)}\">${esc(r.po)}</button>${r.source===\"manual\"?`<div><span class=\"pill a-prod\" style=\"margin-top:4px\">Manual</span></div>`:\"\"}</td>\n  <td data-l=\"Dest.\"><b>${esc(r.wh||\"—\")}</b><div class=\"small\" style=\"margin:0\">${r.wh?(isTPL(r.wh)?\"3PL\":\"Tootsi\"):\"not set\"}</div></td>\n  <td data-l=\"Item\" class=\"wide\"><b>${esc(r.item)}</b> <span class=\"small\" style=\"margin:0\">· ${fmt(r.plt)} plt · ${fmt(r.cs)} cs · ${fmt(r.lbs)} lb</span><div class=\"clamp\">${esc(r.desc)}${r.sub&&r.sub.lot?` · Lot ${esc(r.sub.lot)}`:\"\"}</div></td>\n  <td data-l=\"Status\">${statusCell(r)}${r.sub&&r.sub.comment?`<div class=\"small\" style=\"color:var(--blue)\" title=\"${esc(r.sub.comment)}\">💬 comment</div>`:\"\"}</td>\n  <td data-l=\"Who it's for\">${r.unplanned?\"—\":allocPill(r)}${r.so?`<div class=\"small\">${esc(r.so)}</div>`:\"\"}</td>\n  <td data-l=\"Qty check\"><span class=\"pill ${r.qty.cls}\">${esc(r.qty.t)}</span>${r.qty.sub?`<div class=\"small\">${esc(r.qty.sub)}</div>`:\"\"}</td></tr>`).join(\"\"):`<tr><td colspan=\"7\"><div class=\"empty\" style=\"border:none\">No POs match these filters.</div></td></tr>`}</tbody></table></div>`}`;\n $(\"rc\").textContent=rows.length+\" POs shown\";\n document.querySelectorAll(\"[data-grp]\").forEach(b=>b.onclick=()=>{const k=b.dataset.grp;EXP.has(k)?EXP.delete(k):EXP.add(k);render()});\n document.querySelectorAll(\"tr[data-po]\").forEach(tr=>tr.onclick=e=>{if(e.target.closest(\"[data-open]\"))return;openPO(tr.dataset.po)});\n document.querySelectorAll(\"[data-open]\").forEach(b=>b.onclick=e=>{e.stopPropagation();openPO(b.dataset.open)});\n document.querySelectorAll(\"[data-w]\").forEach(b=>b.onclick=()=>{whF=b.dataset.w;render()});\n document.querySelectorAll(\"[data-k]\").forEach(b=>b.onclick=()=>{kpiF=kpiF===b.dataset.k?\"\":b.dataset.k;render()});\n document.querySelectorAll(\"[data-v]\").forEach(b=>b.onclick=()=>{view=b.dataset.v;render()});\n $(\"more\").onchange=e=>{kpiF=e.target.value;render()};$(\"ownF\").onchange=e=>{ownF=e.target.value;render()};\n const c=$(\"clr\");if(c)c.onclick=()=>{kpiF=whF=ownF=qRaw=\"\";render()};\n const qi=$(\"q\");qi.oninput=e=>{qRaw=e.target.value;clearTimeout(qi.t);qi.t=setTimeout(()=>{render();const n=$(\"q\");n.focus();n.setSelectionRange(n.value.length,n.value.length)},250)};\n $(\"addpo\").onclick=openAddPO;$(\"refresh\").onclick=load;\n $(\"syncnow\").onclick=async()=>{busy(true);try{const r=await call(\"hubAction\",\"\",\"syncNow\",{});await load();const s=r.sync||{};toast(\"Synced. Upcoming: \"+s.upcoming+\" new · slips: \"+s.slips+\" updated · master: \"+(typeof s.master===\"object\"?s.master.seen+\" seen, \"+s.master.preview+\" to fill (preview), \"+s.master.written+\" written\":s.master))}catch(e){toast(e.message)}busy(false)};const wx=$(\"wx\");if(wx)wx.onclick=()=>{welcome=false;try{localStorage.setItem(\"pohub_welcome\",\"0\")}catch(e){}render()};\n}\n\n// ---------- PO page ----------\nasync function act(id,action,d,msg){busy(true);try{await call(\"hubAction\",id,action,d||{});await Promise.all([load(),id?openPO(id,true):null]);if(msg)toast(msg)}catch(e){toast(e.message);const er=$(\"d_err\");if(er)er.textContent=e.message}busy(false)}\nasync function openPO(id,keepScroll){openId=id;const y=keepScroll&&$(\"drawer\").classList.contains(\"open\")?$(\"drawer\").scrollTop:0;let r;\n try{r=await call(\"hubDetailLoad\",id)}catch(e){toast(e.message);return}\n const s=r.sub,is=r.issue,tp=isTPL(r.wh),c=r.counts;\n const cell=(v,k)=>{if(v==null||v===\"\")return `<td style=\"color:var(--muted)\">—</td>`;const exp=numv(r[k]),val=numv(v),skip=(k===\"plt\"&&r.floor)||exp==null;const bad=!skip&&(k===\"lbs\"?(exp&&Math.abs(val-exp)/exp>0.01):val!==exp),d=val-exp;return `<td style=\"${bad?\"color:#b0182f;font-weight:600;background:#fdf6f7\":\"\"}\">${fmt(val)}${bad?`<div style=\"font-size:11px\">${d>0?\"+\":\"\"}${fmt(d)}</div>`:\"\"}</td>`};\n const qrow=(l,o,sub)=>`<tr><td style=\"font-weight:500\">${esc(l)}${sub?`<div class=\"small\" style=\"margin:0\">${esc(sub)}</div>`:\"\"}</td>${o?cell(o.plt,\"plt\")+cell(o.cs,\"cs\")+cell(o.lbs,\"lbs\"):`<td colspan=\"3\" style=\"color:var(--muted)\">Not entered</td>`}</tr>`;\n const sum=is&&is.status===\"Open\"?`⚠ ${is.type} · Owner ${is.owner} · since ${fds(is.opened)}${is.overdue?\" · OVERDUE\":\"\"}`:s?`✓ Arrived ${fd(s.recv)} · ${fmt(s.cs)} cs · ${qcText(r)}`:`ETA ${fd(r.eta)} at ${r.wh||\"?\"}`;\n const av=r.avail;\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"Close\">×</button><div class=\"dsub\">Purchase order · ${esc(r.wh||\"no destination\")} · ${esc(r.supplier)}</div><h2 class=\"mono\" id=\"dtitle\" tabindex=\"-1\">PO ${esc(r.po)}</h2><div>${esc(r.item)} · ${esc(r.desc)}</div>\n  <div style=\"margin-top:8px;display:flex;gap:6px;flex-wrap:wrap\">${r.unplanned?\"\":allocPill(r)}<span class=\"pill ${r.qc===\"Locked\"?\"p-warn\":r.qc===\"Released\"?\"p-done\":\"p-grey\"}\">${qcText(r)}</span>${r.sla&&r.sla.lvl!==\"up\"?slaPill(r):\"\"}</div>\n  <div class=\"summary ${is&&is.status===\"Open\"?\"\":s?\"ok\":\"info\"}\">${esc(sum)}</div>\n  ${!tp&&!s&&!r.closed&&r.wh&&!r.unplanned?`<button type=\"button\" class=\"btn\" id=\"recvtop\" style=\"width:100%;margin-top:10px\">Record receiving at ${esc(r.wh)}</button>`:\"\"}</div>\n <div class=\"db\"><div class=\"ferr\" id=\"d_err\" role=\"alert\"></div>\n  ${r.unplanned?`<div class=\"sec\"><div class=\"card lvl-esc\"><b>Arrived at ${esc(r.wh)} but not on the upcoming list</b><div class=\"small\">Reported by ${esc(r.unplanned.by)} · ${esc(r.unplanned.at)} · container ${esc(r.cntr||\"—\")} · ${fmt(r.plt)} pallets · supplier on paperwork: ${esc(r.supplier)}</div>\n   <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"up_po\">PO number (SAP)</label><input class=\"in\" id=\"up_po\" value=\"${/^NEW-/.test(r.po)?\"\":esc(r.po)}\"></div><div><label for=\"up_item\">Item code</label><input class=\"in\" id=\"up_item\"></div></div>\n   <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"up_cs\">Expected cases</label><input class=\"in\" type=\"number\" min=\"0\" id=\"up_cs\"></div><div><label for=\"up_lbs\">Expected net lb</label><input class=\"in\" type=\"number\" min=\"0\" id=\"up_lbs\"></div></div>\n   <button type=\"button\" class=\"btn\" id=\"up_ok\" style=\"margin-top:8px\">Confirm PO and add to plan</button></div></div>`:\"\"}\n  ${r.delay?`<div class=\"sec\"><div class=\"card\" style=\"border-color:#c9b3e3;background:#f8f4fc\"><b style=\"color:#5b3a91\">⏸ Delay reported by ${esc(r.wh)}: ${esc(r.delay.reasonText)}</b><div style=\"margin-top:6px;font-size:13px\">${esc(r.delay.note)}</div><div class=\"small\">${esc(r.delay.by)} · ${esc(r.delay.at)}${r.delay.newDate?\" · new date \"+fd(r.delay.newDate):\"\"} · owner ${esc(r.delay.owner)}</div></div></div>`\n   :late(r)?`<div class=\"sec\"><div class=\"card lvl-${r.sla.lvl}\">${slaPill(r)} <span class=\"small\">${esc(r.wh)} has not sent the receiving. ${r.sla.h} business hours on the clock.</span>${tp?`<div style=\"margin-top:8px\"><button type=\"button\" class=\"btn ghost\" id=\"nudge\" ${r.nudged===DATA.today?\"disabled\":\"\"}>${r.nudged===DATA.today?\"✓ Nudged today\":\"📨 Nudge \"+esc(r.wh)}</button></div>`:\"\"}</div></div>`:\"\"}\n  ${is?`<div class=\"sec\"><div class=\"issue ${is.status===\"Open\"?\"\":\"closed\"} ${is.overdue?\"overdue\":\"\"}\"><div style=\"display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap\"><b>${is.status===\"Open\"?\"⚠ Open issue\":\"✓ Resolved\"}: ${esc(is.type)}</b><span style=\"font-size:12px\">${is.urgent?\"Urgent\":\"Normal\"} · since ${fds(is.opened)}</span></div>\n   <div class=\"row r3\" style=\"margin-top:10px\"><div><label for=\"is_owner\">Owner (who must act)</label><select class=\"in\" id=\"is_owner\">${OWNERS.map(o=>`<option ${o===is.owner?\"selected\":\"\"}>${o}</option>`).join(\"\")}</select></div>\n    <div><label for=\"is_urg\">Priority</label><select class=\"in\" id=\"is_urg\"><option value=\"1\" ${is.urgent?\"selected\":\"\"}>Urgent</option><option value=\"0\" ${!is.urgent?\"selected\":\"\"}>Normal</option></select></div>\n    <div><label for=\"is_st\">Status</label><select class=\"in\" id=\"is_st\">${[\"Open\",\"Resolved\"].map(o=>`<option ${o===is.status?\"selected\":\"\"}>${o}</option>`).join(\"\")}</select></div></div>\n   ${(s&&s.damage)||is.decision?`<div class=\"row r2\" style=\"margin-top:10px\"><div><label for=\"is_dec\">Dock decision</label><select class=\"in\" id=\"is_dec\">${DECISIONS.map(o=>`<option ${o===(is.decision||\"Pending\")?\"selected\":\"\"}>${o}</option>`).join(\"\")}</select></div><div><label for=\"is_decn\">Reason</label><input class=\"in\" id=\"is_decn\" value=\"${esc(is.decisionNote||\"\")}\"></div></div>`:\"\"}\n   ${r.claim?`<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-top:12px\">${s&&tp&&!r.req?`<button type=\"button\" class=\"btn ghost\" id=\"askre\">↻ Ask ${esc(r.wh)} for another count (${c.length} of 3)</button>`:\"\"}<button type=\"button\" class=\"btn ghost\" id=\"claim\">📋 Build claim package (NCR)</button></div>`:\"\"}\n   <div id=\"askbox\"></div><div id=\"claimbox\"></div></div></div>`:\"\"}\n  ${r.lines&&r.lines.length?`<div class=\"sec\"><h3>${r.cntr?\"Container \"+esc(r.cntr):\"Shipment\"} · ${r.lines.length} item${r.lines.length>1?\"s\":\"\"}${r.shipPlt!=null?\" · \"+fmt(r.shipPlt)+\" pallets\":\"\"}</h3><div class=\"card\" style=\"padding:6px 10px\">\n   ${r.lines.map(x=>`<div style=\"display:flex;gap:8px;justify-content:space-between;padding:6px 0;border-bottom:1px solid #f3f3f0;font-size:13px\"><span>${x.id===r.id?\"▸ \":\"\"}<b class=\"mono\">${esc(x.po)} / ${esc(x.item)}</b> ${esc(x.desc)}</span><span class=\"small\" style=\"margin:0;white-space:nowrap\">${x.sub?fmt(x.sub.cs)+\" cs received\":x.done?\"Before the hub\":(x.cs===\"\"||x.cs==null?\"no expected cs\":fmt(x.cs)+\" cs expected\")}</span></div>`).join(\"\")}\n   ${r.lines.some(x=>!x.done)?`<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-top:10px\"><button type=\"button\" class=\"btn ghost\" id=\"prehub\">Received before the hub</button><span class=\"small\" style=\"align-self:center\">Use this if the shipment arrived before the pilot started. It stops the late alerts.</span></div><div id=\"phbox\"></div>`:\"\"}\n   ${r.closed?`<button type=\"button\" class=\"btn ghost\" id=\"reopen\" style=\"margin-top:8px\">Reopen this line (receiving expected)</button>`:\"\"}</div>\n   <details class=\"sec\" ${r.lines.some(x=>!x.done&&(x.cs===\"\"||x.cs==null))?\"open\":\"\"}><summary>Expected cases for all lines of PO ${esc(r.po)}</summary>\n    ${r.lines.filter(x=>x.po===r.po&&!x.done).map(x=>`<div style=\"display:flex;gap:8px;align-items:center;margin-top:6px\"><label for=\"bc_${esc(x.id)}\" style=\"flex:1;margin:0;font-size:13px\"><b class=\"mono\">${esc(x.item)}</b> ${esc(x.desc)}</label><input class=\"in\" style=\"max-width:120px\" type=\"number\" min=\"0\" id=\"bc_${esc(x.id)}\" data-bc=\"${esc(x.id)}\" value=\"${esc(x.cs)}\"></div>`).join(\"\")}\n    <button type=\"button\" class=\"btn ghost\" id=\"bc_save\" style=\"margin-top:8px\">Save expected cases</button></details></div>`:\"\"}\n  ${r.unplanned?\"\":`<div class=\"sec\"><h3>Who it's for</h3><div class=\"card\">\n   <div class=\"chips\" role=\"group\" aria-label=\"Allocation\">${[\"Production\",\"Sales\",\"Split\",\"\"].map(a=>`<button type=\"button\" data-a=\"${a}\" aria-pressed=\"${r.alloc===a}\" class=\"${r.alloc===a?\"on\":\"\"}\">${a||\"Not set\"}</button>`).join(\"\")}</div>\n   <div id=\"splitbox\" style=\"${r.alloc===\"Split\"?\"\":\"display:none\"}\"><div class=\"row r2\" style=\"margin-top:10px\"><div><label for=\"a_prod\">Cases for production</label><input class=\"in\" type=\"number\" min=\"0\" id=\"a_prod\" value=\"${esc(r.prodQty)}\"></div><div><label for=\"a_sales\">Cases for sales</label><input class=\"in\" type=\"number\" min=\"0\" id=\"a_sales\" value=\"${esc(r.salesQty)}\"></div></div><div class=\"small\">Must add up to ${fmt(r.cs)} cs.</div></div>\n   <div class=\"row r2\" style=\"margin-top:10px\"><div><label for=\"a_so\">Linked sales order</label><input class=\"in\" id=\"a_so\" value=\"${esc(r.so)}\"></div><div><label for=\"a_qc\">QC status</label><select class=\"in\" id=\"a_qc\">${[\"Not received\",\"Locked\",\"Released\"].map(o=>`<option ${o===r.qc?\"selected\":\"\"}>${o}</option>`).join(\"\")}</select></div></div>\n   ${av?`<div class=\"note\" style=\"margin-top:10px\"><b>Available for sales: ${fmt(av.qty)} cs ${av.est?(av.qcLate?\"· QC release overdue (planned \"+fds(av.date)+\")\":\"around \"+fd(av.date)):\"from \"+fd(av.date)}</b></div>`:\"\"}\n   <div style=\"margin-top:10px\"><label>Buyer's PO instructions (from SAP, read-only)</label><div class=\"note\" style=\"white-space:pre-wrap\">${esc(r.instrSap||\"No instructions entered.\")}</div></div>\n   <div style=\"margin-top:10px\"><label for=\"a_extra\">Extra packaging / loading conditions</label><textarea id=\"a_extra\" rows=\"2\">${esc(r.extra)}</textarea></div>\n   ${tp?`<div style=\"margin-top:10px\"><label for=\"a_tpl\">Instructions shown to ${esc(r.wh)} (the only Tootsi text the 3PL sees)</label><textarea id=\"a_tpl\" rows=\"2\">${esc(r.tplInstr)}</textarea></div>`:\"\"}\n   <button type=\"button\" class=\"btn ghost\" id=\"a_save\" style=\"margin-top:8px\">Save</button></div></div>`}\n  <div class=\"sec\"><h3>Supplier packing slip</h3><div class=\"card\">${r.slip&&r.slip.status===\"FOUND\"?`<div style=\"display:flex;gap:8px;flex-wrap:wrap;align-items:center\"><span class=\"pill p-done\">On file</span>${r.slip.has?`<button type=\"button\" class=\"btn ghost\" id=\"slipopen\">📄 Open packing slip</button>`:\"\"}${r.slip.url?`<a href=\"${esc(r.slip.url)}\" target=\"_blank\" rel=\"noopener\" class=\"small\">Open in Drive</a>`:\"\"}</div><div class=\"small\">${esc(r.slip.source)}</div>`\n   :r.slip?`<span class=\"pill p-dmg\">Missing</span><div class=\"small\">${esc(r.slip.source)}</div>${r.slip.action?`<div class=\"note\" style=\"margin-top:6px\">${esc(r.slip.action)}</div>`:\"\"}`:`<span class=\"pill p-grey\">Not in Arash's packing slip log yet</span>`}\n   ${r.masterRecv&&!r.sub&&!r.closed?`<div class=\"note\" style=\"margin-top:10px\"><b>Master sheet shows this line received on ${fd(r.masterRecv)}.</b> If that is right, use \"Received before the hub\" above.</div>`:\"\"}</div></div>\n  <div class=\"sec\"><h3>Quantity check</h3><div class=\"card\" style=\"padding:6px 8px;overflow:auto\"><table class=\"qt\"><thead><tr><th></th><th>Pallets${r.floor?\" *\":\"\"}</th><th>Cases</th><th>Net lb</th></tr></thead><tbody>\n   ${qrow(\"Expected (PO)\",{plt:r.plt,cs:r.cs,lbs:r.lbs})}${qrow(\"Supplier packing list\",r.pl,r.pl?\"entered by \"+r.pl.by:\"\")}${qrow(\"Supplier loading report\",r.lr,r.lr?\"entered by \"+r.lr.by:\"\")}${c.map(z=>qrow(z.l,z.v,z.by)).join(\"\")}</tbody></table></div>\n   ${s&&s.diffNote?`<div class=\"note\" style=\"margin-top:8px\"><b>Receiver note:</b> ${esc(s.diffNote)}</div>`:\"\"}\n   <details class=\"sec\" ${!s&&(r.cs===\"\"||r.cs==null)?\"open\":\"\"}><summary>Edit expected quantities / destination</summary>\n    <div class=\"row r3\" style=\"margin-top:6px\"><div><label for=\"e_plt\">Pallets</label><input class=\"in\" type=\"number\" min=\"0\" id=\"e_plt\" value=\"${esc(r.plt)}\"></div><div><label for=\"e_cs\">Cases</label><input class=\"in\" type=\"number\" min=\"0\" id=\"e_cs\" value=\"${esc(r.cs)}\"></div><div><label for=\"e_lbs\">Net lb</label><input class=\"in\" type=\"number\" min=\"0\" id=\"e_lbs\" value=\"${esc(r.lbs)}\"></div></div>\n    <div class=\"row r2\" style=\"margin-top:6px\"><div><label for=\"e_wh\">Destination</label><select class=\"in\" id=\"e_wh\"><option value=\"\">Not set</option>${WHS.map(w=>`<option ${w===r.wh?\"selected\":\"\"}>${w}</option>`).join(\"\")}</select></div><div><label for=\"e_eta\">Expected date</label><input class=\"in\" type=\"date\" id=\"e_eta\" value=\"${esc(r.eta)}\"></div></div>\n    <button type=\"button\" class=\"btn ghost\" id=\"e_save\" style=\"margin-top:8px\">Save expected data</button></details>\n   <details class=\"sec\"><summary>Enter supplier quantities (packing list / loading report)</summary>${[\"pl\",\"lr\"].map(k=>{const o=r[k];return `<div style=\"margin-top:6px\"><b style=\"font-size:13px\">${k===\"pl\"?\"Packing list\":\"Loading report\"}</b><div class=\"row r3\" style=\"margin-top:6px\"><div><label for=\"${k}_plt\">Pallets</label><input class=\"in\" type=\"number\" min=\"0\" id=\"${k}_plt\" value=\"${o?esc(o.plt):\"\"}\"></div><div><label for=\"${k}_cs\">Cases</label><input class=\"in\" type=\"number\" min=\"0\" id=\"${k}_cs\" value=\"${o?esc(o.cs):\"\"}\"></div><div><label for=\"${k}_lbs\">Net lb</label><input class=\"in\" type=\"number\" min=\"0\" id=\"${k}_lbs\" value=\"${o?esc(o.lbs):\"\"}\"></div></div><button type=\"button\" class=\"btn ghost\" data-saveq=\"${k}\" style=\"margin-top:8px\">Save</button></div>`}).join(\"\")}</details>\n   ${!tp&&!s&&!r.closed&&r.wh&&!r.unplanned?`<button type=\"button\" class=\"btn\" id=\"recvown\" style=\"width:100%;margin-top:12px\">Record receiving at ${esc(r.wh)}${r.lines&&r.lines.filter(x=>!x.done).length>1?\" (\"+r.lines.filter(x=>!x.done).length+\" items)\":\"\"}</button>`:\"\"}</div>\n  <div class=\"sec\"><h3>Timeline</h3><ul class=\"timeline\">${r.timeline.map((z,i)=>`<li><span class=\"dot ${z.done?\"done\":z.bad?\"bad\":\"\"}\" aria-hidden=\"true\"></span><span>${esc(z.t)}</span><span class=\"when\" style=\"${z.bad?\"color:#b0182f\":\"\"}\">${esc(z.when)}</span></li>`).join(\"\")}</ul>\n   ${s&&!r.sap?`<button type=\"button\" class=\"btn ghost\" id=\"sapbtn\" style=\"margin-top:8px\">Mark as entered in SAP / WMS</button>`:\"\"}</div>\n  <details class=\"sec\" ${r.files.length?\"open\":\"\"}><summary>Receiving files (${r.files.length})</summary>\n   ${s?`<div class=\"grid2\" style=\"margin-bottom:10px\"><span class=\"k\">Received</span><span>${fd(s.recv)} · ${esc(s.by)}</span><span class=\"k\">Lot / expiry</span><span class=\"mono\">${esc(s.lot)} / ${esc(s.exp)}</span><span class=\"k\">Condition</span><span>${esc(s.cond)}${s.damage?\" · \"+esc(s.dmgText):\"\"}</span>${s.comment?`<span class=\"k\">Comment</span><span>${esc(s.comment)}</span>`:\"\"}</div>`:\"\"}\n   ${r.files.length?r.files.map(f=>`<div style=\"padding:6px 0;border-bottom:1px solid #f3f3f0;font-size:13px\">📎 <b>${esc(f.cat)}</b>${f.old?' <span class=\"pill p-grey\" style=\"font-size:10.5px;padding:1px 7px\">earlier submission</span>':\"\"} · <a href=\"${esc(f.url)}\" target=\"_blank\" rel=\"noopener\">${esc(f.name)}</a></div>`).join(\"\"):`<div class=\"note\">No files yet.</div>`}</details>\n  <details class=\"sec\" ${r.docs.length?\"open\":\"\"}><summary>Tootsi documents (${r.docs.length})</summary>\n   ${r.docs.map(d=>`<div style=\"padding:6px 0;border-bottom:1px solid #f3f3f0;font-size:13px\">📄 <b>${esc(d.cat)}</b> · <a href=\"${esc(d.url)}\" target=\"_blank\" rel=\"noopener\">${esc(d.name)}</a> <span class=\"small\">${esc(d.by)} · ${esc(d.at)}</span></div>`).join(\"\")}\n   <div class=\"row r2\" style=\"margin-top:10px\"><div><label for=\"i_cat\">Type</label><select class=\"in\" id=\"i_cat\">${CATS.map(cc=>`<option>${cc}</option>`).join(\"\")}</select></div><div><label class=\"drop\" style=\"margin-top:18px\"><input type=\"file\" id=\"i_file\">📎 Add a file</label></div></div></details>\n  <details class=\"sec\" open><summary>Notes (${r.notes.length})</summary>\n   ${r.notes.map(n=>`<div class=\"msg\" style=\"${n.ext===true||n.ext===\"TRUE\"?\"border-left-color:var(--blue)\":\"\"}\"><b>${esc(n.who)}</b> <span class=\"t\">${esc(n.at)}</span><div>${esc(n.text)}</div></div>`).join(\"\")||`<div class=\"note\">No notes yet.</div>`}\n   <div style=\"display:flex;gap:8px;margin-top:8px\"><label for=\"n_txt\" class=\"sr\">Note</label><input class=\"in\" id=\"n_txt\" placeholder=\"Write a note for everyone on this PO…\"><button type=\"button\" class=\"btn\" id=\"n_add\">Post</button></div></details>\n </div>`);\n $(\"drawer\").scrollTop=y;\n const v=id2=>{const el=$(id2);return el?el.value:\"\"};\n document.querySelectorAll(\"[data-a]\").forEach(b=>b.onclick=()=>{document.querySelectorAll(\"[data-a]\").forEach(x=>{x.classList.toggle(\"on\",x===b);x.setAttribute(\"aria-pressed\",x===b)});$(\"splitbox\").style.display=b.dataset.a===\"Split\"?\"\":\"none\";if(b.dataset.a!==\"Split\")$(\"a_save\").click()});\n const sv=$(\"a_save\");if(sv)sv.onclick=()=>{const a=(document.querySelector(\"[data-a].on\")||{dataset:{a:\"\"}}).dataset.a;act(id,\"who\",{alloc:a,prodQty:v(\"a_prod\"),salesQty:v(\"a_sales\"),so:v(\"a_so\"),qc:v(\"a_qc\"),extra:v(\"a_extra\"),...($(\"a_tpl\")?{tplInstr:v(\"a_tpl\")}:{})},\"Saved.\")};\n $(\"e_save\").onclick=()=>act(id,\"expected\",{plt:v(\"e_plt\"),cs:v(\"e_cs\"),lbs:v(\"e_lbs\"),wh:v(\"e_wh\"),eta:v(\"e_eta\")},\"Expected data saved.\");\n document.querySelectorAll(\"[data-saveq]\").forEach(b=>b.onclick=()=>{const k=b.dataset.saveq;act(id,\"supq\",{kind:k,plt:v(k+\"_plt\"),cs:v(k+\"_cs\"),lbs:v(k+\"_lbs\")},\"Saved.\")});\n $(\"n_add\").onclick=()=>{if(v(\"n_txt\").trim())act(id,\"note\",{text:v(\"n_txt\")})};$(\"n_txt\").onkeydown=e=>{if(e.key===\"Enter\")$(\"n_add\").click()};\n $(\"i_file\").onchange=async e=>{const f=e.target.files[0];if(!f)return;busy(true);try{const rf=await readFile(f);busy(false);act(id,\"doc\",{cat:v(\"i_cat\"),file:{name:rf.name,type:rf.type,data:rf.data}},\"File added.\")}catch(er){busy(false);toast(er.message)}};\n const sb=$(\"sapbtn\");if(sb)sb.onclick=()=>act(id,\"sap\",{},\"Marked as entered in SAP.\");\n const nd=$(\"nudge\");if(nd)nd.onclick=()=>act(id,\"nudge\",{},\"Nudge sent.\");\n const uo=$(\"up_ok\");if(uo)uo.onclick=async()=>{const d={po:v(\"up_po\"),item:v(\"up_item\"),cs:v(\"up_cs\"),lbs:v(\"up_lbs\")};busy(true);\n  try{await call(\"hubAction\",id,\"confirmUnplanned\",d);busy(false);await Promise.all([load(),openPO(id,true)]);toast(\"PO confirmed.\")}\n  catch(e){busy(false);const m=String(e.message).match(/^MOVE\\?(\\w+)\\?(\\d+)/);\n   if(!m){toast(e.message);return}\n   // the PO is planned at another warehouse: ask before moving its lines\n   $(\"d_err\").innerHTML=`<div class=\"confirm\">PO ${esc(d.po)} is planned for <b>${m[1]}</b> (${m[2]} line(s)). It arrived at ${esc(r.wh)}. Move those lines to ${esc(r.wh)}?<div style=\"display:flex;gap:8px;margin-top:10px\"><button type=\"button\" class=\"btn\" id=\"mvyes\">Yes, move to ${esc(r.wh)}</button><button type=\"button\" class=\"btn ghost\" id=\"mvno\">Cancel</button></div></div>`;\n   $(\"mvno\").onclick=()=>{$(\"d_err\").innerHTML=\"\"};$(\"mvyes\").onclick=()=>act(id,\"confirmUnplanned\",Object.assign(d,{move:true}),\"PO confirmed and moved.\")}};\n if(is){$(\"is_owner\").onchange=()=>act(id,\"issue\",{owner:v(\"is_owner\")});$(\"is_urg\").onchange=()=>act(id,\"issue\",{urgent:v(\"is_urg\")===\"1\"});$(\"is_st\").onchange=()=>act(id,\"issue\",{status:v(\"is_st\")});\n  const dec=$(\"is_dec\");if(dec)dec.onchange=()=>act(id,\"issue\",{decision:v(\"is_dec\"),decisionNote:v(\"is_decn\")});\n  const a=$(\"askre\");if(a)a.onclick=()=>{$(\"askbox\").innerHTML=`<div class=\"confirm\">Ask ${esc(r.wh)} to count PO ${esc(r.po)} again? Their current count, photos and documents stay on file.<div style=\"display:flex;gap:8px;margin-top:10px\"><button type=\"button\" class=\"btn\" id=\"askyes\">Yes, ask ${esc(r.wh)}</button><button type=\"button\" class=\"btn ghost\" id=\"askno\">Cancel</button></div></div>`;$(\"askyes\").focus();$(\"askno\").onclick=()=>{$(\"askbox\").innerHTML=\"\"};$(\"askyes\").onclick=()=>act(id,\"recount\",{},r.wh+\" will see the recount request.\")};\n  const cl=$(\"claim\");if(cl)cl.onclick=()=>{$(\"claimbox\").innerHTML=`<textarea id=\"ct\" rows=\"16\" style=\"margin-top:10px;font-family:'IBM Plex Mono',monospace;font-size:12px\" aria-label=\"Claim summary\">${esc(r.claim)}</textarea><button type=\"button\" class=\"btn\" id=\"cc\" style=\"margin-top:8px\">Copy for email to supplier</button>`;$(\"cc\").onclick=async()=>{try{await navigator.clipboard.writeText($(\"ct\").value);toast(\"Copied.\")}catch(e){$(\"ct\").select();toast(\"Select all and copy (Ctrl+C).\")}}};}\n const so=$(\"slipopen\");if(so)so.onclick=async()=>{busy(true);try{openFile(await call(\"hubSlip\",id),\"Open packing slip\")}catch(e){toast(e.message)}busy(false)};\n const rt=$(\"recvtop\");if(rt)rt.onclick=()=>$(\"recvown\").click();\n const ro=$(\"recvown\");if(ro)ro.onclick=async()=>{if(!(await ensureReceive(i=>call(\"hubDiag\",i))))return toast(\"Could not load the receiving form. Reload the page.\");HUBLINES=r.lines.filter(x=>!x.done).map(x=>({id:x.id,po:x.po,item:x.item,desc:x.desc}));openReceive({id:r.id,po:r.po,item:r.item,desc:r.desc,eta:r.eta,cntr:r.cntr,plt:r.plt,load:r.load,floor:r.floor,instr:\"\",req:false,shipPlt:r.shipPlt,askP:r.askP})};\n const ph=$(\"prehub\");if(ph)ph.onclick=()=>{$(\"phbox\").innerHTML=`<div class=\"confirm\">Mark the ${r.lines.filter(x=>!x.done).length} open line(s) of this shipment as received before the hub? The 3PL will no longer see them.<div style=\"display:flex;gap:8px;margin-top:10px\"><button type=\"button\" class=\"btn\" id=\"phyes\">Yes, mark as received</button><button type=\"button\" class=\"btn ghost\" id=\"phno\">Cancel</button></div></div>`;$(\"phyes\").focus();$(\"phno\").onclick=()=>{$(\"phbox\").innerHTML=\"\"};$(\"phyes\").onclick=()=>act(id,\"prehub\",{},\"Marked as received before the hub.\")};\n const rop=$(\"reopen\");if(rop)rop.onclick=()=>act(id,\"reopen\",{},\"Reopened.\");\n const bcs=$(\"bc_save\");if(bcs)bcs.onclick=()=>{const d={};document.querySelectorAll(\"[data-bc]\").forEach(i=>{if(i.value!==\"\")d[i.dataset.bc]=i.value});act(id,\"expectedBulk\",{cs:d},\"Expected cases saved.\")};\n}\n\n// ---------- add PO ----------\nfunction openAddPO(){openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"Close\">×</button><h2 id=\"dtitle\" tabindex=\"-1\">Add a PO</h2><div class=\"dsub\">For POs that are not in the upcoming sheet (LTL, small orders, parcels).</div></div>\n <div class=\"db\"><form id=\"nf\" novalidate><div class=\"ferr\" id=\"d_err\" role=\"alert\"></div>\n  <div class=\"row r3\"><div><label for=\"n_po\">PO number *</label><input class=\"in\" id=\"n_po\" name=\"po\"></div><div><label for=\"n_item\">Item code *</label><input class=\"in\" id=\"n_item\" name=\"item\"></div><div><label for=\"n_wh\">Going to *</label><select class=\"in\" id=\"n_wh\" name=\"wh\"><option value=\"\">Choose…</option>${WHS.map(w=>`<option>${w}</option>`).join(\"\")}</select></div></div>\n  <div style=\"margin-top:8px\"><label for=\"n_desc\">Product *</label><input class=\"in\" id=\"n_desc\" name=\"desc\"></div>\n  <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"n_sup\">Supplier *</label><input class=\"in\" id=\"n_sup\" name=\"supplier\"></div><div><label for=\"n_eta\">Expected arrival *</label><input class=\"in\" type=\"date\" id=\"n_eta\" name=\"eta\"></div></div>\n  <div class=\"row r3\" style=\"margin-top:8px\"><div><label for=\"n_plt\">Pallets *</label><input class=\"in\" type=\"number\" min=\"0\" id=\"n_plt\" name=\"plt\"></div><div><label for=\"n_cs\">Cases *</label><input class=\"in\" type=\"number\" min=\"0\" id=\"n_cs\" name=\"cs\"></div><div><label for=\"n_lbs\">Net lb *</label><input class=\"in\" type=\"number\" min=\"0\" id=\"n_lbs\" name=\"lbs\"></div></div>\n  <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"n_load\">Shipment</label><select class=\"in\" id=\"n_load\" name=\"load\"><option>LTL · Palletized</option><option>FTL · Palletized</option><option>FCL 20' · Floor loaded</option><option>FCL 40' · Floor loaded</option><option>FCL 40' · Palletized</option></select></div><div><label for=\"n_cntr\">Container #</label><input class=\"in\" id=\"n_cntr\" name=\"cntr\"></div></div>\n  <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"n_alloc\">Allocation</label><select class=\"in\" id=\"n_alloc\" name=\"alloc\"><option value=\"\">Not set</option><option>Production</option><option>Sales</option></select></div><div><label for=\"n_so\">Sales order</label><input class=\"in\" id=\"n_so\" name=\"so\"></div></div>\n  <div style=\"margin-top:8px\"><label for=\"n_instr\">PO instructions (as in SAP)</label><textarea id=\"n_instr\" name=\"instr\" rows=\"2\"></textarea></div>\n  <button class=\"btn\" type=\"submit\" style=\"width:100%;margin-top:14px\">Add PO</button></form></div>`);\n $(\"nf\").onsubmit=async e=>{e.preventDefault();const f=Object.fromEntries(new FormData(e.target));busy(true);try{await call(\"hubAction\",\"\",\"addPO\",f);closeD();await load();toast(\"PO \"+f.po+\" added.\")}catch(er){$(\"d_err\").textContent=er.message}busy(false)}}\n\nload();\n</script>\n</body></html>\n", "Portal": "<!DOCTYPE html>\n<html lang=\"<?= lang ?>\">\n<head><base target=\"_top\"><meta charset=\"utf-8\"><?!= include('Styles') ?></head>\n<body>\n<header>\n  <?!= include('Logo') ?>\n  <div><h1 id=\"ttl\">3PL <span>Portal</span></h1><div class=\"hsub\" id=\"sub\"></div></div>\n  <div class=\"who\"><div class=\"lang\" role=\"group\" aria-label=\"Language\"><button type=\"button\" data-lang=\"en\">EN</button><button type=\"button\" data-lang=\"fr\">FR</button></div><div class=\"nm\"><div style=\"font-weight:500\" id=\"pname\"></div></div><button type=\"button\" class=\"btn ghost\" id=\"logout\" style=\"display:none;min-height:36px;padding:4px 12px;font-size:13px\"></button></div>\n</header>\n<div id=\"login\" class=\"hide\"></div>\n<main id=\"main\"><div class=\"loading\">…</div></main>\n<div class=\"scrim\" id=\"scrim\"></div>\n<aside class=\"drawer\" id=\"drawer\" role=\"dialog\" aria-modal=\"true\" aria-labelledby=\"dtitle\" tabindex=\"-1\"></aside>\n<div class=\"toast\" id=\"toast\" role=\"status\" aria-live=\"polite\"></div>\n<div class=\"busy\" id=\"busy\" aria-live=\"polite\">…</div>\n<?!= include('Client') ?>\n<?!= include('Lang') ?>\n<?!= include('Receive') ?>\n<script>\nlet TOKEN=null,WH=\"\",NAME=\"\";const HUB_URL=<?!= JSON.stringify(hubUrl) ?>;\nlet lang=<?!= JSON.stringify(lang) ?>;try{lang=localStorage.getItem(\"tootsi_lang\")||lang}catch(e){}\nObject.assign(L.en,{slip:\"Packing slip\",open:\"Open\"});Object.assign(L.fr,{slip:\"Bordereau d'expédition\",open:\"Ouvrir\"});\n/** Server messages are in English: show the 3PL the French version. */\nconst ERR_FR={\"PO not found.\":\"Commande introuvable.\",\"Reason and details are required.\":\"La raison et les détails sont obligatoires.\",\"The new date must be within the next 3 weeks.\":\"La nouvelle date doit être dans les 3 prochaines semaines.\",\n \"Paperwork photo and product are required.\":\"La photo des documents et le produit sont obligatoires.\",\"Choose the load condition.\":\"Choisissez l'état du chargement.\",\"Describe the damage.\":\"Décrivez les dommages.\",\n \"Photos and documents are missing.\":\"Il manque des photos ou des documents.\",\"Explain the difference.\":\"Expliquez l'écart.\",\"Enter the number of pallets.\":\"Entrez le nombre de palettes.\",\n \"A photo was not uploaded. Please add it again.\":\"Une photo n'a pas été envoyée. Ajoutez-la de nouveau.\",\"A photo was already used in another receipt. Please add it again.\":\"Cette photo a déjà servi pour une autre réception. Ajoutez-la de nouveau.\",\n \"Only photos and PDF files can be added.\":\"Seulement des photos et des fichiers PDF.\",\"File too large (10 MB max).\":\"Fichier trop gros (10 Mo max).\",\"Too many photos this hour. Please wait a few minutes.\":\"Trop de photos cette heure-ci. Attendez quelques minutes.\",\n \"No packing slip yet.\":\"Pas encore de bordereau.\",\"The hub is busy saving someone else's change. Please try again in a few seconds.\":\"Le système est occupé. Réessayez dans quelques secondes.\",\n \"Enter every item of this shipment (0 cases if an item did not come).\":\"Entrez chaque article de la livraison (0 caisse si un article n'est pas arrivé).\"};\nERR_FR['This PO is already in your list. Look for it under \"To receive\".']=\"Cette commande est déjà dans votre liste, sous « À recevoir ».\";\nfunction errT(m){m=String(m||\"\");if(lang!==\"fr\")return m;if(ERR_FR[m])return ERR_FR[m];\n let x=m.match(/^Lot and expiry are required for item (\\S+)\\./);if(x)return\"Lot et date d'expiration obligatoires pour l'article \"+x[1]+\". L'expiration doit être après la date de réception.\";\n x=m.match(/^Label photo missing for item (\\S+)\\./);if(x)return\"Photo de l'étiquette manquante pour l'article \"+x[1]+\".\";\n x=m.match(/^Enter cases \\(0 or more\\) for item (\\S+)\\./);if(x)return\"Entrez les caisses (0 ou plus) pour l'article \"+x[1]+\".\";\n if(/date received/i.test(m))return\"Vérifiez la date de réception (pas dans le futur, pas plus de 60 jours).\";\n if(/Failed to fetch|NetworkError|network/i.test(m))return\"Pas de connexion. Vos données sont gardées; réessayez.\";return m}\nObject.assign(L.en,{search:\"Search\",search_ph:\"Search PO, container or product…\"});Object.assign(L.fr,{search:\"Rechercher\",search_ph:\"Rechercher une commande, un conteneur ou un produit…\"});\nObject.assign(L.en,{lg_h:\"Sign in\",lg_s:\"3PL partners: use the username and password Tootsi gave you.\",lg_user:\"Username\",lg_pw:\"Password\",lg_btn:\"Sign in\",lg_bad:\"Wrong username or password.\",lg_locked:\"Too many tries. Wait 10 minutes, or call Tootsi Inventory Control.\",lg_exp:\"Your session ended. Please sign in again.\",lg_staff:\"Tootsi staff\",lg_staffp:\"Use your Tootsi Google account. No extra password.\",lg_staffb:\"Open the PO Hub with Google\",lg_forgot:\"Forgot your password? Call Tootsi Inventory Control. We reset it in one minute.\",lg_tag:\"Tootsi deliveries at your warehouse.\",lg_p:\"See the POs coming to you, receive them and send the photos and counts to Tootsi.\",lg_l1:\"Only your warehouse's deliveries\",lg_l2:\"Blind count with automatic recount\",lg_l3:\"Photos and receiving document in one step\",lg_l4:\"Report a delay in two taps\",logout:\"Sign out\"});\nObject.assign(L.fr,{lg_h:\"Connexion\",lg_s:\"Partenaires 3PL : utilisez le nom d'utilisateur et le mot de passe fournis par Tootsi.\",lg_user:\"Nom d'utilisateur\",lg_pw:\"Mot de passe\",lg_btn:\"Se connecter\",lg_bad:\"Nom d'utilisateur ou mot de passe incorrect.\",lg_locked:\"Trop d'essais. Attendez 10 minutes ou appelez le contrôle d'inventaire Tootsi.\",lg_exp:\"Votre session est terminée. Veuillez vous reconnecter.\",lg_staff:\"Employés Tootsi\",lg_staffp:\"Utilisez votre compte Google Tootsi. Aucun autre mot de passe.\",lg_staffb:\"Ouvrir le PO Hub avec Google\",lg_forgot:\"Mot de passe oublié ? Appelez le contrôle d'inventaire Tootsi. On le réinitialise en une minute.\",lg_tag:\"Les livraisons Tootsi à votre entrepôt.\",lg_p:\"Voyez les commandes qui arrivent chez vous, recevez-les et envoyez photos et quantités à Tootsi.\",lg_l1:\"Seulement les livraisons de votre entrepôt\",lg_l2:\"Comptage à l'aveugle avec recomptage automatique\",lg_l3:\"Photos et bon de réception en une étape\",lg_l4:\"Signaler un retard en deux touches\",logout:\"Déconnexion\"});\nconst t=k=>(L[lang]&&L[lang][k])||L.en[k]||k;\nconst CONDS=[[\"Good\",\"✅\",\"c_Good\"],[\"Shifted load\",\"↔️\",\"c_Shifted\"],[\"Collapsed pallets\",\"⬇️\",\"c_Collapsed\"],[\"Double-stacked\",\"⏫\",\"c_Double\"],[\"Torn / leaking bags\",\"🩹\",\"c_Torn\"],[\"Wet / dirty\",\"💧\",\"c_Wet\"],[\"Other\",\"❓\",\"c_Other\"]];\nconst REASONS=[\"r_notarr\",\"r_customs\",\"r_docs\",\"r_seal\",\"r_wrong\",\"r_dmg\",\"r_dock\",\"r_other\"];\nconst LOAD_FR={\"Palletized\":\"sur palettes\",\"Floor loaded\":\"en vrac au plancher\",\"FTL\":\"Camion complet\",\"LTL\":\"Camion partiel\",\"FCL 40'\":\"Conteneur 40 pi\",\"FCL 20'\":\"Conteneur 20 pi\"};\nconst loadTxt=s=>lang===\"fr\"?String(s||\"\").split(\" · \").map(x=>LOAD_FR[x.trim()]||x).join(\" · \"):s;\nconst ICON={up:\"⚪\",tom:\"🔵\",due:\"🟡\",nudge:\"🟠\",esc:\"🔴\",paused:\"⏸\"};\nlet DATA=null,tab=\"todo\";\n\nfunction setLang(l){lang=l;try{localStorage.setItem(\"tootsi_lang\",l)}catch(e){}document.documentElement.lang=l;LOCALE=l===\"fr\"?\"fr-CA\":\"en-CA\";\n document.querySelectorAll(\"[data-lang]\").forEach(b=>{b.classList.toggle(\"on\",b.dataset.lang===l);b.setAttribute(\"aria-pressed\",b.dataset.lang===l)});\n $(\"ttl\").innerHTML=l===\"fr\"?'Portail <span>3PL</span>':'3PL <span>Portal</span>';$(\"sub\").textContent=t(\"portal_sub\");$(\"logout\").textContent=t(\"logout\");if(!TOKEN)showLogin(loginMsg);else if(DATA)render()}\ndocument.querySelectorAll(\"[data-lang]\").forEach(b=>b.onclick=()=>setLang(b.dataset.lang));\n\nasync function load(){try{ensureReceive(i=>call(\"portalDiag\",TOKEN,i));DATA=await call(\"portalLoad\",TOKEN);WH=DATA.wh;NAME=DATA.name;$(\"pname\").textContent=NAME;$(\"logout\").style.display=\"\";render()}catch(e){if(e.message===\"SESSION\")return signedOut(\"lg_exp\");$(\"main\").innerHTML=`<div class=\"err\">${esc(e.message)}</div>`}}\n// ---------- sign in ----------\nlet loginMsg=\"\";\nconst SKEY=\"tootsi_3pl_session\";\nfunction signedOut(msg){TOKEN=null;DATA=null;try{localStorage.removeItem(SKEY)}catch(e){}closeD&&$(\"drawer\").classList.contains(\"open\")&&closeD();$(\"pname\").textContent=\"\";$(\"logout\").style.display=\"none\";$(\"main\").innerHTML=\"\";showLogin(msg||\"\")}\nfunction showLogin(msg){loginMsg=msg;const lg=$(\"login\");lg.classList.remove(\"hide\");\n lg.innerHTML=`<div class=\"lg-side\"><div style=\"position:relative;z-index:1;font-size:13px;letter-spacing:2px;color:#c9cbcd;text-transform:uppercase\">PO Hub · Tootsi Impex</div>\n  <div><h2>${t(\"lg_tag\")}</h2><p>${t(\"lg_p\")}</p><ul><li>${t(\"lg_l1\")}</li><li>${t(\"lg_l2\")}</li><li>${t(\"lg_l3\")}</li><li>${t(\"lg_l4\")}</li></ul></div>\n  <div style=\"font-size:12px;color:#8d9195;position:relative;z-index:1\">Tootsi Impex · Inventory Control</div></div>\n <div class=\"lg-main\"><form class=\"lg-box\" id=\"lgf\" novalidate>\n  <div class=\"lg-top\">${document.querySelector(\"header img\").outerHTML}<div class=\"lang\" role=\"group\" aria-label=\"Language\">${[\"en\",\"fr\"].map(l=>`<button type=\"button\" data-lglang=\"${l}\" class=\"${lang===l?\"on\":\"\"}\" aria-pressed=\"${lang===l}\">${l.toUpperCase()}</button>`).join(\"\")}</div></div>\n  <h1>${t(\"lg_h\")}</h1><p class=\"s\">${t(\"lg_s\")}</p>\n  <label for=\"lg_u\">${t(\"lg_user\")}</label><input class=\"in\" id=\"lg_u\" autocomplete=\"username\" autocapitalize=\"none\" spellcheck=\"false\">\n  <label for=\"lg_p\">${t(\"lg_pw\")}</label><input class=\"in\" id=\"lg_p\" type=\"password\" autocomplete=\"current-password\">\n  <div class=\"lg-err\" id=\"lg_err\" role=\"alert\">${msg?esc(t(msg)):\"\"}</div>\n  <button class=\"btn\" type=\"submit\">${t(\"lg_btn\")}</button>\n  <p class=\"small\" style=\"margin-top:12px\">${t(\"lg_forgot\")}</p>\n  <div class=\"lg-demo\"><h4>${t(\"lg_staff\")}</h4><p>${t(\"lg_staffp\")}</p><a class=\"btn ghost\" style=\"width:100%;text-decoration:none;display:flex;align-items:center;justify-content:center\" href=\"${esc(HUB_URL||\"#\")}\" target=\"_top\">${t(\"lg_staffb\")}</a></div>\n </form></div>`;\n lg.querySelectorAll(\"[data-lglang]\").forEach(b=>b.onclick=()=>setLang(b.dataset.lglang));\n const u=$(\"lg_u\");setTimeout(()=>u.focus(),50);\n $(\"lgf\").onsubmit=async e=>{e.preventDefault();const er=$(\"lg_err\"),btn=e.target.querySelector(\"button[type=submit]\");er.textContent=\"\";btn.disabled=true;\n  try{const r=await call(\"portalLogin\",$(\"lg_u\").value,$(\"lg_p\").value);TOKEN=r.sid;try{localStorage.setItem(SKEY,r.sid);if(!localStorage.getItem(\"tootsi_lang\"))lang=r.lang}catch(x){}\n   loginMsg=\"\";lg.classList.add(\"hide\");lg.innerHTML=\"\";setLang(lang);$(\"main\").innerHTML='<div class=\"loading\">…</div>';await load()}\n  catch(x){er.textContent=t(x.message===\"LOCKED\"?\"lg_locked\":x.message===\"BADLOGIN\"?\"lg_bad\":\"\")||x.message;btn.disabled=false;$(\"lg_p\").value=\"\";$(\"lg_p\").focus()}}}\n\nconst slaPill=r=>r.sla?`<span class=\"sla sla-${r.sla.lvl}\"><span aria-hidden=\"true\">${ICON[r.sla.lvl]}</span> ${esc(t(\"sla_\"+r.sla.lvl))}</span>`:\"\";\n\nconst LVL=[\"esc\",\"nudge\",\"due\",\"paused\",\"tom\",\"up\"],lvRank=r=>{const i=r.sla?LVL.indexOf(r.sla.lvl):-1;return i<0?99:i};\nlet q=\"\";\n/** Group lines that arrive together (same container, or same PO without a container) into one shipment card. */\nfunction shipments(rows){const m=new Map();rows.forEach(r=>{if(!m.has(r.ship))m.set(r.ship,[]);m.get(r.ship).push(r)});\n return[...m.values()].map(ls=>{const rep=ls.slice().sort((a,b)=>lvRank(a)-lvRank(b))[0];return Object.assign({},rep,{lines:ls,req:ls.some(x=>x.req),delay:ls.map(x=>x.delay).find(Boolean)||null,pos:[...new Set(ls.map(x=>x.po))]})})}\nfunction render(){\n const rows=DATA.rows,todoL=rows.filter(r=>!r.sent),done=rows.filter(r=>r.sent);\n const ql=q.trim().toLowerCase(),hit=g=>!ql||g.lines.some(x=>[x.po,x.cntr,x.item,x.desc].join(\" \").toLowerCase().includes(ql));\n const todo=shipments(todoL).filter(hit);\n const lv=r=>r.sla&&r.sla.lvl;\n const req=todo.filter(r=>r.req),rest=todo.filter(r=>!r.req);\n const late=rest.filter(r=>[\"esc\",\"nudge\"].includes(lv(r))),paused=rest.filter(r=>lv(r)===\"paused\"),today=rest.filter(r=>lv(r)===\"due\"),tom=rest.filter(r=>lv(r)===\"tom\"),later=rest.filter(r=>!lv(r)||lv(r)===\"up\");\n const card=r=>{const many=r.lines&&r.lines.length>1,names=(r.lines||[r]).map(x=>x.desc);\n  return `<div class=\"pocard lvl-${r.sent?\"none\":(lv(r)||\"none\")}\"><button type=\"button\" class=\"po ${r.req?\"req\":\"\"}\" data-po=\"${esc(r.id)}\">\n   <div><div class=\"big mono\">${r.cntr&&!r.sent?esc(r.cntr):\"PO \"+esc((r.pos||[r.po]).join(\", \"))}</div>\n   <div class=\"m\">${r.cntr&&!r.sent?\"PO \"+esc((r.pos||[r.po]).join(\", \"))+\" · \":\"\"}${many?`<b>${r.lines.length} ${t(\"it_n\")}</b> · `:esc(r.item)+\" \"}${esc(names.slice(0,many?3:1).join(\" · \"))}${many&&names.length>3?\" · +\"+(names.length-3):\"\"}</div>\n   <div class=\"m\">${r.shipPlt!=null&&r.shipPlt!==\"\"?fmt(r.shipPlt)+\" \"+t(\"pallets\")+\" · \":\"\"}${esc(loadTxt(r.load))} · ${t(\"expected\")}${t(\"colon\")} <b style=\"color:var(--ink)\">${fd(r.eta)}</b></div>\n   ${!r.sent?`<div style=\"margin-top:6px\">${slaPill(r)}</div>`:\"\"}${r.delay?`<div class=\"m\" style=\"color:#5b3a91;margin-top:4px\">⏸ ${esc(t(r.delay.reason))}${r.delay.newDate?\" · \"+t(\"d_new\")+t(\"colon\")+\" \"+fds(r.delay.newDate):\"\"}</div>`:\"\"}</div>\n   ${r.sent?`<span class=\"pill ${r.sent.damage?\"p-dmg\":\"p-done\"}\">${r.sent.damage?t(\"sent_issue\"):t(\"sent\")}</span>`:`<span class=\"go\">${r.req?t(\"recount_req\"):t(\"receive\")} →</span>`}</button>\n   ${!r.sent?`<div style=\"display:flex\">${(r.lines||[r]).some(x=>x.slip)?`<button type=\"button\" class=\"delaybtn\" style=\"flex:1;color:var(--ink)\" data-slip=\"${esc((r.lines||[r]).find(x=>x.slip).id)}\">📄 ${t(\"slip\")}</button>`:\"\"}<button type=\"button\" class=\"delaybtn\" style=\"flex:1\" data-delay=\"${esc(r.id)}\">${r.delay?t(\"d_solve\"):t(\"delay_btn\")}</button></div>`:\"\"}</div>`};\n const grp=(title,arr,color)=>arr.length?`<div class=\"group\" ${color?`style=\"color:${color}\"`:\"\"}>${title}</div>${arr.map(card).join(\"\")}`:\"\";\n const cap=s=>s.charAt(0).toUpperCase()+s.slice(1);\n const counts=[late.length?`<b style=\"color:#b0182f\">${late.length}</b> ${t(\"late_n\")}`:\"\",today.length?`<b>${today.length}</b> ${t(\"today\")}`:\"\",`<b>${shipments(todoL).length}</b> ${t(\"toRecv\")}`,req.length?`<b style=\"color:var(--amber)\">${req.length}</b> ${t(\"recounts\")}`:\"\"].filter(Boolean).join(\" · \");\n $(\"main\").innerHTML=`<div class=\"hello\"><h2 tabindex=\"-1\">${t(\"hi\")}, ${lang===\"fr\"?t(\"team\")+\" \"+esc(NAME):esc(NAME)+\" \"+t(\"team\")}</h2><p>${t(\"tap\")}</p></div>\n <p class=\"slarule\">⏱ ${t(\"sla_rule\")}</p><p class=\"chip-line\">${counts}</p>\n <button type=\"button\" class=\"unpl\" id=\"unpl\">${t(\"unpl_btn\")}</button>\n <label for=\"pq\" class=\"sr\">${t(\"search\")}</label><input id=\"pq\" class=\"in\" type=\"search\" placeholder=\"${esc(t(\"search_ph\"))}\" value=\"${esc(q)}\" style=\"margin:0 0 12px\">\n <div class=\"tabs\"><button type=\"button\" aria-pressed=\"${tab===\"todo\"}\" data-t2=\"todo\" class=\"${tab===\"todo\"?\"on\":\"\"}\">${t(\"tab_todo\")} (${shipments(todoL).length})</button><button type=\"button\" aria-pressed=\"${tab===\"done\"}\" data-t2=\"done\" class=\"${tab===\"done\"?\"on\":\"\"}\">${t(\"tab_done\")} (${done.length})</button></div>\n <div class=\"list\">${tab===\"todo\"?(todo.length?grp(t(\"g_recount\"),req,\"var(--amber)\")+grp(cap(t(\"late_n\")),late,\"#b0182f\")+grp(t(\"g_today\"),today)+grp(t(\"sla_paused\"),paused,\"#5b3a91\")+grp(t(\"g_tom\"),tom)+grp(t(\"g_later\"),later):`<div class=\"empty\">${t(\"empty_todo\")}</div>`):(done.length?done.map(card).join(\"\"):`<div class=\"empty\">${t(\"empty_done\")}</div>`)}</div>`;\n document.querySelectorAll(\"[data-t2]\").forEach(b=>b.onclick=()=>{tab=b.dataset.t2;render()});\n document.querySelectorAll(\".po\").forEach(el=>el.onclick=async()=>{const r=DATA.rows.find(x=>x.id===el.dataset.po);if(r.sent)return showSent(r);if(!(await ensureReceive(i=>call(\"portalDiag\",TOKEN,i))))return toast(t(\"up_err\"));openReceive(r)});\n const pq=$(\"pq\");pq.oninput=e=>{q=e.target.value;clearTimeout(pq.t);pq.t=setTimeout(()=>{render();const n=$(\"pq\");n.focus();n.setSelectionRange(n.value.length,n.value.length)},250)};\n document.querySelectorAll(\"[data-delay]\").forEach(b=>b.onclick=()=>openDelay(DATA.rows.find(x=>x.id===b.dataset.delay)));\n document.querySelectorAll(\"[data-slip]\").forEach(b=>b.onclick=async()=>{busy(true);try{openFile(await call(\"portalSlip\",TOKEN,b.dataset.slip),t(\"open\"))}catch(e){toast(errT(e.message))}busy(false)});\n $(\"unpl\").onclick=openUnplanned;\n}\n\nconst RCV={lines:r=>DATA.rows.filter(x=>x.ship===r.ship&&!x.sent),askP:r=>!!r.askP,fresh:r=>!!r.fresh,upload:f=>call(\"portalUpload\",TOKEN,f),slip:x=>call(\"portalSlip\",TOKEN,x.id),\n check:(r,c)=>call(\"portalCheckShip\",TOKEN,r.id,c),submit:(r,pl)=>call(\"portalSubmitShip\",TOKEN,r.id,pl),after:(r,next)=>{if(next)tab=\"todo\";load()}};\nfunction showSent(r){const s=r.sent,q=o=>`${fmt(o.plt)} ${t(\"u_plt\")} / ${fmt(o.cs)} ${t(\"u_cs\")}${o.lbs!==\"\"&&o.lbs!=null?\" / \"+fmt(o.lbs)+\" \"+t(\"u_lb\"):\"\"}`;\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><div class=\"dsub\">${t(\"ro\")} · ${esc(s.at)}</div><h2 class=\"mono\" id=\"dtitle\" tabindex=\"-1\">PO ${esc(r.po)}</h2><div>${esc(r.item)} · ${esc(r.desc)}</div></div>\n <div class=\"db\"><div class=\"sec\"><div class=\"grid2 box\"><span class=\"k\">${t(\"date\")}</span><span>${fd(s.recv)}</span>\n  ${s.first?`<span class=\"k\">${t(\"first_c\")}</span><span class=\"mono\">${q(s.first)}</span><span class=\"k\">${t(\"recount_c\")}</span><span class=\"mono\">${q(s)}</span>`:`<span class=\"k\">${t(\"final_c\")}</span><span class=\"mono\">${q(s)}</span>`}\n  ${s.diffNote?`<span class=\"k\">${t(\"note_c\")}</span><span>${esc(s.diffNote)}</span>`:\"\"}<span class=\"k\">${t(\"lot\")}</span><span class=\"mono\">${esc(s.lot)}</span><span class=\"k\">${t(\"exp\")}</span><span>${fdy(s.exp)}</span>\n  <span class=\"k\">${t(\"cond\")}</span><span style=\"${s.damage?\"color:#b0182f\":\"\"}\">${esc(t((CONDS.find(c=>c[0]===s.cond)||[,,s.cond])[2]))}${s.damage?\" · \"+esc(s.dmgText):\"\"}</span>${s.comment?`<span class=\"k\">${t(\"your_com\")}</span><span>${esc(s.comment)}</span>`:\"\"}</div></div>\n  ${s.files.length?`<div class=\"sec small\">${s.files.map(f=>\"📎 \"+esc(f.cat)+\" · \"+esc(f.name)).join(\"<br>\")}</div>`:\"\"}\n  ${r.slip?`<div class=\"sec\"><button type=\"button\" class=\"btn ghost\" id=\"sslip\" style=\"width:100%\">📄 ${t(\"slip\")}</button></div>`:\"\"}\n  <div class=\"sec\"><div class=\"note\">${t(\"ro_change\")}</div></div></div>`);\n const ss=$(\"sslip\");if(ss)ss.onclick=async()=>{busy(true);try{openFile(await call(\"portalSlip\",TOKEN,r.id),t(\"open\"))}catch(e){toast(errT(e.message))}busy(false)}}\n\n// ---------- delays ----------\nlet dform;\nfunction openDelay(r){cur=r;\n if(r.delay){openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><div class=\"dsub\">${t(\"d_cur\")}</div><h2 class=\"mono\" id=\"dtitle\" tabindex=\"-1\">PO ${esc(r.po)}</h2><div class=\"dsub\">${esc(t(r.delay.reason))}</div></div>\n  <div class=\"db\"><div class=\"sec\"><p style=\"font-size:15px\">${t(\"d_solve_q\")}</p><div style=\"display:flex;gap:10px;flex-wrap:wrap\"><button type=\"button\" class=\"btn\" id=\"dy\" style=\"min-height:52px;flex:1\">${t(\"yes\")}</button><button type=\"button\" class=\"btn ghost\" id=\"dn\" style=\"min-height:52px;flex:1\">${t(\"no\")}</button></div></div></div>`);\n  $(\"dn\").onclick=closeD;$(\"dy\").onclick=async()=>{busy(true);try{DATA=await call(\"portalSolveDelay\",TOKEN,r.id);closeD();render();toast(t(\"d_closed\"))}catch(e){toast(errT(e.message))}busy(false)};return}\n dform={reason:\"\",note:\"\",newDate:\"\",pics:[]};drawDelay()}\nfunction drawDelay(){const r=cur,miss=!dform.reason?t(\"d_miss_r\"):dform.note.trim().length<5?t(\"d_miss_n\"):\"\";\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><div class=\"dsub\">${t(\"delay_h\")}</div><h2 class=\"mono\" id=\"dtitle\" tabindex=\"-1\">PO ${esc(r.po)}</h2><div class=\"dsub\">${esc(r.desc)} · ${t(\"expected\")} ${fd(r.eta)}${r.cntr?\" · \"+esc(r.cntr):\"\"}</div></div>\n <div class=\"db\"><p style=\"font-size:14px\">${t(\"delay_p\")}</p><div class=\"step\"><h4>${t(\"d_reason\")}</h4><div class=\"cond\" style=\"grid-template-columns:repeat(auto-fill,minmax(160px,1fr))\">${REASONS.map(x=>`<button type=\"button\" data-rs=\"${x}\" aria-pressed=\"${dform.reason===x}\" class=\"${dform.reason===x?\"on\":\"\"}\">${t(x)}</button>`).join(\"\")}</div></div>\n  <div class=\"step\"><label for=\"d_note\" style=\"font-size:16px\">${t(\"d_note\")}</label><textarea id=\"d_note\" rows=\"3\" placeholder=\"${esc(t(\"d_note_ph\"))}\">${esc(dform.note)}</textarea>\n   <div style=\"margin-top:10px\"><label for=\"d_new\">${t(\"d_new\")}</label><input class=\"in big-in\" type=\"date\" id=\"d_new\" min=\"${DATA.today}\" value=\"${esc(dform.newDate)}\"></div>\n   <label class=\"drop\" style=\"margin-top:10px\"><input type=\"file\" accept=\"image/*,application/pdf\" multiple id=\"d_pic\">${t(\"d_pic\")}</label>${thumbs(dform.pics)}</div></div>\n <div class=\"foot\"><div class=\"prog\" id=\"dmiss\">${miss?`<span style=\"color:#b0182f;font-weight:600\">${t(\"d_miss\")}${t(\"colon\")} ${miss}</span>`:\"\"}</div><button type=\"button\" class=\"btn\" id=\"d_send\" ${miss?\"disabled\":\"\"}>${t(\"d_send\")}</button></div>`);\n const chk=()=>{const m=!dform.reason?t(\"d_miss_r\"):dform.note.trim().length<5?t(\"d_miss_n\"):\"\";$(\"d_send\").disabled=!!m;$(\"dmiss\").innerHTML=m?`<span style=\"color:#b0182f;font-weight:600\">${t(\"d_miss\")}${t(\"colon\")} ${m}</span>`:\"\"};\n document.querySelectorAll(\"[data-rs]\").forEach(b=>b.onclick=()=>{dform.reason=b.dataset.rs;dform.note=$(\"d_note\").value;dform.newDate=$(\"d_new\").value;drawDelay()});\n $(\"d_note\").oninput=e=>{dform.note=e.target.value;chk()};$(\"d_new\").oninput=e=>{dform.newDate=e.target.value};\n $(\"d_pic\").onchange=async e=>{busy(true);for(const f of e.target.files)dform.pics.push(await readFile(f));busy(false);dform.note=$(\"d_note\").value;drawDelay()};\n $(\"d_send\").onclick=async()=>{busy(true);try{DATA=await call(\"portalDelay\",TOKEN,r.id,{reason:dform.reason,note:$(\"d_note\").value,newDate:$(\"d_new\").value,files:dform.pics.map(f=>({name:f.name,type:f.type,data:f.data}))});closeD();render();toast(t(\"d_sent\"))}catch(e){toast(errT(e.message))}busy(false)}}\n\n// ---------- PO arrived that is not on the list ----------\nlet uform;\nfunction openUnplanned(){uform={po:\"\",cntr:\"\",sup:\"\",desc:\"\",plt:\"\",doc:[],load:[]};drawUnplanned()}\nfunction drawUnplanned(){const ok=uform.doc.length&&uform.desc.trim();\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><h2 id=\"dtitle\" tabindex=\"-1\">${t(\"unpl_h\")}</h2></div>\n <div class=\"db\"><p style=\"font-size:14px\">${t(\"unpl_p\")}</p><div class=\"step\"><label class=\"drop\"><input type=\"file\" accept=\"image/*,application/pdf\" multiple data-uf=\"doc\">${t(\"u_doc\")}</label>${thumbs(uform.doc)}\n  <label class=\"drop\" style=\"margin-top:10px\"><input type=\"file\" accept=\"image/*\" multiple data-uf=\"load\">${t(\"u_load\")}</label>${thumbs(uform.load)}</div>\n  <div class=\"step\"><div class=\"row r2\"><div><label for=\"u_po\">${t(\"u_po\")}</label><input class=\"in big-in\" id=\"u_po\" inputmode=\"numeric\" value=\"${esc(uform.po)}\"></div><div><label for=\"u_cntr\">${t(\"u_cntr\")}</label><input class=\"in big-in\" id=\"u_cntr\" value=\"${esc(uform.cntr)}\"></div></div>\n  <div style=\"margin-top:10px\"><label for=\"u_desc\">${t(\"u_desc\")}</label><input class=\"in big-in\" id=\"u_desc\" value=\"${esc(uform.desc)}\"></div>\n  <div class=\"row r2\" style=\"margin-top:10px\"><div><label for=\"u_sup\">${t(\"u_sup\")}</label><input class=\"in big-in\" id=\"u_sup\" value=\"${esc(uform.sup)}\"></div><div><label for=\"u_plt\">${t(\"u_plts\")}</label><input class=\"in big-in\" id=\"u_plt\" type=\"number\" inputmode=\"numeric\" min=\"0\" value=\"${esc(uform.plt)}\"></div></div></div></div>\n <div class=\"foot\"><div class=\"prog\">${ok?\"\":`<span style=\"color:#b0182f;font-weight:600\">${t(\"u_miss\")}</span>`}</div><button type=\"button\" class=\"btn\" id=\"u_send\" ${ok?\"\":\"disabled\"}>${t(\"u_send\")}</button></div>`);\n const keep=()=>[\"po\",\"cntr\",\"sup\",\"desc\",\"plt\"].forEach(k=>uform[k]=$(\"u_\"+k).value);\n [\"po\",\"cntr\",\"sup\",\"desc\",\"plt\"].forEach(k=>$(\"u_\"+k).oninput=()=>{keep();$(\"u_send\").disabled=!(uform.doc.length&&uform.desc.trim())});\n document.querySelectorAll(\"[data-uf]\").forEach(i=>i.onchange=async()=>{keep();busy(true);for(const f of i.files)uform[i.dataset.uf].push(await readFile(f));busy(false);drawUnplanned()});\n $(\"u_send\").onclick=async()=>{keep();busy(true);try{DATA=await call(\"portalUnplanned\",TOKEN,{po:uform.po,cntr:uform.cntr,sup:uform.sup,desc:uform.desc,plt:uform.plt,files:uform.doc.map(f=>({name:f.name,type:f.type,data:f.data})),loadFiles:uform.load.map(f=>({name:f.name,type:f.type,data:f.data}))});closeD();render();toast(t(\"u_sent\"))}catch(e){toast(errT(e.message))}busy(false)}}\n\n// a session can end mid-shift (12 h limit, password reset): send the user back to sign in\nconst _call=call;call=(...a)=>_call(...a).catch(e=>{if(e.message===\"SESSION\"&&a[0]!==\"portalLoad\"){signedOut(\"lg_exp\")}throw e});\n$(\"logout\").onclick=async()=>{const s=TOKEN;signedOut(\"\");try{await _call(\"portalLogout\",s)}catch(e){}};\ntry{TOKEN=localStorage.getItem(SKEY)}catch(e){}\nsetLang(lang);if(TOKEN)load();else showLogin(\"\");\n</script>\n</body></html>\n", "Receive": "<script id=\"rcvjs\">\n/* Receiving form for a whole SHIPMENT (one container, or one PO without a container), shared by the 3PL portal\n   and Tootsi receiving (BM/LY/HB). The page sets:\n   RCV = { lines(row) -> [rows of the same shipment still to receive], askP(row) -> true if pallets are counted now,\n           upload(file) -> {fid}, check(row,count) -> {match,off,limit}, submit(row,payload), after(row,next), title, nextLabel } */\nObject.assign(L.en,{it_h:\"Items in this shipment\",it_n:\"items\",it_cs:\"Cases\",it_lb:\"lb (optional)\",it_lbl:\"Label photo\",it_lot:\"Lot\",it_exp:\"Expiry\",it_zero:\"0 cases: this item did not come (no label needed).\",\n sh_plt:\"Total pallets in the shipment\",s2pm:\"Count the pallets for the whole shipment, then the cases of each item.\",s2rpm:\"These numbers are different from the PO. Count them again:\",\n limit:\"The count still does not match. Explain the difference below, then send.\",up_wait:\"Uploading photos…\",up_err:\"A photo did not upload. Add it again.\",s3m:\"For each item that came: a clear photo of the label, the lot and the expiry date.\",\n m3m:\"label photo, lot and expiry for each item\",off_plt:\"Total pallets\",ok_line:\"matches\",cnt_first:\"first count\"});\nObject.assign(L.fr,{it_h:\"Articles de cette livraison\",it_n:\"articles\",it_cs:\"Caisses\",it_lb:\"lb (facultatif)\",it_lbl:\"Photo de l'étiquette\",it_lot:\"Lot\",it_exp:\"Expiration\",it_zero:\"0 caisse : cet article n'est pas arrivé (pas d'étiquette).\",\n sh_plt:\"Total des palettes de la livraison\",s2pm:\"Comptez les palettes de toute la livraison, puis les caisses de chaque article.\",s2rpm:\"Ces chiffres ne correspondent pas à la commande. Recomptez-les :\",\n limit:\"Le compte ne correspond toujours pas. Expliquez l'écart ci-dessous, puis envoyez.\",up_wait:\"Envoi des photos…\",up_err:\"Une photo n'a pas été envoyée. Ajoutez-la de nouveau.\",s3m:\"Pour chaque article reçu : une photo claire de l'étiquette, le lot et la date d'expiration.\",\n m3m:\"photo, lot et expiration de chaque article\",off_plt:\"Total des palettes\",ok_line:\"correspond\",cnt_first:\"premier compte\"});\nlet form,cur,curLines;\nconst DKEY=r=>\"rcv_draft_\"+r.ship;\n/** The form is kept while the receiver works: closing the screen or reloading the page does not lose counts, lots or uploaded photos. */\nfunction saveDraft(){try{if(!form||!cur)return;const strip=a=>a.filter(f=>f.fid).map(f=>({name:f.name,fid:f.fid,preview:null}));\n const d=Object.assign({},form,{loadPics:strip(form.loadPics),doc:strip(form.doc),dmgPics:strip(form.dmgPics),li:Object.fromEntries(Object.entries(form.li).map(([k,v])=>[k,Object.assign({},v,{label:strip(v.label)})])),up:0,at:Date.now()});\n localStorage.setItem(DKEY(cur),JSON.stringify(d))}catch(e){}}\nfunction loadDraft(r,ids){try{const d=JSON.parse(localStorage.getItem(DKEY(r))||\"null\");if(!d||Date.now()-d.at>2*864e5)return null;\n if(Object.keys(d.li).sort().join()!==ids.slice().sort().join())return null;return d}catch(e){return null}}\nfunction clearDraft(r){try{localStorage.removeItem(DKEY(r))}catch(e){}}\nfunction openReceive(r){cur=r;curLines=RCV.lines(r);const askP=RCV.askP(r),fresh=RCV.fresh?RCV.fresh(r):askP;\n const old=loadDraft(r,curLines.map(x=>x.id));if(old){form=old;form.askP=askP;form.fresh=fresh;drawReceive();return}\n form={recv:DATA.today,askP,fresh,plt:\"\",rplt:\"\",li:{},loadPics:[],doc:[],cond:\"\",dmgText:\"\",dmgPics:[],legible:false,stage:\"count\",checked:false,checkOk:false,off:[],needNote:false,limit:false,diffNote:\"\",comment:\"\",banner:false,up:0};\n curLines.forEach(x=>form.li[x.id]={cs:\"\",lbs:\"\",rcs:\"\",rlbs:\"\",lot:\"\",exp:\"\",label:[]});drawReceive()}\nconst isOff=k=>form.stage===\"recount\"&&form.off.indexOf(k)>=0;\n/** The count that is checked and sent: recount values replace the first count for the lines that were off. */\nfunction curCount(){return{plt:form.askP?numv(isOff(\"pallets\")?form.rplt:form.plt):null,lines:curLines.map(x=>{const v=form.li[x.id],o=isOff(x.id);return{id:x.id,cs:numv(o?v.rcs:v.cs),lbs:numv(o?v.rlbs:v.lbs)}})}}\nconst came=id=>{const c=curCount().lines.find(z=>z.id===id);return c&&c.cs>0};\nfunction steps(){const bad=form.cond&&form.cond!==\"Good\",c=curCount();\n const nums=(!form.askP||(c.plt!=null&&c.plt>=0))&&c.lines.every(z=>z.cs!=null&&z.cs>=0&&(z.lbs==null||z.lbs>=0)),dateOk=!!form.recv&&form.recv<=DATA.today;\n const labelsOk=curLines.every(x=>{if(!came(x.id))return true;const v=form.li[x.id];return v.label.length>0&&v.lot.trim()&&v.exp&&v.exp>form.recv});\n const expErr=curLines.some(x=>{const v=form.li[x.id];return came(x.id)&&v.exp&&v.exp<=form.recv});\n return[{k:\"s1\",ok:!form.fresh||form.loadPics.length>0,m:\"m1\"},\n  {k:\"s2\",ok:nums&&dateOk&&form.checked&&(!form.needNote||form.diffNote.trim().length>=10),m:!nums||!dateOk?(form.stage===\"recount\"?\"m2r\":\"m2\"):!form.checked?\"m2c\":\"m2n\",err:!dateOk&&form.recv?\"e_date\":\"\"},\n  {k:\"s3\",ok:labelsOk&&form.legible,m:\"m3m\",err:expErr?\"e_exp\":\"\"},\n  {k:\"s4\",ok:form.doc.length>0,m:\"m4\"},\n  {k:\"s5\",ok:!!form.cond&&(!bad||(form.dmgText.trim().length>=10&&form.dmgPics.length>0)),m:bad?\"m5n\":\"m5\"}]}\nfunction foot(){const f=$(\"foot\");if(!f)return;const s=steps(),n=s.filter(x=>x.ok).length,miss=s.find(x=>!x.ok),onlyCheck=miss&&miss.m===\"m2c\"&&s.filter(x=>!x.ok).length===1;\n f.innerHTML=`<div class=\"prog\">${s.map(x=>`<i class=\"${x.ok?\"on\":\"\"}\" aria-hidden=\"true\"></i>`).join(\"\")}<span>${form.up?t(\"up_wait\"):t(\"steps\")+t(\"colon\")+\" \"+n+\"/5\"}</span>${miss&&!form.up?`<button type=\"button\" class=\"miss\" data-goto=\"${miss.k}\">${t(\"missing\")}${t(\"colon\")} ${t(miss.m)}</button>`:\"\"}</div>\n <button type=\"button\" class=\"btn\" id=\"send\" ${(miss&&!onlyCheck)||form.up?\"disabled\":\"\"}>${t(\"send\")}</button>`;\n s.forEach(x=>{const el=$(\"st_\"+x.k);if(el)el.classList.toggle(\"ok\",!!x.ok);const er=$(\"err_\"+x.k);if(er)er.textContent=x.err?t(x.err):\"\"});\n const g=f.querySelector(\"[data-goto]\");if(g)g.onclick=()=>{$(\"st_\"+g.dataset.goto).scrollIntoView({behavior:\"smooth\",block:\"start\"})};\n $(\"send\").onmousedown=e=>e.preventDefault();$(\"send\").onclick=submit}\n/** Upload each photo as soon as it is chosen (one server call per file). */\nasync function addFiles(list,arr){for(const file of list){let rf;try{rf=await readFile(file)}catch(e){toast(e.message);continue}\n const item={name:rf.name,preview:rf.preview,fid:null};arr.push(item);form.up++;drawReceive();\n try{const r=await RCV.upload({name:rf.name,type:rf.type,data:rf.data});item.fid=r.fid}catch(e){arr.splice(arr.indexOf(item),1);toast(t(\"up_err\")+\" \"+(typeof errT===\"function\"?errT(e.message):e.message))}\n form.up--;drawReceive()}}\nfunction drawReceive(){\n const r=cur,recount=form.stage===\"recount\",bad=form.cond&&form.cond!==\"Good\",n=i=>`<span class=\"num\" aria-hidden=\"true\">${i}</span>`,many=curLines.length>1;\n const inp=(id,val,mode,ph)=>`<input class=\"in\" id=\"${id}\" type=\"number\" inputmode=\"${mode||\"numeric\"}\" min=\"0\" step=\"any\" value=\"${esc(val)}\" ${ph?`placeholder=\"${esc(ph)}\"`:\"\"}>`;\n const y=$(\"drawer\").classList.contains(\"open\")?$(\"drawer\").scrollTop:0;\n const pos=[...new Set(curLines.map(x=>x.po))].join(\", \");\n const lineCount=x=>{const v=form.li[x.id],o=isOff(x.id);\n  return `<div class=\"itrow ${o?\"warn\":\"\"}\"><div class=\"itname\"><b class=\"mono\">${esc(x.item)}</b> ${esc(x.desc)}${many&&curLines.some(z=>z.po!==x.po)?` <span class=\"small\">PO ${esc(x.po)}</span>`:\"\"}</div>\n   ${recount&&!o?`<div class=\"small\" style=\"margin:0\">${fmt(v.cs)} ${t(\"u_cs\")}${v.lbs!==\"\"?\" · \"+fmt(v.lbs)+\" \"+t(\"u_lb\"):\"\"} · ${t(\"ok_line\")} ✓</div>`\n   :`<div class=\"itin\"><div class=\"unit\"><label for=\"c_${esc(x.id)}\">${t(\"it_cs\")}${o?` <span class=\"small\">(${t(\"cnt_first\")}${t(\"colon\")} ${fmt(v.cs)})</span>`:\"\"}</label>${inp(\"c_\"+x.id,o?v.rcs:v.cs)}<i aria-hidden=\"true\">${t(\"u_cs\")}</i></div>\n    <div class=\"unit\"><label for=\"w_${esc(x.id)}\">${t(\"it_lb\")}</label>${inp(\"w_\"+x.id,o?v.rlbs:v.lbs,\"decimal\")}<i aria-hidden=\"true\">${t(\"u_lb\")}</i></div></div>`}</div>`};\n const lineLabel=x=>{const v=form.li[x.id],z=curCount().lines.find(q=>q.id===x.id);if(z&&z.cs===0)return `<div class=\"itrow\"><div class=\"itname\"><b class=\"mono\">${esc(x.item)}</b> ${esc(x.desc)}</div><div class=\"small\" style=\"margin:0\">${t(\"it_zero\")}</div></div>`;\n  return `<div class=\"itrow\"><div class=\"itname\"><b class=\"mono\">${esc(x.item)}</b> ${esc(x.desc)}</div>\n   <label class=\"drop\"><input type=\"file\" accept=\"image/*\" multiple data-lf=\"${esc(x.id)}\">📷 ${t(\"it_lbl\")}</label>${thumbs(v.label)}\n   <div class=\"row r2\" style=\"margin-top:8px\"><div><label for=\"l_${esc(x.id)}\">${t(\"it_lot\")}</label><input class=\"in\" id=\"l_${esc(x.id)}\" value=\"${esc(v.lot)}\" autocapitalize=\"characters\" autocomplete=\"off\"></div>\n   <div><label for=\"e_${esc(x.id)}\">${t(\"it_exp\")}</label><input class=\"in\" id=\"e_${esc(x.id)}\" type=\"date\" value=\"${esc(v.exp)}\"></div></div></div>`};\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><div class=\"dsub\">${RCV.title||t(\"r_for\")}</div>\n  <h2 class=\"mono\" id=\"dtitle\" tabindex=\"-1\">${r.cntr?esc(r.cntr):\"PO \"+esc(pos)}</h2><div>${r.cntr?\"PO \"+esc(pos)+\" · \":\"\"}${many?curLines.length+\" \"+t(\"it_n\"):esc(r.item)+\" · \"+esc(r.desc)}</div>\n  <div class=\"dsub\" style=\"margin-top:4px\">${esc(loadTxt(r.load))}${form.askP&&r.shipPlt!=null&&r.shipPlt!==\"\"?\" · \"+fmt(r.shipPlt)+\" \"+t(\"pallets\"):\"\"} · ${t(\"expected\")} ${fd(r.eta)}</div></div>\n <div class=\"db\">\n  ${form.banner?`<div class=\"sec\"><div class=\"banner\" role=\"alert\">${t(\"notsent\")}</div></div>`:r.req?`<div class=\"sec\"><div class=\"banner\">${t(\"r_req\")}</div></div>`:\"\"}\n  ${r.instr?`<div class=\"sec\"><div class=\"instr\"><b>${t(\"instr\")}${t(\"colon\")}</b> ${esc(r.instr)}</div></div>`:\"\"}\n  ${RCV.slip&&curLines.some(x=>x.slip)?`<div class=\"sec\"><button type=\"button\" class=\"btn ghost\" id=\"rslip\" style=\"width:100%\">📄 ${t(\"slip\")}</button></div>`:\"\"}\n  <div class=\"sec\">\n   ${form.fresh?`<div class=\"step\" id=\"st_s1\"><h4>${n(1)}${t(\"s1\")}</h4><p>${t(\"s1p\")}</p><label class=\"drop\"><input type=\"file\" accept=\"image/*\" multiple data-f=\"loadPics\">${t(\"s1b\")}</label>${thumbs(form.loadPics)}</div>`:`<div id=\"st_s1\"></div>`}\n   <div class=\"step ${recount?\"warn\":\"\"}\" id=\"st_s2\"><h4>${n(2)}${recount?t(\"s2r\"):t(\"s2\")}</h4>\n    ${recount?`<p><b>${t(\"s2rpm\")}</b><br>${t(\"s2r_saved\")}</p>`:`<p>${many?t(\"s2pm\"):(r.floor?t(\"s2pf\"):t(\"s2p\"))} ${t(\"lbo_h\")}</p><div style=\"margin-bottom:10px\"><label for=\"f_recv\">${t(\"date\")}</label><input class=\"in big-in\" id=\"f_recv\" type=\"date\" value=\"${esc(form.recv)}\" max=\"${DATA.today}\"></div>`}\n    ${form.askP?(recount&&!isOff(\"pallets\")?`<div class=\"small\">${t(\"sh_plt\")}${t(\"colon\")} ${fmt(form.plt)} · ${t(\"ok_line\")} ✓</div>`\n      :`<div class=\"itrow ${isOff(\"pallets\")?\"warn\":\"\"}\"><div class=\"unit\"><label for=\"f_plt\">${many?t(\"sh_plt\"):(r.floor?t(\"pltb\"):t(\"plt\"))}${isOff(\"pallets\")?` <span class=\"small\">(${t(\"cnt_first\")}${t(\"colon\")} ${fmt(form.plt)})</span>`:\"\"}</label>${inp(isOff(\"pallets\")?\"f_rplt\":\"f_plt\",isOff(\"pallets\")?form.rplt:form.plt)}<i aria-hidden=\"true\">${t(\"u_plt\")}</i></div></div>`):\"\"}\n    ${many?`<h5 style=\"margin:12px 0 4px\">${t(\"it_h\")} (${curLines.length})</h5>`:\"\"}${curLines.map(lineCount).join(\"\")}\n    <div class=\"ferr\" id=\"err_s2\" role=\"alert\"></div>\n    ${form.needNote?`<p style=\"margin:12px 0 6px;color:var(--amber)\"><b>${form.limit?t(\"limit\"):t(\"s2n\")}</b></p><label for=\"f_note\" class=\"sr\">${t(\"s2n\")}</label><textarea id=\"f_note\" rows=\"2\" placeholder=\"${esc(t(\"s2n_ph\"))}\">${esc(form.diffNote)}</textarea>`:\"\"}\n    ${form.checked&&form.checkOk?`<p style=\"margin:10px 0 0;color:var(--green);font-weight:500\">${t(\"check_ok\")}</p>`:\"\"}\n    ${!form.checked?`<button type=\"button\" class=\"btn dark\" id=\"chk\" style=\"width:100%;margin-top:12px;min-height:52px\">${t(\"check\")}</button>`:\"\"}</div>\n   <div class=\"step\" id=\"st_s3\"><h4>${n(3)}${t(\"s3\")}</h4><p>${many?t(\"s3m\"):t(\"s3p\")}</p>${curLines.map(lineLabel).join(\"\")}\n    <div class=\"ferr\" id=\"err_s3\" role=\"alert\"></div><label class=\"tick\"><input type=\"checkbox\" id=\"f_leg\" ${form.legible?\"checked\":\"\"}> ${t(\"leg\")}</label></div>\n   <div class=\"step\" id=\"st_s4\"><h4>${n(4)}${t(\"s4\")}</h4><p>${t(\"s4p\")}</p><label class=\"drop\"><input type=\"file\" accept=\"image/*,application/pdf\" multiple data-f=\"doc\">${t(\"s4b\")}</label>${thumbs(form.doc)}</div>\n   <div class=\"step\" id=\"st_s5\"><h4>${n(5)}${t(\"s5\")}</h4><p>${t(\"s5p\")}</p>\n    <div class=\"cond\" role=\"group\" aria-label=\"${esc(t(\"s5\"))}\">${CONDS.map(c=>`<button type=\"button\" data-c=\"${esc(c[0])}\" aria-pressed=\"${form.cond===c[0]}\" class=\"${form.cond===c[0]?\"on\":\"\"} ${c[0]!==\"Good\"?\"badc\":\"\"}\"><span aria-hidden=\"true\">${c[1]}</span>${t(c[2])}</button>`).join(\"\")}</div>\n    ${bad?`<div style=\"margin-top:12px\"><label for=\"f_dmg\" class=\"sr\">${t(\"s5n_ph\")}</label><textarea id=\"f_dmg\" rows=\"3\" placeholder=\"${esc(t(\"s5n_ph\"))}\">${esc(form.dmgText)}</textarea><label class=\"drop\" style=\"margin-top:10px\"><input type=\"file\" accept=\"image/*\" multiple data-f=\"dmgPics\">${t(\"s5b\")}</label>${thumbs(form.dmgPics)}</div>`:\"\"}</div>\n   <div class=\"step\"><h4><span class=\"num\" style=\"background:#6b7075\" aria-hidden=\"true\">💬</span><label for=\"f_com\" style=\"display:inline;margin:0;font-size:16px\">${t(\"com\")}</label> <span style=\"font-weight:400;color:var(--muted);font-size:13px\">· ${t(\"opt\")}</span></h4><p>${t(\"comp\")}</p>\n    <textarea id=\"f_com\" rows=\"3\" maxlength=\"500\" placeholder=\"${esc(t(\"com_ph\"))}\">${esc(form.comment)}</textarea></div>\n  </div></div><div class=\"foot\" id=\"foot\"></div>`);\n $(\"drawer\").scrollTop=y;foot();saveDraft();\n const rd=()=>drawReceive(),reset=()=>{if(form.checked&&!form.limit){form.checked=form.checkOk=form.needNote=false;rd();return true}return false};\n const bindNum=(id,set)=>{const el=$(id);if(el)el.oninput=e=>{set(e.target.value);if(reset())return;foot();saveDraft()}};\n bindNum(\"f_plt\",v=>form.plt=v);bindNum(\"f_rplt\",v=>form.rplt=v);\n curLines.forEach(x=>{const v=form.li[x.id],o=isOff(x.id);bindNum(\"c_\"+x.id,z=>o?v.rcs=z:v.cs=z);bindNum(\"w_\"+x.id,z=>o?v.rlbs=z:v.lbs=z);\n  const lo=$(\"l_\"+x.id);if(lo)lo.oninput=e=>{v.lot=e.target.value;foot();saveDraft()};const ex=$(\"e_\"+x.id);if(ex)ex.oninput=e=>{v.exp=e.target.value;foot();saveDraft()}});\n [[\"f_recv\",\"recv\"],[\"f_note\",\"diffNote\"],[\"f_dmg\",\"dmgText\"],[\"f_com\",\"comment\"]].forEach(([id,k])=>{const el=$(id);if(el)el.oninput=e=>{form[k]=e.target.value;foot();saveDraft()}});\n $(\"f_leg\").onchange=e=>{form.legible=e.target.checked;foot()};\n document.querySelectorAll(\"[data-f]\").forEach(i=>i.onchange=()=>addFiles([...i.files],form[i.dataset.f]));\n document.querySelectorAll(\"[data-lf]\").forEach(i=>i.onchange=()=>addFiles([...i.files],form.li[i.dataset.lf].label));\n document.querySelectorAll(\"[data-c]\").forEach(b=>b.onclick=()=>{form.cond=b.dataset.c;rd()});\n const rs=$(\"rslip\");if(rs)rs.onclick=async()=>{busy(true);try{openFile(await RCV.slip(curLines.find(x=>x.slip)),t(\"open\"))}catch(e){toast(e.message)}busy(false)};\n const ck=$(\"chk\");if(ck){ck.onmousedown=e=>e.preventDefault();ck.onclick=()=>runCheck(false)}\n}\n/** Ask the server if the count matches. It answers which lines are off, never the expected numbers. 2 wrong checks max. */\nasync function runCheck(fromSend){\n const c=curCount();if((form.askP&&(c.plt==null||c.plt<0))||c.lines.some(z=>z.cs==null||z.cs<0)){foot();$(\"st_s2\").scrollIntoView({behavior:\"smooth\",block:\"start\"});return false}\n busy(true);let res;try{res=await RCV.check(cur,c)}catch(e){busy(false);toast(typeof errT===\"function\"?errT(e.message):e.message);return false}busy(false);\n if(res.match){form.checked=form.checkOk=true;form.banner=false;drawReceive();return true}\n if(form.stage===\"count\"&&!res.limit){form.off=res.off||[];form.stage=\"recount\";form.rplt=\"\";curLines.forEach(x=>{form.li[x.id].rcs=form.li[x.id].rlbs=\"\"});form.checked=form.checkOk=false;form.banner=fromSend;drawReceive();\n  setTimeout(()=>{fromSend?$(\"drawer\").scrollTop=0:$(\"st_s2\").scrollIntoView({behavior:\"smooth\",block:\"start\"});const f=$(\"f_rplt\")||document.querySelector(\"#st_s2 input[type=number]\");if(f)f.focus({preventScroll:true})},60);if(fromSend)toast(t(\"notsent\"));return false}\n if(form.stage===\"count\"&&res.limit){form.off=res.off||[]}\n form.checked=true;form.checkOk=false;form.needNote=true;form.limit=!!res.limit;form.banner=false;drawReceive();setTimeout(()=>$(\"f_note\")&&$(\"f_note\").focus(),60);return false}\nasync function submit(){\n if(form.up)return;\n if(!form.checked){const ok=await runCheck(true);if(!ok)return}\n const s=steps();if(s.some(x=>!x.ok)){foot();$(\"st_\"+s.find(x=>!x.ok).k).scrollIntoView({behavior:\"smooth\",block:\"start\"});return}\n const c=curCount(),ids=a=>a.filter(f=>f.fid).map(f=>f.fid);\n const payload={recv:form.recv,plt:c.plt,diffNote:form.diffNote,cond:form.cond,dmgText:form.dmgText,comment:form.comment,load:ids(form.loadPics),doc:ids(form.doc),dmg:ids(form.dmgPics),\n  lines:c.lines.map(z=>{const v=form.li[z.id];return{id:z.id,cs:z.cs,lbs:z.lbs,lot:came(z.id)?v.lot:\"\",exp:came(z.id)?v.exp:\"\",label:came(z.id)?ids(v.label):[]}})};\n busy(true);let res;try{res=await RCV.submit(cur,payload)}catch(e){busy(false);\n  // a lost answer: the receipt may already be saved. Reload and check before showing an error.\n  if(/Already received/.test(e.message)){clearDraft(cur);toast(t(\"done_h\"));closeD();RCV.after(cur,false);return}\n  toast(e.message===\"RECOUNT\"?t(\"notsent\"):(typeof errT===\"function\"?errT(e.message):e.message));return}busy(false);clearDraft(cur);\n const tot=c.lines.reduce((a,z)=>a+(z.cs||0),0);\n openD(`<div class=\"dh\"><button type=\"button\" class=\"close\" id=\"x\" aria-label=\"${esc(t(\"close\"))}\">×</button><h2 id=\"dtitle\" tabindex=\"-1\" class=\"mono\">${cur.cntr?esc(cur.cntr):\"PO \"+esc(cur.po)}</h2></div>\n <div class=\"db\"><div class=\"done-screen\"><div class=\"ok\" aria-hidden=\"true\">✓</div><h2>${t(\"done_h\")}</h2><p class=\"mono\" style=\"font-size:16px\">${c.plt!=null?fmt(c.plt)+\" \"+t(\"u_plt\")+\" · \":\"\"}${fmt(tot)} ${t(\"u_cs\")}${curLines.length>1?\" · \"+curLines.length+\" \"+t(\"it_n\"):\"\"}</p>\n  <p style=\"color:${res.flagged?\"#b0182f\":\"var(--muted)\"}\">${res.flagged?t(\"done_p\"):t(\"done_ok\")}</p><button type=\"button\" class=\"btn\" id=\"nextd\" style=\"min-width:220px;min-height:52px\">${RCV.nextLabel||t(\"next\")}</button></div></div>`);\n $(\"x\").onclick=()=>{closeD();RCV.after(cur,false)};$(\"nextd\").onclick=()=>{closeD();RCV.after(cur,true)};\n}\n</script>\n", "Client": "<script>\n// ---------- shared browser helpers (hub + portal) ----------\nfunction $(id){return document.getElementById(id)}\nconst esc=s=>String(s??\"\").replace(/[&<>\"']/g,c=>({\"&\":\"&amp;\",\"<\":\"&lt;\",\">\":\"&gt;\",'\"':\"&quot;\",\"'\":\"&#39;\"}[c]));\nconst numv=v=>{if(v===\"\"||v==null)return null;const n=+v;return isFinite(n)?n:null};\nlet LOCALE=\"en-CA\";\nconst fmt=n=>n==null||n===\"\"?\"—\":(+n).toLocaleString(LOCALE);\nconst fd=d=>d?new Date(String(d).slice(0,10)+\"T12:00\").toLocaleDateString(LOCALE,{weekday:\"short\",month:\"short\",day:\"numeric\"}):\"—\";\nconst fds=d=>d?new Date(String(d).slice(0,10)+\"T12:00\").toLocaleDateString(LOCALE,{month:\"short\",day:\"numeric\"}):\"—\";\nconst fdy=d=>d?new Date(String(d).slice(0,10)+\"T12:00\").toLocaleDateString(LOCALE,{year:\"numeric\",month:\"short\",day:\"numeric\"}):\"—\";\nconst addDays=(d,n)=>{const x=new Date(d+\"T12:00\");x.setDate(x.getDate()+n);const p=v=>String(v).padStart(2,\"0\");return x.getFullYear()+\"-\"+p(x.getMonth()+1)+\"-\"+p(x.getDate())};\nconst daysBetween=(a,b)=>Math.round((new Date(String(b).slice(0,10)+\"T12:00\")-new Date(String(a).slice(0,10)+\"T12:00\"))/864e5);\n\n/** If the receiving form code did not load inline, load it from the server and report what was damaged. */\nasync function ensureReceive(diag){\n if(typeof openReceive===\"function\")return true;\n try{\n  const code=await call(\"clientCode\",\"Receive\"),el=document.getElementById(\"rcvjs\"),got=el?el.textContent:\"(missing)\";\n  let i=0;while(i<code.length&&i<got.length&&code[i]===got[i])i++;\n  const info=\"inline receive code damaged at char \"+i+\" of \"+code.length+\". sent: \"+JSON.stringify(code.slice(Math.max(0,i-60),i+60))+\" got: \"+JSON.stringify(got.slice(Math.max(0,i-60),i+60))+\" ua: \"+navigator.userAgent.slice(0,80);\n  console.warn(info);try{if(diag)await diag(info)}catch(e){}\n  const s=document.createElement(\"script\");s.textContent=code;document.body.appendChild(s);\n  return typeof openReceive===\"function\";\n }catch(e){console.error(e);return false}}\n/** Call a server function: await call(\"portalLoad\", token) */\nfunction call(fn,...args){return new Promise((res,rej)=>{google.script.run.withSuccessHandler(res).withFailureHandler(e=>rej(e&&e.message?e:new Error(String(e))))[fn](...args)})}\nfunction busy(on){const b=$(\"busy\");if(b)b.classList.toggle(\"on\",!!on)}\nfunction toast(m){const el=$(\"toast\");el.textContent=m;el.classList.add(\"show\");clearTimeout(toast.t);toast.t=setTimeout(()=>el.classList.remove(\"show\"),3600)}\n\n/** Read a file for upload. Photos are shrunk to max 1600 px JPEG (saves Drive space and upload time). */\nfunction readFile(file){return new Promise((res,rej)=>{\n const fr=new FileReader();fr.onerror=()=>rej(new Error(\"Could not read \"+file.name));\n if(!/^image\\//.test(file.type)||/gif|svg/.test(file.type)){if(file.size>15e6)return rej(new Error(file.name+\" is larger than 15 MB.\"));fr.onload=()=>res({name:file.name,type:file.type||\"application/octet-stream\",data:fr.result,preview:null});fr.readAsDataURL(file);return}\n fr.onload=()=>{const img=new Image();img.onload=()=>{const m=1600,r=Math.min(1,m/Math.max(img.width,img.height));const c=document.createElement(\"canvas\");c.width=Math.round(img.width*r);c.height=Math.round(img.height*r);c.getContext(\"2d\").drawImage(img,0,0,c.width,c.height);const data=c.toDataURL(\"image/jpeg\",0.82);res({name:file.name.replace(/\\.\\w+$/,\"\")+\".jpg\",type:\"image/jpeg\",data,preview:data})};img.onerror=()=>res({name:file.name,type:file.type,data:fr.result,preview:null});img.src=fr.result};\n fr.readAsDataURL(file)})}\n/** Open a file sent by the server (packing slip PDF) in a new tab; if the browser blocks it, download it instead. */\n/** Show a file sent by the server (packing slip PDF). Phones block pop-ups opened after a wait, so the user taps a real link. */\nfunction openFile(f,label){const b=Uint8Array.from(atob(f.data),ch=>ch.charCodeAt(0)),u=URL.createObjectURL(new Blob([b],{type:f.type||\"application/pdf\"}));\n let box=$(\"filebox\");if(!box){box=document.createElement(\"div\");box.id=\"filebox\";box.setAttribute(\"role\",\"dialog\");box.style.cssText=\"position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:90;background:#16181a;color:#fff;border-radius:14px;padding:14px 16px;display:flex;gap:12px;align-items:center;box-shadow:0 8px 30px rgba(0,0,0,.3);max-width:92vw\";document.body.appendChild(box)}\n box.innerHTML=`<span style=\"font-size:14px\">📄 ${esc(f.name||\"PDF\")}</span><a href=\"${u}\" target=\"_blank\" rel=\"noopener\" download=\"${esc(f.name||\"packing-slip.pdf\")}\" style=\"background:#d71f3c;color:#fff;padding:10px 14px;border-radius:10px;text-decoration:none;font-weight:600;white-space:nowrap\">${esc(label||\"Open\")}</a><button type=\"button\" aria-label=\"Close\" style=\"background:none;border:none;color:#fff;font-size:20px;cursor:pointer\">×</button>`;\n box.querySelector(\"button\").onclick=()=>box.remove();box.querySelector(\"a\").focus()}\nfunction thumbs(arr){return arr&&arr.length?`<div class=\"thumbs\">${arr.map(f=>f.preview?`<img src=\"${f.preview}\" alt=\"${esc(f.name)}\">`:`<div class=\"file\">📄 ${esc(f.name)}</div>`).join(\"\")}</div>`:\"\"}\n\n// ---------- drawer with focus handling ----------\nconst BG=[\"main\",\"header\"];\nfunction openD(html){const d=$(\"drawer\");if(!d.classList.contains(\"open\"))openD.ret=document.activeElement;d.innerHTML=html;d.scrollTop=0;d.classList.add(\"open\");$(\"scrim\").classList.add(\"open\");BG.forEach(s=>{const el=document.querySelector(s);if(el)el.inert=true});\n const x=$(\"x\");if(x)x.onclick=closeD;setTimeout(()=>{const h=$(\"dtitle\");(h||d).focus({preventScroll:true})},30)}\nfunction closeD(){const d=$(\"drawer\");if(!d.classList.contains(\"open\"))return;d.classList.remove(\"open\");$(\"scrim\").classList.remove(\"open\");BG.forEach(s=>{const el=document.querySelector(s);if(el)el.inert=false});\n setTimeout(()=>{if(!d.classList.contains(\"open\"))d.innerHTML=\"\"},260);const r=openD.ret;if(r&&document.body.contains(r))r.focus()}\ndocument.addEventListener(\"keydown\",e=>{const d=$(\"drawer\");if(!d||!d.classList.contains(\"open\"))return;\n if(e.key===\"Escape\"){closeD();return}\n if(e.key===\"Tab\"){const f=[...d.querySelectorAll('button:not([disabled]),input:not([type=file]),select,textarea,a[href],[tabindex]:not([tabindex=\"-1\"])')].filter(el=>el.offsetParent!==null);if(!f.length)return;\n  const a=f[0],z=f[f.length-1];if(e.shiftKey&&document.activeElement===a){e.preventDefault();z.focus()}else if(!e.shiftKey&&document.activeElement===z){e.preventDefault();a.focus()}else if(!d.contains(document.activeElement)){e.preventDefault();a.focus()}}});\ndocument.addEventListener(\"DOMContentLoaded\",()=>{const s=$(\"scrim\");if(s)s.onclick=closeD});\n</script>\n", "Styles": "<link href=\"https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap\" rel=\"stylesheet\">\n<style>\n:root{--bg:#f6f6f4;--ink:#16181a;--muted:#62676c;--line:#e3e3df;--red:#d71f3c;--red-soft:#fde8eb;--green:#2a7550;--green-soft:#e4f3ea;--amber:#9a5c00;--amber-soft:#fff2db;--blue:#2d5d93;--blue-soft:#e6eef8;--purple:#6b3a91;--purple-soft:#f0e7f7;--shadow:0 1px 2px rgba(16,18,20,.04),0 4px 14px rgba(16,18,20,.05)}\n*{box-sizing:border-box}\nbody{margin:0;font-family:\"IBM Plex Sans\",system-ui,sans-serif;background:var(--bg);color:var(--ink);font-size:14px;-webkit-font-smoothing:antialiased}\n.mono{font-family:\"IBM Plex Mono\",monospace}\n:focus-visible{outline:3px solid var(--blue);outline-offset:2px}\nbutton{font:inherit;color:inherit}\n/* demo bar */\n.demo-bar{background:var(--ink);color:#e9e9e6;padding:6px 24px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;font-size:12.5px}\n.demo-bar button{font-size:12.5px;border:1px solid #555;background:transparent;color:#e9e9e6;padding:7px 14px;border-radius:20px;cursor:pointer;min-height:34px}\n.demo-bar button.on{background:#fff;color:var(--ink);border-color:#fff}\n/* header */\nheader{background:#fff;border-bottom:1px solid var(--line);padding:12px 24px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}\nheader img{height:36px}\nheader .div{width:1px;height:28px;background:var(--line)}\nheader h1{font-size:18px;margin:0;font-weight:600}\nheader h1 span{color:var(--red)}\n.hsub,.dsub{font-size:12.5px;color:var(--muted)}.dsub b{color:var(--ink)}\n.who{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:13px}\n.avatar{width:34px;height:34px;border-radius:50%;background:var(--red);color:#fff;display:grid;place-items:center;font-weight:600;font-size:12px}\n.lang{display:inline-flex;border:1px solid var(--line);border-radius:8px;overflow:hidden}\n.lang button{border:none;background:#fff;padding:7px 10px;font-size:12px;cursor:pointer;min-height:34px}\n.lang button.on{background:var(--ink);color:#fff}\n.signout{font-size:12px;border:1px solid var(--line);background:#fff;padding:8px 12px;border-radius:8px;cursor:pointer;color:var(--muted);min-height:34px}\n.signout:hover{border-color:var(--red);color:var(--red)}\nmain{max-width:1180px;margin:0 auto;padding:20px 24px 60px}\n.hello h2{margin:0 0 4px;font-size:22px;letter-spacing:-.3px}\n.hello p{margin:0 0 16px;color:var(--muted)}\n/* kpis */\n.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:12px}.kpi .l{min-height:1.3em}\n.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 14px;box-shadow:var(--shadow);text-align:left;cursor:pointer;transition:border-color .12s}\n.kpi:hover{border-color:var(--ink)}.kpi.on{border-color:var(--red);box-shadow:0 0 0 2px var(--red-soft)}\n.kpi .n{font-size:24px;font-weight:600}.kpi .l{font-size:12.5px;color:var(--muted)}\n.kpi.alert .n{color:var(--red)}\n.chip-line{font-size:14px;color:var(--muted);margin:-6px 0 14px}\n.chip-line b{color:var(--ink)}\n/* pills */\n.pill{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;white-space:nowrap}\n.pill::before{content:\"\";width:6px;height:6px;border-radius:50%;background:currentColor}\n.p-todo{background:var(--blue-soft);color:var(--blue)}.p-done{background:var(--green-soft);color:var(--green)}.p-dmg{background:var(--red-soft);color:#b0182f}\n.p-warn{background:var(--amber-soft);color:var(--amber)}.p-grey{background:#efefec;color:#55595d}.a-prod{background:var(--purple-soft);color:var(--purple)}\n.small{font-size:11.5px;color:var(--muted);margin-top:4px}\n/* buttons */\n.btn{font-weight:500;border:none;background:var(--red);color:#fff;padding:11px 18px;border-radius:10px;cursor:pointer;min-height:42px}\n.btn:hover{filter:brightness(.95)}\n.btn:disabled{background:#e4e4e0;color:#55595d;cursor:not-allowed}\n.btn.ghost{background:#fff;color:var(--ink);border:1px solid var(--line)}\n.btn.ghost:hover{border-color:var(--ink)}\n.btn.dark{background:var(--ink)}\n/* bar & filters */\n.bar{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px}\n.bar #q{flex:1;min-width:240px;font:inherit;padding:10px 12px;border:1px solid #8f938f;border-radius:10px;box-shadow:var(--shadow)}\n.bar select{font:inherit;padding:9px 10px;border:1px solid #8f938f;border-radius:10px;background:#fff;min-height:40px}\n.chips{display:flex;gap:6px;flex-wrap:wrap}\n.chips button{font-size:12.5px;border:1px solid var(--line);background:#fff;padding:8px 13px;border-radius:20px;cursor:pointer;min-height:36px}\n.chips button:hover{border-color:var(--ink)}.chips button.on{background:var(--ink);color:#fff;border-color:var(--ink)}\n.chips button small{opacity:.65;margin-left:3px}\n/* table */\n.tw{overflow:auto;border-radius:12px;box-shadow:var(--shadow);background:#fff;border:1px solid var(--line)}\ntable{width:100%;border-collapse:collapse;background:#fff}\nth{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);font-weight:500;padding:10px 12px;background:#fbfbfa;border-bottom:1px solid var(--line);white-space:nowrap}\ntd{padding:11px 12px;border-bottom:1px solid #f0f0ed;font-size:13px;vertical-align:top}\ntr[data-po]{cursor:pointer}tr[data-po]:hover td,tr[data-po]:focus-within td{background:#fdf6f7}\n.pobtn{background:none;border:none;padding:0;color:#b0182f;font-weight:600;font-family:\"IBM Plex Mono\",monospace;font-size:13.5px;cursor:pointer}\n.clamp{display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;color:var(--muted);font-size:12px;min-width:200px}\n.need{font-size:12.5px;font-weight:500;color:var(--ink)}.need::before{content:\"→ \";color:var(--muted)}.need.u{color:#b0182f}.need.u::before{content:\"● \";color:#b0182f}\ntr[data-po]:focus-within td:first-child{box-shadow:inset 3px 0 0 var(--blue)}\n.sr{position:absolute!important;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}\n/* 3pl list */\n.list{display:grid;gap:10px}\n.group{font-size:12px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin:12px 0 2px;font-weight:600}\n.po{background:#fff;border:1px solid var(--line);border-radius:14px;padding:16px 18px;box-shadow:var(--shadow);display:grid;grid-template-columns:1fr auto;gap:14px;align-items:center;cursor:pointer;width:100%;text-align:left;min-height:76px}\n.po:hover{border-color:var(--red)}\n.po .big{font-size:19px;font-weight:600}.po .m{font-size:13px;color:var(--muted);margin-top:3px}\n.po .go{background:var(--red);color:#fff;border-radius:10px;padding:12px 18px;font-weight:500;white-space:nowrap}\n.po.req{border-color:var(--amber);background:#fffaf0}\n.tabs{display:flex;gap:6px;margin-bottom:12px}\n.tabs button{border:1px solid var(--line);background:#fff;padding:10px 16px;border-radius:24px;cursor:pointer;min-height:44px}\n.tabs button.on{background:var(--ink);color:#fff;border-color:var(--ink)}\n/* drawer */\n.scrim{position:fixed;inset:0;background:rgba(16,18,20,.35);display:none;z-index:9}\n.scrim.open{display:block}\n.drawer{position:fixed;top:0;right:0;height:100%;width:min(620px,100%);background:#fff;transform:translateX(100%);transition:transform .22s;overflow:auto;z-index:10;box-shadow:-10px 0 40px rgba(0,0,0,.15);display:flex;flex-direction:column}\n.drawer{visibility:hidden;transition:transform .22s,visibility 0s .22s}.drawer.open{transform:none;visibility:visible;transition:transform .22s}\n.dh{padding:18px 22px;border-bottom:1px solid var(--line);position:sticky;top:0;background:#fff;z-index:2}\n.dh h2{margin:2px 0;font-size:22px}\n.close{float:right;width:44px;height:44px;border-radius:50%;border:1px solid var(--line);background:#fff;cursor:pointer;font-size:20px;color:var(--muted)}\n.close:hover{border-color:var(--red);color:var(--red)}\n.back{border:none;background:none;color:var(--blue);cursor:pointer;padding:0;font-size:13px;margin-bottom:6px}\n.db{padding:4px 22px 30px;flex:1}\n.summary{margin-top:10px;padding:10px 12px;border-radius:10px;background:#fdf6f7;color:#7a1b2b;font-size:13px;font-weight:500}\n.summary.ok{background:var(--green-soft);color:var(--green)}\n.summary.info{background:#f3f4f5;color:#3d4247}\n.sec{margin-top:18px}.sec h3{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin:0 0 10px;font-weight:600}\ndetails.sec{border:1px solid var(--line);border-radius:12px;padding:0 14px}\ndetails.sec>summary{cursor:pointer;padding:13px 0;font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);font-weight:600;list-style:none;display:flex;justify-content:space-between}\ndetails.sec>summary::after{content:\"＋\";font-size:14px}\ndetails.sec[open]>summary::after{content:\"－\"}\ndetails.sec[open]{padding-bottom:14px}\n.grid2{display:grid;grid-template-columns:150px 1fr;gap:8px 14px;font-size:13px}\n.grid2.box{border:1px solid var(--line);border-radius:12px;padding:14px 16px}\n.grid2 .k{color:var(--muted)}\n.card{border:1px solid var(--line);border-radius:12px;padding:14px 16px}\n.issue{border:1px solid #f3b9c2;background:#fdf6f7;border-radius:12px;padding:14px 16px;color:#7a1b2b}\n.issue.closed{border-color:#bfe0cc;background:#fafdfb;color:var(--green)}\n.issue.overdue{border-color:var(--red);box-shadow:0 0 0 2px var(--red-soft)}\nlabel{font-size:12px;color:var(--muted);display:block;margin-bottom:4px}\ntextarea,.in,select.in{width:100%;font:inherit;padding:10px 12px;border:1px solid #8f938f;border-radius:9px;background:#fff;color:var(--ink)}\ntextarea:focus,.in:focus{border-color:var(--blue)}\n.row{display:grid;gap:10px}\n.r2{grid-template-columns:1fr 1fr}.r3{grid-template-columns:1fr 1fr 1fr}\n.note{background:#f6f6f3;border-radius:8px;padding:10px 12px;font-size:12.5px;color:#43484c}\n.msg{border-left:3px solid var(--line);padding:6px 10px;margin-bottom:8px;font-size:13px}.msg .t{font-size:11.5px;color:var(--muted)}\n.confirm{margin-top:10px;padding:12px;border-radius:10px;background:#fff;border:1px solid var(--amber);color:var(--ink);font-size:13px}\n.timeline{list-style:none;padding:0;margin:0}\n.timeline li{display:flex;gap:12px;padding:6px 0;font-size:13px;position:relative}\n.timeline li:not(:last-child)::after{content:\"\";position:absolute;left:6px;top:21px;bottom:-7px;width:2px;background:var(--line)}\n.dot{width:14px;height:14px;border-radius:50%;margin-top:2px;flex:none;border:2px solid #cfcfca;background:#fff;z-index:1}\n.dot.done{background:var(--green);border-color:var(--green)}.dot.now{border-color:var(--red);box-shadow:0 0 0 4px var(--red-soft)}.dot.bad{background:var(--red);border-color:var(--red)}\n.timeline .when{color:var(--muted);margin-left:auto;font-size:12px;text-align:right}\n.qt{width:100%;border-collapse:collapse;font-size:13px}\n.qt th{background:none;padding:6px 8px}.qt td{padding:8px;border-bottom:1px solid #f0f0ed}\n/* receiving form (big touch targets) */\n.step{border:1px solid var(--line);border-radius:14px;padding:16px;margin-bottom:12px;scroll-margin-top:120px}\n.step.ok{border-color:#b7dcc6;background:#fafdfb}\n.step.warn{border-color:#efc27a;background:#fffaf0}\n.step h4{margin:0 0 6px;font-size:16px;display:flex;align-items:center;gap:10px}\n.num{width:28px;height:28px;border-radius:50%;background:var(--ink);color:#fff;font-size:14px;display:grid;place-items:center;flex:none}\n.step.ok .num{background:var(--green)}\n.step.warn .num{background:var(--amber)}\n.step p{margin:0 0 12px;font-size:14px;color:#4b5055}.step label{font-size:14px;color:var(--ink);font-weight:500}\n.drop{display:flex;align-items:center;justify-content:center;gap:8px;border:2px dashed #cfcfca;border-radius:12px;min-height:64px;padding:14px;text-align:center;font-size:15px;color:var(--ink);cursor:pointer;background:#fff;font-weight:500}\n.drop:hover,.drop:focus-within{border-color:var(--red);color:var(--red)}\n.drop input{position:absolute;opacity:0;width:1px;height:1px}\n.thumbs{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}\n.thumbs img,.thumbs .file{width:84px;height:84px;object-fit:cover;border-radius:8px;border:1px solid var(--line)}\n.thumbs .file{display:grid;place-items:center;font-size:11px;color:var(--muted);padding:4px;text-align:center;background:#fafaf8;overflow:hidden}\n.count{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}\n.count .in,.big-in{font-size:20px;min-height:56px;padding:10px 12px}\n.unit{position:relative}.unit .in{padding-right:40px}.unit i{position:absolute;right:12px;bottom:17px;font-style:normal;color:var(--muted);font-size:14px}\n.tick{display:flex;gap:12px;align-items:center;min-height:52px;padding:8px 12px;border:1px solid var(--line);border-radius:10px;background:#fff;cursor:pointer;font-size:14px;color:var(--ink);margin:10px 0 0}\n.tick input{width:24px;height:24px;accent-color:var(--green);flex:none}\n.cond{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:8px}\n.cond button{min-height:56px;border:1px solid var(--line);background:#fff;border-radius:12px;cursor:pointer;font-size:14px;padding:8px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px}\n.cond button span{font-size:20px}\n.cond button.on{background:var(--ink);color:#fff;border-color:var(--ink)}\n.cond button.on.badc{background:#b0182f;border-color:#b0182f}\n.foot{position:sticky;bottom:0;background:#fff;border-top:1px solid var(--line);padding:12px 22px;z-index:2;box-shadow:0 -6px 20px rgba(0,0,0,.05)}\n.prog{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--muted);margin-bottom:8px;flex-wrap:wrap}\n.prog i{width:12px;height:12px;border-radius:50%;background:#dcdcd7;display:inline-block}\n.prog i.on{background:var(--green)}\n.prog .miss{color:#b0182f;font-weight:600;background:#fdf0f2;border:1px solid #f3b9c2;border-radius:10px;padding:10px 12px;cursor:pointer;font-size:14px;min-height:44px;width:100%;text-align:left;margin-top:6px}\n.foot .btn{width:100%;min-height:56px;font-size:16px}\n.done-screen{text-align:center;padding:40px 10px}\n.done-screen .ok{width:76px;height:76px;border-radius:50%;background:var(--green);color:#fff;font-size:40px;display:grid;place-items:center;margin:0 auto 16px}\n.done-screen h2{margin:0 0 6px}\n\n/* receiving clock colors (like the 3PL master) */\n.sla{display:inline-flex;align-items:center;gap:4px;font-size:12px;font-weight:500;padding:3px 9px;border-radius:12px;white-space:nowrap}\n.sla-up{background:#f1f1ee;color:#55595d}.sla-tom{background:#eef3fa;color:#2d5d93}.sla-due{background:#fff6cc;color:#6b4f00}.sla-nudge{background:#ffe2c2;color:#8a4200}.sla-esc{background:#fbd5da;color:#a3132b}.sla-paused{background:#efe7f7;color:#5b3a91}\ntr.lvl-due td:first-child{box-shadow:inset 4px 0 0 #c99a00}tr.lvl-nudge td:first-child{box-shadow:inset 4px 0 0 #b86a00}tr.lvl-esc td:first-child{box-shadow:inset 4px 0 0 #d71f3c}tr.lvl-paused td:first-child{box-shadow:inset 4px 0 0 #7a4fb5}\ntr.lvl-esc td{background:#fef6f7}\n.legend{font-size:12px;color:var(--muted);margin:0 0 10px}\n.legend summary{cursor:pointer;display:inline}\n.legend div{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}\n.pocard{display:grid;gap:0}\n.pocard.lvl-due .po{border-left:6px solid #c99a00}.pocard.lvl-nudge .po{border-left:6px solid #b86a00;background:#fffaf3}.pocard.lvl-esc .po{border-left:6px solid #d71f3c;background:#fff6f7;animation:pulsecard 1.6s ease-in-out infinite}.pocard.lvl-paused .po{border-left:6px solid #7a4fb5}\n@keyframes pulsecard{50%{box-shadow:0 0 0 4px rgba(215,31,60,.18)}}\n@media (prefers-reduced-motion:reduce){.pocard.lvl-esc .po{animation:none}}\n.pocard .po{border-bottom-left-radius:0;border-bottom-right-radius:0}\n.delaybtn{border:1px solid var(--line);border-top:none;background:#fafaf8;border-radius:0 0 14px 14px;min-height:48px;cursor:pointer;font-size:15px;color:#5b3a91;font-weight:500}\n.delaybtn:hover{background:#f6f0fc}\n.slarule{font-size:14px;color:#43484c;background:#fff;border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:0 0 12px}\n.card.lvl-due{background:#fffcee;border-color:#c99a00}.card.lvl-nudge{background:#fff4e8;border-color:#b86a00}.card.lvl-esc{background:#fdeef0;border-color:#d71f3c}\n.instr{background:#eef3fa;border:1px solid #c7d7ec;border-radius:12px;padding:12px 14px;font-size:14px;color:#1f3f66}\n.banner{background:#fff4e8;border:1px solid #b86a00;color:#6a3500;border-radius:12px;padding:12px 14px;font-size:14.5px;font-weight:500}\n.welcome{background:#fff;border:1px solid var(--line);border-left:5px solid var(--red);border-radius:12px;padding:14px 16px;margin-bottom:14px;box-shadow:var(--shadow)}\n.welcome h3{margin:0 0 6px;font-size:16px}.welcome ol{margin:6px 0 0;padding-left:20px;font-size:13.5px;line-height:1.7}\n.welcome .lnk{background:none;border:none;color:#1d4e89;text-decoration:underline;cursor:pointer;padding:0;font-size:13.5px}\n.view{display:inline-flex;border:1px solid #8f938f;border-radius:10px;overflow:hidden}.view button{border:none;background:#fff;padding:9px 14px;cursor:pointer;min-height:40px}.view button.on{background:var(--ink);color:#fff}\n.week{display:grid;grid-template-columns:repeat(7,minmax(150px,1fr));gap:8px;overflow:auto;padding-bottom:6px}\n.wcol{background:#fff;border:1px solid var(--line);border-radius:12px;padding:8px;min-height:120px}\n.wcol h4{margin:2px 4px 8px;font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted)}\n.wcol.today{border-color:var(--ink)}\n.wcard{display:block;width:100%;text-align:left;border:1px solid var(--line);border-left:4px solid #cfcfca;background:#fff;border-radius:8px;padding:7px 8px;margin-bottom:6px;cursor:pointer;font-size:12px}\n.wcard b{font-family:\"IBM Plex Mono\",monospace;font-size:13px}.wcard:hover{border-color:var(--ink)}\n.wcard.lvl-due{border-left-color:#c99a00}.wcard.lvl-nudge{border-left-color:#b86a00}.wcard.lvl-esc{border-left-color:#d71f3c;background:#fef6f7}.wcard.lvl-paused{border-left-color:#7a4fb5}.wcard.done{border-left-color:var(--green);opacity:.8}\n.unpl{width:100%;min-height:52px;border:2px dashed #8f938f;background:#fff;border-radius:12px;font-size:15px;font-weight:500;cursor:pointer;margin-bottom:12px}\n.unpl:hover{border-color:var(--red);color:var(--red)}\n/* toast */\n.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:var(--ink);color:#fff;padding:12px 18px;border-radius:10px;font-size:13px;opacity:0;transition:opacity .2s;z-index:60;pointer-events:none;max-width:90vw}\n.toast.show{opacity:1}\n.ichip{display:inline-flex;font-size:11px;background:var(--amber-soft);color:var(--amber);padding:3px 9px;border-radius:12px;font-weight:500}\n.empty{padding:30px;text-align:center;color:var(--muted);background:#fff;border:1px dashed var(--line);border-radius:12px}\n.ferr{color:#b0182f;font-size:12px;margin-top:4px;min-height:0}\n/* login */\n#login{position:fixed;inset:0;z-index:50;display:grid;grid-template-columns:1.1fr 1fr;background:#fff;overflow:auto}\n#login.hide{display:none}\n.lg-side{background:var(--ink);color:#fff;padding:48px;display:flex;flex-direction:column;justify-content:space-between;position:relative;overflow:hidden}\n.lg-side::after{content:\"\";position:absolute;right:-120px;bottom:-120px;width:420px;height:420px;border-radius:50%;background:var(--red);opacity:.9}\n.lg-side h2{font-size:34px;line-height:1.15;margin:0 0 14px;letter-spacing:-.6px;max-width:440px;position:relative;z-index:1}\n.lg-side p{color:#c9cbcd;max-width:420px;line-height:1.5;position:relative;z-index:1}\n.lg-side ul{list-style:none;padding:0;margin:22px 0 0;position:relative;z-index:1}\n.lg-side li{padding:6px 0;color:#e5e6e7;font-size:14px}.lg-side li::before{content:\"✓\";color:#ff6b80;margin-right:10px;font-weight:600}\n.lg-main{display:flex;align-items:center;justify-content:center;padding:32px}\n.lg-box{width:100%;max-width:390px}\n.lg-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:28px}\n.lg-top img{height:46px}\n.lg-box h1{font-size:24px;margin:0 0 6px;letter-spacing:-.3px}\n.lg-box .s{color:var(--muted);margin:0 0 20px}\n.lg-box label{font-size:13.5px;margin:12px 0 5px;color:var(--ink);font-weight:500}.lg-demo button:focus-visible{border-color:var(--red)}.lg-mtag{display:none;color:var(--muted);margin:-12px 0 18px;font-size:13.5px}\n.lg-box .in{min-height:50px;font-size:15px}\n.lg-box .btn{width:100%;margin-top:18px;min-height:52px;font-size:15px}\n.lg-err{color:#b0182f;font-size:13px;margin-top:8px}.lg-err:empty{display:none}\n.lg-link{background:none;border:none;color:var(--blue);cursor:pointer;padding:0;font-size:13px;margin-top:12px}\n.lg-demo{margin-top:24px;padding-top:18px;border-top:1px solid var(--line)}\n.lg-demo h4{margin:0 0 4px;font-size:13px}\n.lg-demo p{font-size:12px;color:var(--muted);margin:0 0 10px}\n.lg-demo div{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}\n.lg-demo button{min-height:58px;font-size:13px;border:1px solid var(--line);background:#fff;border-radius:12px;cursor:pointer;line-height:1.3}\n.lg-demo button:hover{border-color:var(--red);color:var(--red)}\n.lg-demo button small{display:block;color:var(--muted);font-size:11px}\n@media(max-width:860px){\n #login{grid-template-columns:1fr}.lg-side{display:none}\n /* internal table becomes cards */\n .tw{border:none;background:none;box-shadow:none}\n .tw thead{display:none}\n .tw tr[data-po]{display:grid;grid-template-columns:1fr 1fr;gap:6px 12px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:10px;box-shadow:var(--shadow)}\n .tw td{border:none;padding:0}\n .tw td[data-l]::before{content:attr(data-l);display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:2px}\n .tw td.wide{grid-column:1/-1}\n .clamp{min-width:0}\n}\n@media(max-width:860px){\n .tw tr.lvl-esc{background:#fef6f7;border-left:4px solid #d71f3c}.tw tr.lvl-nudge{border-left:4px solid #b86a00}.tw tr.lvl-due{border-left:4px solid #c99a00}.tw tr.lvl-paused{border-left:4px solid #7a4fb5}\n .tw tr[class*=lvl-] td{background:none!important;box-shadow:none!important}\n .tw td.nona{display:none}\n .kpis{grid-template-columns:1fr 1fr}\n .lg-mtag{display:block}\n}\n@media(max-width:640px){\n .who .avatar{display:none}.signout{min-height:44px}.lang button{min-height:44px;padding:8px 12px}\n header{padding:10px 14px;gap:10px}header .div,.hsub{display:none}header img{height:30px}\n .who .nm{display:none}\n main{padding:14px 14px 50px}\n .demo-bar{padding:6px 12px}.demo-bar .dhint{display:none}\n .r3{grid-template-columns:1fr 1fr}\n .count{grid-template-columns:1fr 1fr 1fr;gap:8px}\n .count .in{font-size:18px;padding:10px 8px}\n .unit .in{padding-right:28px}.unit i{right:8px;font-size:12px}\n .grid2{grid-template-columns:110px 1fr}\n .toast{top:12px;bottom:auto}\n .po .big{font-size:18px}\n .po{grid-template-columns:1fr}\n .po .go{text-align:center}\n .db,.dh,.foot{padding-left:16px;padding-right:16px}\n}\n\n.loading{padding:40px;text-align:center;color:var(--muted)}\n.err{background:#fdeef0;border:1px solid #d71f3c;color:#7a1b2b;border-radius:10px;padding:12px 14px;margin:10px 0}\n.busy{position:fixed;inset:0;background:rgba(255,255,255,.55);display:none;z-index:70;place-items:center;font-weight:600}\n.busy.on{display:grid}\n.itrow{border-top:1px solid #ecece8;padding:10px 0}.itrow:first-of-type{border-top:none}.itrow.warn{background:#fff8ec;margin:0 -10px;padding:10px}\n.itname{font-size:14px;margin-bottom:6px;line-height:1.3}.itin{display:grid;grid-template-columns:1fr 1fr;gap:10px}\ntr.grp td{background:#f7f7f4;padding:6px 12px}.grpbtn{background:none;border:none;cursor:pointer;font:inherit;font-size:13px;color:var(--ink);padding:4px 0;min-height:32px}\n</style>\n", "Lang": "<script>\nconst L={\nen:{lg_tag:\"Every PO. Every document. One place.\",lg_p:\"Purchasing, Fulfillment, QC, Logistics, Receiving and our 3PL partners work from the same PO page.\",lg_l1:\"Know who each PO is for: production or sales\",lg_l2:\"See when it arrives, and when it really arrived\",lg_l3:\"Counts, photos and receiving docs on the PO\",lg_l4:\"Claims built in one click\",\n lg_h:\"Sign in to Tootsi PO Hub\",lg_s:\"Use your work email.\",lg_email:\"Email\",lg_pw:\"Password\",lg_btn:\"Sign in\",lg_forgot:\"Forgot password?\",lg_dh:\"Demo accounts\",lg_dp:\"Demo only: any password works. @tootsi.com opens the internal hub, @trans-plus opens TP, @deslauriers opens DL.\",\n lg_err_empty:\"Enter your email and password.\",lg_err_dom:\"This email is not set up. Use a demo account below.\",lg_forgot_t:\"In the real app this sends a reset link to your email.\",\n signout:\"Sign out\",signout2:\"Tap again to sign out\",close:\"Close\",portal_sub:\"Tootsi deliveries to your warehouse\",colon:\":\",\n hi:\"Hi\",team:\"team\",tap:\"Tap a delivery when the truck arrives.\",today:\"today\",toRecv:\"to receive\",recounts:\"recounts asked\",late_n:\"late\",\n tab_todo:\"To receive\",tab_done:\"Sent\",g_today:\"Today\",g_tom:\"Tomorrow\",g_later:\"Later\",g_recount:\"Tootsi asked for a recount\",\n receive:\"Receive\",sent:\"Sent\",sent_issue:\"Sent · issue\",recount_req:\"Recount\",\n empty_todo:\"No deliveries waiting. New Tootsi POs show up here automatically.\",empty_done:\"Nothing sent yet.\",\n pallets:\"pallets\",expected:\"Expected\",container:\"Container\",\n r_for:\"Receiving for Tootsi\",r_req:\"Tootsi asked for a recount. Please count the full load again.\",instr:\"Tootsi instructions\",\n s1:\"Photo of the load before unloading\",s1p:\"Open the trailer or container. Take a photo before touching anything.\",s1b:\"📷 Take photo of the load\",\n s2:\"Count what you received\",s2p:\"Count every pallet and case.\",s2pf:\"Floor-loaded: count every bag or case, then the pallets you built.\",\n date:\"Date received\",plt:\"Pallets\",pltb:\"Pallets built\",cs:\"Cases\",lbo:\"Net weight (lb)\",lbo_h:\"Weight is optional.\",\n u_plt:\"plt\",u_cs:\"cs\",u_lb:\"lb\",\n check:\"Check count\",check_ok:\"✓ Count matches. Continue with the next steps.\",\n s2r:\"Count again\",s2rp:\"Your count is different from the PO. Recount the pallets and cases, then enter the new numbers.\",s2r_saved:\"Your first count is saved.\",\n notsent:\"Not sent yet. Your count is different from the PO. Please recount (step 2).\",\n s2n:\"Still different after the recount. Write what you found. Purchasing will be alerted.\",s2n_ph:\"e.g. Recounted twice. 20 cases short on pallet 7.\",\n s3:\"Label, lot and expiry\",s3p:\"Take a clear, close photo of the label. Then type the lot number and expiry.\",s3b:\"📷 Take label photo\",lot:\"Lot number\",exp:\"Expiry date\",leg:\"The lot number and expiry are easy to read in the photo.\",\n s4:\"Receiving document\",s4p:\"Signed receiving report or BOL. PDF or photo.\",s4b:\"📄 Add receiving document\",\n s5:\"Load condition\",s5p:\"How did the load arrive?\",s5n_ph:\"What, how many, where. e.g. 3 bags torn on pallet 4, product spilled.\",s5b:\"📷 Add damage photos\",\n c_Good:\"Good\",c_Shifted:\"Load shifted\",c_Collapsed:\"Collapsed pallets\",c_Double:\"Double-stacked\",c_Torn:\"Torn or leaking bags\",c_Wet:\"Wet or dirty\",c_Other:\"Other\",\n com:\"Comments for Tootsi\",opt:\"optional\",comp:\"Delays, seal number, driver notes, questions for Purchasing.\",com_ph:\"e.g. Truck arrived 2 h late. Seal intact.\",\n steps:\"Steps\",missing:\"Missing\",send:\"Check and send to Tootsi\",\n m1:\"load photo\",m2:\"count\",m2c:\"check the count\",m2r:\"recount\",m2n:\"note on difference\",m3:\"label photo, lot and expiry\",m4:\"receiving document\",m5:\"load condition\",m5n:\"damage note and photos\",\n e_date:\"The date received can't be in the future.\",e_exp:\"Expiry must be after the date received.\",\n done_h:\"Sent to Tootsi\",done_p:\"Purchasing was alerted about the difference.\",done_ok:\"Thank you. Tootsi has your receiving information.\",next:\"Next delivery\",\n ro:\"Sent\",ro_change:\"To change something, contact Tootsi Inventory Control.\",your_com:\"Your comment\",cond:\"Condition\",first_c:\"First count\",recount_c:\"Recount\",final_c:\"Count\",note_c:\"Your note\",\n sla_up:\"Upcoming\",sla_tom:\"Tomorrow\",sla_due:\"Expected today\",sla_nudge:\"1 business day late\",sla_esc:\"2+ business days late\",sla_paused:\"Delay reported\",\n sla_rule:\"Receive and send on the expected day. After 1 business day the PO turns orange, after 2 it is flagged to Tootsi. If you can't, report a delay.\",\n delay_btn:\"⚠ Report a delay or issue\",delay_h:\"Report a delay or issue\",delay_p:\"Tell Tootsi why this PO can't be received yet. The delay is paused while Tootsi checks.\",\n d_reason:\"Reason\",d_note:\"Details\",d_note_ph:\"e.g. Container held at customs since Friday. Broker says Wednesday.\",d_new:\"New expected date\",d_pic:\"📷 Add photo (optional)\",d_send:\"Send to Tootsi\",d_sent:\"Delay sent. Tootsi was alerted.\",d_miss:\"Missing\",d_miss_r:\"reason\",d_miss_n:\"details\",\n d_cur:\"Delay reported\",d_solve:\"✓ Problem solved?\",d_solve_q:\"Is the problem solved? The PO goes back to your receiving list.\",yes:\"Yes, solved\",no:\"No\",d_closed:\"Delay closed. You can receive this PO.\",\n unpl_btn:\"＋ A PO arrived that is not on this list\",unpl_h:\"PO arrived, not on the list\",unpl_p:\"A truck or container arrived for Tootsi but it is not on your list. Tell us what is on the paperwork. Tootsi will check it right away.\",\n u_po:\"PO number on the paperwork (if any)\",u_cntr:\"Container or trailer number\",u_sup:\"Supplier on the paperwork\",u_desc:\"Product\",u_plts:\"Pallets\",u_doc:\"📄 Photo of BOL / paperwork\",u_load:\"📷 Photo of the load\",u_send:\"Send to Tootsi\",u_sent:\"Sent. Tootsi Logistics was alerted.\",u_miss:\"Add the paperwork photo and the product.\",\n r_notarr:\"Truck or container not arrived\",r_customs:\"Held at port or customs\",r_docs:\"Missing documents (BOL, packing list)\",r_seal:\"Seal broken or missing\",r_wrong:\"Wrong product or labels\",r_dmg:\"Damaged load, waiting for a decision\",r_dock:\"Dock full or no staff\",r_other:\"Other\"},\nfr:{lg_tag:\"Chaque commande. Chaque document. Un seul endroit.\",lg_p:\"Achats, Exécution, CQ, Logistique, Réception et nos partenaires 3PL travaillent sur la même page.\",lg_l1:\"Savoir pour qui est chaque commande : production ou ventes\",lg_l2:\"Voir quand elle arrive, et quand elle est vraiment arrivée\",lg_l3:\"Comptes, photos et documents de réception sur la commande\",lg_l4:\"Réclamations prêtes en un clic\",\n lg_h:\"Connexion à Tootsi PO Hub\",lg_s:\"Utilisez votre courriel de travail.\",lg_email:\"Courriel\",lg_pw:\"Mot de passe\",lg_btn:\"Se connecter\",lg_forgot:\"Mot de passe oublié ?\",lg_dh:\"Comptes de démonstration\",lg_dp:\"Démo seulement : tout mot de passe fonctionne. @tootsi.com ouvre le portail interne, @trans-plus ouvre TP, @deslauriers ouvre DL.\",\n lg_err_empty:\"Entrez votre courriel et votre mot de passe.\",lg_err_dom:\"Ce courriel n'est pas configuré. Utilisez un compte de démonstration.\",lg_forgot_t:\"Dans la vraie application, un lien est envoyé à votre courriel.\",\n signout:\"Déconnexion\",signout2:\"Touchez encore pour vous déconnecter\",close:\"Fermer\",portal_sub:\"Livraisons Tootsi à votre entrepôt\",colon:\" :\",\n hi:\"Bonjour\",team:\"équipe\",tap:\"Touchez une livraison quand le camion arrive.\",today:\"aujourd'hui\",toRecv:\"à recevoir\",recounts:\"recomptes demandés\",late_n:\"en retard\",\n tab_todo:\"À recevoir\",tab_done:\"Envoyés\",g_today:\"Aujourd'hui\",g_tom:\"Demain\",g_later:\"Plus tard\",g_recount:\"Tootsi demande un recompte\",\n receive:\"Recevoir\",sent:\"Envoyé\",sent_issue:\"Envoyé · problème\",recount_req:\"Recompte\",\n empty_todo:\"Aucune livraison en attente. Les nouvelles commandes Tootsi apparaissent ici automatiquement.\",empty_done:\"Rien d'envoyé pour l'instant.\",\n pallets:\"palettes\",expected:\"Prévu\",container:\"Conteneur\",\n r_for:\"Réception pour Tootsi\",r_req:\"Tootsi demande un recompte. Recomptez tout le chargement.\",instr:\"Consignes Tootsi\",\n s1:\"Photo du chargement avant déchargement\",s1p:\"Ouvrez la remorque ou le conteneur. Prenez une photo avant de toucher quoi que ce soit.\",s1b:\"📷 Photo du chargement\",\n s2:\"Comptez ce que vous avez reçu\",s2p:\"Comptez chaque palette et caisse.\",s2pf:\"Chargement en vrac : comptez chaque sac ou caisse, puis les palettes montées.\",\n date:\"Date de réception\",plt:\"Palettes\",pltb:\"Palettes montées\",cs:\"Caisses\",lbo:\"Poids net (lb)\",lbo_h:\"Le poids est facultatif.\",\n u_plt:\"pal.\",u_cs:\"caisses\",u_lb:\"lb\",\n check:\"Vérifier le compte\",check_ok:\"✓ Le compte correspond. Continuez avec les étapes suivantes.\",\n s2r:\"Recomptez\",s2rp:\"Votre compte ne correspond pas à la commande. Recomptez les palettes et les caisses, puis entrez les nouveaux chiffres.\",s2r_saved:\"Votre premier compte est enregistré.\",\n notsent:\"Pas encore envoyé. Votre compte ne correspond pas à la commande. Recomptez (étape 2).\",\n s2n:\"Toujours différent après le recompte. Écrivez ce que vous avez trouvé. Les Achats seront avisés.\",s2n_ph:\"ex. Recompté deux fois. 20 caisses manquantes sur la palette 7.\",\n s3:\"Étiquette, lot et expiration\",s3p:\"Prenez une photo claire et proche de l'étiquette. Entrez ensuite le lot et la date d'expiration.\",s3b:\"📷 Photo de l'étiquette\",lot:\"Numéro de lot\",exp:\"Date d'expiration\",leg:\"Le lot et la date d'expiration sont faciles à lire sur la photo.\",\n s4:\"Document de réception\",s4p:\"Rapport de réception signé ou connaissement. PDF ou photo.\",s4b:\"📄 Ajouter le document\",\n s5:\"État du chargement\",s5p:\"Dans quel état est arrivé le chargement ?\",s5n_ph:\"Quoi, combien, où. ex. 3 sacs déchirés sur la palette 4, produit renversé.\",s5b:\"📷 Photos des dommages\",\n c_Good:\"Bon\",c_Shifted:\"Chargement déplacé\",c_Collapsed:\"Palettes écrasées\",c_Double:\"Palettes empilées en double\",c_Torn:\"Sacs déchirés ou percés\",c_Wet:\"Mouillé ou sale\",c_Other:\"Autre\",\n com:\"Commentaires pour Tootsi\",opt:\"facultatif\",comp:\"Retards, numéro de scellé, notes du chauffeur, questions pour les Achats.\",com_ph:\"ex. Camion arrivé 2 h en retard. Scellé intact.\",\n steps:\"Étapes\",missing:\"Il manque\",send:\"Vérifier et envoyer à Tootsi\",\n m1:\"photo du chargement\",m2:\"compte\",m2c:\"vérifier le compte\",m2r:\"recompte\",m2n:\"note sur l'écart\",m3:\"photo, lot et expiration\",m4:\"document de réception\",m5:\"état du chargement\",m5n:\"note et photos des dommages\",\n e_date:\"La date de réception ne peut pas être dans le futur.\",e_exp:\"L'expiration doit être après la date de réception.\",\n done_h:\"Envoyé à Tootsi\",done_p:\"Les Achats ont été avisés de l'écart.\",done_ok:\"Merci. Tootsi a reçu vos informations.\",next:\"Livraison suivante\",\n ro:\"Envoyé\",ro_change:\"Pour une modification, contactez le Contrôle d'inventaire Tootsi.\",your_com:\"Votre commentaire\",cond:\"État\",first_c:\"Premier compte\",recount_c:\"Recompte\",final_c:\"Compte\",note_c:\"Votre note\",\n sla_up:\"À venir\",sla_tom:\"Demain\",sla_due:\"Prévu aujourd'hui\",sla_nudge:\"1 jour ouvrable de retard\",sla_esc:\"2 jours ouvrables ou plus de retard\",sla_paused:\"Retard signalé\",\n sla_rule:\"Recevez et envoyez le jour prévu. Après 1 jour ouvrable, la commande devient orange ; après 2, elle est signalée à Tootsi. Sinon, signalez un retard.\",\n delay_btn:\"⚠ Signaler un retard ou un problème\",delay_h:\"Signaler un retard ou un problème\",delay_p:\"Dites à Tootsi pourquoi cette commande ne peut pas encore être reçue. Le délai est suspendu pendant que Tootsi vérifie.\",\n d_reason:\"Raison\",d_note:\"Détails\",d_note_ph:\"ex. Conteneur retenu aux douanes depuis vendredi. Le courtier dit mercredi.\",d_new:\"Nouvelle date prévue\",d_pic:\"📷 Ajouter une photo (facultatif)\",d_send:\"Envoyer à Tootsi\",d_sent:\"Retard envoyé. Tootsi a été avisé.\",d_miss:\"Il manque\",d_miss_r:\"la raison\",d_miss_n:\"les détails\",\n d_cur:\"Retard signalé\",d_solve:\"✓ Problème réglé ?\",d_solve_q:\"Le problème est réglé ? La commande revient dans votre liste à recevoir.\",yes:\"Oui, réglé\",no:\"Non\",d_closed:\"Retard fermé. Vous pouvez recevoir cette commande.\",\n unpl_btn:\"＋ Une commande est arrivée mais n'est pas sur la liste\",unpl_h:\"Commande arrivée, absente de la liste\",unpl_p:\"Un camion ou un conteneur est arrivé pour Tootsi, mais il n'est pas sur votre liste. Indiquez ce qui est écrit sur les documents. Tootsi vérifie tout de suite.\",\n u_po:\"Numéro de commande sur les documents (si présent)\",u_cntr:\"Numéro de conteneur ou de remorque\",u_sup:\"Fournisseur sur les documents\",u_desc:\"Produit\",u_plts:\"Palettes\",u_doc:\"📄 Photo du connaissement / documents\",u_load:\"📷 Photo du chargement\",u_send:\"Envoyer à Tootsi\",u_sent:\"Envoyé. La Logistique Tootsi a été avisée.\",u_miss:\"Ajoutez la photo des documents et le produit.\",\n r_notarr:\"Camion ou conteneur pas arrivé\",r_customs:\"Retenu au port ou aux douanes\",r_docs:\"Documents manquants (connaissement, liste de colisage)\",r_seal:\"Scellé brisé ou absent\",r_wrong:\"Mauvais produit ou étiquettes\",r_dmg:\"Chargement endommagé, en attente de décision\",r_dock:\"Quai plein ou pas de personnel\",r_other:\"Autre\"}};\n</script>\n", "Logo": "<img src=\"data:image/png;base64,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\" alt=\"Tootsi Impex\"><div class=\"div\"></div>", "Message": "<!DOCTYPE html>\n<html><head><base target=\"_top\"><?!= include('Styles') ?></head>\n<body><main style=\"max-width:520px;margin:60px auto;text-align:center\">\n<h2>Tootsi PO Hub</h2><div class=\"err\"><?= msg ?></div>\n<p class=\"small\">Need access? Contact Tootsi Inventory Control.</p></main></body></html>\n"};
