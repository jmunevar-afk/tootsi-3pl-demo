/**
 * Setup.gs — run these once from the Apps Script editor (Run ▶).
 *
 *  setup()            creates every tab with headers, the Drive folder, and default settings.
 *  addThreePL('TP','Trans-Plus','receiving@…')  creates a private portal link for a 3PL. Run once per 3PL.
 *  loadSampleData()   optional: fills a few sample POs so you can try it right away.
 *  installTriggers()  turns on the daily jobs (import from the upcoming sheet, reminders, digest email).
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the PO Hub Google Sheet (Extensions > Apps Script).');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(TABLES).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var cols = TABLES[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#16181a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  var s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ROOT_FOLDER_ID')) {
    var parent = DriveApp.getFileById(ss.getId()).getParents();
    var folder = (parent.hasNext() ? parent.next() : DriveApp.getRootFolder()).createFolder('PO Hub files');
    props.setProperty('ROOT_FOLDER_ID', folder.getId());
  }
  var me = Session.getActiveUser().getEmail();
  var defaults = {
    DOMAIN: (me.split('@')[1] || 'tootsi.com'),
    SLA_NUDGE_H: 9, SLA_ESC_H: 18, LB_TOLERANCE: 0.01, QC_DAYS: 2,
    HOLIDAYS: '2026-12-25,2026-12-28,2027-01-01,2027-04-02,2027-05-24,2027-06-24,2027-07-01,2027-09-06,2027-10-11',
    EMAIL_ON: 'no', PILOT_ONLY_ME: 'yes', FALLBACK_EMAIL: me, CC_ALWAYS: '', EMAIL_3PL: 'no',
    TEAM_Purchasing: '', TEAM_Logistics: '', TEAM_QC: '', TEAM_Receiving: '', TEAM_InventoryControl: me, TEAM_Fulfillment: '',
    LANG_TP: 'fr', LANG_DL: 'fr', HUB_URL: '', PORTAL_URL: '',
    UPCOMING_SHEET_ID: '' /* paste the ID of the UPCOMING FTL FCL sheet */, UPCOMING_TAB: 'UPCOMING FTL FCL', DIGEST_TO: me
  };
  var have = config_();
  Object.keys(defaults).forEach(function (k) { if (!(k in have)) append('Config', { key: k, value: defaults[k] }); });
  SpreadsheetApp.getUi && toastSafe_('Setup done. Next: addThreePL, then Deploy (two web apps).');
}

/** Create a private portal link for a 3PL. Example: addThreePL('TP', 'Trans-Plus', 'receiving@trans-plus.com') */
function addThreePL(wh, name, email) {
  wh = wh || 'TP'; name = name || (wh === 'TP' ? 'Trans-Plus' : 'Deslauriers');
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  append('Users', { token: token, wh: wh, name: name, email: email || '', active: true });
  var url = (config_().PORTAL_URL || '<PORTAL_URL>') + '?k=' + token;
  Logger.log('Private link for ' + name + ': ' + url);
  return url;
}

/** Turn off a 3PL link (for example when a receiver leaves). */
function disableThreePL(token) { var u = readAll('Users').filter(function (r) { return r.token === token; })[0]; if (u) { u.active = false; upsert('Users', 'token', u); } }

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(7).nearMinute(40).everyDays(1).inTimezone('America/Toronto').create();
  ScriptApp.newTrigger('importUpcoming').timeBased().everyHours(2).create();
}

function toastSafe_(m) { try { SpreadsheetApp.getActiveSpreadsheet().toast(m, 'PO Hub', 8); } catch (e) {} }

/** Optional: a few sample POs to try the pilot. Safe to delete the rows afterwards. */
function loadSampleData() {
  var t = today_(), d = function (n) { return addDays(t, n); };
  [
    { id: '90101|10201', po: '90101', item: '10201', wh: 'TP', desc: 'RAW WHOLE CASHEWS (W-320)', supplier: 'Cashew exporter · Vietnam', eta: d(1), cntr: 'TCLU4826103', plt: 28, cs: 1100, lbs: 55000, load: "FCL 40' · Palletized", alloc: 'Production', qc: 'Not received', instrSap: 'FOR PRODUCTION ONLY. No double stacking.', tplInstr: 'Do not accept double-stacked pallets. Photograph any torn bag before unloading.', source: 'sample' },
    { id: '90106|10220', po: '90106', item: '10220', wh: 'TP', desc: 'CASHEW PIECES (SSP2)', supplier: 'Cashew processor · Ghana', eta: d(-4), cntr: 'UETU3091547', plt: 18, cs: 700, lbs: 35000, load: "FCL 20' · Palletized", alloc: 'Sales', so: 'SO 118441', qc: 'Not received', source: 'sample' },
    { id: '90203|14015', po: '90203', item: '14015', wh: 'DL', desc: 'SULTANA RAISINS (TURKEY)', supplier: 'Raisin packer · Türkiye', eta: d(0), cntr: 'MSBU1188204', plt: 20, cs: 1600, lbs: 44096, load: "FCL 20' · Floor loaded", alloc: 'Production', qc: 'Not received', tplInstr: 'Floor-loaded. Photograph the open container before unloading.', source: 'sample' },
    { id: '90301|10501', po: '90301', item: '10501', wh: 'BM', desc: 'FANCY JUMBO PECAN HALVES', supplier: 'Pecan sheller · Mexico', eta: d(0), cntr: '', plt: 22, cs: 1680, lbs: 41580, load: 'FTL · Palletized', alloc: 'Sales', so: 'SO 118455', qc: 'Not received', source: 'sample' }
  ].forEach(function (p) { p.updatedAt = nowStamp_(); upsert('POs', 'id', p); });
}
