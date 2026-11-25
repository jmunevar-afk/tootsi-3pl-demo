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
  var c = { dec: col(cfg.UPCOMING_WH_COL || 'WH DECISION (MAJID)'), po: col('PO#'), ven: col('Vendor Name'), skid: col('#Skid'), size: col('Ship Size'), csize: col('CSize'), pltzd: col('PLTZD'), cntr: col('CNTR/TrackNum'), whs: col('WHS'), eta: col('ETA Destination'), item: col('Item No.'), desc: col('Item Description'), wt: col('Line Weight') };
  if (c.po < 0 || c.item < 0) throw new Error('Could not find the PO# / Item No. columns in the upcoming sheet.');
  var lines = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], po = normPO(r[c.po]), item = String(r[c.item]).trim();
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
  // The warehouse is often written on the first line of a PO only. Give the other lines of that PO the same warehouse.
  var poWh = {};
  lines.forEach(function (L) { if (L.wh && !poWh[L.po]) poWh[L.po] = L.wh; });
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), have = {}, added = 0, changed = 0, notesByPO = {};
    var removedPO = {};
    st.pos.forEach(function (p) { have[p.id] = p; if (p.wh && p.source === 'sheet' && !poWh[p.po]) poWh[p.po] = p.wh; if (/^REMOVED/.test(String(p.closed || ''))) removedPO[p.po] = true; });
    lines.forEach(function (L) { if (!L.wh && poWh[L.po]) L.wh = poWh[L.po]; });
    lines.forEach(function (L) {
      var p = have[L.id];
      if (!p && removedPO[L.po]) return; // removed from the plan in the hub
      if (!p) {
        L.cs = ''; L.qc = 'Not received'; L.source = 'sheet'; L.addedBy = 'Import'; L.updatedAt = ctx.stamp;
        if (L.plt == null) L.plt = ''; if (L.lbs == null) L.lbs = '';
        append_('POs', L); have[L.id] = L; added++; return;
      }
      // keep the hub in step with the sheet: ETA, destination, container, pallets, weight, load type
      // never after receiving, and never for a field someone changed by hand in the hub
      if ((p.source !== 'sheet' && p.source !== 'sap') || st.sub[p.id] || p.closed) return;
      var manual = String(p.manual || '').split(',').filter(String), diff = [], drop = [];
      ['eta', 'wh', 'cntr', 'plt', 'lbs', 'load'].forEach(function (k) {
        var nv = L[k] == null ? '' : L[k], ov = p[k] == null ? '' : p[k];
        if (manual.indexOf(k) >= 0 || nv === '') return; // changed by hand in the hub
        // the SAP report replaced the sheet value: keep SAP's value while the sheet still shows the old one; once the sheet changes, the sheet wins
        var mk = manual.filter(function (m) { return m.indexOf(k + '@sheet=') === 0; })[0];
        if (mk) { if (String(nv).replace(/,/g, ' ') === mk.slice((k + '@sheet=').length)) return; drop.push(mk); }
        if (String(nv) === String(ov)) return;
        diff.push(k + ' ' + (ov === '' ? '(blank)' : ov) + ' → ' + nv); p[k] = nv;
      });
      if (drop.length) { p.manual = manual.filter(function (m) { return drop.indexOf(m) < 0; }).join(','); if (!diff.length) savePO_(p, ctx); }
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
  try { out.sap = importSapReport_(); } catch (err) { out.sap = 'error: ' + err.message; logAction_('Sync', 'error', 'sap', err.message); }
  try { out.slips = importSlips_(); } catch (err) { out.slips = 'error: ' + err.message; logAction_('Sync', 'error', 'slips', err.message); }
  try { out.master = syncMaster_(); } catch (err) { out.master = 'error: ' + err.message; logAction_('Sync', 'error', 'master', err.message); }
  // "last synced" for the hub header, and when the Upcoming sheet itself last changed (how fresh the ETAs are)
  var info = { LAST_SYNC: nowStamp_(), LAST_SYNC_OK: /error/.test(String(out.upcoming)) ? 'no' : 'yes' };
  try { var id = config_().UPCOMING_SHEET_ID; if (id) info.UPCOMING_AT = Utilities.formatDate(DriveApp.getFileById(id).getLastUpdated(), 'America/Toronto', 'yyyy-MM-dd HH:mm'); } catch (e) {}
  PropertiesService.getScriptProperties().setProperties(info);
  return out;
}

/**
 * SAP "Upcoming Deliveries to Tootsi" report (SAP server email, 10 AM and 3 PM, in the hub owner's inbox).
 * It has what the FTL FCL sheet lacks: the warehouse on every line, the cases, and LTL / LCL deliveries.
 * It looks only ~2 days ahead, so it confirms and completes the sheet; it never removes anything.
 * Fields it sets are marked "<field>@sap" so the sheet import does not put the old value back.
 * A field someone changed by hand in the hub is never touched.
 */
function importSapReport_() {
  var cfg = config_(); if (String(cfg.SAP_REPORT || 'yes').toLowerCase() === 'no') return 'off';
  var q = String(cfg.SAP_REPORT_QUERY || 'from:sap@tootsi.com subject:"Upcoming Deliveries to Tootsi" newer_than:3d');
  var threads = GmailApp.search(q, 0, 10), msgs = [];
  var sender = String(cfg.SAP_REPORT_FROM || 'sap@tootsi.com').toLowerCase();
  threads.forEach(function (t) { t.getMessages().forEach(function (m) {
    var from = String(m.getFrom() || '').toLowerCase(), addr = (from.match(/<([^>]+)>/) || [, from])[1].trim();
    if (addr === sender && /Upcoming Deliveries to Tootsi/i.test(m.getSubject())) msgs.push(m); // only the SAP server itself, not a forward or a look-alike
  }); });
  if (!msgs.length) return 'no report';
  msgs.sort(function (a, b) { return b.getDate() - a.getDate(); });
  var msg = msgs[0], props = PropertiesService.getScriptProperties();
  var when = Utilities.formatDate(msg.getDate(), 'America/Toronto', 'yyyy-MM-dd HH:mm'), label = (String(msg.getSubject()).match(/\(([^)]+)\)/) || [, ''])[1];
  var body = msg.getBody(), rows = parseSapReport_(body);
  if (!rows.length && !/PO#/i.test(body)) throw new Error('SAP report of ' + when + ': no table found.');
  var res = applySapRows_(rows, when, label);
  props.setProperties({ SAP_REPORT_AT: when, SAP_REPORT_LABEL: label, SAP_REPORT_LINES: String(rows.length) });
  return res;
}
/** The report is an HTML table: header row, then one row per PO line. */
function parseSapReport_(html) {
  var cell = function (s) { return String(s).replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;|&apos;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#(\d+);/g, function (m, c) { return String.fromCharCode(+c); }).replace(/\s+/g, ' ').trim(); };
  var trs = String(html).match(/<tr[\s\S]*?<\/tr>/gi) || [], head = null, out = [];
  trs.forEach(function (tr) {
    var tds = (tr.match(/<td[\s\S]*?<\/td>/gi) || []).map(cell);
    if (!head) { if (tds.some(function (x) { return /^PO#$/i.test(x); })) head = tds.map(function (x) { return x.toLowerCase(); }); return; }
    if (tds.length < head.length - 1) return;
    var o = {}; head.forEach(function (h, i) { o[h] = tds[i] == null ? '' : tds[i]; }); out.push(o);
  });
  var dash = function (v) { v = String(v || '').trim(); return v === '-' ? '' : v; };
  var n = function (v) { v = dash(v).replace(/,/g, ''); return v === '' || isNaN(Number(v)) ? null : Number(v); };
  return out.map(function (o) {
    var d = dash(o['planned delivery']).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\b/);
    var size = dash(o['ship size']), csize = dash(o['csize']), pal = dash(o['pltzd']) === '1';
    var load = size ? (size === 'FCL' ? 'FCL ' + csize : size) + ' · ' + (size === 'FCL' && !pal ? 'Floor loaded' : 'Palletized') : '';
    return { po: normPO(o['po#']), item: dash(o['item code']), supplier: dash(o['name']), desc: dash(o['item description']),
      cs: n(o['item quantity']), lbs: n(o['item total weight(lb)']), plt: n(o['#skid']), load: load.replace(/\s+/g, ' ').trim(),
      cntr: dash(o['cntr/tracknum']).toUpperCase(), wh: dash(o['whs']).toUpperCase(), eta: d ? d[3] + '-' + ('0' + d[2]).slice(-2) + '-' + ('0' + d[1]).slice(-2) : '' };
  }).filter(function (r) { return /^\d{3,10}(-\d+)?$/.test(r.po) && r.item; });
}
function applySapRows_(rows, when, label) {
  var WH = TPL_WH.concat(OWN_WH);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), have = {}, added = 0, changed = 0, notesByPO = {}, removedPO = {}, capped = false;
    st.pos.forEach(function (p) { have[p.id] = p; if (/^REMOVED/.test(String(p.closed || ''))) removedPO[p.po] = true; });
    rows.forEach(function (r) {
      var id = r.po + '|' + r.item, p = have[id], wh = WH.indexOf(r.wh) >= 0 ? r.wh : '';
      if (!p && removedPO[r.po]) return; // someone removed this PO from the plan: do not bring it back
      if (!p && added >= 60) { capped = true; return; } // circuit breaker: never more than 60 new lines from one report
      if (!p) {
        var np = { id: id, po: r.po, item: r.item, wh: wh, desc: r.desc, supplier: r.supplier, eta: r.eta, cntr: r.cntr, plt: r.plt == null ? '' : r.plt,
          cs: r.cs == null ? '' : r.cs, lbs: r.lbs ? r.lbs : '', load: r.load, qc: 'Not received', source: 'sap', addedBy: 'SAP report' };
        append_('POs', np); have[id] = np; added++;
        (notesByPO[r.po] = notesByPO[r.po] || { id: id, d: {} }).d['added from the SAP report (' + (r.load || 'delivery') + ', ' + (wh || 'no WHS') + ')'] = true;
        return;
      }
      if (st.sub[p.id] || p.closed) return; // received, removed or closed: the report never reopens a line
      var man = String(p.manual || '').split(',').filter(function (m) { return m && !/@sap$/.test(m); }), diff = []; // v19 markers are replaced by <field>@sheet=
      var set = function (k, nv) {
        if (nv == null || nv === '' || man.indexOf(k) >= 0) return; // empty in SAP, or changed by hand in the hub
        if (String(nv) === String(p[k] == null ? '' : p[k])) return;
        diff.push('item ' + p.item + ' ' + k + ' ' + (p[k] === '' || p[k] == null ? '(blank)' : p[k]) + ' → ' + nv);
        if (!man.some(function (m) { return m.indexOf(k + '@sheet=') === 0; })) man.push(k + '@sheet=' + String(p[k] == null ? '' : p[k]).replace(/,/g, ' '));
        p[k] = nv;
      };
      set('wh', wh); set('cs', r.cs); set('eta', r.eta); set('cntr', r.cntr);
      if (r.lbs) set('lbs', r.lbs); if (r.plt != null) set('plt', r.plt);
      if (diff.length || man.join(',') !== String(p.manual || '')) { p.manual = man.join(','); savePO_(p, ctx); }
      if (diff.length) { (notesByPO[p.po] = notesByPO[p.po] || { id: p.id, d: {} }); diff.forEach(function (t) { notesByPO[p.po].d[t] = true; }); changed++; }
    });
    Object.keys(notesByPO).forEach(function (po) { note_(notesByPO[po].id, 'SAP report', 'SAP ' + (label || 'report') + ' ' + when + ': ' + Object.keys(notesByPO[po].d).join(', '), ctx); });
    if (added || changed) { logAction_('Import', 'importSapReport', when, added + ' new PO lines, ' + changed + ' updated'); sync_(loadState_(), ctx); }
    if (capped) logAction_('Import', 'error', 'sap', 'More than 60 new lines in one SAP report: the rest were not added. Check the report.');
    return { lines: rows.length, added: added, updated: changed, report: when, capped: capped };
  });
}

/** Packing slips: Arash's extractor keeps a "PO Status" tab (PO, FOUND/MISSING, file link, action). Copy it into the Slips tab. */
function importSlips_() {
  var cfg = config_(); if (!cfg.SLIP_LOG_ID) return 0;
  var sh = SpreadsheetApp.openById(cfg.SLIP_LOG_ID).getSheetByName(cfg.SLIP_TAB || 'PO Status'); if (!sh) throw new Error('Packing slip tab not found.');
  var v = sh.getDataRange().getDisplayValues(); if (v.length < 2) return 0;
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase(); }), c = function (n) { return h.indexOf(n); };
  var ci = { po: c('po'), st: c('packing slip'), file: c('packing slip file'), src: c('source / why it is missing'), act: c('action needed') };
  if (ci.po < 0 || ci.st < 0) throw new Error('Packing slip log: columns PO / Packing slip not found.');
  var rows = v.slice(1).filter(function (r) { return /^\d{4,6}(-\d+)?$/.test(normPO(r[ci.po])); }).map(function (r) {
    var url = ci.file >= 0 ? String(r[ci.file]).trim() : '', m = url.match(/\/d\/([\w-]{20,})/);
    return { po: normPO(r[ci.po]), status: String(r[ci.st]).trim().toUpperCase(), url: url, fileId: m ? m[1] : '', source: ci.src >= 0 ? String(r[ci.src]).slice(0, 300) : '', action: ci.act >= 0 ? String(r[ci.act]).slice(0, 300) : '', checked: today_() };
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
  var C = { po: col('PO Number'), wh: col('Warehouse'), item: col('Product'), desc: col('Description'), recv: col('Receiving Date'), exp: col('Expiry Date'), units: col('Units shipped to 3PL'), docs: col('Receiving Docs'), status: col('Report Alert Status'), plt: col('Pallets shipped to 3PL'), ext: col('Extended Product Code'), stock: col('Stock Status') };
  if (!C.po || !C.wh) throw new Error('Master: PO Number / Warehouse columns not found.');
  var last = sh.getLastRow(); if (last <= hr) return 0;
  var n = last - hr;
  // one read per column, everything else is compared in memory (no slow cell-by-cell reads)
  var disp = function (c) { return c ? sh.getRange(hr + 1, c, n, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : []; };
  var raw = function (c) { return c ? sh.getRange(hr + 1, c, n, 1).getValues().map(function (r) { return r[0]; }) : []; };
  var M = { po: disp(C.po), wh: disp(C.wh), item: disp(C.item), desc: disp(C.desc), recv: raw(C.recv), recvT: disp(C.recv), exp: disp(C.exp), expRaw: raw(C.exp), units: disp(C.units), docs: disp(C.docs), status: disp(C.status), plt: disp(C.plt), ext: disp(C.ext), stock: disp(C.stock) };
    var rowsFor = function (po, wh) { var out = []; for (var i = 0; i < n; i++) if (normPO(M.po[i]) === normPO(po) && M.wh[i].toUpperCase() === wh) out.push(i); return out; };
  var ctx = ctx_(), seen = 0, written = 0, previewed = 0, closedN = 0, autoClose = String(cfg.MASTER_AUTOCLOSE || 'yes').toLowerCase() !== 'no';
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
      // the master shows it received (units + a real date): the hub marks the line received too, unless someone reopened it by hand
      if (ds && autoClose && String(p.manual || '').split(',').indexOf('open=' + ds) < 0 && String(p.manual || '').split(',').indexOf('open') < 0) {
        p.closed = ds + ' · received per master sheet (row ' + (hr + 1 + hit) + ', ' + M.units[hit] + ' units)';
        if (!p.qc || p.qc === 'Not received') { p.qc = 'Locked'; p.qcAt = ds; }
        savePO_(p, ctx); note_(p.id, 'Master sheet', 'Received at ' + p.wh + ' on ' + ds + ': ' + M.units[hit] + ' units in the master sheet (row ' + (hr + 1 + hit) + ').', ctx); closedN++;
      }
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
      // a row still marked "Receiving pending" holds the PLANNED date: the real receiving date replaces it
      if (q.field === 'recv' && cur !== '' && /receiving pending/i.test(M.stock[row] || '') && !(M.recv[row] instanceof Date && Utilities.formatDate(M.recv[row], 'America/Toronto', 'yyyy-MM-dd') === String(q.value))) cur = '';
      if (cur !== '') {
        // dates are compared as dates (the master shows 23/Nov/26, the hub writes 2026-11-23)
        var rawD = q.field === 'recv' ? M.recv[row] : q.field === 'exp' ? M.expRaw[row] : null;
        var curD = rawD instanceof Date && !isNaN(rawD) ? Utilities.formatDate(rawD, 'America/Toronto', 'yyyy-MM-dd') : '';
        var same = cur === String(q.value) || (curD && curD === String(q.value)) || ((q.field === 'units' || q.field === 'plt') && num(cur.replace(/[^\d.]/g, '')) === num(q.value)) || (q.field === 'ext' && cur.replace(/\s+/g, '') === String(q.value).replace(/\s+/g, ''));
        q.status = same ? 'already the same' : 'differs: master has ' + cur.slice(0, 40); save(); return;
      }
      if (mode !== 'on') { q.status = 'preview: would write'; previewed++; save(); return; }
      // right before writing, make sure the row was not moved (someone may sort or insert rows in the master)
      var chkWh = String(sh.getRange(hr + 1 + row, C.wh).getDisplayValue()).trim().toUpperCase(), chkIt = C.item ? String(sh.getRange(hr + 1 + row, C.item).getDisplayValue()).trim() : '';
      if (normPO(sh.getRange(hr + 1 + row, C.po).getDisplayValue()) !== q.po || chkWh !== q.wh || (chkIt !== '' && chkIt !== q.item)) { q.status = 'master rows moved: will retry'; save(); return; }
      var v = String(q.value), m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      var val = (q.field === 'recv' || q.field === 'exp') && m ? new Date(+m[1], +m[2] - 1, +m[3])
        : q.field === 'units' || q.field === 'plt' || (q.field === 'item' && /^\d+$/.test(v)) ? num(v)
        : q.field === 'docs' && /^https:\/\/drive\.google\.com\//.test(v) ? '=HYPERLINK("' + v.replace(/"/g, '') + '","Docs ' + Utilities.formatDate(new Date(), 'America/Toronto', 'MMM d') + '")'
        : "'" + v; // text always as text: never a formula in Majid's sheet
      sh.getRange(hr + 1 + row, c).setValue(val);
      if (q.field === 'recv') { M.recvT[row] = v; M.recv[row] = val; } else { if (M[q.field]) M[q.field][row] = v; if (q.field === 'exp') M.expRaw[row] = val; }
      q.status = 'written'; written++; save();
    });
  });
  return { seen: seen, written: written, preview: previewed, closed: closedN };
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
