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
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var col = function (name) { return h.indexOf(name.toLowerCase()); };
  var c = { po: col('PO#'), ven: col('Vendor Name'), skid: col('#Skid'), size: col('Ship Size'), csize: col('CSize'), pltzd: col('PLTZD'), cntr: col('CNTR/TrackNum'), whs: col('WHS'), eta: col('ETA Destination'), item: col('Item No.'), desc: col('Item Description'), wt: col('Line Weight') };
  if (c.po < 0 || c.item < 0) throw new Error('Could not find the PO# / Item No. columns in the upcoming sheet.');
  var lines = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], po = String(r[c.po]).trim(), item = String(r[c.item]).trim();
    if (!po || !item) continue;
    var eta = r[c.eta] instanceof Date ? Utilities.formatDate(r[c.eta], srcTz, 'yyyy-MM-dd') : (/^\d{4}-\d{2}-\d{2}/.test(String(r[c.eta])) ? String(r[c.eta]).slice(0, 10) : '');
    var whs = String(r[c.whs] || '').trim().toUpperCase();
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
  var C = { po: col('PO Number'), wh: col('Warehouse'), item: col('Product'), desc: col('Description'), recv: col('Receiving Date'), exp: col('Expiry Date'), units: col('Units shipped to 3PL'), docs: col('Receiving Docs') };
  if (!C.po || !C.wh) throw new Error('Master: PO Number / Warehouse columns not found.');
  var last = sh.getLastRow(); if (last <= hr) return 0;
  var n = last - hr, read = function (c) { return c ? sh.getRange(hr + 1, c, n, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : []; };
  var M = { po: read(C.po), wh: read(C.wh), item: read(C.item), recv: read(C.recv) };
  var normPO = function (x) { return String(x).replace(/^PO\s*#?/i, '').trim(); };
  var rowsFor = function (po, wh) { var out = []; for (var i = 0; i < n; i++) if (normPO(M.po[i]) === po && M.wh[i].toUpperCase() === wh) out.push(i); return out; };
  var ctx = ctx_(), st = loadState_(), seen = 0, written = 0, previewed = 0;
  withLock_(function () {
    // READ: master receiving dates for open hub lines
    var have = {}; readAll_('MasterSeen').forEach(function (s) { have[s.id] = s; });
    st.pos.forEach(function (p) {
      if (!isTPL(p.wh) || isDone(p, st)) return;
      var hit = rowsFor(p.po, p.wh).filter(function (i) { return !C.item || M.item[i] === p.item || M.item[i] === ''; })[0];
      if (hit == null) return;
      var d = new Date(M.recv[hit]); if (isNaN(d)) return;
      var ds = Utilities.formatDate(d, 'America/Toronto', 'yyyy-MM-dd');
      if (ds < '2026-01-01' || ds > ctx.today) return; // the master has corrupted dates: keep plausible ones only
      if (!have[p.id] || have[p.id].recv !== ds) { upsert_('MasterSeen', 'id', { id: p.id, recv: ds, row: String(hr + 1 + hit), checked: ctx.today }); seen++; }
    });
    // WRITE: pending fills from receipts
    readAll_('MasterSync').filter(function (q) { return q.status === 'pending' || /^preview/.test(q.status) || q.status === 'no master row'; }).forEach(function (q) {
      var cands = rowsFor(q.po, q.wh), withItem = cands.filter(function (i) { return M.item[i] === q.item; }), blank = cands.filter(function (i) { return M.item[i] === ''; });
      var row = withItem.length === 1 ? withItem[0] : (!withItem.length && blank.length === 1 ? blank[0] : null);
      if (row == null) { q.status = withItem.length > 1 || blank.length > 1 ? 'several master rows: fill by hand' : 'no master row'; upsertWhere_('MasterSync', function (r) { return r.key === q.key; }, q); return; }
      var c = C[q.field]; if (!c) { q.status = 'column not in master'; upsertWhere_('MasterSync', function (r) { return r.key === q.key; }, q); return; }
      var cell = sh.getRange(hr + 1 + row, c), cur = String(cell.getDisplayValue()).trim();
      if (cur !== '') { q.status = 'already filled: ' + cur.slice(0, 40); }
      else if (mode === 'on') { cell.setValue(/^(desc|docs)$/.test(q.field) ? "'" + q.value : q.value); q.status = 'written'; written++; }
      else { q.status = 'preview: would write'; previewed++; }
      q.row = String(hr + 1 + row); upsertWhere_('MasterSync', function (r) { return r.key === q.key; }, q);
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
