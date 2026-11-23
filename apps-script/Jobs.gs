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
    var ctx = ctx_(), st = loadState_(), have = {}, added = 0, changed = 0;
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
      if (diff.length) { savePO_(p, ctx); note_(p.id, 'Upcoming sheet', 'Updated from the Upcoming sheet: ' + diff.join(', '), ctx); changed++; }
    });
    if (added || changed) logAction_('Import', 'importUpcoming', '', added + ' new PO lines, ' + changed + ' updated');
    return added;
  });
}

function dailyJob(e) {
  requireOwner_(e);
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
  var body = 'PO Hub · ' + ctx.today + '\n\nLate at the warehouse (' + late.length + '):\n' + late.map(line).join('\n') +
    '\n\nOpen issues (' + open.length + '):\n' + open.map(line).join('\n') + '\n\nArrived 2+ days, not in SAP (' + sap.length + '):\n' + sap.map(line).join('\n') + '\n\n' + (cfg.HUB_URL || '');
  MailApp.sendEmail({ to: cfg.DIGEST_TO, subject: 'PO Hub daily · ' + late.length + ' late · ' + open.length + ' issues', body: body });
}
