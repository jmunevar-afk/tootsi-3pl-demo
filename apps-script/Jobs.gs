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
function importUpcoming() {
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
  var have = {}; readAll('POs').forEach(function (p) { have[p.id] = true; });
  var stamp = nowStamp_(), n = 0;
  for (var i = 1; i < v.length; i++) {
    var r = v[i], po = String(r[c.po]).trim(), item = String(r[c.item]).trim();
    if (!po || !item) continue;
    var id = po + '|' + item; if (have[id]) continue;
    var eta = r[c.eta] instanceof Date ? Utilities.formatDate(r[c.eta], srcTz, 'yyyy-MM-dd') : String(r[c.eta]).slice(0, 10);
    var whs = String(r[c.whs] || '').trim().toUpperCase();
    var size = String(r[c.size] || '').trim(), csize = String(r[c.csize] || '').trim(), pal = String(r[c.pltzd] || '').trim() === '1';
    var load = (size === 'FCL' ? 'FCL ' + (csize && csize !== '-' ? csize : '') : size) + ' · ' + (size === 'FCL' && !pal ? 'Floor loaded' : 'Palletized');
    append('POs', { id: id, po: po, item: item, wh: (['TP', 'DL', 'BM', 'LY', 'HB'].indexOf(whs) >= 0 ? whs : ''), desc: String(r[c.desc]).trim(), supplier: String(r[c.ven]).trim(), eta: eta,
      cntr: String(r[c.cntr] || '').trim(), plt: num(r[c.skid]), cs: '', lbs: num(r[c.wt]), load: load.trim(), qc: 'Not received', source: 'sheet', addedBy: 'Import', updatedAt: stamp });
    have[id] = true; n++;
  }
  // fill the destination later when Camilla's sheet decides it (only if still empty)
  readAll('POs').forEach(function (p) {
    if (p.wh || p.source !== 'sheet') return;
    for (var j = 1; j < v.length; j++) { var r2 = v[j]; if (String(r2[c.po]).trim() + '|' + String(r2[c.item]).trim() === p.id) { var w = String(r2[c.whs] || '').trim().toUpperCase(); if (['TP', 'DL', 'BM', 'LY', 'HB'].indexOf(w) >= 0) { p.wh = w; p.updatedAt = stamp; upsert('POs', 'id', p); } break; } }
  });
  if (n) logAction_('Import', 'importUpcoming', '', n + ' new PO lines');
  return n;
}

function dailyJob() {
  try { importUpcoming(); } catch (e) { logAction_('Import', 'error', '', e.message); }
  var ctx = ctx_(), st = sync_(loadState_(), ctx), cfg = ctx.cfg;
  var late = st.pos.filter(function (p) { return isLate(p, st, ctx); });
  late.forEach(function (p) { if (isTPL(p.wh) && p.nudged !== ctx.today && sla(p, st, ctx).lvl === 'nudge') { p.nudged = ctx.today; savePO_(p, ctx); notify3pl_(p, 'nudge', ctx); } });
  var open = st.pos.filter(function (p) { return openIssue(p, st); });
  var sap = st.pos.filter(function (p) { return notInSap(p, st, ctx); });
  if (!cfg.DIGEST_TO || String(cfg.EMAIL_ON || 'no') !== 'yes') return;
  var line = function (p) { var i = st.issue[p.id]; return '• PO ' + p.po + (p.item ? ' / ' + p.item : '') + ' · ' + p.wh + ' · ' + p.desc + (i && i.status === 'Open' ? ' — ' + i.type + ' (owner ' + i.owner + ')' : ''); };
  var body = 'PO Hub · ' + ctx.today + '\n\nLate at the warehouse (' + late.length + '):\n' + late.map(line).join('\n') +
    '\n\nOpen issues (' + open.length + '):\n' + open.map(line).join('\n') + '\n\nArrived 2+ days, not in SAP (' + sap.length + '):\n' + sap.map(line).join('\n') + '\n\n' + (cfg.HUB_URL || '');
  MailApp.sendEmail({ to: cfg.DIGEST_TO, subject: 'PO Hub daily · ' + late.length + ' late · ' + open.length + ' issues', body: body });
}
