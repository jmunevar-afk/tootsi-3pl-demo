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
function tplUser_(token) {
  var t = String(token || '');
  if (t.length < 20) throw new Error('Invalid link.');
  var u = readAll('Users').filter(function (r) { return r.token === t && truthy(r.active); })[0];
  if (!u) throw new Error('This link is not active. Contact Tootsi Inventory Control.');
  return u; // {wh, name, email}
}

// ---------- state ----------
function ctx_() {
  var cfg = config_(), d = new Date();
  return { cfg: cfg, today: today_(), nowH: Number(Utilities.formatDate(d, 'America/Toronto', 'H')), stamp: nowStamp_() };
}
function loadState_() {
  clearCache_();
  return indexState({ pos: readAll('POs'), subs: readAll('Submissions'), issues: readAll('Issues'), notes: readAll('Notes'), delays: readAll('Delays'), recounts: readAll('Recounts'), docs: readAll('Docs'), supq: readAll('SupplierQty') });
}
/** Recompute issues and save what changed. Also emails new urgent issues. */
function sync_(st, ctx) {
  var r = syncIssues(st, ctx);
  r.writes.forEach(function (is) {
    var isNew = is._new, esc = is._esc; delete is._new; delete is._esc;
    upsert('Issues', 'id', is);
    if (isNew || esc) notifyIssue_(is, ctx); // email only when an issue is created or becomes urgent
  });
  r.notes.forEach(function (n) { append('Notes', n); });
  return st;
}
function withLock_(fn) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function findPO_(st, id) { var p = st.pos.filter(function (x) { return String(x.id) === String(id); })[0]; if (!p) throw new Error('PO not found.'); return p; }
function savePO_(p, ctx) { p.updatedAt = ctx.stamp; upsert('POs', 'id', p); }
function note_(id, who, text, ctx, ext) { append('Notes', { id: id, who: who, text: text, at: ctx.stamp, ext: !!ext }); }

// =====================================================================
// 3PL PORTAL
// =====================================================================
function portalLoad(token) {
  var u = tplUser_(token), ctx = ctx_();
  var st = sync_(loadState_(), ctx);
  var rows = st.pos.filter(function (p) { return p.wh === u.wh && (!p.unplannedJson || truthy(p.planOk) || st.sub[p.id]); })
    .map(function (p) { return portalRow(p, st, ctx); });
  return { wh: u.wh, name: u.name, today: ctx.today, rows: rows };
}

/** Blind check: answers only "matches" or "does not match". Never which field, never the expected numbers. */
function portalCheck(token, id, count) {
  var u = tplUser_(token), ctx = ctx_(), st = loadState_(), p = findPO_(st, id);
  if (p.wh !== u.wh) throw new Error('Not your PO.');
  return { match: diffs(p, { plt: num(count.plt), cs: num(count.cs), lbs: num(count.lbs) }, ctx).length === 0 };
}

function portalSubmit(token, id, f) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = findPO_(st, id);
    if (p.wh !== u.wh) throw new Error('Not your PO.');
    return submitReceipt_(p, st, ctx, f, u.name + ' receiving', true);
  });
}

function portalDelay(token, id, d) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = findPO_(st, id);
    if (p.wh !== u.wh) throw new Error('Not your PO.');
    if (!DELAY_OWNER[d.reason] || String(d.note || '').trim().length < 5) throw new Error('Reason and details are required.');
    var files = (d.files || []).map(function (f) { var s = saveFile_(id, f); s.cat = 'Delay photo'; return s; });
    var newDate = d.newDate && d.newDate >= ctx.today ? d.newDate : '';
    append('Delays', { delayId: Utilities.getUuid(), id: id, reason: d.reason, note: String(d.note).trim(), newDate: newDate, by: u.name + ' receiving', at: ctx.stamp, date: ctx.today, closedAt: '', filesJson: files });
    var is = st.issue[id]; if (is && is.status !== 'Open') { is.status = 'Open'; is.opened = ctx.today; upsert('Issues', 'id', is); }
    note_(id, u.name + ' receiving', 'Delay reported: ' + REASON_EN[d.reason] + '. ' + String(d.note).trim() + (newDate ? ' New date: ' + newDate : ''), ctx, true);
    logAction_(u.name, 'delay', id, d.reason);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

function portalSolveDelay(token, id) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = findPO_(st, id), dl = st.delay[id];
    if (p.wh !== u.wh) throw new Error('Not your PO.');
    if (!dl) return portalLoad(token);
    var from = dl.newDate && dl.newDate > ctx.today ? dl.newDate : ctx.today;
    dl.closedAt = ctx.stamp; upsert('Delays', 'delayId', dl);
    p.clockFrom = from; savePO_(p, ctx);
    note_(id, u.name + ' receiving', 'Delay solved (' + REASON_EN[dl.reason] + '). Receiving clock restarts ' + from + '.', ctx, true);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

function portalUnplanned(token, d) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_();
    if (!(d.files && d.files.length) || !String(d.desc || '').trim()) throw new Error('Paperwork photo and product are required.');
    var n = st.pos.filter(function (p) { return /^NEW-/.test(p.id); }).length + 1;
    var po = String(d.po || '').replace(/[^\d-]/g, '');
    var id = 'NEW-' + n;
    var files = d.files.map(function (f) { var s = saveFile_(id, f); s.cat = 'Paperwork (arrived not on list)'; return s; })
      .concat((d.loadFiles || []).map(function (f) { var s = saveFile_(id, f); s.cat = 'Load photo (arrived not on list)'; return s; }));
    var p = { id: id, po: po || id, item: '', wh: u.wh, desc: String(d.desc).trim().toUpperCase(), supplier: String(d.sup || '').trim() || 'Unknown', eta: ctx.today, clockFrom: '', cntr: String(d.cntr || '').trim().toUpperCase(),
      plt: num(d.plt), cs: '', lbs: '', load: 'Unknown', alloc: '', qc: 'Not received', source: 'unplanned', addedBy: u.name + ' receiving',
      unplannedJson: { by: u.name + ' receiving', at: ctx.stamp, files: files }, planOk: false };
    savePO_(p, ctx);
    note_(id, u.name + ' receiving', 'Arrived at ' + u.wh + ' but not on the upcoming list. Paperwork attached.', ctx, true);
    logAction_(u.name, 'unplanned', id, d.desc);
    sync_(loadState_(), ctx);
    return portalLoad(token);
  });
}

/** Shared by the 3PL portal and Tootsi receiving at BM/LY/HB. Re-checks everything on the server. */
function submitReceipt_(p, st, ctx, f, by, ext) {
  var id = p.id;
  if (st.sub[id]) throw new Error('Already received. Contact Tootsi Inventory Control to change it.');
  var c = { plt: num(f.plt), cs: num(f.cs), lbs: num(f.lbs) };
  if (c.plt == null || c.cs == null || c.plt < 0 || c.cs < 0 || (c.lbs != null && c.lbs < 0)) throw new Error('Enter pallets and cases (0 or more).');
  if (!f.recv || f.recv > ctx.today) throw new Error('The date received can\'t be in the future.');
  if (!String(f.lot || '').trim() || !f.exp || f.exp <= f.recv) throw new Error('Lot and expiry are required. Expiry must be after the date received.');
  if (!f.cond) throw new Error('Choose the load condition.');
  var bad = f.cond !== 'Good';
  if (bad && String(f.dmgText || '').trim().length < 10) throw new Error('Describe the damage.');
  var d = diffs(p, c, ctx);
  if (d.length && !f.first) throw new Error('RECOUNT'); // client must go through the recount step
  if (d.length && String(f.diffNote || '').trim().length < 10) throw new Error('Explain the difference.');
  if (!(f.loadFiles || []).length || !(f.labelFiles || []).length || !(f.docFiles || []).length || (bad && !(f.dmgFiles || []).length)) throw new Error('Photos and documents are missing.');
  var files = [];
  var put = function (arr, cat) { (arr || []).forEach(function (x) { var s = saveFile_(id, x); s.cat = cat; files.push(s); }); };
  put(f.loadFiles, 'Load photo (before unloading)'); put(f.labelFiles, 'Label photo'); put(f.docFiles, 'Receiving document'); put(f.dmgFiles, 'Damage photo');
  var hist = st.hist[id] || [];
  append('Submissions', { subId: Utilities.getUuid(), id: id, seq: hist.length + 1, superseded: false, by: by, at: ctx.stamp, recv: f.recv, plt: c.plt, cs: c.cs, lbs: c.lbs == null ? '' : c.lbs,
    firstJson: f.first || null, diffNote: d.length ? String(f.diffNote).trim() : '', lot: String(f.lot).trim(), exp: f.exp, cond: f.cond, damage: bad, dmgText: bad ? String(f.dmgText).trim() : '', comment: String(f.comment || '').trim().slice(0, 500), filesJson: files });
  if (st.req[id]) { st.req[id].active = false; upsert('Recounts', 'id', st.req[id]); }
  if (st.delay[id]) { var dl = st.delay[id]; dl.closedAt = ctx.stamp; upsert('Delays', 'delayId', dl); note_(id, by, 'Received. Delay closed.', ctx, ext); }
  if (!p.qc || p.qc === 'Not received') { p.qc = 'Locked'; p.qcAt = f.recv; savePO_(p, ctx); }
  if (String(f.comment || '').trim()) note_(id, by, String(f.comment).trim().slice(0, 500), ctx, ext);
  logAction_(by, 'receipt', id, c.plt + ' plt / ' + c.cs + ' cs' + (bad ? ' / ' + f.cond : ''));
  sync_(loadState_(), ctx);
  return { ok: true, flagged: d.length > 0 || bad };
}

// =====================================================================
// INTERNAL HUB (Tootsi domain only)
// =====================================================================
function hubLoad() {
  var me = hubUser_(), ctx = ctx_();
  var st = sync_(loadState_(), ctx);
  return { me: me, today: ctx.today, rows: st.pos.map(function (p) { return hubRow(p, st, ctx); }), cfg: { nudgeH: ctx.cfg.SLA_NUDGE_H || 9, escH: ctx.cfg.SLA_ESC_H || 18 } };
}
function hubDetailLoad(id) {
  hubUser_(); var ctx = ctx_(), st = loadState_();
  return hubDetail(findPO_(st, id), st, ctx);
}
/** One entry point for all internal edits. action = what to do, d = data. */
function hubAction(id, action, d) {
  var me = hubUser_();
  return withLock_(function () {
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
        ['plt', 'cs', 'lbs'].forEach(function (k) { if (d[k] !== undefined && d[k] !== '') { var v = num(d[k]); if (v == null || v < 0) throw new Error('Quantities must be 0 or more.'); p[k] = v; } });
        if (d.wh) p.wh = d.wh; if (d.eta) p.eta = d.eta; savePO_(p, ctx); note_(id, me, 'Expected data updated', ctx); break;
      case 'supq': // packing list / loading report quantities
        var v = { plt: num(d.plt), cs: num(d.cs), lbs: num(d.lbs) };
        if ([v.plt, v.cs, v.lbs].some(function (z) { return z == null || z < 0; })) throw new Error('Enter pallets, cases and lb (0 or more).');
        upsertSupQ_(id, d.kind === 'lr' ? 'lr' : 'pl', v, me, ctx); note_(id, me, (d.kind === 'lr' ? 'Loading report' : 'Packing list') + ' quantities: ' + v.plt + ' plt / ' + v.cs + ' cs / ' + v.lbs + ' lb', ctx); break;
      case 'issue':
        if (!is) throw new Error('No issue on this PO.');
        if (d.owner) { is.owner = d.owner; is.ownerManual = true; }
        if (d.urgent !== undefined) { is.urgent = !!d.urgent; is.urgentManual = true; }
        if (d.status) { is.status = d.status; if (d.status === 'Open') { is.opened = ctx.today; is.openManual = true; } else is.openManual = false; }
        if (d.decision !== undefined) { is.decision = d.decision; is.decisionNote = String(d.decisionNote || '').trim(); }
        is.updatedAt = ctx.stamp; upsert('Issues', 'id', is); note_(id, me, 'Issue updated: ' + JSON.stringify(d), ctx);
        if (d.owner) notifyIssue_(is, ctx); break;
      case 'recount':
        var s = st.sub[id]; if (!s || !isTPL(p.wh)) throw new Error('Nothing to recount.');
        s.superseded = true; upsert('Submissions', 'subId', s);
        upsert('Recounts', 'id', { id: id, at: ctx.stamp, from: ctx.today, active: true }); note_(id, me, 'Asked ' + p.wh + ' for another count', ctx); notify3pl_(p, 'recount', ctx); break;
      case 'nudge':
        if (p.nudged === ctx.today) throw new Error('Already nudged today.');
        p.nudged = ctx.today; savePO_(p, ctx); note_(id, me, 'Nudge sent to ' + p.wh, ctx); notify3pl_(p, 'nudge', ctx); break;
      case 'sap': p.sap = ctx.today; savePO_(p, ctx); note_(id, me, 'Entered in SAP / WMS', ctx); break;
      case 'note': if (String(d.text || '').trim()) note_(id, me, String(d.text).trim(), ctx); break;
      case 'doc': var f = saveFile_(id, d.file); append('Docs', { id: id, cat: d.cat || 'Other', name: f.name, url: f.url, by: me, at: ctx.stamp }); break;
      case 'receive': // Tootsi receiving at BM / LY / HB
        if (isTPL(p.wh)) throw new Error('3PL POs are received by the 3PL.');
        return submitReceipt_(p, st, ctx, d, 'Tootsi receiving (' + p.wh + ') · ' + me, false);
      case 'check': return { match: diffs(p, { plt: num(d.plt), cs: num(d.cs), lbs: num(d.lbs) }, ctx).length === 0 };
      case 'confirmUnplanned':
        var np = String(d.po || '').trim(); if (!/^\d{3,10}(-\d+)?$/.test(np)) throw new Error('Enter the PO number from SAP.');
        var cs = num(d.cs), lb = num(d.lbs); if (cs == null || cs < 0 || lb == null || lb < 0) throw new Error('Enter expected cases and lb.');
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
}
function upsertSupQ_(id, kind, v, me, ctx) {
  upsertWhere('SupplierQty', function (r) { return r.id === id && r.kind === kind; }, { id: id, kind: kind, plt: v.plt, cs: v.cs, lbs: v.lbs, by: me, at: ctx.stamp });
}

// ---------- email (MailApp is included with Google Workspace) ----------
function notifyIssue_(is, ctx) {
  var cfg = ctx.cfg; if (String(cfg.EMAIL_ON || 'no') !== 'yes') return;
  var to = String(cfg['TEAM_' + is.owner.replace(/\s/g, '')] || cfg.FALLBACK_EMAIL || '').trim();
  if (String(cfg.PILOT_ONLY_ME || 'yes') === 'yes') to = cfg.FALLBACK_EMAIL;
  if (!to) return;
  var cc = String(cfg.PILOT_ONLY_ME || 'yes') === 'yes' ? '' : (cfg.CC_ALWAYS || '');
  MailApp.sendEmail({ to: to, cc: cc, subject: (truthy(is.urgent) ? '[URGENT] ' : '') + 'PO Hub · ' + String(is.id).replace('|', ' item ') + ' · ' + is.type,
    body: 'Owner: ' + is.owner + '\nOpened: ' + is.opened + '\n\nOpen the PO Hub: ' + (cfg.HUB_URL || '') });
}
function notify3pl_(p, kind, ctx) {
  var cfg = ctx.cfg; if (String(cfg.EMAIL_3PL || 'no') !== 'yes') return;
  var u = readAll('Users').filter(function (r) { return r.wh === p.wh && truthy(r.active) && r.email; })[0]; if (!u) return;
  var fr = String(cfg['LANG_' + p.wh] || 'fr') === 'fr';
  var subj = kind === 'recount' ? (fr ? 'Tootsi demande un recompte · PO ' : 'Tootsi asks for a recount · PO ') : (fr ? 'Rappel de réception · PO ' : 'Receiving reminder · PO ');
  MailApp.sendEmail({ to: u.email, subject: subj + p.po, body: (fr ? 'Ouvrez votre portail Tootsi pour recevoir ou signaler un retard : ' : 'Open your Tootsi portal to receive or report a delay: ') + (cfg.PORTAL_URL || '') + '?k=' + u.token });
}
