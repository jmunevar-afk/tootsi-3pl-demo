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
