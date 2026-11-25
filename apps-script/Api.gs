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
  var lines = openLines_(p, st), key = shipKey(p), fromReport = isTPL(p.wh) && wh === 'HUB';
  if (!lines.length) throw new Error('Already received. Contact Tootsi Inventory Control to change it.');
  var byId = {}; (f.lines || []).forEach(function (L) { byId[L.id] = L; });
  if (lines.some(function (x) { return !byId[x.id]; })) throw new Error('Enter every item of this shipment (0 cases if an item did not come).');
  if (!f.recv || !/^\d{4}-\d{2}-\d{2}$/.test(f.recv) || f.recv > ctx.today || f.recv < addDays(ctx.today, -60)) throw new Error('Check the date received (not in the future, not more than 60 days ago).');
  var askP = palletsAsked_(p, st);
  if (askP && (num(f.plt) == null || num(f.plt) < 0 || num(f.plt) > 200)) throw new Error('Enter the number of pallets.');
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs), lb = num(L.lbs);
    if (cs == null || cs < 0 || cs > 1000000 || (lb != null && (lb < 0 || lb > 200000))) throw new Error('Enter cases (0 or more) for item ' + x.item + '.');
    var expOk = L.exp ? (/^\d{4}-\d{2}-\d{2}$/.test(String(L.exp)) && L.exp > f.recv && L.exp <= addDays(f.recv, 3660)) : fromReport; // expiry optional when Tootsi types in a 3PL report
    if (cs > 0 && (!String(L.lot || '').trim() || !expOk)) throw new Error('Lot and expiry are required for item ' + x.item + '. Expiry must be after the date received.');
    if (cs > 0 && !(L.label || []).length) throw new Error('Label photo missing for item ' + x.item + '.');
  });
  if (!f.cond) throw new Error('Choose the load condition.');
  var bad = f.cond !== 'Good';
  if (bad && String(f.dmgText || '').trim().length < 10) throw new Error('Describe the damage.');
  var fresh = shipFresh(p, st) && !fromReport; // the "before unloading" photo only makes sense on the first receiving
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
  queueMaster_(lines, byId, f, isTPL(p.wh) ? p.wh : wh, ctx, holder);
  sync_(loadState_(), ctx);
  var report = fromReport ? { wh: p.wh, plt: askP ? num(f.plt) : null, holder: holder, files: Object.keys(files),
    rows: lines.map(function (x) { var L = byId[x.id]; return { po: x.po, item: x.item, lot: String(L.lot || '').trim(), cw: caseWeight_(x, L), cs: num(L.cs), holder: x.id === holder }; }) } : null;
  return { ok: true, flagged: !!offAll || bad, n: lines.length, _moves: moves, _report: report };
}

/** After a 3PL receipt: queue the master-sheet fills (only empty cells will ever be filled; see syncMaster_ in Jobs). */
/** Internal pilot: Tootsi staff type in 3PL receipts from the 3PL's email (Config TPL_BY_STAFF, default yes). */
function staffEnters3pl_() { return String(config_().TPL_BY_STAFF || 'yes').toLowerCase() !== 'no'; }
/** Case weight in lb: received lb / cases, else the PO line weight / cases. 40000 lb / 1600 cases = 25. */
function caseWeight_(x, L) {
  var cs = num(L.cs), lb = num(L.lbs), base = lb > 0 ? lb : num(x.lbs);
  if (!(cs > 0) || !(base > 0)) return '';
  return String(Math.round(base / cs * 100) / 100);
}
function queueMaster_(lines, byId, f, wh, ctx, holder) {
  if (!isTPL(wh) || !config_().MASTER_SHEET_ID || String(config_().MASTER_MODE || 'preview') === 'off') return;
  var docs = ''; try { docs = poFolder_(lines[0].id).getUrl(); } catch (e) {}
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs); if (!(cs > 0)) return;
    var cw = caseWeight_(x, L), lot = String(L.lot || '').trim();
    var vals = { item: x.item, desc: x.desc, recv: f.recv, exp: L.exp, units: cs, docs: docs,
      plt: x.id === holder && num(f.plt) != null ? num(f.plt) : '',
      ext: lot && cw ? x.item + ' #' + lot + '#' + x.po + '#' + cw : '' };
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
  return { me: me, today: ctx.today, rows: st.pos.map(function (p) { return hubRow(p, st, ctx); }), cfg: { nudgeH: ctx.cfg.SLA_NUDGE_H || 9, escH: ctx.cfg.SLA_ESC_H || 18, portalUrl: String(ctx.cfg.PORTAL_URL || ''), goLive: String(ctx.cfg.GO_LIVE || '') }, sync: syncInfo_() };
}
function syncInfo_() {
  var p = PropertiesService.getScriptProperties();
  return { at: p.getProperty('LAST_SYNC') || '', ok: p.getProperty('LAST_SYNC_OK') !== 'no', upcomingAt: p.getProperty('UPCOMING_AT') || '' };
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
      case 'receive': // Tootsi receiving at BM / LY / HB, or a 3PL receipt typed in by Tootsi from the 3PL's report
        if (isTPL(p.wh) && !staffEnters3pl_()) throw new Error('3PL POs are received by the 3PL.');
        return submitShip_(p, st, ctx, d, (isTPL(p.wh) ? 'Tootsi, from ' + p.wh + ' report' : 'Tootsi receiving (' + p.wh + ')') + ' · ' + me, false, 'HUB');
      case 'check': if (isTPL(p.wh) && !staffEnters3pl_()) throw new Error('3PL POs are received by the 3PL.'); return checkShipCall_(p, st, ctx, d, me);
      case 'confirmUnplanned':
        if (!p.unplannedJson || truthy(p.planOk)) throw new Error('This is not an unplanned arrival.');
        var np = normPO(d.po); if (!/^\d{3,10}(-\d+)?$/.test(np)) throw new Error('Enter the PO number from SAP.');
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
        var po = normPO(d.po), item = String(d.item || '').trim();
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
  if (out && out._report) {
    var rpt = out._report; delete out._report;
    try { out.master = syncMaster_(); } catch (e) { out.master = 'error: ' + e.message; logAction_(me, 'error', 'master', e.message); }
    try { if (String(config_().RCV_DRAFT || 'yes').toLowerCase() !== 'no') out.draft = receiptDraft_(rpt); } catch (e) { out.draftErr = e.message; logAction_(me, 'error', 'draft', e.message); }
  }
  return out;
}

/** Gmail draft for the receiving team, like Juan's usual "PO ... – Receiving Documents from DL" email, with the photos and documents attached. */
function receiptDraft_(r) {
  var cfg = config_(), rows = r.rows.filter(function (x) { return x.cs > 0; });
  if (!rows.length) return null;
  var pos = []; rows.forEach(function (x) { if (pos.indexOf(x.po) < 0) pos.push(x.po); });
  var subject = 'PO ' + pos.join(', ') + ' – Receiving Documents from ' + r.wh;
  var e = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var th = 'border:1px solid #000;background:#ead1dc;font-weight:bold;text-align:center;padding:3px 10px;font-family:Arial;font-size:11pt';
  var td = 'border:1px solid #000;text-align:center;padding:3px 10px;font-family:Calibri,Arial;font-size:11pt';
  var head = ['3PL', 'Product', 'PO', 'Lot', 'Weight (LB)', 'No. Pallets', 'No. Units'];
  var body = rows.map(function (x) { return [r.wh, x.item, x.po, x.lot, x.cw, x.holder && r.plt != null ? r.plt : (rows.length === 1 && r.plt != null ? r.plt : ''), x.cs]; });
  var intro = 'Hi Receiving Team,<br><br>The inventory detailed below has been received from ' + e(r.wh) + '.<br>Please update the system as soon as possible with this information.<br><br>' +
    'Additionally, please attach the enclosed receiving documents and pictures to the relevant PO, and forward them to QC for review and approval.<br><br>';
  var table = '<table style="border-collapse:collapse"><tr>' + head.map(function (h) { return '<th style="' + th + '">' + h + '</th>'; }).join('') + '</tr>' +
    body.map(function (b) { return '<tr>' + b.map(function (v) { return '<td style="' + td + '">' + e(v) + '</td>'; }).join('') + '</tr>'; }).join('') + '</table>';
  var sig = String(cfg.EMAIL_SIGNATURE || 'Regards,<br><b>Juan Munevar</b><br>Spécialiste en Contrôle Inventaire / Inventory Control Specialist<br>(514) 381-9790 Ext 147');
  var html = '<div style="font-family:Arial;font-size:10pt">' + intro + table + '<br>' + sig + '</div>';
  var plain = 'Hi Receiving Team,\n\nThe inventory detailed below has been received from ' + r.wh + '.\nPlease update the system as soon as possible with this information.\n\n' +
    'Additionally, please attach the enclosed receiving documents and pictures to the relevant PO, and forward them to QC for review and approval.\n\n' +
    head.join('\t') + '\n' + body.map(function (b) { return b.join('\t'); }).join('\n') + '\n\nRegards,\nJuan Munevar';
  var att = [], size = 0;
  (r.files || []).forEach(function (fid) {
    try { var b = DriveApp.getFileById(fid).getBlob(), n = b.getBytes().length; if (size + n > 20 * 1024 * 1024) return; size += n; att.push(b); } catch (x) {}
  });
  var d = GmailApp.createDraft(String(cfg.RCV_EMAIL_TO || 'hbreception@tootsi.com,whb1@tootsi.com'), subject, plain,
    { cc: String(cfg.RCV_EMAIL_CC || 'mjavdany@tootsi.com,jcelis@tootsi.com'), htmlBody: html, attachments: att });
  return { subject: subject, files: att.length, url: 'https://mail.google.com/mail/u/0/#drafts' };
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
