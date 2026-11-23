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
  return indexState({ pos: readAll_('POs'), subs: readAll_('Submissions'), issues: readAll_('Issues'), notes: readAll_('Notes'), delays: readAll_('Delays'), recounts: readAll_('Recounts'), docs: readAll_('Docs'), supq: readAll_('SupplierQty') });
}
/** Recompute issues and save what changed. Also emails new urgent issues. */
function sync_(st, ctx) {
  var r = syncIssues(st, ctx), mail = [];
  r.writes.forEach(function (is) {
    var isNew = is._new, esc = is._esc, chg = is._chg; delete is._new; delete is._esc; delete is._chg;
    upsert_('Issues', 'id', is);
    if (isNew || esc || chg) mail.push(is); // email only when an issue is created or becomes urgent
  });
  notifyIssues_(mail, ctx);
  r.notes.forEach(function (n) { append_('Notes', n); });
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
function myPO_(st, id, u) { var p = st.pos.filter(function (x) { return String(x.id) === String(id); })[0]; if (!p || p.wh !== u.wh || p.closed) throw new Error('PO not found.'); return p; }

/** Step 1 of any upload: one file per call, so big containers never hit Google's size limit. Returns a file id, never a link. */
function portalUpload(token, f) { var u = tplUser_(token); return saveIncoming_(u.wh, checkFile_(f)); }

/** Blind check for a whole shipment. Answers which lines (and whether the pallet total) are off, never the expected numbers.
 *  At most 2 wrong checks per receiving round, recorded on the server. */
function portalCheckShip(token, id, f) {
  var u = tplUser_(token);
  return withLock_(function () { var ctx = ctx_(), st = loadState_(); return checkShipCall_(myPO_(st, id, u), st, ctx, f || {}, u.name + ' receiving'); });
}
function portalSubmitShip(token, id, f) {
  var u = tplUser_(token);
  return withLock_(function () { var ctx = ctx_(), st = loadState_(); return submitShip_(myPO_(st, id, u), st, ctx, f || {}, u.name + ' receiving', true, u.wh); });
}

function portalDelay(token, id, d) {
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_(), p = myPO_(st, id, u);
    if (!DELAY_OWNER[d.reason] || String(d.note || '').trim().length < 5) throw new Error('Reason and details are required.');
    var newDate = String(d.newDate || '');
    if (newDate && (!/^\d{4}-\d{2}-\d{2}$/.test(newDate) || newDate < ctx.today || newDate > addDays(ctx.today, 21))) throw new Error('The new date must be within the next 3 weeks.');
    var lines = shipLines(p, st).filter(function (x) { return x.wh === u.wh && !isDone(x, st) && !st.delay[x.id]; });
    var files = (d.files || []).slice(0, 10).map(function (f) { var s = saveFile_(p.id, checkFile_(f)); s.cat = 'Delay photo'; return s; });
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
  var u = tplUser_(token);
  return withLock_(function () {
    var ctx = ctx_(), st = loadState_();
    if (!(d.files && d.files.length) || !String(d.desc || '').trim()) throw new Error('Paperwork photo and product are required.');
    var po = String(d.po || '').replace(/[^\d-]/g, '').slice(0, 15);
    // the PO may already be in the hub without a destination: then it is simply "for this warehouse"
    var match = po ? st.pos.filter(function (x) { return x.po === po && !isDone(x, st) && (!x.wh || x.wh === u.wh); }) : [];
    if (match.length) {
      var saved = d.files.slice(0, 10).map(function (f) { var s = saveFile_(match[0].id, checkFile_(f)); return s; })
        .concat((d.loadFiles || []).slice(0, 10).map(function (f) { return saveFile_(match[0].id, checkFile_(f)); }));
      saved.forEach(function (s) { append_('Docs', { id: match[0].id, cat: 'Paperwork from ' + u.wh + ' (arrived)', name: s.name, url: s.url, by: u.name + ' receiving', at: ctx.stamp }); });
      match.forEach(function (x) {
        var m = String(x.manual || '').split(',').filter(String);
        if (!x.wh) { x.wh = u.wh; if (m.indexOf('wh') < 0) m.push('wh'); }
        x.clockFrom = ctx.today; x.manual = m.join(','); savePO_(x, ctx);
      });
      notePOs_(match, u.name + ' receiving', 'Arrived at ' + u.wh + '. The PO was in the hub, now assigned to ' + u.wh + '. Paperwork attached.', ctx, true);
      logAction_(u.name, 'unplanned-matched', match[0].id, po);
      sync_(loadState_(), ctx);
      return portalLoad(token);
    }
    var n = st.pos.filter(function (p) { return /^NEW-/.test(p.id); }).length + 1;
    var id = 'NEW-' + n;
    var files = d.files.slice(0, 10).map(function (f) { var s = saveFile_(id, checkFile_(f)); s.cat = 'Paperwork (arrived not on list)'; return s; })
      .concat((d.loadFiles || []).slice(0, 10).map(function (f) { var s = saveFile_(id, checkFile_(f)); s.cat = 'Load photo (arrived not on list)'; return s; }));
    var p = { id: id, po: po || id, item: '', wh: u.wh, desc: String(d.desc).trim().toUpperCase().slice(0, 200), supplier: String(d.sup || '').trim().slice(0, 200) || 'Unknown', eta: ctx.today, clockFrom: '', cntr: String(d.cntr || '').trim().toUpperCase().slice(0, 30),
      plt: num(d.plt), cs: '', lbs: '', load: 'Unknown', alloc: '', qc: 'Not received', source: 'unplanned', addedBy: u.name + ' receiving',
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
function palletsAsked_(p, st) { return shipFresh(p, st); }
function roundStart_(lines, st) { var t = ''; lines.forEach(function (x) { if (st.req[x.id] && String(st.req[x.id].at) > t) t = String(st.req[x.id].at); }); return t; }
function wrongChecks_(key, since) { return readAll_('Checks').filter(function (c) { return c.key === key && c.off !== '' && String(c.at) >= since; }); }

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
  var r = shipDiff_(p, st, ctx, f), key = shipKey(p), since = roundStart_(r.lines, st);
  var wrong = wrongChecks_(key, since);
  if (wrong.length >= 2) { var last = wrong[wrong.length - 1]; return { match: false, limit: true, off: String(last.off).split(',').filter(String) }; }
  var offList = r.off.slice(); if (r.plt) offList.push('pallets');
  if (offList.length) append_('Checks', { key: key, at: ctx.stamp, by: by, off: offList.join(','), countJson: { plt: num(f.plt), lines: (f.lines || []).map(function (L) { return { id: L.id, cs: num(L.cs), lbs: num(L.lbs) }; }) } });
  return { match: !offList.length, off: offList, tries: wrong.length + (offList.length ? 1 : 0) };
}

function submitShip_(p, st, ctx, f, by, ext, wh) {
  var r = shipDiff_(p, st, ctx, f), lines = r.lines, key = shipKey(p);
  if (!lines.length) throw new Error('Already received. Contact Tootsi Inventory Control to change it.');
  var byId = {}; (f.lines || []).forEach(function (L) { byId[L.id] = L; });
  if (lines.some(function (x) { return !byId[x.id]; })) throw new Error('Enter every item of this shipment (0 cases if an item did not come).');
  if (!f.recv || !/^\d{4}-\d{2}-\d{2}$/.test(f.recv) || f.recv > ctx.today) throw new Error('The date received can\'t be in the future.');
  if (r.askP && (num(f.plt) == null || num(f.plt) < 0)) throw new Error('Enter the number of pallets.');
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs), lb = num(L.lbs);
    if (cs == null || cs < 0 || (lb != null && lb < 0)) throw new Error('Enter cases (0 or more) for item ' + x.item + '.');
    if (cs > 0 && (!String(L.lot || '').trim() || !/^\d{4}-\d{2}-\d{2}$/.test(String(L.exp || '')) || L.exp <= f.recv)) throw new Error('Lot and expiry are required for item ' + x.item + '. Expiry must be after the date received.');
    if (cs > 0 && !(L.label || []).length) throw new Error('Label photo missing for item ' + x.item + '.');
  });
  if (!f.cond) throw new Error('Choose the load condition.');
  var bad = f.cond !== 'Good';
  if (bad && String(f.dmgText || '').trim().length < 10) throw new Error('Describe the damage.');
  if ((r.askP && !(f.load || []).length) || !(f.doc || []).length || (bad && !(f.dmg || []).length)) throw new Error('Photos and documents are missing.');
  var offAll = r.off.length || r.plt, wrong = wrongChecks_(key, roundStart_(lines, st));
  if (offAll && !wrong.length) throw new Error('RECOUNT'); // the count must go through the check (and recount) first
  if (offAll && String(f.diffNote || '').trim().length < 10) throw new Error('Explain the difference.');
  // first count = the first wrong check of this round, taken from the server (not from the browser)
  var first = wrong.length ? (wrong[0].countJson || {}) : null, firstLine = {};
  if (first) (first.lines || []).forEach(function (L) { firstLine[L.id] = L; });
  // files: already uploaded one by one; now move them into the PO folder
  var take = function (ids, cat, poId) { return (ids || []).slice(0, 20).map(function (fid) { var s = takeFile_(fid, wh, poId); s.cat = cat; return s; }); };
  var lead = lines[0].id;
  var shared = take(f.load, 'Load photo (before unloading)', lead).concat(take(f.doc, 'Receiving document', lead), take(f.dmg, 'Damage photo', lead));
  var grp = lines.length > 1 || shipLines(p, st).length > 1 ? r.grp : '';
  var plt = r.askP ? num(f.plt) : '';
  var comment = String(f.comment || '').trim().slice(0, 500), note = String(f.diffNote || '').trim().slice(0, 1000);
  lines.forEach(function (x) {
    var L = byId[x.id], cs = num(L.cs), lb = num(L.lbs);
    var fl = firstLine[x.id], lineOff = r.off.indexOf(x.id) >= 0;
    var hist = st.hist[x.id] || [];
    append_('Submissions', { subId: Utilities.getUuid(), id: x.id, seq: hist.length + 1, superseded: false, by: by, at: ctx.stamp, recv: f.recv, plt: plt, cs: cs, lbs: lb == null ? '' : lb,
      firstJson: first && (lineOff || r.plt) ? { plt: first.plt, cs: fl ? fl.cs : null, lbs: fl ? fl.lbs : null } : null,
      diffNote: (lineOff || r.plt) ? note : '', lot: String(L.lot || '').trim().slice(0, 60), exp: L.exp || '', cond: f.cond, damage: bad, dmgText: bad ? String(f.dmgText).trim().slice(0, 1000) : '', comment: comment,
      filesJson: shared.concat(take(L.label, 'Label photo', x.id)), ship: key, grpPlt: grp == null ? '' : grp });
    if (st.req[x.id]) { st.req[x.id].active = false; upsert_('Recounts', 'id', st.req[x.id]); }
    if (st.delay[x.id]) { var dl = st.delay[x.id]; dl.closedAt = ctx.stamp; upsert_('Delays', 'delayId', dl); }
    if (!x.qc || x.qc === 'Not received') { x.qc = 'Locked'; x.qcAt = f.recv; savePO_(x, ctx); }
  });
  if (comment) notePOs_(lines, by, comment, ctx, ext);
  logAction_(by, 'receipt', key, lines.length + ' lines' + (r.askP ? ' / ' + plt + ' plt' : '') + (bad ? ' / ' + f.cond : ''));
  sync_(loadState_(), ctx);
  return { ok: true, flagged: !!offAll || bad, n: lines.length };
}

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
function hubUpload(f) { hubUser_(); return saveIncoming_('HUB', checkFile_(f)); }
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
        // fields changed by hand are marked "manual": the Upcoming sheet import will not overwrite them
        var man = String(p.manual || '').split(',').filter(String), mark = function (k) { if (man.indexOf(k) < 0) man.push(k); }, ch = [];
        ['plt', 'cs', 'lbs'].forEach(function (k) { if (d[k] !== undefined && d[k] !== '') { var v = num(d[k]); if (v == null || v < 0) throw new Error('Quantities must be 0 or more.'); if (String(v) !== String(p[k])) { ch.push(k + ' ' + v); p[k] = v; if (k !== 'cs') mark(k); } } });
        if (d.wh && d.wh !== p.wh) { if (OWN_WH.concat(TPL_WH).indexOf(d.wh) < 0) throw new Error('Unknown warehouse.'); p.wh = d.wh; mark('wh'); ch.push('destination ' + d.wh); }
        if (d.eta && d.eta !== p.eta) { if (!/^\d{4}-\d{2}-\d{2}$/.test(d.eta)) throw new Error('Bad date.'); p.eta = d.eta; mark('eta'); ch.push('ETA ' + d.eta); }
        p.manual = man.join(','); savePO_(p, ctx); note_(id, me, 'Expected data updated' + (ch.length ? ': ' + ch.join(', ') : ''), ctx); break;
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
        upsert_('Recounts', 'id', { id: id, at: ctx.stamp, from: ctx.today, active: true }); note_(id, me, 'Asked ' + p.wh + ' for another count', ctx); notify3pl_(p, 'recount', ctx); break;
      case 'nudge':
        if (p.nudged === ctx.today) throw new Error('Already nudged today.');
        p.nudged = ctx.today; savePO_(p, ctx); note_(id, me, 'Nudge sent to ' + p.wh, ctx); notify3pl_(p, 'nudge', ctx); break;
      case 'sap': p.sap = ctx.today; savePO_(p, ctx); note_(id, me, 'Entered in SAP / WMS', ctx); break;
      case 'note': if (String(d.text || '').trim()) note_(id, me, String(d.text).trim(), ctx); break;
      case 'doc': var f = saveFile_(id, checkFile_(d.file)); append_('Docs', { id: id, cat: d.cat || 'Other', name: f.name, url: f.url, by: me, at: ctx.stamp }); break;
      case 'receive': // Tootsi receiving at BM / LY / HB, a whole shipment at once
        if (isTPL(p.wh)) throw new Error('3PL POs are received by the 3PL.');
        return submitShip_(p, st, ctx, d, 'Tootsi receiving (' + p.wh + ') · ' + me, false, 'HUB');
      case 'check': if (isTPL(p.wh)) throw new Error('3PL POs are received by the 3PL.'); return checkShipCall_(p, st, ctx, d, me);
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
