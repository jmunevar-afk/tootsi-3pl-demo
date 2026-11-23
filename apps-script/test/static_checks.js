// static_checks.js — runs on the ONE-FILE bundle (dist/PO_Hub.gs). Two checks the browser mock cannot do:
// 1) Security audit: Apps Script lets a web page call ANY top-level function whose name has no trailing "_".
//    Every such function that can reach data (Sheets, Drive, mail, triggers, cache) must check who is calling first.
// 2) Real Store.gs against a fake Sheet that parses input the way Google Sheets does.
const fs = require('fs'), vm = require('vm');
const code = fs.readFileSync(__dirname + '/../dist/PO_Hub.gs', 'utf8');
let pass = 0, fail = 0;
const ok = (n, c, info) => { if (c) { pass++; console.log('PASS ' + n); } else { fail++; console.log('FAIL ' + n + (info !== undefined ? ' · ' + JSON.stringify(info) : '')); } };

// ---------- 1. public function audit ----------
const fnRe = /^function (\w+)\s*\(([^)]*)\)\s*\{/gm, bodies = Object.create(null);
let m; const starts = [];
while ((m = fnRe.exec(code))) starts.push({ name: m[1], at: m.index });
starts.forEach((s, i) => { bodies[s.name] = code.slice(s.at, i + 1 < starts.length ? starts[i + 1].at : code.indexOf('var HTML_FILES')); });
const SENSITIVE = /\b(SpreadsheetApp|DriveApp|MailApp|ScriptApp|PropertiesService|CacheService|LockService|readAll_|append_|upsert_|upsertWhere_|deleteWhere_|saveFile_|saveIncoming_|takeFile_|config_|sheet_)\b/;
const reach = Object.create(null);
function reaches(n, seen = new Set()) {
  if (n in reach) return reach[n]; if (seen.has(n)) return false; seen.add(n);
  const b = bodies[n]; if (!b) return false;
  if (SENSITIVE.test(b.replace(/^function [^{]+\{/, ''))) return (reach[n] = true);
  const calls = [...b.matchAll(/\b(\w+)\s*\(/g)].map(x => x[1]).filter(x => x !== n && bodies[x]);
  return (reach[n] = calls.some(c => reaches(c, seen)));
}
const GUARD = /^function [^{]+\{\s*(?:\/\/[^\n]*\n\s*)*(?:var \w+ = )?(hubUser_|tplUser_|requireOwner_)\(|^function [^{]+\{\s*requireOwner_\(|^function [^{]+\{\s*try \{ (tplUser_|hubUser_)\(|^function [^{]+\{\s*\w[\w\s=,]*\b(hubUser_|tplUser_)\(/;
const EXEMPT = { doGet: 'routes after checking the Google account', include: 'returns page HTML only', portalLogin: 'is the login itself' };
const pub = starts.map(s => s.name).filter(n => !/_$/.test(n));
const exposed = pub.filter(n => reaches(n) && !EXEMPT[n] && !GUARD.test(bodies[n]));
ok('every browser-callable function that reaches data checks the caller first', exposed.length === 0, exposed);
ok('Store helpers are private', ['readAll', 'append', 'upsert', 'deleteWhere', 'upsertWhere'].every(n => !bodies[n]));
console.log('   public functions that reach data (all guarded): ' + pub.filter(n => reaches(n)).join(', '));

// ---------- 2. real Store.gs on a Sheets-like fake ----------
function parseInput(v) { // what Google Sheets stores when a script writes v with setValues / appendRow
  if (typeof v !== 'string') return v;
  if (v.startsWith("'")) return v.slice(1);             // apostrophe = keep as text
  if (v.startsWith('=')) return '#FORMULA:' + v;         // would become a live formula
  if (/^[-+]\d[\w\s]*[a-z]/i.test(v)) return '#ERROR!';  // "-2 cs" is parsed as a broken formula
  if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return new Date(v + 'T00:00:00');
  if (/^\d{1,2}\/\d{1,2}$/.test(v)) return new Date('2026-' + v.replace('/', '-').padStart(5, '0') + 'T00:00:00');
  if (v === 'TRUE' || v === 'FALSE') return v === 'TRUE';
  return v;
}
function FakeSheet(cols) { this.rows = [cols.slice()]; }
FakeSheet.prototype = {
  getLastRow() { let n = this.rows.length; while (n > 1 && this.rows[n - 1].every(x => x === '')) n--; return n; },
  getRange(r, c, nr, nc) { const sh = this; nr = nr || 1; nc = nc || 1;
    return { getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = sh.rows[r - 1 + i] || []; const o = []; for (let j = 0; j < nc; j++) o.push(row[c - 1 + j] === undefined ? '' : row[c - 1 + j]); out.push(o); } return out; },
      setValues(v) { v.forEach((row, i) => { while (sh.rows.length < r + i) sh.rows.push([]); row.forEach((x, j) => { sh.rows[r - 1 + i][c - 1 + j] = parseInput(x); }); }); return this; },
      setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; } }; },
  appendRow(v) { this.rows.push(v.map(parseInput)); },
  deleteRow(n) { this.rows.splice(n - 1, 1); },
  clearRow(n) { this.rows[n - 1] = this.rows[n - 1].map(() => ''); }
};
const sheets = {};
const fmt2 = d => d.toISOString().slice(0, 10);
const ctx = vm.createContext({
  console, JSON, Math, Date, Number, String, Object, Array, isFinite, RegExp, Error,
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'fake', setProperty() {} }) },
  SpreadsheetApp: { openById: () => ({ getSheetByName: n => sheets[n] || null, getSpreadsheetTimeZone: () => 'America/Toronto' }), getActiveSpreadsheet: () => null },
  Utilities: { formatDate: (d, tz, f) => f === 'H' ? '10' : (f.length > 10 ? fmt2(d) + ' 10:00' : fmt2(d)), getUuid: () => Math.random().toString(16).slice(2) + Math.random().toString(16).slice(2) },
  Session: { getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => 'owner@tootsi.com' }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
  Logger: { log() {} }, HtmlService: {}, ScriptApp: { getProjectTriggers: () => [] }
});
vm.runInContext(code, ctx);
const T = vm.runInContext('TABLES', ctx);
Object.keys(T).forEach(n => sheets[n] = new FakeSheet(T[n]));
const run = s => vm.runInContext(s, ctx);
run(`append_('Config', {key:'DOMAIN', value:'tootsi.com'}); clearCache_();`);
run(`['A','B','C','D'].forEach(function(x){ append_('Notes', {id:'4001|1', who:'t', text:x, at:'2026-11-23 10:00', ext:false}); });`);
sheets.Notes.clearRow(3); // someone clears row "B" by hand in the Sheet
run(`clearCache_(); var n = JSON.parse(JSON.stringify(readAll_('Notes').filter(function(r){return r.text==='C';})[0])); n.text='C2'; upsertWhere_('Notes', function(r){return r.text==='C';}, n); clearCache_();`);
const texts = sheets.Notes.rows.slice(1).map(r => r[2]);
ok('a cleared row does not shift saves onto the wrong row', JSON.stringify(texts) === JSON.stringify(['A', '', 'C2', 'D']), texts);
run(`append_('Notes', {id:'4001|1', who:'3PL', text:'=IMPORTRANGE("x","Users!A:Z")', at:'2026-11-23 10:00', ext:true});
     append_('Notes', {id:'4001|1', who:'3PL', text:'-2 cs on pallet 4', at:'2026-11-23 10:00', ext:true});
     append_('Notes', {id:'4001|1', who:'3PL', text:'12/11', at:'2026-11-23 10:00', ext:true}); clearCache_();`);
const last3 = run(`readAll_('Notes').slice(-3).map(function(r){return r.text;})`);
ok('typed text stays text: no formula, no #ERROR!, no date', JSON.stringify(last3) === JSON.stringify(['=IMPORTRANGE("x","Users!A:Z")', '-2 cs on pallet 4', '12/11']), last3);
run(`append_('POs', {id:'0042|00123', po:'0042', item:'00123', plt:'21', cs:21, eta:'2026-11-25', unplannedJson:{a:"x'y", n:[1,2]}}); clearCache_();`);
const po = run(`readAll_('POs')[0]`);
ok('PO / item keep leading zeros, numbers stay numbers, dates stay text, JSON round-trips',
  po.po === '0042' && po.item === '00123' && po.plt === 21 && po.cs === 21 && po.eta === '2026-11-25' && po.unplannedJson && po.unplannedJson.a === "x'y", po);
let threw = '';
try { run(`setup()`); } catch (e) { threw = e.message; }
ok('setup() refuses an anonymous web caller', /owner/i.test(threw), threw);
try { threw = ''; run(`addThreePL('TP','x','','hacker')`); } catch (e) { threw = e.message; }
ok('addThreePL() refuses an anonymous web caller', /owner/i.test(threw), threw);
console.log(pass + '/' + (pass + fail) + ' static checks passed');
process.exit(fail ? 1 : 0);
