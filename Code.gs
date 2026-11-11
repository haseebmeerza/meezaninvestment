/**
 * Meezan Investment Tracker v3.0.0 – Google Sheet backend (Apps Script, bound to the Sheet).
 * Sheet "Rates": A Date, B NAV (you / the app enter these), C:K formulas (script writes them), M:N parameters.
 * Sheet "Plans": investment plans added from the app.
 * Rule: a Friday rate is applied automatically to Saturday and Sunday.
 */
const RATES = 'Rates', PLANS = 'Plans', TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
const START = '2026-06-17', CAL_START = '2026-06-01', END = '2026-12-31';   // START = first seed rate; calendar rows begin 1 June 2026   // calendar rows, like the Excel (rates blank until entered)
const SEED = [170.1922, 170.1922, 170.1922, 170.1922, 170.1922, 169.1158, 168.9046, 168.9046, 168.9046, 168.9046, 168.9046, 168.9046, 168.6041, 169.7, 172.0482, 172.6587, 173.5081, 173.5081, 173.5081, 175.3126, 173.5889, 169.023, 168.8494, 169.6963, 169.6963, 169.6963, 167.1206, 161.1499, 162.6919, 165.6168, 163.0619, 163.0619, 163.0619, 163.0872, 163.063, 161.5665, 159.0563, 158.5053, 158.5053, 158.5053, 165.3995, 164.9907, 163.085, 162.0444, 162.6489, 162.6489, 162.6489, 164.9567, 163.8439, 166.501, 168.0725, 167.95, 167.95, 167.95, 167.7894, 166.0406, 166.3348, 166.2893, 166.2893, 166.2893, 166.2893, 167.0441, 164.1014, 162.9894, 163.1252, 164.696, 164.696, 164.696, 164.5455, 164.5741, 164.3279, 164.3279, 164.8071, 164.8071, 164.8071, 164.8302, 164.8722, 163.1833, 163.241, 163.6147, 163.6147, 163.6147, 161.8187, 160.8015, 160.0133, 156.6644, 158.6359, 158.6359, 158.6359, 156.0864, 157.0741, 155.6762, 156.5435, 159.6763, 159.6763, 159.6763, 158.4533, 158.9749, 160.1294, 160.13, 158.156, 158.156, 158.156, 158.068, 157.5153, 157.9976, 156.7499, 156.2384, 156.2384, 156.2384, 154.3667, 156.9964, 156.8103, 155.8041, 155.4522, 155.4522, 155.4522, 153.6078, 153.4495, 154.1391, 154.8773, 154.7994, 154.7994, 154.7994, 154.49, 153.6336, 153.96, 152.076, 150.3916, 150.3916, 150.3916, 151.1222, 151.8735, 152.7288, 154.04, 153.1358, 153.1358, 153.1358, 153.0154];
const EVENTS = 'Events', EV_COLS = ['date','title','note','type'];
const SITE = 'https://www.almeezangroup.com/fund-prices', FUND_NAME = 'Meezan Islamic Fund';
const DEFAULT_HOLIDAYS = [['2026-08-14','Independence Day'],['2026-11-09','Iqbal Day'],['2026-12-25','Quaid-e-Azam Day / Christmas']];
let HOL = {};   // holiday dates (yyyy-mm-dd) loaded per request from the Events sheet (type = Holiday)
const PLAN_COLS = ['name','date','mAmt','dAmt','nav','mUnits','dUnits','units','status'];
const VERSION = '3.0.0';
const ADMIN_ONLY = { reset: 1, deleteRate: 1 };           // actions only the admin password may run
// Private figures live ONLY here (server side) and are sent to the app after a successful login.
const CONFIG = { invest: 100000, charges: 1955, tax: 293.5, unitsTotal: 1155.5629, portfolio: '902300-1', bank: {port:'9023••-1',holder:'MARIA HASEEB MIRZA',amc:'AL MEEZAN INVESTMENT MANAGEMENT LTD',asat:'2026-09-30',printed:'2026-10-01',
 funds:[{n:'Meezan Islamic Fund (MIF)',t:'GROWTH-B',u:1155.5629,nav:157.9976,val:182576,gl:-17424,fy:-14549},{n:'Meezan Islamic Income Fund (MIIF)',t:'GROWTH-B',u:18.9207,nav:53.3381,val:1009,gl:9,fy:9}],
 tx:[['2026-07-21','PURCHASE (MIIF)',1000,0,'52.2512','18.9207'],['2026-07-21','LOAD CHARGED (MIIF)',0,10,'0.0000','–'],['2026-07-21','GOVT. TAX (MIIF)',0,1,'0.0000','–']]} };

/** ---------- Security: password login (hashes in Script Properties, session token in cache) ---------- */
function sha(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }
/** Run ONCE from the editor: type the two real passwords below, Run, then put the placeholders back. */
function setupPasswords() {
  const USER_PW = 'CHANGE-ME-MARIA', ADMIN_PW = 'CHANGE-ME-ADMIN';
  if (USER_PW.indexOf('CHANGE-ME') === 0 || ADMIN_PW.indexOf('CHANGE-ME') === 0) throw new Error('Edit USER_PW and ADMIN_PW first, then run again');
  if (USER_PW === ADMIN_PW) throw new Error('Use two different passwords');
  PropertiesService.getScriptProperties().setProperties({ PW_USER: sha(USER_PW), PW_ADMIN: sha(ADMIN_PW) });
}
function login(pw) {
  const cache = CacheService.getScriptCache(), P = PropertiesService.getScriptProperties();
  if (!P.getProperty('PW_ADMIN')) return { ok: false, error: 'Passwords not set – run setupPasswords once in Apps Script' };
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 6) return { ok: false, error: 'Too many wrong attempts – wait 10 minutes' };
  const h = sha(String(pw || ''));
  const role = h === P.getProperty('PW_ADMIN') ? 'admin' : (h === P.getProperty('PW_USER') ? 'user' : null);
  if (!role) { cache.put('fails', String(fails + 1), 600); Utilities.sleep(700); return { ok: false, error: 'Wrong password' }; }
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('t_' + token, role, 21600);                    // session lasts 6 hours
  return { ok: true, token: token, role: role };
}
function roleOf(token) { return token ? CacheService.getScriptCache().get('t_' + String(token)) : null; }

/** Run ONCE from the editor (Run > setup) to prepare the sheets and authorise the script. */
function setup() {
  sh(RATES); sh(PLANS);
  const p = sh(PLANS);
  if (p.getLastRow() < 1) p.getRange(1, 1, 1, PLAN_COLS.length).setValues([PLAN_COLS]).setFontWeight('bold');
  sh(EVENTS);
  if (readEvents().length === 0) writeEvents(DEFAULT_HOLIDAYS.map(function (h) { return { date: h[0], title: h[1], note: 'Bank holiday', type: 'Holiday' }; }));
  loadHolidays();
  fixParams();
  const r = readRates(); fillWk(r); writeRates(r);
  const s = sh(RATES);
  s.getRange('N8').setFormula('=LOOKUP(2,1/(B2:B1000<>""),B2:B1000)');
  s.getRange('N9').setFormula('=LOOKUP(2,1/(B2:B1000<>""),H2:H1000)');
  s.getRange('N10').setFormula('=LOOKUP(2,1/(B2:B1000<>""),J2:J1000)');
}

/** Run from the editor to load all rates (17 Jun – 2 Nov 2026) into the sheet. Keeps plans and any other rates you added. */
/** Units per bank statement (30-Sep-2026): 1,155.5629 units ÷ 2 partners. */
function fixParams() {
  const s = sh(RATES);
  s.getRange('M5').setValue('Buy NAV (derived)'); s.getRange('N5').setFormula('=(N2-N3-N4)/N6');
  s.getRange('M6').setValue('Units each (bank statement ÷ 2)'); s.getRange('N6').setValue(577.78145);
}
function importRates() {
  fixParams();
  const r = readRates();
  SEED.forEach(function (v, i) { r[addDays(START, i)] = v; });
  fillWk(r); writeRates(r);
  const s = sh(RATES);
  s.getRange('N8').setFormula('=LOOKUP(2,1/(B2:B1000<>""),B2:B1000)');
  s.getRange('N9').setFormula('=LOOKUP(2,1/(B2:B1000<>""),H2:H1000)');
  s.getRange('N10').setFormula('=LOOKUP(2,1/(B2:B1000<>""),J2:J1000)');
}

function sh(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { if (name === RATES) { s = ss.getSheets()[0]; s.setName(RATES); } else { s = ss.insertSheet(name); } }
  return s;
}
function iso(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).slice(0, 10); }
function addDays(d, k) { const t = new Date(d + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() + k); return t.toISOString().slice(0, 10); }
function isFri(d) { return new Date(d + 'T00:00:00Z').getUTCDay() === 5; }
function isNonBiz(d) { const w = new Date(d + 'T00:00:00Z').getUTCDay(); return w === 0 || w === 6 || HOL[d] === 1; }
/** After saving a rate for d: following Sat/Sun/holiday days take that rate (overwrites). Friday -> Sat+Sun; day before a holiday -> holiday. */
function wkSet(o, d) { let x = addDays(d, 1); while (isNonBiz(x)) { o[x] = o[d]; x = addDays(x, 1); } }
/** Fill (only empty) weekend / holiday days that follow a business-day rate. Returns true if anything was added. */
function fillWk(o) {
  let added = false;
  Object.keys(o).sort().forEach(function (d) {
    if (isNonBiz(d)) return;
    let x = addDays(d, 1);
    while (isNonBiz(x)) { if (o[x] === undefined) { o[x] = o[d]; added = true; } x = addDays(x, 1); }
  });
  return added;
}
function loadHolidays() { HOL = {}; readEvents().forEach(function (e) { if (e.type === 'Holiday') HOL[e.date] = 1; }); }

/** Read today's MIF rate from the Al Meezan fund-prices page (server-side, so no browser CORS limits). */
function fetchSite() {
  const res = UrlFetchApp.fetch(SITE, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-US,en;q=0.9', 'Referer': 'https://www.almeezangroup.com/' } });
  if (res.getResponseCode() !== 200) throw new Error('Bank website refused automatic access (HTTP ' + res.getResponseCode() + ') – enter the rate under Add a rate; weekends and holidays fill automatically');
  const text = res.getContentText().replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ').replace(/<\/(td|th)>/gi, ' | ').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
  const re = new RegExp('(?:^|\\|)\\s*' + FUND_NAME + '\\s*\\|\\s*\\d{1,2} \\w{3} \\d{4}\\s*\\|\\s*(\\d{1,2}) (\\w{3}) (\\d{4})\\s*\\|\\s*([\\d.]+)\\s*\\|\\s*([\\d.]+)\\s*\\|\\s*([\\d.]+)', 'i');
  const m = text.match(re);
  if (!m) throw new Error('Could not find the "' + FUND_NAME + '" row on the bank website (page layout may have changed)');
  const mon = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 }[m[2].toLowerCase()];
  const date = m[3] + '-' + ('0' + mon).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  const repurchase = Number(m[4]), nav = Number(m[6]);
  const rate = repurchase > 0 ? repurchase : nav;          // open-end funds publish the NAV in the Repurchase column
  if (!(rate > 20 && rate < 1000)) throw new Error('Unexpected NAV value on bank website: ' + rate);
  return { date: date, rate: rate };
}
/** Fetch the bank rate, store it for its validity date, carry weekends/holidays, audit gaps. Safe to call repeatedly. */
function syncSite() {
  loadHolidays();
  const r = readRates(), info = { ok: false, msg: '', at: new Date().toISOString() };
  try {
    const s = fetchSite();
    info.date = s.date; info.rate = s.rate;
    if (r[s.date] !== s.rate) { r[s.date] = s.rate; info.updated = true; }
    wkSet(r, s.date);
    info.ok = true; info.msg = info.updated ? 'Bank rate saved for ' + s.date : 'Already up to date (' + s.date + ')';
  } catch (e) { info.msg = String(e.message); }
  fillWk(r); writeRates(r);
  info.missing = missingDays(r);
  return info;
}
/** Weekdays (not holidays) with no rate between the first stored rate and today. */
function missingDays(r) {
  const ks = Object.keys(r).sort(), out = [];
  if (!ks.length) return out;
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  for (let d = ks[0]; d <= today; d = addDays(d, 1)) if (!isNonBiz(d) && r[d] === undefined) out.push(d);
  return out;
}
/** Time-driven trigger target (see installTrigger). */
function dailyJob() { syncSite(); }
/** Run ONCE: checks the bank website every evening (19:00 and 23:00, sheet time zone). */
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t); });
  [19, 23].forEach(function (h) { ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(h).create(); });
}

function readRates() {
  const s = sh(RATES), n = s.getLastRow(), o = {};
  if (n < 2) return o;
  s.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { if (r[0] !== '' && r[1] !== '' && !isNaN(Number(r[1]))) o[iso(r[0])] = Number(r[1]); });
  return o;
}
function calendarDates(o) {
  const set = {}; Object.keys(o).forEach(function (d) { set[d] = 1; });
  const ks = Object.keys(o).sort(), last = ks.length && ks[ks.length - 1] > END ? ks[ks.length - 1] : END;
  for (let d = CAL_START; d <= last; d = addDays(d, 1)) set[d] = 1;
  return Object.keys(set).sort();
}
function writeRates(o) {
  const s = sh(RATES), ds = calendarDates(o), n = ds.length;
  if (s.getMaxRows() < n + 2) s.insertRowsAfter(s.getMaxRows(), n + 2 - s.getMaxRows());
  s.getRange(2, 1, s.getMaxRows() - 1, 12).clearContent(); s.getRange('L1').setValue('Day');
  s.getRange(2, 1, n, 1).setNumberFormat('dd/mm/yyyy');
  s.getRange(2, 1, n, 2).setValues(ds.map(function (d) { return [d, o[d] === undefined ? '' : o[d]]; }));
  const f = ds.map(function (d, i) {
    const r = i + 2, B = 'B' + r, E = function (x) { return '=IF(' + B + '="","",' + x + ')'; };
    return [
      i === 0 ? E('0') : E('IF(B' + (r - 1) + '="",0,(B' + r + '-B' + (r - 1) + ')/B' + (r - 1) + ')'),
      E('$N$6'), E('$N$6'), E('D' + r + '*' + B), E('E' + r + '*' + B), E('F' + r + '+G' + r),
      E('2*$N$2'), E('H' + r + '-I' + r), E('J' + r + '/I' + r)
    ];
  });
  s.getRange(2, 3, n, 9).setFormulas(f);
  s.getRange(2, 12, n, 1).setFormulas(ds.map(function (d, i) { return ['=TEXT(A' + (i + 2) + ',"ddd")']; }));
  s.getRange(2, 3, n, 1).setNumberFormat('0.00%'); s.getRange(2, 11, n, 1).setNumberFormat('0.00%');
}
function readPlans() {
  const s = sh(PLANS), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, PLAN_COLS.length).getValues().filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const p = {}; PLAN_COLS.forEach(function (c, i) { p[c] = (c === 'date') ? iso(r[i]) : r[i]; });
    ['mAmt','dAmt','nav','mUnits','dUnits','units'].forEach(function (c) { p[c] = Number(p[c]); });
    return p;
  });
}
function writePlans(list) {
  const s = sh(PLANS);
  s.clear();
  s.getRange(1, 1, 1, PLAN_COLS.length).setValues([PLAN_COLS]).setFontWeight('bold');
  s.getRange(2, 2, 500, 1).setNumberFormat('dd/mm/yyyy');
  if (list.length) s.getRange(2, 2, list.length, 1).setNumberFormat('dd/mm/yyyy');
  if (list.length) s.getRange(2, 1, list.length, PLAN_COLS.length).setValues(list.map(function (p) { return PLAN_COLS.map(function (c) { return p[c] === undefined ? '' : p[c]; }); }));
}

function readEvents() {
  const s = sh(EVENTS), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, EV_COLS.length).getValues().filter(function (r) { return r[0] !== '' || r[1] !== ''; }).map(function (r) { return { date: iso(r[0]), title: String(r[1]), note: String(r[2]), type: r[3] === 'Holiday' ? 'Holiday' : 'Event' }; });
}
function writeEvents(list) {
  const s = sh(EVENTS);
  s.clear();
  s.getRange(1, 1, 1, EV_COLS.length).setValues([EV_COLS]).setFontWeight('bold');
  if (list.length) {
    s.getRange(2, 1, list.length, 1).setNumberFormat('dd/mm/yyyy');
    s.getRange(2, 1, list.length, EV_COLS.length).setValues(list.map(function (e) { return [e.date, e.title, e.note || '', e.type === 'Holiday' ? 'Holiday' : 'Event']; }));
  }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function state() { loadHolidays(); const r = readRates(); fillWk(r); return { rates: r, plans: readPlans(), events: readEvents(), config: CONFIG, version: VERSION }; }
function doGet() { return out({ ok: true, app: 'Meezan Investment Tracker', version: VERSION }); }   // no data without login

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ error: 'Bad request' }); }
  if (b.action === 'login') return out(login(b.password));
  const role = roleOf(b.token);
  if (!role) return out({ error: 'Session expired – please log in again', auth: false });
  if (ADMIN_ONLY[b.action] && role !== 'admin') return out({ error: 'Admin only action' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    loadHolidays();
    let sync = null;
    if (b.action === 'saveRate') {
      const nav = Number(b.nav);
      if (!(nav > 0) || !/^\d{4}-\d\d-\d\d$/.test(b.date)) throw new Error('Invalid date or rate');
      const r = readRates(); r[b.date] = nav; wkSet(r, b.date); writeRates(r);
    } else if (b.action === 'saveRates') {
      const m = b.rates || {}, r = readRates();
      Object.keys(m).forEach(function (d) {
        const nav = Number(m[d]);
        if (!(nav > 0) || !/^\d{4}-\d\d-\d\d$/.test(d)) throw new Error('Invalid date or rate: ' + d);
        r[d] = nav; wkSet(r, d);
      });
      writeRates(r);
    } else if (b.action === 'deleteRate') {
      const r = readRates(); const v = r[b.date]; delete r[b.date];
      { let x = addDays(b.date, 1); while (isNonBiz(x) && r[x] === v) { delete r[x]; x = addDays(x, 1); } }
      writeRates(r);
    } else if (b.action === 'savePlans') {
      if (!Array.isArray(b.plans)) throw new Error('Invalid plans');
      writePlans(b.plans);
    } else if (b.action === 'saveEvents') {
      if (!Array.isArray(b.events)) throw new Error('Invalid events');
      writeEvents(b.events);
    } else if (b.action === 'reset') {
      const r = {}; SEED.forEach(function (v, i) { r[addDays(START, i)] = v; }); fillWk(r); writeRates(r); writePlans([]);
    } else if (b.action === 'sync') {
      sync = syncSite();
    } else if (b.action === 'audit') {
      const r = readRates(); if (fillWk(r)) writeRates(r); sync = { ok: true, msg: 'Audit complete', missing: missingDays(r) };
    } else if (b.action === 'state') { /* read only */ } else { throw new Error('Unknown action'); }
    SpreadsheetApp.flush();
    const st = state(); st.role = role; if (sync) st.sync = sync;
    return out(st);
  } catch (err) {
    return out({ error: String(err.message) });
  } finally { lock.releaseLock(); }
}
