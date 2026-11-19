/**
 * Meezan Investment Tracker v3.4.0 – Google Sheet backend (Apps Script, bound to the Sheet).
 * Sheet "Rates": A Date, B NAV (you / the app enter these), C:K formulas (script writes them), M:N parameters.
 * Sheet "Plans": investment plans added from the app.
 * Rule: a Friday rate is applied automatically to Saturday and Sunday.
 */
const RATES = 'Rates', PLANS = 'Plans', TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
const START = '2026-06-17', CAL_START = '2026-06-01', END = '2026-12-31';   // START = first seed rate; calendar rows begin 1 June 2026   // calendar rows, like the Excel (rates blank until entered)
const SEED = [170.1922, 170.1922, 170.1922, 170.1922, 170.1922, 169.1158, 168.9046, 168.9046, 168.9046, 168.9046, 168.9046, 168.9046, 168.6041, 169.7, 172.0482, 172.6587, 173.5081, 173.5081, 173.5081, 175.3126, 173.5889, 169.023, 168.8494, 169.6963, 169.6963, 169.6963, 167.1206, 161.1499, 162.6919, 165.6168, 163.0619, 163.0619, 163.0619, 163.0872, 163.063, 161.5665, 159.0563, 158.5053, 158.5053, 158.5053, 165.3995, 164.9907, 163.085, 162.0444, 162.6489, 162.6489, 162.6489, 164.9567, 163.8439, 166.501, 168.0725, 167.95, 167.95, 167.95, 167.7894, 166.0406, 166.3348, 166.2893, 166.2893, 166.2893, 166.2893, 167.0441, 164.1014, 162.9894, 163.1252, 164.696, 164.696, 164.696, 164.5455, 164.5741, 164.3279, 164.3279, 164.8071, 164.8071, 164.8071, 164.8302, 164.8722, 163.1833, 163.241, 163.6147, 163.6147, 163.6147, 161.8187, 160.8015, 160.0133, 156.6644, 158.6359, 158.6359, 158.6359, 156.0864, 157.0741, 155.6762, 156.5435, 159.6763, 159.6763, 159.6763, 158.4533, 158.9749, 160.1294, 160.13, 158.156, 158.156, 158.156, 158.068, 157.5153, 157.9976, 156.7499, 156.2384, 156.2384, 156.2384, 154.3667, 156.9964, 156.8103, 155.8041, 155.4522, 155.4522, 155.4522, 153.6078, 153.4495, 154.1391, 154.8773, 154.7994, 154.7994, 154.7994, 154.49, 153.6336, 153.96, 152.076, 150.3916, 150.3916, 150.3916, 151.1222, 151.8735, 152.7288, 154.04, 153.1358, 153.1358, 153.1358, 153.0154];
const EVENTS = 'Events', EV_COLS = ['date','title','note','type'];
const SITE = 'https://www.almeezangroup.com/fund-prices/', FUND_NAME = 'Meezan Islamic Fund';
const DEFAULT_HOLIDAYS = [['2026-08-14','Independence Day'],['2026-11-09','Iqbal Day'],['2026-12-25','Quaid-e-Azam Day / Christmas']];
let HOL = {};   // holiday dates (yyyy-mm-dd) loaded per request from the Events sheet (type = Holiday)
const PLAN_COLS = ['name','date','mAmt','dAmt','nav','mUnits','dUnits','units','status'];
const VERSION = '3.4.0';
const ADMIN_ONLY = { reset: 1, deleteRate: 1 };           // actions only the admin password may run
// Private figures live ONLY here (server side) and are sent to the app after a successful login.
const CONFIG = { invest: 100000, charges: 1955, tax: 293.5, unitsTotal: 1155.5629, portfolio: '902300-1', bank: {port:'9023••-1',holder:'MARIA HASEEB MIRZA',amc:'AL MEEZAN INVESTMENT MANAGEMENT LTD',asat:'2026-10-30',printed:'2026-11-03',
 funds:[{n:'Meezan Islamic Fund (MIF)',t:'GROWTH-B',u:1155.5629,nav:153.1358,val:176958,gl:-23042,fy:-20167},{n:'Meezan Islamic Income Fund (MIIF)',t:'GROWTH-B',u:18.9207,nav:53.7525,val:1017,gl:17,fy:17}],
 tx:[['2026-07-21','PURCHASE (MIIF)',1000,0,'52.2512','18.9207'],['2026-07-21','LOAD CHARGED (MIIF)',0,10,'0.0000','–'],['2026-07-21','GOVT. TAX (MIIF)',0,1,'0.0000','–']]}, banks: [{port:'9023••-1',holder:'MARIA HASEEB MIRZA',amc:'AL MEEZAN INVESTMENT MANAGEMENT LTD',asat:'2026-09-30',printed:'2026-10-01',
 funds:[{n:'Meezan Islamic Fund (MIF)',t:'GROWTH-B',u:1155.5629,nav:157.9976,val:182576,gl:-17424,fy:-14549},{n:'Meezan Islamic Income Fund (MIIF)',t:'GROWTH-B',u:18.9207,nav:53.3381,val:1009,gl:9,fy:9}],
 tx:[['2026-07-21','PURCHASE (MIIF)',1000,0,'52.2512','18.9207'],['2026-07-21','LOAD CHARGED (MIIF)',0,10,'0.0000','–'],['2026-07-21','GOVT. TAX (MIIF)',0,1,'0.0000','–']]}, {port:'9023••-1',holder:'MARIA HASEEB MIRZA',amc:'AL MEEZAN INVESTMENT MANAGEMENT LTD',asat:'2026-10-30',printed:'2026-11-03',
 funds:[{n:'Meezan Islamic Fund (MIF)',t:'GROWTH-B',u:1155.5629,nav:153.1358,val:176958,gl:-23042,fy:-20167},{n:'Meezan Islamic Income Fund (MIIF)',t:'GROWTH-B',u:18.9207,nav:53.7525,val:1017,gl:17,fy:17}],
 tx:[['2026-07-21','PURCHASE (MIIF)',1000,0,'52.2512','18.9207'],['2026-07-21','LOAD CHARGED (MIIF)',0,10,'0.0000','–'],['2026-07-21','GOVT. TAX (MIIF)',0,1,'0.0000','–']]}] };

/** ---------- Login: simple fixed passwords (change them here any time, then Deploy > New version) ---------- */
const USER_PW = 'maria123', ADMIN_PW = 'maria123';   // same password = full access for everyone who knows it
function login(pw) {
  const p = String(pw == null ? '' : pw).trim();
  const role = p === ADMIN_PW ? 'admin' : (p === USER_PW ? 'user' : null);
  if (!role) { Utilities.sleep(500); return { ok: false, error: 'Wrong password' }; }
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('t_' + token, role, 21600);   // stays logged in for 6 hours
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
  latestFormulas();
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
  latestFormulas();
}

/** Latest NAV / value / gain-loss cells (N8:N10). ARRAYFORMULA is required in Google Sheets, otherwise the cells show #N/A. */
function latestFormulas() {
  const s = sh(RATES);
  s.getRange('M8').setValue('Latest NAV'); s.getRange('M9').setValue('Latest Total Value'); s.getRange('M10').setValue('Latest Gain/Loss');
  [['N8','B'],['N9','H'],['N10','J']].forEach(function (c) { s.getRange(c[0]).setFormula('=IFERROR(ARRAYFORMULA(LOOKUP(2,1/(B2:B1000<>""),' + c[1] + '2:' + c[1] + '1000)),"")'); });
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

/** Public text relay, used ONLY if the bank site refuses Google's servers (HTTP 403). Set to '' to disable. */
const RELAY = 'https://r.jina.ai/';
function pageText(url, headers) { const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: headers }); return { code: res.getResponseCode(), text: res.getContentText() }; }
/** Table rows as arrays of cell text, from HTML (<tr>/<td>) or from a markdown table (| a | b |). */
function tableRows(text) {
  const rows = [];
  if (/<tr[\s>]/i.test(text)) {
    (text.match(/<tr[\s\S]*?<\/tr>/gi) || []).forEach(function (tr) {
      const cells = (tr.match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || []).map(function (c) { return c.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim(); });
      if (cells.length) rows.push(cells);
    });
  } else {
    text.split(/\r?\n/).forEach(function (l) { l = l.trim(); if (l.charAt(0) === '|') rows.push(l.replace(/^\||\|$/g, '').split('|').map(function (c) { return c.trim(); })); });
  }
  return rows;
}
/** Read the "Repurchase (Rs.)" rate of Meezan Islamic Fund (Equity Funds) and its Validity Date from the Al Meezan fund-prices page. */
function fetchSite() {
  const H = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-US,en;q=0.9', 'Referer': 'https://www.almeezangroup.com/' };
  let r = pageText(SITE, H), via = 'bank website';
  if ((r.code !== 200 || !/Repurchase/i.test(r.text)) && RELAY) {
    try { const r2 = pageText(RELAY + SITE, { 'Accept': 'text/plain' }); if (r2.code === 200 && /Repurchase/i.test(r2.text)) { r = r2; via = 'relay'; } } catch (e) { /* keep first result */ }
  }
  if (r.code !== 200) throw new Error('Bank website refused automatic access (HTTP ' + r.code + ') – enter the rate under Add a rate; weekends and holidays fill automatically');
  const rows = tableRows(r.text);
  const head = rows.filter(function (c) { return c.some(function (x) { return /repurchase/i.test(x); }); })[0];
  if (!head) throw new Error('Could not find the "Repurchase (Rs.)" column on the bank website (page layout may have changed)');
  const iRep = head.findIndex(function (x) { return /repurchase/i.test(x); }), iVal = head.findIndex(function (x) { return /validity/i.test(x); });
  const row = rows.filter(function (c) { return String(c[0]).toLowerCase() === FUND_NAME.toLowerCase(); })[0];
  if (!row) throw new Error('Could not find the "' + FUND_NAME + '" row on the bank website (page layout may have changed)');
  const m = String(row[iVal]).match(/(\d{1,2}) (\w{3}) (\d{4})/);
  if (!m) throw new Error('Could not read the Validity Date for ' + FUND_NAME);
  const mon = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 }[m[2].toLowerCase()];
  const date = m[3] + '-' + ('0' + mon).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  const rate = Number(row[iRep]);
  if (!(rate > 20 && rate < 1000)) throw new Error('Unexpected Repurchase value on bank website: ' + row[iRep]);
  return { date: date, rate: rate, via: via };
}
/** ---------- Market page data: researched snapshot + live refresh from the Al Meezan website ---------- */
const MARKET = {"asof":"2026-11-19","policy":{"rate":12,"from":"2026-10-27"},"news":[{"d":"2026-11-13","t":"Launch of Meezan Paidaar Munafa Plan-60"},{"d":"2026-11-10","t":"Launch of Meezan Paidaar Munafa Plan-59"},{"d":"2026-11-04","t":"Launch of Meezan Paidaar Munafa Plan-58 (subscription 28 Oct – 3 Nov)"},{"d":"2026-10-09","t":"Maturity Update: Meezan Paidaar Munafa Plan-53 managed under Meezan Fixed Term Fund"},{"d":"2026-10-08","t":"Launch of Meezan Paidaar Munafa Plan-56"},{"d":"2026-10-07","t":"Maturity Update: Meezan Paidaar Munafa Plan-43 managed under Meezan Fixed Term Fund"},{"d":"2026-10-06","t":"Notice of Amendment to the Offering Document of Meezan Gold Fund"},{"d":"2026-09-25","t":"Launch of Meezan Paidaar Munafa Plan-55"},{"d":"2026-09-22","t":"Your Concerns Deserve To Be Heard"},{"d":"2026-09-16","t":"Notice of Amendment to the Offering Document of Meezan Strategic Allocation Fund"},{"d":"2026-09-09","t":"Launch of Meezan Paidaar Munafa Plan-54"},{"d":"2026-09-07","t":"Relocation of Hyderabad Branch"},{"d":"2026-09-01","t":"Notice of Amendment to the Offering Document of Meezan Tahaffuz Pension Fund"},{"d":"2026-08-31","t":"Launch of Meezan Islamic Government of Balochistan Pension Fund (MIGOBPF)"}],"macro":[{"d":"2026-11-18","t":"National Savings revises profit rates","x":"Special Savings Certificates 11.8% on the first five payouts; Regular Income Certificates 12.12%; Behbood, Pensioners' Benefit and Shuhada accounts 12.96%; Sarwa Islamic Savings Account 12.04%; Sarwa Islamic Term 1-year 12.04%, 3-year 11.98%, 5-year 12.12%.","s":"Minute Mirror / Daily Pakistan, 18 Nov 2026"},{"d":"2026-10-26","t":"State Bank raises policy rate to 12%","x":"The MPC raised the policy rate by 50 bps to 12%, effective 27 Oct 2026 – the first increase since April 2026. Overnight reverse-repo ceiling is 13%.","s":"State Bank of Pakistan / ARY News"},{"d":"2026-10-31","t":"Meezan Paidaar Munafa Plan-58 raises Rs 25.76 billion","x":"Subscription ran 28 Oct – 3 Nov 2026 and the plan launched on 4 Nov 2026; Rs 25.76 billion had been received by 31 Oct.","s":"Al Meezan Fund Manager's Report, Oct 2026"},{"d":"2026-06-30","t":"Al Meezan funds under management reach Rs 702.84 billion","x":"Al Meezan reported Rs 702.84 billion under management at 30 Jun 2026 and carries an AM1 (stable outlook) management-quality rating.","s":"Al Meezan Fund Manager's Reports, Jun & Oct 2026"},{"d":"2026-10-02","t":"Islamic banks declare Sept profit after month end","x":"Meezan Bank paid 6.63% on individual savings (up to Rs 10 million) for September 2026 and declared 9.55% on its 1-year certificate for October. Islamic bank profit is declared after the month, not promised in advance.","s":"Pakera (bank rate sheets), Oct–Nov 2026"}],"rates":[["Equity Funds","Al Meezan Mutual Fund","30 Oct 2026",47.089,48.172,0.0,-3.0,-8.63,-7.94,29.87,64.7,10536.76],["Equity Funds","Meezan Islamic Fund","30 Oct 2026",153.1358,156.6579,0.0,-3.08,-10.23,-8.39,33.02,59.22,3669.21],["Equity Funds","Meezan Energy Fund","30 Oct 2026",63.8151,66.0167,0.0,-2.29,3.39,2.22,26.23,53.99,146.53],["Equity Funds","Meezan Dedicated Equity Fund","30 Oct 2026",89.6806,92.7745,0.0,-2.51,-8.07,-8.21,27.96,65.72,64.56],["Index Tracker Fund","KSE Meezan Index Fund","30 Oct 2026",170.1844,174.5241,0.0,-2.42,-7.87,-5.08,38.09,43.52,647.9],["Balanced Fund","Meezan Balanced Fund","30 Oct 2026",26.8016,27.418,0.0,-1.17,-3.39,-1.01,20.77,36.6,1447.04],["Asset Allocation Fund","Meezan Dynamic Asset Allocation Fund MDY","30 Oct 2026",81.5007,83.3752,0.0,-1.48,-6.14,-2.76,35.5,45.92,85.58],["Asset Allocation Fund","MCPF-III-Meezan Capital Secure Plan-I","30 Oct 2026",49.7836,51.5011,0.0,-1.4,-0.6,-0.05,0.55,0.0,-0.05],["Asset Allocation Fund","Meezan Asset Allocation Fund","30 Oct 2026",110.7008,114.5199,0.0,-2.24,-6.75,-7.53,24.94,58.21,84.76],["Asset Allocation Fund","MCPF-III-Meezan Capital Secure Plan-II","30 Oct 2026",49.6872,0.0,0.0,-0.58,-0.63,-0.63,0.0,0.0,-0.63],["Income Funds","Meezan Islamic Income Fund Type B Units","30 Oct 2026",53.7525,54.3706,0.0,9.45,10.39,9.57,9.03,12.54,29.53],["Income Funds","Meezan Islamic Income Fund Type C Units","30 Oct 2026",53.7525,53.7525,0.0,9.45,10.39,9.57,9.03,12.54,10.39],["Income Funds","Meezan Sovereign Fund","30 Oct 2026",54.0456,54.6671,0.0,7.67,8.4,8.28,8.4,12.0,22.66],["Income Funds","Meezan Daily Income Plan – I","30 Oct 2026",50.0,50.575,0.0,9.78,9.82,9.57,9.37,13.38,8.54],["Income Funds","Meezan Daily Income Fund-MMMP","30 Oct 2026",0.0,0.0,50.0,10.02,9.92,9.75,9.5,11.15,9.96],["Income Funds","MDIF-Meezan Sehl Account Plan","30 Oct 2026",0.0,0.0,52.3169,10.18,9.94,9.66,8.88,10.34,14.75],["Income Funds","MDIF-Meezan Munafa Plan-I","30 Oct 2026",52.3053,52.9068,0.0,10.43,10.41,10.16,9.94,12.95,16.34],["Income Funds","MDIF-Meezan Super Saver Plan","30 Oct 2026",0.0,0.0,52.2587,9.9,10.03,9.58,9.27,12.75,12.58],["Income Funds","MGSF–Meezan Government Securities Plan-I","30 Oct 2026",51.809,52.4048,0.0,9.89,11.1,11.1,0.0,0.0,11.1],["Money Market Funds","Meezan Cash Fund","30 Oct 2026",0.0,0.0,53.3489,10.21,10.11,9.82,9.55,13.34,24.28],["Money Market Funds","Meezan Rozana Amdani Fund","30 Oct 2026",0.0,0.0,50.0,10.21,10.18,9.86,9.49,13.62,5.61],["Money Market Funds","Meezan Islamic Asaan Cash Fund","30 Oct 2026",0.0,0.0,51.7969,10.3,10.21,10.29,10.02,0.0,10.29],["Commodities Fund","Meezan Gold Fund","29 Oct 2026",262.916,271.5922,0.0,-4.36,-6.64,-6.62,22.15,42.72,501.38],["Fund of Funds Scheme","MFPF-Very Conservative Allocation Plan","29 Oct 2026",54.2307,54.8543,0.0,9.46,9.75,9.47,9.19,10.62,14.25],["Fund of Funds Scheme","Aggressive Allocation Plan","29 Oct 2026",177.1262,181.2001,0.0,-1.17,-4.41,-2.54,25.93,47.75,437.67],["Fund of Funds Scheme","Moderate Allocation Plan","29 Oct 2026",143.8905,146.3726,0.0,-0.59,-2.16,0.76,20.45,36.25,399.61],["Fund of Funds Scheme","Conservative Allocation Plan","29 Oct 2026",126.6232,128.0793,0.0,0.15,0.85,5.32,15.26,23.76,308.14],["Fund of Funds Scheme","Meezan Asset Allocation Plan-I","29 Oct 2026",134.2436,0.0,0.0,-2.0,-7.58,-7.11,22.26,46.45,237.55],["Fund of Funds Scheme","Meezan Strategic Allocation Plan-III","29 Oct 2026",85.6125,0.0,0.0,-1.56,-5.96,-5.71,21.6,47.12,128.7],["Fund of Funds Scheme","Meezan Strategic Allocation Plan-IV","29 Oct 2026",90.3737,0.0,0.0,-1.7,-6.53,-5.25,22.67,45.49,131.68],["Fund of Funds Scheme","Meezan Strategic Allocation Plan-V","29 Oct 2026",99.4296,0.0,0.0,-1.64,-6.38,-7.52,19.73,44.09,152.01],["Fund of Funds Scheme","Meezan Strategic Allocation Plan-I","29 Oct 2026",90.2985,0.0,0.0,-1.53,-5.89,-5.32,21.23,50.03,165.57],["Fund of Funds Scheme","Meezan Strategic Allocation Plan-II","29 Oct 2026",77.41,0.0,0.0,-1.66,-6.39,-6.16,22.04,49.58,139.74],["Pension Scheme","Meezan Tahaffuz Pension Fund – Gold sub fund","29 Oct 2026",0.0,0.0,489.2036,-4.31,-6.57,-6.41,21.74,39.51,389.22],["Pension Scheme","Meezan Tahaffuz Pension Fund – Equity Sub Fund","30 Oct 2026",0.0,0.0,1697.4892,-2.97,-8.79,-7.83,30.36,65.72,1597.55],["Pension Scheme","Meezan Tahaffuz Pension Fund – Debt sub fund","30 Oct 2026",0.0,0.0,545.7938,10.23,10.39,9.88,9.5,14.38,24.92],["Pension Scheme","Meezan Tahaffuz Pension Fund – Money Market sub fund","30 Oct 2026",0.0,0.0,544.8692,10.47,10.5,10.31,9.97,15.43,24.89],["Pension Scheme","Meezan GOKP Pension Fund-MMSF","30 Oct 2026",0.0,0.0,148.928,10.27,11.03,10.87,10.81,15.77,16.99],["Pension Scheme","MIGOPPF-MMSF","30 Oct 2026",0.0,0.0,110.5503,10.53,14.06,11.27,9.23,0.0,11.23],["Pension Scheme","MIGOBPF-MMSF","30 Oct 2026",0.0,0.0,102.97,10.86,10.84,10.84,0.0,0.0,10.84],["Exchange Traded Fund","Meezan Pakistan Exchange Traded Fund","30 Oct 2026",0.0,0.0,16.9939,-2.03,-6.31,-3.49,39.59,31.76,112.25],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 55","28 Oct 2026",0.0,0.0,50.3124,11.4,11.4,11.4,0.0,0.0,11.4],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 48","29 Oct 2026",0.0,0.0,51.8727,10.45,11.2,11.3,11.59,0.0,11.3],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 56","29 Oct 2026",0.0,0.0,50.2493,11.37,11.37,11.37,0.0,0.0,11.37],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 49","29 Oct 2026",0.0,0.0,51.8307,11.06,11.23,11.23,0.0,0.0,11.23],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 57","29 Oct 2026",0.0,0.0,50.1396,11.32,11.32,11.32,0.0,0.0,11.32],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 50","29 Oct 2026",0.0,0.0,51.4104,10.97,11.19,11.19,0.0,0.0,11.19],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 52","28 Oct 2026",0.0,0.0,51.3449,11.62,11.69,11.69,0.0,0.0,11.69],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 34","29 Oct 2026",0.0,0.0,52.2773,13.16,13.63,12.26,10.44,0.0,12.26],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 54","28 Oct 2026",0.0,0.0,50.701,12.08,11.9,11.9,0.0,0.0,11.9],["Fixed Term Fund (Open End CIS)","MFTF-Meezan Paidar Munafa Plan 47","28 Oct 2026",0.0,0.0,51.854,10.3,11.18,11.16,10.71,0.0,11.16]],"ratesAsOf":"2026-10-30","ns":[["Behbood Savings Certificates","12.96%","Monthly","Seniors 60+, widows, disabled","No"],["Pensioners' Benefit Account","12.96%","Monthly","Pensioners","No"],["Shuhada Family Welfare Account","12.96%","Monthly","Martyrs' families","No"],["Regular Income Certificates","12.12%","Monthly","All Pakistanis","No"],["Special Savings Certificates","11.8% (first 5 payouts), 6th 12.8%","Every 6 months","3 years","No"],["Short-Term Savings Certificates","3m 11.56% · 6m 11.90% · 1y 12.14%","At maturity","3 / 6 / 12 months","No"],["Savings Account","10.50%","Periodic","Any amount","No"],["Sarwa Islamic Savings Account","12.04% (expected)","Monthly","Open to resident adults","Yes"],["Sarwa Islamic Term Account – 1 year","12.04% (expected)","At maturity","1 year","Yes"],["Sarwa Islamic Term Certificate – 3 years","11.98% (expected)","Every 6 months","3 years","Yes"],["Sarwa Islamic Term Certificate – 5 years","12.12% (expected)","Monthly","5 years","Yes"],["Defence Savings Certificates","Unchanged (cumulative, up to ~215% over 10 years)","At maturity","10 years","No"]],"banks":[["Meezan Bank","Islamic Savings (individuals up to Rs 10m)","6.63%","Declared monthly (Mudarabah)","Sep 2026","Yes"],["Meezan Bank","Certificate of Islamic Investment – 1 year (min Rs 50,000)","9.55%","At maturity; declared after month end","Oct 2026","Yes"],["Dubai Islamic Bank Pakistan","Savings (E-Savings)","6.45%","Monthly pool profit","Sep 2026","Yes"],["Dubai Islamic Bank Pakistan","Premium Term Deposit – 12 months","up to 6.85% (6.50% typical)","At maturity","Sep 2026","Yes"],["BankIslami","1-year term deposit (monthly profit)","8.50%","70:30 profit share","Aug 2026","Yes"],["BankIslami","Savings","6.78%","Monthly pool profit","Aug 2026","Yes"],["Faysal Bank (Islamic)","Islamic Savings","6.37%","Monthly pool profit","Oct 2026","Yes"],["Bank Alfalah","Alfa digital 1-year deposit","11.60%","Conventional (interest-based)","Nov 2026","No"],["HBL, UBL, MCB, NBP, Allied, Bank Alfalah, Al Habib, BOP, Standard Chartered","Standard savings account","about 10% (SBP minimum floor)","Conventional (interest-based)","Oct 2026","No"]],"banklist":[["Meezan Bank","Full Islamic bank","Largest Islamic network; Al Meezan's sister company and distributor"],["Dubai Islamic Bank Pakistan","Full Islamic bank","Useful for UAE-linked Pakistanis"],["BankIslami Pakistan","Full Islamic bank","Low-balance Islamic accounts"],["Faysal Bank","Islamic (largest Islamic share) with conventional history","Islamic savings and Islamic funds through Faysal AMC"],["Al Baraka Bank (Pakistan)","Full Islamic bank","Rates not researched – ask the bank"],["HBL / UBL / MCB / Bank Alfalah / Allied","Conventional banks with Islamic windows or AMCs","Check that the product is the Islamic one"],["National Savings (CDNS)","Government savings","Sarwa Islamic products are the Shariah-compliant ones"],["Al Meezan / NBP / HBL / Faysal / Atlas / ABL / Alfalah asset managers","Mutual-fund managers","Islamic funds sold direct or via banks"]],"amcs":[["NBP Islamic Money Market Fund","AA(f)","Money market","11.13% YTD · 11.27% MTD (annualised)","18 Nov 2026"],["HBL Islamic Money Market Fund","AA+(f)","Money market","10.09% YTD · 9.74% MTD (annualised)","11 Nov 2026"],["Faysal Islamic Cash Fund","AA(f)","Money market","9.97% YTD · 11.19% MTD (annualised)","17 Nov 2026"],["Atlas Islamic Money Market Fund","–","Money market","10.32% (FY26, 1-year)","FY26"],["ABL Islamic Money Market Plan-I","–","Money market","10.34% (FY26, 1-year)","FY26"],["NBP Riba Free Savings Fund","–","Income","9.9% (FY26, 1-year)","FY26"],["Faysal Halal Amdani Fund","–","Income / money market","9.29% (FY26, 1-year)","FY26"],["HBL Islamic Stock Fund","–","Equity (high risk)","+27.45% (FY26, 1-year)","FY26"],["Faysal Islamic Stock Fund","–","Equity (high risk)","+31.08% (FY26, 1-year)","FY26"]],"plans":[{"n":"Emergency & liquidity","g":"3–6 months of expenses, available any day","r":"Low","i":"Islamic savings account, Meezan Cash Fund / Rozana Amdani, NBP or HBL Islamic Money Market Fund"},{"n":"Monthly income","g":"A steady monthly cheque","r":"Low–Medium","i":"Sarwa Islamic Savings or 5-year Term (monthly), Behbood / Regular Income (if eligible and you accept conventional), Islamic income funds with monthly payout"},{"n":"Short term (3–12 months)","g":"Money you need within a year","r":"Low","i":"Meezan Paidar Munafa fixed-term plans, Sarwa Islamic 1-year, Islamic 1-year certificate"},{"n":"Medium term (1–3 years)","g":"A planned purchase or goal","r":"Low–Medium","i":"Meezan Islamic Income / Sovereign Fund, Government Securities Plan, Sarwa Islamic 3-year"},{"n":"Long-term growth (5+ years)","g":"Wealth building, retirement, children's future","r":"High","i":"KSE Meezan Index Fund, Meezan Islamic Fund, Dynamic / Asset Allocation funds – invested monthly (SIP style) to average the entry price"},{"n":"Retirement","g":"Long-term pension savings","r":"Medium–High","i":"Meezan Tahaffuz Pension Fund (equity / debt / money-market sub-funds); ask about tax credit rules that apply to you"},{"n":"Gold hedge","g":"Small diversifier against inflation / currency","r":"High","i":"Meezan Gold Fund – a small slice only; its price can fall as well as rise"}],"steps":[["1. Decide goal, horizon and risk","Write down what the money is for, when you need it and how much loss you can accept. Money needed within a year belongs in low-risk options; equity funds suit 5+ years."],["2. Check it is Shariah-compliant and well rated","Look for a Shariah advisor's certificate, the AMC rating (Al Meezan is AM1) and the fund's own stability rating for income and money-market funds."],["3. Compare costs","Read the factsheet: sales load (the bank statement notes up to 3% when buying direct or through third-party portals and up to 1.5% through the AMC's own portal), management fee, total expense ratio and back-end load."],["4. Open the account","Mutual funds: apply with Al Meezan (online, branch or via Meezan Bank) with your CNIC, bank account proof (cheque/IBAN), nominee, zakat declaration and risk-profile form. Banks: apply at the branch or app. National Savings: apply at a National Savings centre or authorised bank. Ask each provider for the current document list."],["5. Fund it and choose units","Send money by IBFT / Raast / cheque in favour of the fund, and pick the unit type (growth or income) and fund. Keep the payment slip."],["6. Confirm and keep records","Check the confirmation and monthly statement, quote your portfolio number in all correspondence, and report any error within 7 working days as stated on the statement."],["7. Review monthly, not daily","Read the monthly Fund Manager's Report, compare with your goal and rebalance once or twice a year. Daily NAV moves are noise for long-term money."],["8. Know taxes and zakat","Profit may be subject to withholding tax and zakat; rules differ by product, so confirm before investing."]],"src":[["Al Meezan – Fund Prices","https://www.almeezangroup.com/fund-prices/"],["Al Meezan – News & Announcements","https://www.almeezangroup.com/news-announcement/"],["Al Meezan – Fund Manager's Report Oct 2026","https://www.almeezangroup.com/assets/uploads/2026/02/FMR-October-2026.pdf"],["State Bank of Pakistan – policy rate","https://www.sbp.org.pk/"],["Minute Mirror – National Savings revised rates","https://minutemirror.com.pk/national-savings-revises-profit-rates-on-multiple-savings-schemes-662172/"],["Daily Pakistan – National Savings Nov 2026","https://en.dailypakistan.com.pk/18-Nov-2026/pakistan-revises-national-savings-profit-rates-starting-november-2026"],["Pakera – fixed deposit rates Nov 2026","https://pakera.pk/fixed-deposit-rates-banks-pakistan-november-2026/"],["Pakera – Islamic bank savings (Meezan 6.63%)","https://pakera.pk/islamic-bank-savings-profit-75-percent-pool-rule-meezan-pakistan/"],["Pakera – Dubai Islamic Bank rates","https://pakera.pk/dubai-islamic-bank-profit-rates-2026-savings-term-deposit/"],["MUFAP – daily fund performance","https://www.mufap.com.pk/Industry/IndustryStatDaily?tab=1"],["HalalWallet – halal funds compared","https://www.halalwallet.pk/blog/best-halal-mutual-funds-pakistan-2026"]],"mix":{"Conservative":[35,50,10,5],"Balanced":[20,35,40,5],"Growth":[10,20,65,5]},"buckets":[["Cash & liquidity","Islamic savings account, Meezan Cash Fund, Islamic money-market funds"],["Fixed income","Sarwa Islamic certificates, Islamic income / sovereign funds, fixed-term plans"],["Equity","KSE Meezan Index Fund, Meezan Islamic Fund, asset-allocation plans"],["Gold","Meezan Gold Fund (small slice)"]]};
function getPage(url) {
  const H = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36', 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-US,en;q=0.9', 'Referer': 'https://www.almeezangroup.com/' };
  let r = { code: 0, text: '' };
  try { r = pageText(url, H); } catch (e) { /* try relay */ }
  if ((r.code !== 200 || r.text.length < 500) && RELAY) { try { const r2 = pageText(RELAY + url, { 'Accept': 'text/plain' }); if (r2.code === 200) r = r2; } catch (e) { /* keep */ } }
  return r.code === 200 ? r.text : null;
}
const MON3 = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
function isoOf(s) { const m = String(s).match(/(\d{1,2}) (\w{3}) (\d{4})/); return m && MON3[m[2].toLowerCase()] ? m[3] + '-' + ('0' + MON3[m[2].toLowerCase()]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : ''; }
function parseRates(text) {
  const rows = tableRows(text), out = [], TOP = /(Equity Funds|Index Tracker|Balanced Fund|Asset Allocation|Income Funds|Money Market|Commodities|Fund of Funds|Pension Scheme|Exchange Traded|Open End CIS)/i;
  const hi = rows.findIndex(function (c) { return c.some(function (x) { return /repurchase/i.test(x); }); });
  if (hi < 0) return out;
  const h = rows[hi], ix = function (re) { return h.findIndex(function (x) { return re.test(x); }); };
  const I = { val: ix(/validity/i), rep: ix(/repurchase/i), off: ix(/offer/i), nav: ix(/^nav/i), mtd: ix(/mtd/i), fy: ix(/fytd/i), cy: ix(/cytd/i), f25: ix(/fy25/i), f24: ix(/fy24/i), si: ix(/since/i) };
  const num = function (c, k) { const v = Number(String(c[I[k]] == null ? '' : c[I[k]]).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
  let cat = 'Funds';
  rows.slice(hi + 1).forEach(function (c) {
    if (c.length < 3 || c.slice(1).every(function (x) { return x === ''; })) { if (c[0] && TOP.test(c[0])) cat = c[0]; return; }
    const iso = isoOf(c[I.val]); if (!iso || !c[0]) return;
    const d = new Date(iso + 'T00:00:00Z');
    out.push([cat, c[0], d.getUTCDate() + ' ' + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getUTCMonth()] + ' ' + d.getUTCFullYear(), num(c,'rep'), num(c,'off'), num(c,'nav'), num(c,'mtd'), num(c,'fy'), num(c,'cy'), num(c,'f25'), num(c,'f24'), num(c,'si')]);
  });
  return out;
}
function parseNews(text) {
  const lines = text.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, '\n').replace(/&nbsp;/g, ' ').split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
  const out = [], seen = {};
  for (let i = 0; i < lines.length - 1; i++) {
    if (!/^\d{1,2} [A-Za-z]{3} \d{4}$/.test(lines[i])) continue;
    let t = lines[i + 1].replace(/^#+\s*/, '').replace(/\*+/g, '').replace(/\[([^\]]+)\]\([^)]*\)/g, '$1').trim();
    if (!t || /^learn more$/i.test(t) || /^\d{1,2} [A-Za-z]{3} \d{4}$/.test(t)) continue;
    const d = isoOf(lines[i]), k = d + t; if (!d || seen[k]) continue; seen[k] = 1; out.push({ d: d, t: t });
  }
  return out.sort(function (a, b) { return a.d < b.d ? 1 : -1; }).slice(0, 40);
}
function liveMarket() {
  const o = {};
  try { const t = getPage(SITE); if (t) { const r = parseRates(t); if (r.length > 10) { o.rates = r; o.ratesAsOf = r.map(function (x) { return isoOf(x[2]); }).sort().pop(); } } } catch (e) { /* snapshot */ }
  try { const n = getPage('https://www.almeezangroup.com/news-announcement/'); if (n) { const l = parseNews(n); if (l.length > 3) o.news = l; } } catch (e) { /* snapshot */ }
  return (o.rates || o.news) ? o : null;
}
function marketData(force) {
  const cache = CacheService.getScriptCache(), out = JSON.parse(JSON.stringify(MARKET));
  let live = null; const c = force ? null : cache.get('mk_live');
  if (c) { live = JSON.parse(c); } else { live = liveMarket(); if (live) { try { cache.put('mk_live', JSON.stringify(live), 21600); } catch (e) { /* too big */ } } }
  if (live && live.rates) { out.rates = live.rates; out.ratesAsOf = live.ratesAsOf; out.ratesLive = true; }
  if (live && live.news) { out.news = live.news; out.newsLive = true; }
  return out;
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
    info.ok = true; info.msg = (info.updated ? 'Bank rate saved for ' + s.date : 'Already up to date (' + s.date + ')') + ' · Repurchase rate via ' + s.via;
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
  latestFormulas();
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

const RATINGS = 'Ratings', RT_COLS = ['date','user','stars','status','gl','note'];
function readRatings() {
  const s = sh(RATINGS), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, RT_COLS.length).getValues().filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { date: iso(r[0]), user: String(r[1]), stars: Number(r[2]), status: r[3] === 'plus' ? 'plus' : 'minus', gl: Number(r[4]), note: String(r[5] || '') };
  }).slice(-50);
}
function addRating(role, b) {
  const st = Math.round(Number(b.stars));
  if (!(st >= 1 && st <= 5)) throw new Error('Rating must be 1 to 5 stars');
  const s = sh(RATINGS);
  if (s.getLastRow() < 1) s.getRange(1, 1, 1, RT_COLS.length).setValues([RT_COLS]).setFontWeight('bold');
  s.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), role === 'admin' ? 'Admin' : 'Maria', st, b.status === 'plus' ? 'plus' : 'minus', Math.round(Number(b.gl) || 0), String(b.note || '').slice(0, 200)]);
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
function state() { loadHolidays(); const r = readRates(); fillWk(r); return { rates: r, plans: readPlans(), events: readEvents(), ratings: readRatings(), config: CONFIG, version: VERSION }; }
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
    } else if (b.action === 'market') {
      return out({ market: marketData(!!b.force) });
    } else if (b.action === 'saveRating') {
      addRating(role, b);
    } else if (b.action === 'state') { /* read only */ } else { throw new Error('Unknown action'); }
    SpreadsheetApp.flush();
    const st = state(); st.role = role; if (sync) st.sync = sync;
    return out(st);
  } catch (err) {
    return out({ error: String(err.message) });
  } finally { lock.releaseLock(); }
}
