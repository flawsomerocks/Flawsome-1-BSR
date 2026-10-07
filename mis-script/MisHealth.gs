/**
 * MIS HEALTH  -  SEPARATE NEW FILE (nothing else is touched)   [add as "Health.gs"]
 * Run  HEALTH_SETUP  once. Every day ~12 noon (after all fills) it checks the whole MIS system:
 *   triggers, DashData / SpendData freshness, month columns, error cells, every tab's channel / product / SKU totals
 *   against DashData, Blinkit (RO) against Invoice Log, Master Data against SpendData.
 * If something is wrong it first tries to FIX it (re-runs the fills), checks again, and only if the problem is still
 * there it sends you an EMAIL. Everything is also written to the "MIS Log" tab.
 * Needs MisAutoFill.gs (helper functions) in the same project; other files are optional.
 */
const HC = {
  HOUR: 12,
  EMAIL: [],                 // extra recipients; the owner of the sheet always gets the mail
  STALE_DAYS: 2,             // DashData / SpendData older than this many days = problem
  TOL: 1,                    // rupees of difference that is ignored
  TABS: [
    { name: 'MIS as per 1 Total', words: ['mis', 'total'], measure: 'total', kind: 'dash' },
    { name: 'MIS as per ITEM PRICE', words: ['mis', 'item'], measure: 'item price', kind: 'dash' },
    { name: 'MIS as per RO', key: 'misasperro', measure: 'total', kind: 'ro' }
  ],
  MASTER_KEY: 'misaspermasterdata'
};

function HEALTH_SETUP() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'healthDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('healthDaily').timeBased().everyDays(1).atHour(HC.HOUR).create();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = hcRun_(ss);
  hcLog_(ss, r, 'SETUP');
  hcMail_(ss, r, 'MIS Health: setup ho gaya (test mail)', true);
  ss.toast(r.problems.length ? r.problems.length + ' problem mili. "MIS Log" dekho.' : 'Sab theek hai.', 'MIS Health', 10);
}

// daily: check -> try to fix -> check again -> mail only if still broken
function healthDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let r = hcRun_(ss), fixed = [];
  if (r.problems.length) {
    const fixes = [
      ['fillDaily', () => { if (typeof fillDaily === 'function') { fillDaily(); return true; } }],
      ['rowsDaily', () => { if (typeof rowsDaily === 'function') { rowsDaily(); return true; } }],
      ['masterDaily', () => { if (typeof masterDaily === 'function') { masterDaily(); return true; } }],
      ['roDaily', () => { if (typeof roDaily === 'function') { roDaily(); return true; } }]
    ];
    const fixTriggers = [['fillDaily', 9], ['rowsDaily', 8], ['masterDaily', 10], ['roDaily', 11]];
    fixTriggers.forEach(x => {
      try {
        let exists = false;
        if (x[0] === 'fillDaily') exists = typeof fillDaily === 'function';
        if (x[0] === 'rowsDaily') exists = typeof rowsDaily === 'function';
        if (x[0] === 'masterDaily') exists = typeof masterDaily === 'function';
        if (x[0] === 'roDaily') exists = typeof roDaily === 'function';
        const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === x[0]);
        if (exists && !has) { ScriptApp.newTrigger(x[0]).timeBased().everyDays(1).atHour(x[1]).create(); fixed.push('trigger ' + x[0] + ' dobara bana'); }
      } catch (e) { r.notes.push('trigger fix ' + x[0] + ' fail: ' + e.message); }
    });
    fixes.forEach(f => {
      try { if (f[1]()) fixed.push(f[0]); } catch (e) { r.notes.push('auto-fix ' + f[0] + ' fail: ' + e.message); }
    });
    try { const b = hcFixBlinkit_(ss); if (b) fixed.push(b); } catch (e) { r.notes.push('Blinkit fix fail: ' + e.message); }
    SpreadsheetApp.flush();
    const before = r.problems.length;
    r = hcRun_(ss);
    r.notes.push('Auto-fix chalaya (' + fixed.join(', ') + '): problems ' + before + ' -> ' + r.problems.length);
  }
  hcLog_(ss, r, r.problems.length ? 'PROBLEM' : 'OK');
  if (r.problems.length) hcMail_(ss, r, 'MIS Health: ' + r.problems.length + ' problem hai', false);
}

function hcRun_(ss) {
  const res = { problems: [], notes: [] };
  const P = m => res.problems.push(m), N = m => res.notes.push(m);
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  // 1. triggers
  const have = {};
  ScriptApp.getProjectTriggers().forEach(t => { have[t.getHandlerFunction()] = 1; });
  [['fillDaily', typeof fillDaily === 'function', '9am MIS fill (Code.gs)'],
   ['rowsDaily', typeof rowsDaily === 'function', '8am Auto rows (ROWS_SETUP)'],
   ['masterDaily', typeof masterDaily === 'function', '10am Master Data (MASTER_SETUP)'],
   ['roDaily', typeof roDaily === 'function', '11am MIS as per RO (RO_SETUP)']].forEach(x => {
    if (x[1] && !have[x[0]]) P('Trigger band/missing: ' + x[2] + '. Us file ka SETUP function dobara chalao.');
  });

  // 2. DashData
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!dd) { P('DashData tab nahi mila.'); return res; }
  const values = dd.getDataRange().getValues();
  let maxD = null;
  const head = (values[0] || []).map(h => String(h).trim().toLowerCase());
  const iD = head.indexOf('date');
  if (iD < 0) { P('DashData me "Date" header nahi mila.'); return res; }
  let noDate = 0;
  for (let r = 1; r < values.length; r++) {
    const d = values[r][iD];
    if (hcIsDate_(d) && d.getFullYear() >= 2000) { if (!maxD || d > maxD) maxD = d; } else if (values[r].some(c => c !== '')) noDate++;
  }
  if (!maxD) P('DashData me koi valid date nahi.');
  else {
    const age = Math.round((today - new Date(maxD.getFullYear(), maxD.getMonth(), maxD.getDate())) / 86400000);
    if (age > HC.STALE_DAYS) P('DashData purana hai: last date ' + hcFmt_(maxD) + ' (' + age + ' din pehle). Sales sheet ka DashData trigger / portal tabs check karo.');
    else N('DashData last date: ' + hcFmt_(maxD));
  }
  if (noDate) N('DashData me ' + noDate + ' rows bina valid date ki (count nahi hoti).');

  // 3. SpendData
  const sd = ss.getSheetByName(MISC.SPEND_TAB);
  let spendVals = null;
  if (sd) {
    const a1 = String(sd.getRange(1, 1).getValue());
    if (/^#|error|loading/i.test(a1)) P('SpendData A1 par error: "' + a1 + '". A1 par click karke "Allow access" dabao.');
    else {
      spendVals = sd.getDataRange().getValues();
      let mx = null; spendVals.forEach(r => { if (hcIsDate_(r[0]) && (!mx || r[0] > mx)) mx = r[0]; });
      if (!mx) P('SpendData me koi date nahi (Ads tracker ka _Daily khaali ya link toota).');
      else {
        const age = Math.round((today - new Date(mx.getFullYear(), mx.getMonth(), mx.getDate())) / 86400000);
        if (age > HC.STALE_DAYS) P('SpendData purana: last date ' + hcFmt_(mx) + '. College wale Ads tracker ka daily refresh band ho sakta hai.');
        else N('SpendData last date: ' + hcFmt_(mx));
      }
    }
  } else N('SpendData tab nahi (Ad Spend link nahi kiya).');

  // 4. each MIS tab for the running month and the previous month
  const months = [[now.getFullYear(), now.getMonth() - 1], [now.getFullYear(), now.getMonth()]].map(x => {
    const f = new Date(x[0], x[1], 1); return [f.getFullYear(), f.getMonth()];
  });
  HC.TABS.forEach(spec => {
    const tab = spec.key ? ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === spec.key) : misFindSheet_(ss, spec.words, false);
    if (!tab) { N('"' + spec.name + '" tab nahi mila (skip).'); return; }
    try { hcCheckTab_(ss, tab, spec, values, months, res); } catch (e) { P('[' + spec.name + '] check me error: ' + e.message); }
  });

  // 5. Master Data tab vs SpendData
  const mt = ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === HC.MASTER_KEY);
  if (mt && spendVals) {
    try { hcCheckMaster_(mt, spendVals, months, res); } catch (e) { P('[MIS as per Master Data] check me error: ' + e.message); }
  }
  return res;
}

function hcCheckTab_(ss, tab, spec, values, months, res) {
  const nm = '[' + spec.name + '] ';
  const P = m => res.problems.push(nm + m), N = m => res.notes.push(nm + m);
  if (tab.getLastColumn() < 4 || tab.getLastRow() < 20) { P('tab khaali / chhota hai.'); return; }
  const hr = misHeaderRow_(tab);
  const grid = tab.getRange(1, 1, tab.getLastRow(), 4).getValues();
  const lab = grid.map(r => String(r[1]).trim().toLowerCase());
  // channel block
  let st = -1; grid.forEach((r, i) => { if (st < 0 && String(r[0]).trim().toLowerCase() === 'channel wise revenue split') st = i; });
  const chRows = []; if (st >= 0) for (let i = st; i < grid.length && lab[i]; i++) chRows.push(i);
  const idx = l => { const i = grid.findIndex(r => r.some(c => String(c).trim().toLowerCase() === l)); return i; };
  const catRows = Object.keys(MISC.CATS).map(k => idx(MISC.CATS[k])).filter(i => i >= 0);
  const skuRows = [];
  grid.forEach((r, i) => r.forEach(c => { if (/^F[A-Z]+\d+$/.test(String(c).trim().toUpperCase())) skuRows.push(i); }));
  const restRow = idx(MISC.REST_SKU), revRow = idx(MISC.REV), blinkRow = idx('blinkit');
  months.forEach(([y, m]) => {
    const tag = (m + 1) + '/' + y + ': ';
    const agg = misAggregate_(values, y, m, spec.measure);
    const col = misFindMonthCol_(tab, hr, y, m);
    if (!agg.rows) { N(tag + 'DashData me data nahi.'); return; }
    if (!col) { P(tag + 'month ka column nahi hai.'); return; }
    const colVals = tab.getRange(1, col, tab.getLastRow(), 1).getValues().map(r => r[0]);
    const errs = colVals.filter(v => typeof v === 'string' && /^#(REF|VALUE|N\/A|DIV|NAME|ERROR|NUM)/i.test(v)).length;
    if (errs) P(tag + errs + ' cell me formula error (#REF!/#VALUE! jaisa).');
    const num = i => Number(colVals[i]) || 0;
    // channel rows: what each row SHOULD hold from DashData
    let chSum = 0, chExp = 0;
    chRows.forEach(i => {
      const l = lab[i];
      chSum += num(i);
      if (spec.kind === 'ro' && l === 'blinkit') return;
      chExp += MISC.CHANNELS[l] ? (agg.channels[l] || 0) : (agg.unmappedPortals[l] || 0);
    });
    const chSumNoBlink = chSum - (spec.kind === 'ro' && blinkRow >= 0 ? num(blinkRow) : 0);
    if (Math.abs(chSumNoBlink - chExp) > HC.TOL) P(tag + 'channel rows ka total ' + Math.round(chSumNoBlink) + ' hai, DashData se ' + Math.round(chExp) + ' hona chahiye.');
    // portals that have NO row at all
    const rowSet = {}; chRows.forEach(i => { rowSet[lab[i]] = 1; });
    const lost = Object.keys(agg.unmappedPortals).filter(p => p && p.charAt(0) !== '(' && !rowSet[p] && agg.unmappedPortals[p] >= 1);
    if (lost.length) {
      const txt = lost.map(p => p + ' ' + Math.round(agg.unmappedPortals[p])).join(', ');
      if (spec.kind === 'ro') N(tag + 'in portals ki row nahi (RO tab me jaan-boojh kar): ' + txt);
      else P(tag + 'in portals ki channel row nahi hai, Revenue kam aa raha: ' + txt + '. (AutoRows chalao)');
    }
    // revenue row
    if (revRow >= 0 && chRows.length) {
      const rv = num(revRow);
      if (Math.abs(rv - chSum) > HC.TOL) P(tag + 'Revenue row (' + Math.round(rv) + ') channel rows ke total (' + Math.round(chSum) + ') se alag hai.');
    }
    // product + SKU consistency
    const catSum = catRows.reduce((s, i) => s + num(i), 0);
    if (catRows.length === 6 && Math.abs(catSum - agg.revenue) > HC.TOL) P(tag + 'Product wise split ' + Math.round(catSum) + ', DashData ' + Math.round(agg.revenue) + '.');
    if (skuRows.length && restRow >= 0) {
      const sk = skuRows.reduce((s, i) => s + num(i), 0) + num(restRow);
      if (Math.abs(sk - agg.revenue) > HC.TOL) P(tag + 'SKU wise split + Rest SKU ' + Math.round(sk) + ', DashData ' + Math.round(agg.revenue) + '.');
    }
    // units
    const ui = idx(MISC.UNITS);
    if (ui >= 0 && Math.abs(num(ui) - agg.units) > HC.TOL) P(tag + 'Units ' + num(ui) + ', DashData ' + agg.units + '.');
    // Blinkit from RO invoices
    if (spec.kind === 'ro' && blinkRow >= 0 && typeof roInvoiceInfo_ === 'function') {
      const info = roInvoiceInfo_(ss);
      if (info.err) N(tag + 'Invoice Log check skip: ' + info.err);
      else {
        const exp = info.expected[y + '-' + m] || 0, got = num(blinkRow);
        if (Math.abs(got - exp) > HC.TOL) P(tag + 'Blinkit RO: sheet me ' + Math.round(got) + ', Invoice Log me ' + Math.round(exp) + '. (Invoice tool me "Sync now" dabao)');
      }
    }
    if (!res.problems.some(p => p.indexOf(nm + tag) === 0)) N(tag + 'theek (Revenue ' + Math.round(colVals[revRow] || 0) + ').');
  });
}

function hcCheckMaster_(tab, sp, months, res) {
  const nm = '[MIS as per Master Data] ';
  if (tab.getLastColumn() < 5) return;
  const hr = 2;
  const grid = tab.getRange(1, 1, Math.min(tab.getLastRow(), 12), 1).getValues();
  const revRow = grid.findIndex(r => String(r[0]).trim().toLowerCase() === 'revenue') + 1;
  const mktRow = grid.findIndex(r => String(r[0]).trim().toLowerCase().indexOf('marketing') === 0) + 1;
  months.forEach(([y, m]) => {
    const tag = (m + 1) + '/' + y + ': ';
    const col = misFindMonthCol_(tab, hr, y, m);
    let rev = 0, mkt = 0, rows = 0;
    sp.forEach(r => { if (hcIsDate_(r[0]) && r[0].getFullYear() === y && r[0].getMonth() === m) { rows++; rev += Number(r[2]) || 0; mkt += Number(r[5]) || 0; } });
    if (!rows) return;
    if (!col) { res.problems.push(nm + tag + 'month ka column nahi hai (MASTER_SETUP / masterDaily).'); return; }
    const v = tab.getRange(revRow, col).getValue(), k = tab.getRange(mktRow, col).getValue();
    if (typeof v === 'string' && v[0] === '#') res.problems.push(nm + tag + 'Revenue me formula error.');
    else if (Math.abs((Number(v) || 0) - rev) > HC.TOL) res.problems.push(nm + tag + 'Revenue ' + Math.round(Number(v) || 0) + ', SpendData me ' + Math.round(rev) + '.');
    else if (Math.abs((Number(k) || 0) - mkt) > HC.TOL) res.problems.push(nm + tag + 'Marketing ' + Math.round(Number(k) || 0) + ', SpendData me ' + Math.round(mkt) + '.');
    else res.notes.push(nm + tag + 'theek.');
  });
}

function hcIsDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
function hcFmt_(d) { return d.getDate() + '-' + (d.getMonth() + 1) + '-' + d.getFullYear(); }

function hcText_(r, tag) {
  return 'MIS HEALTH ' + (tag || '') + ' - ' + (r.problems.length ? r.problems.length + ' PROBLEM' : 'SAB THEEK') +
    (r.problems.length ? '\n\nPROBLEMS:\n- ' + r.problems.join('\n- ') : '') +
    (r.notes.length ? '\n\nDETAIL:\n- ' + r.notes.join('\n- ') : '');
}

function hcLog_(ss, r, tag) { const t = hcText_(r, tag); Logger.log(t); misLog_(ss, t); }

function hcMail_(ss, r, subject, always) {
  try {
    const me = Session.getEffectiveUser().getEmail();
    const to = [me].concat(HC.EMAIL).filter(Boolean).join(',');
    if (!to) return;
    MailApp.sendEmail(to, subject, hcText_(r, '') + '\n\nSheet: ' + ss.getUrl());
  } catch (e) { Logger.log('mail fail: ' + e.message); }
}


// RO tab: a plain number (not a formula) sitting in the Blinkit row where Invoice Log has invoices is replaced by the
// verified invoice formula. A cell that already holds a formula is never touched here.
function hcFixBlinkit_(ss) {
  if (typeof roInvoiceInfo_ !== 'function' || typeof roFillBlinkitFromInvoices_ !== 'function') return '';
  const tab = ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === 'misasperro');
  if (!tab || tab.getLastRow() < 20) return '';
  const info = roInvoiceInfo_(ss);
  if (info.err) return '';
  const hr = misHeaderRow_(tab);
  const grid = tab.getRange(1, 1, tab.getLastRow(), 4).getValues();
  let row = 0; grid.forEach((r, i) => r.forEach(c => { if (!row && String(c).trim().toLowerCase() === 'blinkit') row = i + 1; }));
  if (!row) return '';
  const now = new Date(), done = [];
  [[now.getFullYear(), now.getMonth() - 1], [now.getFullYear(), now.getMonth()]].forEach(x => {
    const f = new Date(x[0], x[1], 1), y = f.getFullYear(), m = f.getMonth();
    const col = misFindMonthCol_(tab, hr, y, m);
    if (!col) return;
    const cell = tab.getRange(row, col), exp = info.expected[y + '-' + m] || 0;
    if (exp > 0 && !cell.getFormula() && Math.abs((Number(cell.getValue()) || 0) - exp) > HC.TOL) {
      cell.clearContent();
      roFillBlinkitFromInvoices_(tab, cell, y, m);
      done.push((m + 1) + '/' + y);
    }
  });
  return done.length ? 'Blinkit RO cell invoice formula se badla: ' + done.join(', ') : '';
}
