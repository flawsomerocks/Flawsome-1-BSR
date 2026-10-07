/**
 * MIS AUTO FILL  -  ONE FILE, paste this whole thing in the "MIS DEEP ANALISIS" sheet
 * (Extensions > Apps Script), Save, run  START_HERE  once.  That's it.
 *
 * Reads tab "DashData" and fills one month column in BOTH MIS tabs:
 *   "MIS as per 1 Total"      <- DashData "Total" column
 *   "MIS as per ITEM PRICE"   <- DashData "ITEM PRICE" column (blank -> Total is used)
 * Fills: Revenue (=SUM of channel rows), Units Sold, Channel / Product / SKU split.
 * NOT filled (not in DashData): Marketing Spends, Platform Margin, Opex -> manual.
 * Auto-runs EVERY DAY at ~9am: refreshes the previous month and the running month (MTD).
 */
const MISC = {
  DATA_TAB: 'DashData',
  SPEND_TAB: 'SpendData',   // IMPORTRANGE of the Ads tracker's hidden _Daily tab (see SETUP_SPEND)
  SPEND_LABEL: 'marketing spends (e-commerce)',
  // tab found by words in its name (case-insensitive), so small spelling changes are ok
  TABS: [
    { label: 'MIS as per 1 Total', words: ['mis', 'total'], measure: 'total' },
    { label: 'MIS as per ITEM PRICE', words: ['mis', 'item'], measure: 'item price' }
  ],
  // MIS channel label -> DashData "Portal" names
  CHANNELS: {
    'amazon': ['amz fba', 'amazon vendor', 'amazon dropship'],
    'flipkart': ['flipkart'], 'myntra': ['myntra'], 'snapdeal': ['snapdeal'],
    'tata 1mg': ['1mg'], 'first cry': ['firstcry'], 'tata cliq': ['tatacliq'],
    'pharmeasy': ['pharmeasy'], 'website': ['website'], 'apollo': ['apollo'],
    'cred': ['cred'], 'wh smith': ['wh smith'], 'meesho': ['meesho'],
    'swiggy instamrt': ['swiggy'], 'pop club': ['pop'], 'blinkit': ['blinkit']
  },
  CATS: {
    padCotton: 'sanitary pads cotton based', padPlant: 'sanitary pads plant based',
    plCotton: 'panty liner cotton based', plPlant: 'panty liner plant based',
    pants: 'period pants', combos: 'combos'
  },
  REV: 'revenue', UNITS: 'units sold', REST_SKU: "rest sku's (combo's)",
  FORMULA_LABELS: ['cogs', 'cm1', 'platform margin', 'cm2', 'cm2 %', 'roas (blended)']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 MIS')
    .addItem('START HERE (fill now + daily auto-run 9am)', 'START_HERE')
    .addItem('Fill previous month only', 'fillPreviousMonth')
    .addItem('Fill a specific month…', 'fillSpecificMonth')
    .addItem('Link Ad Spend (one time)', 'SETUP_SPEND')
    .addToUi();
}

// run this once: builds empty tabs, fills previous + running month, sets the daily 9am trigger
function START_HERE() {
  installMisTrigger_();
  buildEmptyTabs_();     // empty MIS tabs -> format + all past months from DashData
  fillDaily();           // previous month + running month (MTD)
}

function buildEmptyTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  const msgs = [];
  if (!dd) { misLog_(ss, 'ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const values = dd.getDataRange().getValues();
  MISC.TABS.forEach(t => {
    const mis = misFindSheet_(ss, t.words, false);
    if (!mis || mis.getLastColumn() >= 3 && mis.getLastRow() >= 5) return;   // missing or already has content
    try { msgs.push(misBuildTab_(mis, t, values)); } catch (e) { msgs.push('ERROR build [' + t.label + ']: ' + e.message); }
  });
  if (msgs.length) { misLog_(ss, msgs.join('\n\n')); SpreadsheetApp.flush(); }
}

function installMisTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === 'fillPreviousMonth' || h === 'fillDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fillDaily').timeBased().everyDays(1).atHour(9).create();
}

// daily job: previous month (settles late data) + running month (MTD)
function fillDaily() {
  const n = new Date();
  fillMonths_([[n.getFullYear(), n.getMonth() - 1], [n.getFullYear(), n.getMonth()]]);
}

function fillPreviousMonth() {
  const n = new Date();
  fillMonths_([[n.getFullYear(), n.getMonth() - 1]]);
}

function fillSpecificMonth() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Month daalo (yyyy-mm), jaise 2026-09');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim().match(/^(\d{4})-(\d{1,2})$/);
  if (!m) { ui.alert('Format galat. Example: 2026-09'); return; }
  fillMonths_([[Number(m[1]), Number(m[2]) - 1]]);
}

function fillMonths_(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  const msgs = [];
  if (!dd) {
    msgs.push('ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.');
  } else {
    const values = dd.getDataRange().getValues();
    list.forEach(ym => {
      const first = new Date(ym[0], ym[1], 1);
      const year = first.getFullYear(), monthIdx = first.getMonth();
      MISC.TABS.forEach(t => {
        const mis = misFindSheet_(ss, t.words, false);
        if (!mis) { msgs.push('ERROR: tab nahi mila: "' + t.label + '"'); return; }
        try {
          msgs.push(fillOneMis_(mis, t, misAggregate_(values, year, monthIdx, t.measure), year, monthIdx));
        } catch (e) { msgs.push('ERROR [' + t.label + ']: ' + e.message); }
      });
    });
  }
  const msg = msgs.join('\n\n');
  Logger.log(msg);
  misLog_(ss, msg);
  ss.toast('Report: "MIS Log" tab dekho', 'MIS', 10);
}

function misLog_(ss, msg) {
  let sh = ss.getSheetByName('MIS Log');
  if (!sh) {
    sh = ss.insertSheet('MIS Log');
    sh.getRange(1, 1, 1, 2).setValues([['Time', 'Report']]).setFontWeight('bold');
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 700);
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 2).setValues([[new Date(), msg]]);
  sh.getRange(2, 2).setWrap(true);
  sh.getRange(2, 1, 1, 2).setVerticalAlignment('top');
}

function fillOneMis_(mis, tab, agg, year, monthIdx) {
  const title = '[' + mis.getName() + '] ' + (monthIdx + 1) + '/' + year;
  if (mis.getLastColumn() < 3 || mis.getLastRow() < 5) {
    return title + '\nYe tab khaali hai. Pehle MIS ka format (Month header row, Revenue, Channel Wise, SKU rows) is tab me daalo (xlsx import / copy-paste), phir dobara run karo.';
  }
  if (!agg.rows) return title + '\nDashData me is mahine ka data nahi mila. Kuch nahi likha.';

  const hr = misHeaderRow_(mis);
  let c = misFindMonthCol_(mis, hr, year, monthIdx);
  let created = false;
  if (!c) { c = misInsertMonthCol_(mis, hr, year, monthIdx); created = true; }
  const L = misColLetter_(c);

  // row finder: label can sit in any of columns A-D; SKU rows by SKU code (FSPCL10 ...)
  const grid = mis.getRange(1, 1, mis.getLastRow(), 4).getValues();
  const rowOf = {}, skuRows = {};
  grid.forEach((r, i) => r.forEach(cell => {
    const t = String(cell).trim().toLowerCase();
    if (t && rowOf[t] === undefined) rowOf[t] = i + 1;
    const code = t.toUpperCase();
    if (/^F[A-Z]+\d+$/.test(code) && !skuRows[code]) skuRows[code] = i + 1;
  }));

  const missing = [];
  const put = (label, v) => {
    const row = rowOf[label];
    if (!row) { missing.push(label); return; }
    mis.getRange(row, c).setValue(Math.round(v * 100) / 100);
  };
  let chFirst = 1e9, chLast = 0;
  Object.keys(MISC.CHANNELS).forEach(k => {
    if (rowOf[k]) { chFirst = Math.min(chFirst, rowOf[k]); chLast = Math.max(chLast, rowOf[k]); }
    put(k, agg.channels[k] || 0);
  });
  Object.keys(MISC.CATS).forEach(k => put(MISC.CATS[k], agg.cats[k] || 0));
  Object.keys(skuRows).forEach(k => mis.getRange(skuRows[k], c).setValue(Math.round((agg.skus[k] || 0) * 100) / 100));
  Object.keys(agg.skus).forEach(k => { if (!skuRows[k]) agg.restSku += agg.skus[k]; });
  put(MISC.REST_SKU, agg.restSku);
  put(MISC.UNITS, agg.units);
  if (chLast && rowOf[MISC.REV]) mis.getRange(rowOf[MISC.REV], c).setFormula('=SUM(' + L + chFirst + ':' + L + chLast + ')');
  else missing.push(MISC.REV);

  const spendOn = !!mis.getParent().getSheetByName(MISC.SPEND_TAB);
  if (spendOn && rowOf[MISC.SPEND_LABEL]) {
    mis.getRange(rowOf[MISC.SPEND_LABEL], c).setFormula(misSpendFormula_(year, monthIdx + 1));
  }
  if (created) {   // new column: copy the formulas of the previous month column (COGS, CM1 ...)
    MISC.FORMULA_LABELS.forEach(lb => {
      const row = rowOf[lb];
      if (!row) return;
      const f = mis.getRange(row, c - 1).getFormulaR1C1();
      if (f) mis.getRange(row, c).setFormulaR1C1(f);
    });
  }

  return [
    title + ' -> column ' + L + (created ? ' (naya column banaya)' : ' (purana column update)'),
    'Rows: ' + agg.rows + ' | Revenue(DashData): ' + Math.round(agg.revenue) + ' | Units: ' + agg.units,
    agg.fallbackRows ? 'ITEM PRICE khaali tha -> Total use hua: ' + agg.fallbackRows + ' rows (' + Math.round(agg.fallbackAmt) + ')' : '',
    Object.keys(agg.unmappedPortals).length ? 'Ye portal MIS me nahi hai (channel split me nahi gaya): ' +
      Object.keys(agg.unmappedPortals).map(p => p + '=' + Math.round(agg.unmappedPortals[p])).join(', ') : '',
    missing.length ? 'MIS me ye row nahi mili (skip): ' + missing.join(', ') : '',
    'Rest SKU (combos / unmatched): ' + Math.round(agg.restSku),
    spendOn ? 'Marketing Spends: SpendData se linked (formula). Manual bharna hai: Platform Margin, Opex.'
            : 'Manual bharna hai: Marketing Spends, Platform Margin, Opex. (Spend link karne ke liye menu > Link Ad Spend)'
  ].filter(Boolean).join('\n');
}

// ---------- calculation (no Sheets calls) ----------
function misAggregate_(values, year, monthIdx, measure) {
  const head = values[0].map(h => String(h).trim().toLowerCase());
  const ix = n => head.indexOf(n);
  const iDate = ix('date'), iU = ix('sku'), iQ = ix('qty'), iP = ix('portal'), iT = ix('total'), iM = ix(measure || 'total');
  if ([iDate, iU, iQ, iP, iT, iM].some(i => i < 0)) throw new Error('DashData headers nahi mile: ' + head.join(' | '));

  const portalToCh = {};
  Object.keys(MISC.CHANNELS).forEach(k => MISC.CHANNELS[k].forEach(p => { portalToCh[p] = k; }));

  const a = { fallbackRows: 0, fallbackAmt: 0, rows: 0, revenue: 0, units: 0, channels: {}, cats: {}, skus: {}, restSku: 0, unmappedPortals: {} };
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    let d = row[iDate];
    if (!(d instanceof Date)) d = new Date(d);
    if (isNaN(d.getTime()) || d.getFullYear() !== year || d.getMonth() !== monthIdx) continue;
    let rev = Number(row[iM]);
    if (iM !== iT && !(rev > 0)) { rev = Number(row[iT]) || 0; a.fallbackRows++; a.fallbackAmt += rev; }
    rev = rev || 0;
    const qty = Number(row[iQ]) || 0;
    const portal = String(row[iP]).trim().toLowerCase();
    a.rows++; a.revenue += rev; a.units += qty;

    const ch = portalToCh[portal];
    if (ch) a.channels[ch] = (a.channels[ch] || 0) + rev;
    else a.unmappedPortals[portal] = (a.unmappedPortals[portal] || 0) + rev;

    const key = String(row[iU]).trim().toUpperCase();
    const cat = misCatOf_(key);
    a.cats[cat] = (a.cats[cat] || 0) + rev;
    if (key) a.skus[key] = (a.skus[key] || 0) + rev;
  }
  return a;
}

// FPLPLA.. panty liner plant | FPLC.. panty liner cotton | FPP.. period pants
// FSPC.. pads cotton | FSPP.. pads plant | anything else = combos
function misCatOf_(code) {
  if (code.indexOf('FPLPLA') === 0) return 'plPlant';
  if (code.indexOf('FPLC') === 0) return 'plCotton';
  if (code.indexOf('FPP') === 0) return 'pants';
  if (code.indexOf('FSPC') === 0) return 'padCotton';
  if (code.indexOf('FSPP') === 0) return 'padPlant';
  return 'combos';
}

// ---------- sheet helpers ----------
function misFindSheet_(ss, words, exact) {
  const sheets = ss.getSheets();
  for (const s of sheets) {
    const n = s.getName().toLowerCase();
    if (exact ? n.trim() === words[0] : words.every(w => n.indexOf(w) >= 0)) return s;
  }
  return null;
}

// header row = first of rows 1-6 whose column A says "Month" (default 2)
function misHeaderRow_(mis) {
  const a = mis.getRange(1, 1, 6, 1).getValues();
  for (let i = 0; i < a.length; i++) if (String(a[i][0]).trim().toLowerCase() === 'month') return i + 1;
  return 2;
}

// month headers are dates typed like "May-26" (read as 26-May) or real month dates (1-Jun-2026)
function misHeaderToYm_(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  const day = v.getDate();
  return { y: day === 1 ? v.getFullYear() : 2000 + day, m: v.getMonth() };
}

function misFindMonthCol_(mis, hr, y, m) {
  const hdr = mis.getRange(hr, 1, 1, mis.getLastColumn()).getValues()[0];
  for (let i = 2; i < hdr.length; i++) {
    const ym = misHeaderToYm_(hdr[i]);
    if (ym && ym.y === y && ym.m === m) return i + 1;
  }
  return 0;
}

function misInsertMonthCol_(mis, hr, y, m) {
  const hdr = mis.getRange(hr, 1, 1, mis.getLastColumn()).getValues()[0];
  let tot = hdr.findIndex(h => String(h).trim().toLowerCase() === 'total') + 1;
  if (!tot) tot = mis.getLastColumn() + 1;
  mis.insertColumnBefore(tot);
  mis.getRange(1, tot - 1, mis.getMaxRows(), 1)
    .copyTo(mis.getRange(1, tot, mis.getMaxRows(), 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  mis.getRange(hr, tot).setValue(new Date(y, m, 1)).setNumberFormat('mmm-yy');
  return tot;
}

function misColLetter_(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}


// ---------- builds the MIS format in an EMPTY tab and fills every past month ----------
const MISL = {
  CH: ['Amazon', 'Flipkart', 'Myntra', 'Snapdeal', 'TATA 1MG', 'First cry', 'Tata Cliq', 'Pharmeasy',
       'Website', 'Apollo', 'Cred', 'Wh smith', 'Meesho', 'Swiggy INSTAMRT', 'PoP Club', 'Blinkit'],
  CAT: ['Sanitary Pads Cotton Based', 'Sanitary Pads Plant Based', 'Panty Liner Cotton Based',
        'Panty Liner Plant Based', 'Period Pants', 'Combos'],
  CAT_KEYS: ['padCotton', 'padPlant', 'plCotton', 'plPlant', 'pants', 'combos'],
  SKUS: [
    ['FPLPLA30', 'Flawsome Panty Liner (Pack of 30) Plant-Based'],
    ['FPLC30', 'Flawsome Panty Liner (Pack of 30) Organic Cotton'],
    ['FPPL4', 'Flawsome Period Pants (Pack of 4, Size L)'],
    ['FPPXL4', 'Flawsome Period Pants (Pack of 4, Size XL)'],
    ['FSPCL10', 'Flawsome Organic Sanitary Pads (Pack of 10, Size L)'],
    ['FSPCL20', 'Flawsome Organic Sanitary Pads (Pack of 20, Size L)'],
    ['FSPCL30', 'Flawsome Organic Sanitary Pads (Pack of 30, Size L)'],
    ['FSPCXL10', 'Flawsome Organic Sanitary Pads (Pack of 10, Size XL)'],
    ['FSPCXL20', 'Flawsome Organic Sanitary Pads (Pack of 20, Size XL)'],
    ['FSPCXL30', 'Flawsome Organic Sanitary Pads (Pack of 30, Size XL)'],
    ['FSPCXXL10', 'Flawsome Organic Sanitary Pads (Pack of 10, Size XXL)'],
    ['FSPCXXL20', 'Flawsome Organic Sanitary Pads (Pack of 20, Size XXL)'],
    ['FSPCXXL30', 'Flawsome Organic Sanitary Pads (Pack of 30, Size XXL)'],
    ['FSPPLAL10', 'Flawsome Sensitive Sanitary Pads (Pack of 10, Size L)'],
    ['FSPPLAL20', 'Flawsome Sensitive Sanitary Pads (Pack of 20, Size L)'],
    ['FSPPLAL30', 'Flawsome Sensitive Sanitary Pads (Pack of 30, Size L)'],
    ['FSPPLAXL10', 'Flawsome Sensitive Sanitary Pads (Pack of 10, Size XL)'],
    ['FSPPLSXL20', 'Flawsome Sensitive Sanitary Pads (Pack of 20, Size XL)'],
    ['FSPPLSXL30', 'Flawsome Sensitive Sanitary Pads (Pack of 30, Size XL)'],
    ['FSPPLAXXL10', 'Flawsome Sensitive Sanitary Pads (Pack of 10, Size XXL)'],
    ['FSPPLAXXL20', 'Flawsome Sensitive Sanitary Pads (Pack of 20, Size XXL)'],
    ['FSPPLAXXL30', 'Flawsome Sensitive Sanitary Pads (Pack of 30, Size XXL)']
  ]
};

function misMonthsFrom_(values) {
  const iDate = values[0].map(h => String(h).trim().toLowerCase()).indexOf('date');
  let min = null;
  for (let r = 1; r < values.length; r++) {
    const d = values[r][iDate];
    if (d instanceof Date && !isNaN(d.getTime()) && d.getFullYear() >= 2000 && (!min || d < min)) min = d;
  }
  const out = [];
  if (!min) return out;
  const now = new Date(), end = new Date(now.getFullYear(), now.getMonth() - 1, 1);   // previous month
  for (let d = new Date(min.getFullYear(), min.getMonth(), 1); d <= end && out.length < 36; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
    out.push(d);
  }
  return out;
}

function misBuildTab_(mis, tab, values) {
  const months = misMonthsFrom_(values);
  if (!months.length) return '[' + mis.getName() + '] DashData me koi valid date nahi mili. Tab nahi bana.';
  const N = 66, FIRST = 4, TOTAL = FIRST + months.length;
  const spendOn = !!mis.getParent().getSheetByName(MISC.SPEND_TAB);
  const rows = [];
  for (let i = 0; i < N; i++) rows.push([]);

  // labels: A = section / line item, B = channel / product name, C = SKU code
  const A = (r, v) => { rows[r - 1][0] = v; }, B = (r, v) => { rows[r - 1][1] = v; }, C = (r, v) => { rows[r - 1][2] = v; };
  A(1, 'XLEAP CARE PRIVATE LIMITED'); A(2, 'Month'); C(2, 'SKU Code');
  A(4, 'Revenue'); A(6, 'COGS'); A(8, 'CM1'); A(10, 'Marketing Spends (E-Commerce)'); A(11, 'Platform Margin');
  A(13, 'CM2'); A(14, 'CM2 %'); A(16, 'ROAS (Blended)'); A(18, 'Units Sold');
  A(20, 'Channel Wise Revenue Split'); MISL.CH.forEach((n, i) => B(20 + i, n));
  A(37, 'Product Wise Revenue Split'); MISL.CAT.forEach((n, i) => B(37 + i, n));
  A(44, 'SKU Wise Revenue Split'); MISL.SKUS.forEach((s, i) => { B(44 + i, s[1]); C(44 + i, s[0]); });
  B(66, "Rest SKU's (Combo's)");

  const sumRows = [4, 6, 8, 10, 11, 13, 18];
  for (let i = 0; i < 16; i++) sumRows.push(20 + i);
  for (let i = 0; i < 6; i++) sumRows.push(37 + i);
  for (let i = 0; i < MISL.SKUS.length; i++) sumRows.push(44 + i);
  sumRows.push(66);

  months.forEach((m, k) => {
    const c = FIRST + k, L = misColLetter_(c);
    const agg = misAggregate_(values, m.getFullYear(), m.getMonth(), tab.measure);
    const put = (r, v) => { rows[r - 1][c - 1] = v; };
    put(2, m);
    put(4, '=SUM(' + L + '20:' + L + '35)');
    put(6, '=' + L + '4*30%'); put(8, '=' + L + '4-' + L + '6');
    put(11, '=' + L + '4*30%');
    if (spendOn) put(10, misSpendFormula_(m.getFullYear(), m.getMonth() + 1));
    put(13, '=' + L + '8-' + L + '10-' + L + '11');
    put(14, '=IFERROR(' + L + '13/' + L + '4,"")'); put(16, '=IFERROR(' + L + '4/' + L + '10,"")');
    put(18, agg.units);
    MISL.CH.forEach((n, i) => put(20 + i, Math.round((agg.channels[n.toLowerCase()] || 0) * 100) / 100));
    MISL.CAT_KEYS.forEach((key, i) => put(37 + i, Math.round((agg.cats[key] || 0) * 100) / 100));
    let rest = agg.restSku;
    const known = {};
    MISL.SKUS.forEach((s, i) => { known[s[0]] = 1; put(44 + i, Math.round((agg.skus[s[0]] || 0) * 100) / 100); });
    Object.keys(agg.skus).forEach(k2 => { if (!known[k2]) rest += agg.skus[k2]; });
    put(66, Math.round(rest * 100) / 100);
  });

  // Total column
  const TL = misColLetter_(TOTAL);
  rows[1][TOTAL - 1] = 'Total';
  sumRows.forEach(r => { rows[r - 1][TOTAL - 1] = '=SUM(D' + r + ':INDEX($' + r + ':$' + r + ',COLUMN()-1))'; });
  rows[13][TOTAL - 1] = '=IFERROR(' + TL + '13/' + TL + '4,"")';
  rows[15][TOTAL - 1] = '=IFERROR(' + TL + '4/' + TL + '10,"")';

  const width = TOTAL;
  rows.forEach(r => { for (let j = 0; j < width; j++) if (r[j] === undefined) r[j] = ''; });
  mis.getRange(1, 1, N, width).setValues(rows);

  // look & feel
  mis.getRange(4, FIRST, N - 3, months.length + 1).setNumberFormat('#,##0');
  mis.getRange(14, FIRST, 1, months.length + 1).setNumberFormat('0.0%');
  mis.getRange(16, FIRST, 1, months.length + 1).setNumberFormat('0.00');
  mis.getRange(2, FIRST, 1, months.length).setNumberFormat('mmm-yy');
  mis.getRange(2, 1, 1, width).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  mis.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  [4, 8, 13].forEach(r => mis.getRange(r, 1, 1, width).setFontWeight('bold'));
  [20, 37, 44].forEach(r => mis.getRange(r, 1).setFontWeight('bold'));
  mis.getRange(10, FIRST, 1, months.length).setBackground('#fff2cc');   // manual input row
  mis.setFrozenRows(2); mis.setFrozenColumns(3);
  mis.setColumnWidth(1, 210); mis.setColumnWidth(2, 330); mis.setColumnWidth(3, 120);

  return '[' + mis.getName() + '] FORMAT BANA DIYA: ' + months.length + ' mahine (' +
    (months[0].getMonth() + 1) + '/' + months[0].getFullYear() + ' se ' +
    (months[months.length - 1].getMonth() + 1) + '/' + months[months.length - 1].getFullYear() + ') DashData se bhare.\n' +
    'Peela row (Marketing Spends) manual bharna hai. COGS 30% aur Platform Margin 30% default formula hai, apne hisaab se badal lena.';
}


// ---------- CHECK: portal-wise totals of the previous month (compare with your daily report) ----------
function MIS_CHECK() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!dd) { misLog_(ss, 'ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const msg = misCheckText_(dd.getDataRange().getValues(), new Date());
  Logger.log(msg);
  misLog_(ss, msg);
  ss.toast('Check report: "MIS Log" tab dekho', 'MIS', 10);
}

function misCheckText_(values, now) {
  const first = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const y = first.getFullYear(), m = first.getMonth();
  const head = values[0].map(h => String(h).trim().toLowerCase());
  const iD = head.indexOf('date'), iP = head.indexOf('portal'), iQ = head.indexOf('qty'),
        iT = head.indexOf('total'), iI = head.indexOf('item price');
  if ([iD, iP, iQ, iT, iI].some(i => i < 0)) return 'ERROR: DashData headers nahi mile: ' + head.join(' | ');
  const by = {}, all = { rows: 0, units: 0, total: 0, item: 0, itemBlank: 0 };
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const d = row[iD] instanceof Date ? row[iD] : new Date(row[iD]);
    if (isNaN(d.getTime()) || d.getFullYear() !== y || d.getMonth() !== m) continue;
    const p = String(row[iP]).trim() || '(blank portal)';
    const o = by[p] = by[p] || { rows: 0, units: 0, total: 0, item: 0, itemBlank: 0 };
    const t = Number(row[iT]) || 0, it = Number(row[iI]) || 0, q = Number(row[iQ]) || 0;
    [o, all].forEach(x => { x.rows++; x.units += q; x.total += t; x.item += it; if (!(it > 0)) x.itemBlank++; });
  }
  const f = n => Math.round(n).toLocaleString('en-IN');
  const lines = Object.keys(by).sort((a, b) => by[b].total - by[a].total).map(p =>
    p + ' | rows ' + by[p].rows + ' | units ' + by[p].units + ' | Total ' + f(by[p].total) +
    ' | ITEM PRICE ' + f(by[p].item) + ' (blank rows ' + by[p].itemBlank + ')');
  return 'CHECK ' + (m + 1) + '/' + y + ' (DashData, portal wise)\n' + lines.join('\n') +
    '\nALL | rows ' + all.rows + ' | units ' + all.units + ' | Total ' + f(all.total) +
    ' | ITEM PRICE ' + f(all.item) + ' (blank rows ' + all.itemBlank + ')';
}


// ---------- AD SPEND LINK ----------
// SpendData = IMPORTRANGE of the Ads tracker's hidden "_Daily" tab: A Date | B Portal | C Sales | D Units | E Orders | F Ad Spend | G Ad Sales | H Impressions
function misSpendFormula_(year, month1) {
  const S = MISC.SPEND_TAB + '!';
  const d = 'DATE(' + year + ',' + month1 + ',1)';
  return '=SUMIFS(' + S + '$F:$F,' + S + '$A:$A,">="&' + d + ',' + S + '$A:$A,"<="&EOMONTH(' + d + ',0))';
}

function SETUP_SPEND() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Ads tracker ka Google Sheet link paste karo (jisme "Master Data" aur hidden "_Daily" tab hai)');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (!/\/d\/[A-Za-z0-9_-]+/.test(url)) { ui.alert('Ye Google Sheet ka link nahi lag raha. /d/.... wala poora link daalo.'); return; }
  misSetupSpend_(SpreadsheetApp.getActiveSpreadsheet(), url);
  ui.alert('SpendData tab ban gaya.\n\nZARURI: SpendData tab me A1 cell par click karo, jo "#REF!" ya "Allow access" dikhe use dabao.\n' +
    'Uske baad MIS tabs ka Marketing Spends row apne aap bhar jayega.');
}

function misSetupSpend_(ss, url) {
  let sh = ss.getSheetByName(MISC.SPEND_TAB);
  if (!sh) sh = ss.insertSheet(MISC.SPEND_TAB);
  sh.clear();
  sh.getRange('A1').setFormula('=IMPORTRANGE("' + url.replace(/"/g, '') + '","_Daily!A:H")');
  const done = [];
  MISC.TABS.forEach(t => {
    const mis = misFindSheet_(ss, t.words, false);
    if (!mis || mis.getLastColumn() < 3 || mis.getLastRow() < 5) return;
    const grid = mis.getRange(1, 1, mis.getLastRow(), 4).getValues();
    let row = 0;
    grid.forEach((rw, i) => rw.forEach(cell => { if (!row && String(cell).trim().toLowerCase() === MISC.SPEND_LABEL) row = i + 1; }));
    if (!row) { done.push(mis.getName() + ': "Marketing Spends (E-Commerce)" row nahi mili'); return; }
    const hr = misHeaderRow_(mis);
    const hdr = mis.getRange(hr, 1, 1, mis.getLastColumn()).getValues()[0];
    let n = 0;
    for (let i = 2; i < hdr.length; i++) {
      const ym = misHeaderToYm_(hdr[i]);
      if (ym) { mis.getRange(row, i + 1).setFormula(misSpendFormula_(ym.y, ym.m + 1)); n++; }
    }
    done.push(mis.getName() + ': ' + n + ' mahino ka spend link hua');
  });
  misLog_(ss, 'SPEND LINK: ' + done.join(' | '));
  return done;
}
