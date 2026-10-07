/**
 * MIS AUTO FILL  -  ONE FILE, paste this whole thing in the "MIS DEEP ANALISIS" sheet
 * (Extensions > Apps Script), Save, run  START_HERE  once.  That's it.
 *
 * Reads tab "DashData" and fills one month column in BOTH MIS tabs:
 *   "MIS as per 1 Total"      <- DashData "Total" column
 *   "MIS as per ITEM PRICE"   <- DashData "ITEM PRICE" column (blank -> Total is used)
 * Fills: Revenue (=SUM of channel rows), Units Sold, Channel / Product / SKU split.
 * NOT filled (not in DashData): Marketing Spends, Platform Margin, Opex -> manual.
 * Auto-runs on the 7th of every month (fills the previous month).
 */
const MISC = {
  DATA_TAB: 'DashData',
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
    .addItem('START HERE (fill last month + auto-run 7th)', 'START_HERE')
    .addItem('Fill previous month only', 'fillPreviousMonth')
    .addItem('Fill a specific month…', 'fillSpecificMonth')
    .addToUi();
}

// run this once: fills last month and sets the 7th-of-month trigger
function START_HERE() {
  installMisTrigger_();
  fillPreviousMonth();
}

function installMisTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fillPreviousMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fillPreviousMonth').timeBased().onMonthDay(7).atHour(9).create();
}

function fillPreviousMonth() {
  const n = new Date();
  fillMis_(n.getFullYear(), n.getMonth() - 1);
}

function fillSpecificMonth() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Month daalo (yyyy-mm), jaise 2026-09');
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim().match(/^(\d{4})-(\d{1,2})$/);
  if (!m) { ui.alert('Format galat. Example: 2026-09'); return; }
  fillMis_(Number(m[1]), Number(m[2]) - 1);
}

function fillMis_(year, monthIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const first = new Date(year, monthIdx, 1);
  year = first.getFullYear(); monthIdx = first.getMonth();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  const msgs = [];
  if (!dd) {
    msgs.push('ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.');
  } else {
    const values = dd.getDataRange().getValues();
    MISC.TABS.forEach(t => {
      const mis = misFindSheet_(ss, t.words, false);
      if (!mis) { msgs.push('ERROR: tab nahi mila: "' + t.label + '"'); return; }
      try {
        msgs.push(fillOneMis_(mis, t, misAggregate_(values, year, monthIdx, t.measure), year, monthIdx));
      } catch (e) { msgs.push('ERROR [' + t.label + ']: ' + e.message); }
    });
  }
  const msg = msgs.join('\n\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { ss.toast(msg.slice(0, 200), 'MIS', 10); }
}

function fillOneMis_(mis, tab, agg, year, monthIdx) {
  const title = '[' + mis.getName() + '] ' + (monthIdx + 1) + '/' + year;
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
    'Manual bharna hai: Marketing Spends, Platform Margin, Opex.'
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
