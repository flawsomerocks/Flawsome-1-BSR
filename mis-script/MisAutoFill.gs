/**
 * MIS AUTO FILL  (put this in the "MIS DEEP ANALISIS" sheet: Extensions > Apps Script)
 * Reads tab "DashData" (IMPORTRANGE from the sales sheet) and fills one month column of tab "MIS":
 *   Revenue (=SUM of channel rows), Units Sold, Channel Wise, Product Wise, SKU Wise.
 * Marketing Spends / Platform Margin / Opex etc. are NOT in DashData -> fill those by hand.
 * Runs by itself on the 7th of every month (previous month) after installMisTrigger().
 */
const MISC = {
  DATA: 'DashData',
  MIS: 'MIS',
  HEADER_ROW: 2,       // row with month headers in MIS
  FIRST_MONTH_COL: 3,  // column C
  REV_ROW: 4,
  UNITS_ROW: 18,
  FORMULA_ROWS: [6, 8, 11, 13, 14, 16],  // copied from previous month column when a new column is created
  // MIS channel label (col B, lowercase) -> DashData Portal names (lowercase)
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
  REST_SKU: "rest sku's (combo's)"
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 MIS')
    .addItem('Fill previous month', 'fillPreviousMonth')
    .addItem('Fill a specific month…', 'fillSpecificMonth')
    .addItem('Auto-run on 7th of every month', 'installMisTrigger')
    .addToUi();
}

function installMisTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'fillPreviousMonth') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fillPreviousMonth').timeBased().onMonthDay(7).atHour(9).create();
  try { SpreadsheetApp.getUi().alert('Done: har mahine ki 7 tarik ~9am previous month MIS me fill hoga.'); } catch (e) {}
}

function fillPreviousMonth() {
  const n = new Date();
  fillMis_(n.getFullYear(), n.getMonth() - 1);   // JS handles -1 => December of last year
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
  const dd = ss.getSheetByName(MISC.DATA), mis = ss.getSheetByName(MISC.MIS);
  if (!dd) throw new Error('"' + MISC.DATA + '" tab nahi mila.');
  if (!mis) throw new Error('"' + MISC.MIS + '" tab nahi mila (xlsx import karke MIS tab banao).');
  const first = new Date(year, monthIdx, 1);
  year = first.getFullYear(); monthIdx = first.getMonth();

  const agg = misAggregate_(dd.getDataRange().getValues(), year, monthIdx);

  // ---- locate / create the month column ----
  let c = misFindMonthCol_(mis, year, monthIdx);
  let created = false;
  if (!c) { c = misInsertMonthCol_(mis, year, monthIdx); created = true; }
  const L = colToLetter_(c);

  // ---- write rows (found by label in col B) ----
  const lastRow = mis.getLastRow();
  const labels = mis.getRange(1, 2, lastRow, 1).getValues().map(r => String(r[0]).trim().toLowerCase());
  const rowOf = {};
  labels.forEach((t, i) => { if (t && rowOf[t] === undefined) rowOf[t] = i + 1; });
  // SKU rows are matched by parsed key (labels are long product names)
  const skuRows = {};
  labels.forEach((t, i) => { const k = misNameKey_(t); if (k && !skuRows[k]) skuRows[k] = i + 1; });

  const put = (row, v) => { if (row) mis.getRange(row, c).setValue(Math.round(v * 100) / 100); };
  let chFirst = 1e9, chLast = 0;
  Object.keys(MISC.CHANNELS).forEach(k => {
    const row = rowOf[k];
    if (!row) return;
    chFirst = Math.min(chFirst, row); chLast = Math.max(chLast, row);
    put(row, agg.channels[k] || 0);
  });
  Object.keys(MISC.CATS).forEach(k => put(rowOf[MISC.CATS[k]], agg.cats[k] || 0));
  Object.keys(skuRows).forEach(k => put(skuRows[k], agg.skus[k] || 0));
  // SKUs that have no row in MIS (e.g. panty liner 60) go to the "Rest SKU" row
  Object.keys(agg.skus).forEach(k => { if (!skuRows[k]) agg.restSku += agg.skus[k]; });
  put(rowOf[MISC.REST_SKU], agg.restSku);
  put(MISC.UNITS_ROW, agg.units);
  if (chLast) mis.getRange(MISC.REV_ROW, c).setFormula('=SUM(' + L + chFirst + ':' + L + chLast + ')');

  // ---- report ----
  const msg = [
    'MIS ' + (monthIdx + 1) + '/' + year + ' -> column ' + L + (created ? ' (new column)' : ''),
    'Rows used: ' + agg.rows + ' | Revenue(DashData Total): ' + Math.round(agg.revenue) + ' | Units: ' + agg.units,
    'Rows without valid date in this period are not counted.',
    Object.keys(agg.unmappedPortals).length ? 'Portal MIS me nahi hai (revenue channel split me nahi gaya): ' +
      Object.keys(agg.unmappedPortals).map(p => p + '=' + Math.round(agg.unmappedPortals[p])).join(', ') : '',
    'Rest SKU (combos / unmatched): ' + Math.round(agg.restSku),
    'Manual bharna hai: Marketing Spends, Platform Margin, Opex.'
  ].filter(Boolean).join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { ss.toast('MIS filled: ' + L, 'MIS', 5); }
}

// ---------- pure logic (no Sheets calls) ----------
function misAggregate_(values, year, monthIdx) {
  const head = values[0].map(h => String(h).trim().toLowerCase());
  const ix = n => head.indexOf(n);
  const iDate = ix('date'), iU = ix('universal sku'), iQ = ix('qty'), iP = ix('portal'), iT = ix('total');
  if ([iDate, iU, iQ, iP, iT].some(i => i < 0)) throw new Error('DashData headers nahi mile: ' + head.join(' | '));

  const portalToCh = {};
  Object.keys(MISC.CHANNELS).forEach(k => MISC.CHANNELS[k].forEach(p => { portalToCh[p] = k; }));

  const a = { rows: 0, revenue: 0, units: 0, channels: {}, cats: {}, skus: {}, restSku: 0, unmappedPortals: {} };
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    let d = row[iDate];
    if (!(d instanceof Date)) d = new Date(d);
    if (isNaN(d.getTime()) || d.getFullYear() !== year || d.getMonth() !== monthIdx) continue;
    const rev = Number(row[iT]) || 0, qty = Number(row[iQ]) || 0;
    const portal = String(row[iP]).trim().toLowerCase();
    a.rows++; a.revenue += rev; a.units += qty;

    const ch = portalToCh[portal];
    if (ch) a.channels[ch] = (a.channels[ch] || 0) + rev;
    else a.unmappedPortals[portal] = (a.unmappedPortals[portal] || 0) + rev;

    const key = misUniKey_(row[iU]);
    const cat = misCatOf_(key);
    a.cats[cat] = (a.cats[cat] || 0) + rev;
    if (key) a.skus[key] = (a.skus[key] || 0) + rev; else a.restSku += rev;
  }
  return a;
}

// Universal SKU -> key like "pad|10|xl|cotton", "pl|30|plant", "pp|4|xl" (null = combo / unknown)
function misUniKey_(u) {
  const s = String(u).toLowerCase().replace(/\s+/g, '');
  let m = s.match(/^pantyliner(cotton|plant)0?(\d+)$/);
  if (m) return 'pl|' + m[2] + '|' + m[1];
  m = s.match(/^periodpants(xl|l)\((\d+)\)$/);
  if (m) return 'pp|' + m[2] + '|' + m[1];
  m = s.match(/^(xxl|xxi|xl|xi|l)0?(\d+)(cotton|plant)$/);
  if (m) {
    const size = (m[1] === 'xxl' || m[1] === 'xxi') ? 'xxl' : (m[1] === 'xl' || m[1] === 'xi') ? 'xl' : 'l';
    return 'pad|' + m[2] + '|' + size + '|' + m[3];
  }
  return null;
}

// MIS row label (long product name) -> same key
function misNameKey_(t) {
  const s = String(t).toLowerCase();
  if (s.indexOf('flawsome') < 0) return null;
  const pack = (s.match(/pack of (\d+)/) || [])[1];
  const size = (s.match(/size\s+(xxl|xl|l)\b/) || [])[1];
  if (s.indexOf('panty liner') >= 0) return 'pl|' + pack + '|' + (s.indexOf('plant-based') >= 0 ? 'plant' : 'cotton');
  if (s.indexOf('period pants') >= 0) return 'pp|' + pack + '|' + size;
  if (s.indexOf('organic sanitary') >= 0) return 'pad|' + pack + '|' + size + '|cotton';
  if (s.indexOf('sensitive sanitary') >= 0) return 'pad|' + pack + '|' + size + '|plant';
  return null;
}

function misCatOf_(key) {
  if (!key) return 'combos';
  if (key.indexOf('pl|') === 0) return key.endsWith('plant') ? 'plPlant' : 'plCotton';
  if (key.indexOf('pp|') === 0) return 'pants';
  return key.endsWith('plant') ? 'padPlant' : 'padCotton';
}

// ---------- sheet helpers ----------
// Header cells are dates typed like "May-26" (read as 26-May) or real month dates (1-Jun-2026)
function misHeaderToYm_(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  const day = v.getDate();
  return { y: day === 1 ? v.getFullYear() : 2000 + day, m: v.getMonth() };
}

function misFindMonthCol_(mis, y, m) {
  const hdr = mis.getRange(MISC.HEADER_ROW, 1, 1, mis.getLastColumn()).getValues()[0];
  for (let i = MISC.FIRST_MONTH_COL - 1; i < hdr.length; i++) {
    const ym = misHeaderToYm_(hdr[i]);
    if (ym && ym.y === y && ym.m === m) return i + 1;
  }
  return 0;
}

function misInsertMonthCol_(mis, y, m) {
  const hdr = mis.getRange(MISC.HEADER_ROW, 1, 1, mis.getLastColumn()).getValues()[0];
  let tot = hdr.findIndex(h => String(h).trim().toLowerCase() === 'total') + 1;
  if (!tot) tot = mis.getLastColumn() + 1;
  mis.insertColumnBefore(tot);
  const c = tot, src = c - 1;
  mis.getRange(1, src, mis.getMaxRows(), 1)
    .copyTo(mis.getRange(1, c, mis.getMaxRows(), 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  MISC.FORMULA_ROWS.forEach(r => {
    const f = mis.getRange(r, src).getFormulaR1C1();
    if (f) mis.getRange(r, c).setFormulaR1C1(f);
  });
  mis.getRange(MISC.HEADER_ROW, c).setValue(new Date(y, m, 1)).setNumberFormat('mmm-yy');
  return c;
}

function colToLetter_(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
