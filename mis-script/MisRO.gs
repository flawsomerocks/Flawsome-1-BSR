/**
 * "MIS as per RO"  -  SEPARATE NEW FILE (old MIS code is not touched)   [paste into INVOIVE.gs]
 * Add this as a new file in the MIS DEEP ANALISIS Apps Script project, Save, run  RO_SETUP  once.
 *
 * Fills every channel / product / SKU / Units / Marketing-spend cell of the tab "MIS as per RO" from DashData
 * (Total column), exactly like "MIS as per 1 Total", BUT the Blinkit row is NEVER changed:
 * it belongs to the Invoice tool (SUMIFS on 'Invoice Log'). Existing Blinkit formulas are put back after every run.
 * New month -> a new column is added before Total (Blinkit cell of that new column stays empty until the
 * Invoice tool syncs). Old months are never touched.
 * Needs the old MIS file (MisAutoFill.gs) in the same project: this file uses its helper functions.
 */
const RO = {
  TAB_KEY: 'misasperro',     // tab name ignoring spaces and capitals
  INVOICE_TAB_KEY: 'invoicelog',   // tab written by the Invoice tool's web app
  MEASURE: 'item price',     // the RO tab is on the ITEM PRICE basis (blank ITEM PRICE -> Total)
  SKIP_LABEL: 'blinkit',     // row that the Invoice tool owns
  BUYER_WORD: 'blink',       // invoices whose buyer contains this count as Blinkit
  INVOICE_VALUE: 'taxable',  // 'taxable' or 'total' (all your invoices have 0 GST, so same)
  HOUR: 11
};

// run once from the function dropdown: fills now + sets the daily trigger
function RO_SETUP() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'roDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('roDaily').timeBased().everyDays(1).atHour(RO.HOUR).create();
  roDaily();
}

// running month every day; previous month only until the 10th, then it is frozen
function roDaily() {
  const n = new Date();
  const list = [[n.getFullYear(), n.getMonth()]];
  if (n.getDate() <= 10) list.unshift([n.getFullYear(), n.getMonth() - 1]);
  roFill_(list);
}

function roFindTab_(ss) {
  return ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === RO.TAB_KEY) || null;
}

function roFill_(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = roFindTab_(ss);
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!tab) {
    const names = ss.getSheets().map(s => '"' + s.getName() + '"').join(', ');
    misLog_(ss, 'RO ERROR: tab "MIS as per RO" nahi mila. Tabs: ' + names);
    return;
  }
  if (!dd) { misLog_(ss, 'RO ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const values = dd.getDataRange().getValues();
  const msgs = [];
  list.forEach(ym => {
    const first = new Date(ym[0], ym[1], 1);
    const y = first.getFullYear(), m = first.getMonth();
    try { msgs.push(roFillMonth_(tab, values, y, m)); }
    catch (e) { msgs.push('RO ERROR ' + (m + 1) + '/' + y + ': ' + e.message); }
  });
  const msg = msgs.join('\n\n');
  Logger.log(msg);
  misLog_(ss, msg);
}

function roFillMonth_(tab, values, y, m) {
  const agg = misAggregate_(values, y, m, RO.MEASURE);
  if (!agg.rows || tab.getLastColumn() < 3 || tab.getLastRow() < 5) {
    return '[MIS as per RO] ' + (m + 1) + '/' + y + ': DashData me data nahi / tab khaali. Kuch nahi likha.';
  }
  const hr = misHeaderRow_(tab);
  // find the Blinkit row (label in columns A-D)
  const grid = tab.getRange(1, 1, tab.getLastRow(), 4).getValues();
  let row = 0;
  grid.forEach((r, i) => r.forEach(c => { if (!row && String(c).trim().toLowerCase() === RO.SKIP_LABEL) row = i + 1; }));

  // remember what the Invoice tool wrote in this month's Blinkit cell
  const col0 = misFindMonthCol_(tab, hr, y, m);
  let keep = null;
  if (row && col0) {
    const cell = tab.getRange(row, col0);
    keep = { f: cell.getFormula(), v: cell.getValue() };
  }

  // product / SKU split must add up to the channel total, which holds the Blinkit INVOICE value:
  // share that invoice value over Blinkit's own products / SKUs (in the ratio of its DashData sales)
  let V = 0;
  if (keep && (keep.f || (keep.v !== '' && keep.v !== null))) V = Number(keep.v) || 0;
  else { const inf = roInvoiceInfo_(tab.getParent()); V = inf.err ? 0 : (inf.expected[y + '-' + m] || 0); }
  const alloc = roAllocateBlinkit_(values, y, m, agg, V);

  const info = fillOneMis_(tab, { label: 'MIS as per RO', measure: RO.MEASURE }, agg, y, m);

  // put the Blinkit cell back
  const col1 = misFindMonthCol_(tab, hr, y, m);
  let note = '';
  if (row && col1) {
    const cell = tab.getRange(row, col1);
    const had = keep && (keep.f || (keep.v !== '' && keep.v !== null));
    if (had) {
      if (keep.f) cell.setFormula(keep.f); else cell.setValue(keep.v);
      note = 'Blinkit row jyon ki tyon (invoice tool ki).';
    } else {
      cell.clearContent();
      note = roFillBlinkitFromInvoices_(tab, cell, y, m);
    }
  } else if (!row) {
    note = 'WARNING: "Blinkit" row nahi mili.';
  }
  return info + '\n' + alloc + '\n' + note;
}

// moves the Blinkit part of agg (products / SKUs) from the DashData value to the invoice value V
function roAllocateBlinkit_(values, y, m, agg, V) {
  const head = values[0].map(h => String(h).trim().toLowerCase());
  const iP = head.indexOf('portal');
  const only = [values[0]].concat(values.slice(1).filter(r => String(r[iP]).trim().toLowerCase() === RO.SKIP_LABEL));
  const b = misAggregate_(only, y, m, RO.MEASURE);
  if (!(b.revenue > 0)) {
    if (V > 0) { agg.restSku += V; agg.cats.combos = (agg.cats.combos || 0) + V; }
    return 'Split: Blinkit DashData me nahi, invoice ' + Math.round(V) + ' Rest SKU / Combos me gaya.';
  }
  const k = V / b.revenue - 1;
  Object.keys(b.skus).forEach(s => { agg.skus[s] = (agg.skus[s] || 0) + b.skus[s] * k; });
  Object.keys(b.cats).forEach(c => { agg.cats[c] = (agg.cats[c] || 0) + b.cats[c] * k; });
  agg.restSku += b.restSku * k;
  return 'Split: Blinkit DashData ' + Math.round(b.revenue) + ' -> invoice ' + Math.round(V) + ' (product/SKU me usi ratio se baanta).';
}


// ---------- Blinkit cell from the "Invoice Log" tab (only when the cell is empty) ----------
function roNorm_(v) { return String(v === null || v === undefined ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function roIsDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }

// reads Invoice Log, detects columns by header name, returns what is needed to build + verify a formula
function roInvoiceInfo_(ss) {
  const sh = ss.getSheets().find(s => roNorm_(s.getName()) === RO.INVOICE_TAB_KEY);
  if (!sh) return { err: 'Invoice Log tab nahi mila' };
  const vals = sh.getDataRange().getValues();
  let hr = -1;
  for (let i = 0; i < Math.min(vals.length, 6); i++) {
    if (vals[i].some(c => roNorm_(c).indexOf('status') >= 0)) { hr = i; break; }
  }
  if (hr < 0) return { err: 'Invoice Log me "Status" header nahi mila' };
  const head = vals[hr].map(roNorm_);
  const find = fn => head.findIndex(fn);
  const cStatus = find(h => h.indexOf('status') >= 0);
  const cBuyer = find(h => /buyer|customer|party|billto/.test(h));
  let cValue = RO.INVOICE_VALUE === 'taxable' ? find(h => h.indexOf('taxable') >= 0) : -1;
  if (cValue < 0) cValue = find(h => h === 'total' || h === 'invoicetotal' || h === 'totalamount' || h === 'amount');
  const cDate = find(h => h === 'invoicedate' || h === 'date');
  const cKey = find(h => h === 'monthkey');
  const cMonth = find(h => h.indexOf('month') >= 0 && h !== 'monthkey');
  if (cBuyer < 0 || cValue < 0 || (cDate < 0 && cMonth < 0 && cKey < 0)) {
    return { err: 'Invoice Log ke headers samajh nahi aaye: ' + vals[hr].join(' | ') };
  }
  const data = vals.slice(hr + 1).filter(r => r.some(c => c !== '' && c !== null));

  // which column can be used for the month, and what kind of values does it hold?
  let col = -1, kind = '';
  const samples = [];
  [cKey, cDate, cMonth].filter(c => c >= 0).forEach(c => {
    if (col >= 0) return;
    const cells = data.map(r => r[c]).filter(v => v !== '' && v !== null);
    if (!cells.length) return;
    samples.push(vals[hr][c] + ': ' + cells.slice(0, 3).map(v => (roIsDate_(v) ? 'DATE ' : typeof v + ' ') + String(v)).join(' | '));
    if (cells.every(v => /^\d{6}$/.test(String(v).trim()) && +String(v).trim().slice(4) >= 1 && +String(v).trim().slice(4) <= 12)) { col = c; kind = 'key6'; }
    else if (cells.every(roIsDate_)) { col = c; kind = 'date'; }
    else if (cells.every(v => typeof v === 'string' && /^\d{4}-\d{2}(-\d{2})?/.test(v.trim()))) { col = c; kind = 'iso'; }
    else if (cells.every(v => typeof v === 'string' && /^[A-Za-z]{3,}[ -]\d{2,4}$/.test(v.trim()))) { col = c; kind = 'monyy'; }
  });
  if (col < 0) return { err: 'Invoice Log me date / month key ka format samajh nahi aaya. Namune -> ' + samples.join(' || ') };

  const expected = {}, label = {};
  let activeText = 'Active', unparsed = '';
  data.forEach(r => {
    const st = String(r[cStatus]).trim();
    if (roNorm_(st).indexOf('active') !== 0) return;
    activeText = st;
    if (roNorm_(r[cBuyer]).indexOf(RO.BUYER_WORD) < 0) return;
    const ym = roYm_(r[col]);
    if (!ym) { unparsed = unparsed || String(r[col]); return; }
    const k = ym.y + '-' + ym.m;
    expected[k] = (expected[k] || 0) + (Number(r[cValue]) || 0);
    if (!label[k]) label[k] = String(r[col]).trim();
  });
  if (unparsed) return { err: 'Invoice Log me ek tarikh padh nahi payi: "' + unparsed + '"' };

  const L = n => { let s = ''; while (n > 0) { const x = (n - 1) % 26; s = String.fromCharCode(65 + x) + s; n = Math.floor((n - 1) / 26); } return s; };
  const rng = c => "'" + sh.getName() + "'!$" + L(c + 1) + ':$' + L(c + 1);
  const common = ',' + rng(cStatus) + ',"' + activeText.replace(/"/g, '') + '",' + rng(cBuyer) + ',"*' + RO.BUYER_WORD.toUpperCase() + '*")';
  const formulaFor = (y, m) => {
    const d = 'DATE(' + y + ',' + (m + 1) + ',1)';
    if (kind === 'key6') {
      return '=SUMIFS(' + rng(cValue) + ',' + rng(col) + ',' + (y * 100 + m + 1) + common;
    }
    if (kind === 'date') {
      return '=SUMIFS(' + rng(cValue) + ',' + rng(col) + ',">="&' + d + ',' + rng(col) + ',"<="&EOMONTH(' + d + ',0)' + common;
    }
    if (kind === 'iso') {
      return '=SUMIFS(' + rng(cValue) + ',' + rng(col) + ',"' + y + '-' + ('0' + (m + 1)).slice(-2) + '*"' + common;
    }
    const lb = label[y + '-' + m];
    return lb ? '=SUMIFS(' + rng(cValue) + ',' + rng(col) + ',"' + lb.replace(/"/g, '') + '"' + common : null;
  };
  return { expected: expected, formulaFor: formulaFor, kind: kind };
}

// year + month (0-based) from a Date, an ISO text, dd/mm/yyyy text or "Sep-26" style text
function roYm_(v) {
  if (roIsDate_(v)) return { y: v.getFullYear(), m: v.getMonth() };
  const s = String(v).trim();
  if (/^\d{6}$/.test(s) && +s.slice(4) >= 1 && +s.slice(4) <= 12) return { y: +s.slice(0, 4), m: +s.slice(4) - 1 };
  let x = s.match(/^(\d{4})[-\/.](\d{1,2})/);
  if (x && +x[2] >= 1 && +x[2] <= 12) return { y: +x[1], m: +x[2] - 1 };
  x = s.match(/^(\d{1,2})[-\/. ](\d{1,2})[-\/. ](\d{4})/);
  if (x && +x[2] >= 1 && +x[2] <= 12) return { y: +x[3], m: +x[2] - 1 };
  const MON = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  x = s.match(/^(?:\d{1,2}[ -])?([A-Za-z]{3,})[ ,-]+(\d{2,4})$/);
  if (x) { const mi = MON.indexOf(x[1].slice(0, 3).toLowerCase()); if (mi >= 0) return { y: +x[2] < 100 ? 2000 + +x[2] : +x[2], m: mi }; }
  return null;
}

// one click: shows what the Invoice Log really looks like (open "MIS Log" tab afterwards)
function RO_INVOICE_DEBUG() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().find(s => roNorm_(s.getName()) === RO.INVOICE_TAB_KEY);
  if (!sh) { misLog_(ss, 'RO DEBUG: Invoice Log tab nahi mila'); return; }
  const v = sh.getRange(1, 1, Math.min(sh.getLastRow(), 4), sh.getLastColumn()).getValues();
  const t = x => (roIsDate_(x) ? 'DATE ' : typeof x + ' ') + String(x);
  misLog_(ss, 'RO DEBUG Invoice Log:\n' + v.map((r, i) => 'row ' + (i + 1) + ': ' + r.map(t).join(' | ')).join('\n'));
}

// writes the formula, checks that the sheet's own answer equals the script's sum, otherwise removes it
function roFillBlinkitFromInvoices_(tab, cell, y, m) {
  const info = roInvoiceInfo_(tab.getParent());
  if (info.err) return 'Blinkit cell khaali: ' + info.err + '. (Invoice tool me "Sync now" dabao.)';
  const exp = info.expected[y + '-' + m] || 0;
  if (!(exp > 0)) return 'Blinkit cell khaali: is mahine ka koi active Blinkit invoice nahi.';
  const f = info.formulaFor(y, m);
  if (!f) return 'Blinkit cell khaali: is mahine ka formula nahi ban paya. (Invoice tool me "Sync now" dabao.)';
  cell.setFormula(f);
  SpreadsheetApp.flush();
  const got = Number(cell.getValue());
  if (isFinite(got) && Math.abs(got - exp) < 1) {
    return 'Blinkit: Invoice Log se formula likha aur check kiya (' + Math.round(got) + ' = invoices ka total).';
  }
  cell.clearContent();
  return 'Blinkit cell khaali chhoda: formula ka jawab (' + got + ') invoices ke total (' + Math.round(exp) + ') se match nahi hua. Invoice tool me "Sync now" dabao.';
}
