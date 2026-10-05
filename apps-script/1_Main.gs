/** FLAWSOME SALES DASHBOARD - file 1 of 4 (Main). Run setupDashboard() once. */
const CFG = {
  SOURCE: 'Sales', DASH: 'Dashboard', DATA: 'Data', HEADER_ROW: 1,
  HEADERS: {
    date: ['date'],
    portal: ['portal'],
    sku: ['universal sku', 'universal_sku', 'sku'],
    qty: ['qty', 'quantity'],
    rev: ['revenue', 'total', 'amount']
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Dashboard')
    .addItem('Rebuild Dashboard', 'setupDashboard').addToUi();
}

function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CFG.SOURCE);
  if (!src) throw new Error('"' + CFG.SOURCE + '" naam ki sheet nahi mili.');
  const col = findColumns_(src);
  [CFG.DASH, CFG.DATA].forEach(n => {
    const old = ss.getSheetByName(n);
    if (old) ss.deleteSheet(old);
  });
  const data = ss.insertSheet(CFG.DATA);
  const d = ss.insertSheet(CFG.DASH);
  buildData_(data, col);
  buildDash_(d);
  buildCharts_(d);
  d.setActiveSelection('B3');
  ss.setActiveSheet(d);
  ss.moveActiveSheet(1);
}

function findColumns_(sheet) {
  const heads = sheet.getRange(CFG.HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0].map(h => String(h).trim().toLowerCase());
  const out = {}, missing = [];
  Object.keys(CFG.HEADERS).forEach(key => {
    let idx = -1;
    for (const a of CFG.HEADERS[key]) { idx = heads.indexOf(a); if (idx >= 0) break; }
    if (idx < 0) missing.push(key + ' (' + CFG.HEADERS[key].join(' / ') + ')');
    else out[key] = colLetter_(idx + 1);
  });
  if (missing.length) {
    throw new Error('Sales me ye headers nahi mile: ' + missing.join(', ') +
      '\nMile huye headers: ' + heads.join(' | '));
  }
  return out;
}

function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
