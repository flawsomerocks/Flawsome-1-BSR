/**
 * FLAWSOME SALES DASHBOARD  (one-time setup, then fully automatic)
 *
 * Source sheet : "Sales"  (header in row 1)
 * Needed headers (any order, case-insensitive): Date, Universal SKU, Qty, Portal, Revenue (or Total)
 *
 * How it works
 *  - Sales sheet is NEVER modified.
 *  - A hidden "Data" sheet cleans Date / Qty / Revenue with array formulas
 *    (handles real dates AND text dates, ignores junk years like 0206).
 *  - "Dashboard" only reads the dropdowns (Period / Portal / Custom dates)
 *    and everything else updates by itself when new rows are added to Sales.
 *
 * Use: paste in Extensions > Apps Script, run setupDashboard() once.
 * Afterwards use the menu  "📊 Dashboard > Rebuild Dashboard" only if you break something.
 */

const CFG = {
  SOURCE: 'Sales',
  DASH: 'Dashboard',
  DATA: 'Data',
  HEADER_ROW: 1,
  // header aliases -> first match wins
  HEADERS: {
    date: ['date'],
    portal: ['portal'],
    sku: ['universal sku', 'universal_sku', 'sku'],
    qty: ['qty', 'quantity'],
    rev: ['revenue', 'total', 'amount']
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Dashboard')
    .addItem('Rebuild Dashboard', 'setupDashboard')
    .addToUi();
}

function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CFG.SOURCE);
  if (!src) throw new Error('"' + CFG.SOURCE + '" naam ki sheet nahi mili.');

  const col = findColumns_(src);            // {date:'B', portal:'H', ...}
  const S = "'" + CFG.SOURCE + "'!";
  const R = c => S + c + '2:' + c;          // open-ended range, e.g. 'Sales'!B2:B

  // ---------- fresh sheets ----------
  [CFG.DASH, CFG.DATA].forEach(n => {
    const old = ss.getSheetByName(n);
    if (old) ss.deleteSheet(old);
  });
  const data = ss.insertSheet(CFG.DATA);
  const d = ss.insertSheet(CFG.DASH);

  // ---------- DATA sheet (cleaned copy, all formulas) ----------
  data.getRange('A1:E1').setValues([['Date', 'Portal', 'Universal SKU', 'Qty', 'Revenue']]);
  data.getRange('G1:K1').setValues([['F Date', 'F Portal', 'F Universal SKU', 'F Qty', 'F Revenue']]);

  // Date: real date or text date -> date. Anything before year 2000 (e.g. 0206) => blank.
  data.getRange('A2').setFormula(
    '=ARRAYFORMULA(LET(x,IF(ISNUMBER(' + R(col.date) + '),' + R(col.date) +
    ',IFERROR(DATEVALUE(' + R(col.date) + '),0)),IF(x>=36526,INT(x),"")))');
  data.getRange('B2').setFormula('=ARRAYFORMULA(TRIM(' + R(col.portal) + '))');
  data.getRange('C2').setFormula('=ARRAYFORMULA(TRIM(' + R(col.sku) + '))');
  data.getRange('D2').setFormula('=ARRAYFORMULA(IFERROR(VALUE(' + R(col.qty) + '),0))');
  data.getRange('E2').setFormula('=ARRAYFORMULA(IFERROR(VALUE(' + R(col.rev) + '),0))');

  // Filtered rows (by Dashboard period + portal)
  data.getRange('G2').setFormula(
    '=IFERROR(FILTER({A2:A,B2:B,C2:C,D2:D,E2:E},' +
    'A2:A<>"",A2:A>=Dashboard!$B$5,A2:A<=Dashboard!$B$6,' +
    '(Dashboard!$B$4="All")+(B2:B=Dashboard!$B$4)>0),"")');
  data.getRange('A2:A').setNumberFormat('dd-mmm-yyyy');
  data.getRange('G2:G').setNumberFormat('dd-mmm-yyyy');
  data.hideSheet();

  // ---------- DASHBOARD ----------
  d.setHiddenGridlines(true);
  d.getRange('A1:K1').merge()
    .setValue('FLAWSOME SALES DASHBOARD')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');

  d.getRange('A3:A6').setValues([['Period'], ['Portal'], ['Start Date'], ['End Date']]).setFontWeight('bold');
  d.getRange('B3').setValue('MTD');
  d.getRange('B4').setValue('All');

  d.getRange('D3:D5').setValues([['Custom Start'], ['Custom End'], ['Latest Data Date']]).setFontWeight('bold');
  d.getRange('E3').setFormula('=E5');   // starting suggestion; overwrite when using "Custom"
  d.getRange('E4').setFormula('=E5');
  d.getRange('E5').setFormula('=IFERROR(MAX(Data!A2:A),"")');

  d.getRange('B5').setFormula(
    '=IFERROR(SWITCH($B$3,' +
    '"Last Day",$E$5,' +
    '"Last 7 Days",$E$5-6,' +
    '"Last 30 Days",$E$5-29,' +
    '"MTD",DATE(YEAR($E$5),MONTH($E$5),1),' +
    '"Last Month",DATE(YEAR($E$5),MONTH($E$5)-1,1),' +
    '"All Time",MIN(Data!A2:A),' +
    '"Custom",$E$3,' +
    '$E$5),"")');
  d.getRange('B6').setFormula(
    '=IFERROR(SWITCH($B$3,"Last Month",EOMONTH($E$5,-1),"Custom",$E$4,$E$5),"")');

  // portal list (hidden helper)
  d.getRange('Z1').setFormula(
    '={"All";IFERROR(SORT(UNIQUE(FILTER(Data!B2:B,Data!B2:B<>""))),"")}');
  d.hideColumns(26);

  d.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Last Day', 'Last 7 Days', 'Last 30 Days', 'MTD', 'Last Month', 'All Time', 'Custom'], true)
    .setAllowInvalid(false).build());
  d.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(d.getRange('Z1:Z100'), true).setAllowInvalid(false).build());
  d.getRange('B3:B4').setBackground('#fff2cc');
  d.getRange('E3:E4').setBackground('#fff2cc');
  d.getRange('B5:B6').setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('left');
  d.getRange('E3:E5').setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('left');
  d.getRange('B3:B4').setHorizontalAlignment('left');

  // KPI cards
  d.getRange('A8:F8')
    .setValues([['Total Revenue', 'Total Qty', 'Avg per Qty', 'Last Day Revenue', 'Last Day Qty', 'Rows without Date']])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');

  const lastDay = col2 =>
    '=IFERROR(SUM(FILTER(Data!' + col2 + '2:' + col2 + ',Data!A2:A=$E$5,' +
    '($B$4="All")+(Data!B2:B=$B$4)>0)),0)';

  d.getRange('A9').setFormula('=SUM(Data!K2:K)');
  d.getRange('B9').setFormula('=SUM(Data!J2:J)');
  d.getRange('C9').setFormula('=IFERROR(A9/B9,0)');
  d.getRange('D9').setFormula(lastDay('E'));
  d.getRange('E9').setFormula(lastDay('D'));
  d.getRange('F9').setFormula('=SUMPRODUCT((Data!A2:A="")*(Data!E2:E<>0))');
  d.getRange('A9:F9').setFontSize(16).setFontWeight('bold').setNumberFormat('#,##0').setHorizontalAlignment('right');

  // Tables
  const Q = 'Data!G2:K';
  d.getRange('A11').setValue('PORTAL WISE').setFontWeight('bold');
  d.getRange('E11').setValue('DATE WISE').setFontWeight('bold');
  d.getRange('I11').setValue('TOP 10 SKU').setFontWeight('bold');

  d.getRange('A12').setFormula(
    '=IFERROR(QUERY(' + Q + ',"select Col2, sum(Col4), sum(Col5) where Col1 is not null ' +
    'group by Col2 order by sum(Col5) desc label Col2 \'Portal\', sum(Col4) \'Qty\', sum(Col5) \'Revenue\'",0),"")');
  d.getRange('E12').setFormula(
    '=IFERROR(QUERY(' + Q + ',"select Col1, sum(Col4), sum(Col5) where Col1 is not null ' +
    'group by Col1 order by Col1 label Col1 \'Date\', sum(Col4) \'Qty\', sum(Col5) \'Revenue\'",0),"")');
  d.getRange('I12').setFormula(
    '=IFERROR(QUERY(' + Q + ',"select Col3, sum(Col4), sum(Col5) where Col1 is not null and Col3 <> \'\' ' +
    'group by Col3 order by sum(Col5) desc limit 10 label Col3 \'Universal SKU\', sum(Col4) \'Qty\', sum(Col5) \'Revenue\'",0),"")');

  d.getRange('E13:E').setNumberFormat('dd-mmm-yy').setHorizontalAlignment('left');
  d.getRangeList(['B13:C', 'F13:G', 'J13:K']).setNumberFormat('#,##0');
  d.getRangeList(['A12:C12', 'E12:G12', 'I12:K12']).setFontWeight('bold').setBackground('#d9eaf7');

  // Charts (placed right of the tables)
  d.insertChart(d.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(d.getRange('E12:E131')).addRange(d.getRange('G12:G131'))
    .setNumHeaders(1)
    .setOption('title', 'Daily Revenue').setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 260)
    .setPosition(3, 13, 0, 0).build());

  d.insertChart(d.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(d.getRange('A12:A31')).addRange(d.getRange('C12:C31'))
    .setNumHeaders(1)
    .setOption('title', 'Portal-wise Revenue')
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 520).setOption('height', 280)
    .setPosition(17, 13, 0, 0).build());

  d.insertChart(d.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(d.getRange('I12:I22')).addRange(d.getRange('K12:K22'))
    .setNumHeaders(1)
    .setOption('title', 'Top 10 SKU (Revenue)').setOption('legend', { position: 'none' })
    .setOption('width', 520).setOption('height', 300)
    .setPosition(33, 13, 0, 0).build());

  d.setColumnWidths(1, 11, 125);
  d.setFrozenRows(1);
  d.setActiveSelection('B3');
  ss.setActiveSheet(d);
  ss.moveActiveSheet(1);
}

/** Maps required fields to column letters using the Sales header row. */
function findColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const heads = sheet.getRange(CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim().toLowerCase());
  const out = {};
  const missing = [];
  Object.keys(CFG.HEADERS).forEach(key => {
    let idx = -1;
    for (const alias of CFG.HEADERS[key]) {
      idx = heads.indexOf(alias);
      if (idx >= 0) break;
    }
    if (idx < 0) missing.push(key + ' (' + CFG.HEADERS[key].join(' / ') + ')');
    else out[key] = colLetter_(idx + 1);
  });
  if (missing.length) {
    throw new Error('Sales sheet me ye headers nahi mile: ' + missing.join(', ') +
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
