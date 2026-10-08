// Fills the "DashData" sheet straight from every portal tab. Never touches "Compile".
const DD = {
  OUT: 'DashData',
  HEAD: ['Month', 'Date', 'Portal sku', 'Universal SKU', 'Qty', 'SKU', 'Asp', 'Portal', 'Total', 'ITEM PRICE'],
  // default column positions (0 = A). Override per tab with m:{...} if a tab is different.
  DEF: { month: 0, date: 1, psku: 2, usku: 3, qty: 4, sku: 5, asp: 6, total: 8, item: 9 },
  SHEETS: [
    { tab: 'Jio mart(d-1)', label: 'Jio Mart' },
    { tab: 'Cred', label: 'Cred' },
    { tab: 'Apollo(d-1)', label: 'Apollo' },
    { tab: 'WH Smith ', label: 'WH Smith' },
    { tab: 'Firstcry', label: 'Firstcry' },
    { tab: '1MG(d-1)', label: '1MG' },
    { tab: 'smytten(d-1)', label: 'Smytten' },
    { tab: 'Pharmeasy(d-1)', label: 'Pharmeasy' },
    { tab: 'Meesho(d-1)', label: 'Meesho', m: { total: 6, item: 6 } },
    { tab: 'Tatacliq', label: 'Tatacliq' },
    { tab: 'POP', label: 'POP', m: { total: 6, item: 6 } },
    { tab: 'Amazon Vendor central(d-2)', label: 'Amazon Vendor' },
    { tab: 'Amazon Dropship(d-1)', label: 'Amazon Dropship', m: { asp: 12, total: 15, item: 15 } },
    { tab: 'FBA AMZ', label: 'AMZ FBA', m: { total: 10, item: 10 } },
    { tab: 'Flipkart(d-1)', label: 'Flipkart' },
    { tab: 'Snapdeal(d-1)', label: 'Snapdeal' },
    { tab: 'Blinkit', label: 'Blinkit' },
    { tab: 'Myntra(d-1)', label: 'Myntra' },
    { tab: 'Website', label: 'Website' },
    { tab: 'Swiggy', label: 'Swiggy' }
  ]
};

function buildDashData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = ss.getSheetByName(DD.OUT) || ss.insertSheet(DD.OUT);
  const rows = [], report = [];

  DD.SHEETS.forEach(cfg => {
    const sh = ss.getSheetByName(cfg.tab);
    if (!sh) { report.push(cfg.tab + ': SHEET NOT FOUND'); return; }
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) { report.push(cfg.label + ': 0 rows'); return; }
    const m = Object.assign({}, DD.DEF, cfg.m || {});
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let n = 0, noDate = 0, sum = 0;
    data.forEach(r => {
      const usku = ddGet_(r, m.usku), psku = ddGet_(r, m.psku);
      if (usku === '' && psku === '') return;
      const date = ddDate_(ddGet_(r, m.date));
      const month = date ? new Date(date.getFullYear(), date.getMonth(), 1) : '';
      const total = ddNum_(ddGet_(r, m.total));
      if (!date) noDate++;
      sum += Number(total) || 0;
      n++;
      rows.push([month, date, psku, usku, ddNum_(ddGet_(r, m.qty)), ddGet_(r, m.sku),
        ddNum_(ddGet_(r, m.asp)), cfg.label, total, ddNum_(ddGet_(r, m.item))]);
    });
    report.push(cfg.label + ': ' + n + ' rows, total ' + Math.round(sum) + ', without date ' + noDate);
  });

  rows.sort((a, b) => (a[1] ? a[1].getTime() : 0) - (b[1] ? b[1].getTime() : 0));

  out.clearContents();
  out.getRange(1, 1, 1, DD.HEAD.length).setValues([DD.HEAD]).setFontWeight('bold')
    .setBackground('#434343').setFontColor('#ffffff');
  if (rows.length) {
    out.getRange(2, 1, rows.length, DD.HEAD.length).setValues(rows);
    out.getRange(2, 1, rows.length, 1).setNumberFormat('mmm-yy');
    out.getRange(2, 2, rows.length, 1).setNumberFormat('dd-mmm-yyyy');
  }
  out.setFrozenRows(1);
  Logger.log(report.join('\n') + '\nTOTAL ' + rows.length + ' rows');
  ss.toast(rows.length + ' rows -> DashData. Details: View > Logs / Execution log', 'DashData', 8);
}

// optional: run buildDashData automatically every day (~7am)
function installDashDataTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'buildDashData') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildDashData').timeBased().everyDays(1).atHour(7).create();
}

function ddGet_(row, i) {
  const v = row[i];
  return (v === undefined || v === null) ? '' : (typeof v === 'string' ? v.trim() : v);
}

function ddNum_(v) {
  if (typeof v === 'number') return v;
  if (typeof v !== 'string' || v === '') return '';
  const n = parseFloat(v.replace(/[₹,\s]/g, ''));
  return isNaN(n) ? '' : n;
}

function ddDate_(v) {
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'string' && v !== '') d = new Date(v);
  if (!d || isNaN(d.getTime()) || d.getFullYear() < 2000) return '';
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
