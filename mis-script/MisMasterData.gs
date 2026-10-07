/**
 * "MIS as per Master Data"  -  SEPARATE NEW FILE (old MIS code is not touched)
 * Add this as a new file in the MIS DEEP ANALISIS Apps Script project, Save, run  MASTER_SETUP  once.
 *
 * Builds the tab "MIS as per Master Data" from tab "SpendData" (= IMPORTRANGE of the Ads tracker's hidden _Daily).
 * Everything in the tab is a live formula, so it updates by itself when the Ads tracker updates.
 * Old months are never touched. A new month gets its own column automatically (inserted before Total).
 * Needs: tab "SpendData" with columns A Date | B Portal | C Total Sales | D Units | E Orders | F Ad Spend | G Ad Sales | H Impressions
 */
const MM = {
  TAB: 'MIS as per Master Data',
  SPEND_TAB: 'SpendData',
  LOG_TAB: 'MIS Log',
  // [label in MIS, Portal name used in the Ads tracker]
  PORTALS: [['Amazon', 'Amazon'], ['Flipkart', 'Flipkart'], ['Swiggy INSTAMRT', 'Swiggy'], ['Website', 'Website'], ['Blinkit', 'Blinkit']],
  R: { rev: 4, cogs: 6, cm1: 8, mkt: 10, plat: 11, cm2: 13, cm2p: 14, roas: 16, units: 18, orders: 19,
       salesHead: 21, spendHead: 28, adSalesHead: 35, adSalesTot: 41, tacos: 42, roasAd: 43 },
  ROWS: 44, FIRST_COL: 4, HEADER_ROW: 2
};

// run once from the function dropdown: builds the tab and sets the daily trigger
function MASTER_SETUP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const msg = mmBuild_(ss);
  mmLog_(ss, msg);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'masterDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('masterDaily').timeBased().everyDays(1).atHour(10).create();
  ss.toast('Master MIS ready. Report: "MIS Log" tab dekho', 'MIS', 10);
}

// daily trigger: only adds the running month column when a new month starts
function masterDaily() {
  mmEnsureCurrentMonth_(SpreadsheetApp.getActiveSpreadsheet());
}

function mmMonthsFrom_(values) {
  let min = null;
  for (let r = 0; r < values.length; r++) {
    const d = values[r][0];
    if (d instanceof Date && !isNaN(d.getTime()) && d.getFullYear() >= 2000 && (!min || d < min)) min = d;
  }
  const out = [];
  if (!min) return out;
  const now = new Date(), end = new Date(now.getFullYear(), now.getMonth(), 1);   // up to the running month
  for (let d = new Date(min.getFullYear(), min.getMonth(), 1); d <= end && out.length < 36; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) out.push(d);
  return out;
}

function mmBuild_(ss) {
  const sd = ss.getSheetByName(MM.SPEND_TAB);
  if (!sd) return 'ERROR: "' + MM.SPEND_TAB + '" tab nahi mila. Pehle Ads tracker ko SpendData me link karo (SETUP_SPEND).';
  const months = mmMonthsFrom_(sd.getDataRange().getValues());
  if (!months.length) return 'ERROR: SpendData me koi date nahi mili. SpendData A1 par "Allow access" dabaya? (#REF! to nahi dikh raha?)';
  let sh = ss.getSheetByName(MM.TAB);
  if (!sh) sh = ss.insertSheet(MM.TAB);
  sh.clear();

  const R = MM.R, N = MM.ROWS, FIRST = MM.FIRST_COL, TOTAL = FIRST + months.length;
  const rows = [];
  for (let i = 0; i < N; i++) rows.push([]);
  const A = (r, v) => { rows[r - 1][0] = v; }, B = (r, v) => { rows[r - 1][1] = v; }, C = (r, v) => { rows[r - 1][2] = v; };
  A(1, 'XLEAP CARE PRIVATE LIMITED  -  as per Master Data (Ads tracker)'); A(2, 'Month'); C(2, 'Master Portal');
  A(R.rev, 'Revenue'); A(R.cogs, 'COGS'); A(R.cm1, 'CM1'); A(R.mkt, 'Marketing Spends (E-Commerce)'); A(R.plat, 'Platform Margin');
  A(R.cm2, 'CM2'); A(R.cm2p, 'CM2 %'); A(R.roas, 'ROAS (Blended)'); A(R.units, 'Units Sold'); A(R.orders, 'Orders');
  A(R.salesHead, 'Channel Wise Revenue Split');
  A(R.spendHead, 'Channel Wise Ad Spend');
  A(R.adSalesHead, 'Channel Wise Ad Sales');
  MM.PORTALS.forEach((p, i) => {
    [R.salesHead, R.spendHead, R.adSalesHead].forEach(h => { B(h + i, p[0]); C(h + i, p[1]); });
  });
  A(R.adSalesTot, 'Ad Sales (attributed, excl. Website)'); A(R.tacos, 'TACOS (Ad Spend / Revenue)');
  A(R.roasAd, 'ROAS (Ad Sales / Ad Spend excl. Website)');

  const S = MM.SPEND_TAB + '!';
  const win = L => S + '$A:$A,">="&' + L + '$2,' + S + '$A:$A,"<="&EOMONTH(' + L + '$2,0)';
  const sumAll = (L, col) => '=SUMIFS(' + S + '$' + col + ':$' + col + ',' + win(L) + ')';
  const sumPortal = (L, col, r) => '=SUMIFS(' + S + '$' + col + ':$' + col + ',' + S + '$B:$B,$C' + r + ',' + win(L) + ')';
  const web = R.spendHead + 3;   // Website row inside the Ad Spend block

  months.forEach((m, k) => {
    const c = FIRST + k, L = mmColLetter_(c);
    const put = (r, v) => { rows[r - 1][c - 1] = v; };
    put(2, m);
    put(R.rev, sumAll(L, 'C'));
    put(R.cogs, '=' + L + R.rev + '*30%'); put(R.cm1, '=' + L + R.rev + '-' + L + R.cogs);
    put(R.mkt, sumAll(L, 'F')); put(R.plat, '=' + L + R.rev + '*30%');
    put(R.cm2, '=' + L + R.cm1 + '-' + L + R.mkt + '-' + L + R.plat);
    put(R.cm2p, '=IFERROR(' + L + R.cm2 + '/' + L + R.rev + ',"")');
    put(R.roas, '=IFERROR(' + L + R.rev + '/' + L + R.mkt + ',"")');
    put(R.units, sumAll(L, 'D')); put(R.orders, sumAll(L, 'E'));
    MM.PORTALS.forEach((p, i) => {
      put(R.salesHead + i, sumPortal(L, 'C', R.salesHead + i));
      put(R.spendHead + i, sumPortal(L, 'F', R.spendHead + i));
      put(R.adSalesHead + i, sumPortal(L, 'G', R.adSalesHead + i));
    });
    put(R.adSalesTot, '=SUM(' + L + R.adSalesHead + ':' + L + (R.adSalesHead + 4) + ')');
    put(R.tacos, '=IFERROR(' + L + R.mkt + '/' + L + R.rev + ',"")');
    put(R.roasAd, '=IFERROR(' + L + R.adSalesTot + '/(' + L + R.mkt + '-' + L + web + '),"")');
  });

  const TL = mmColLetter_(TOTAL);
  rows[1][TOTAL - 1] = 'Total';
  const sumRows = [R.rev, R.cogs, R.cm1, R.mkt, R.plat, R.cm2, R.units, R.orders, R.adSalesTot];
  for (let i = 0; i < 5; i++) sumRows.push(R.salesHead + i, R.spendHead + i, R.adSalesHead + i);
  sumRows.forEach(r => { rows[r - 1][TOTAL - 1] = '=SUM(D' + r + ':INDEX($' + r + ':$' + r + ',COLUMN()-1))'; });
  rows[R.cm2p - 1][TOTAL - 1] = '=IFERROR(' + TL + R.cm2 + '/' + TL + R.rev + ',"")';
  rows[R.roas - 1][TOTAL - 1] = '=IFERROR(' + TL + R.rev + '/' + TL + R.mkt + ',"")';
  rows[R.tacos - 1][TOTAL - 1] = '=IFERROR(' + TL + R.mkt + '/' + TL + R.rev + ',"")';
  rows[R.roasAd - 1][TOTAL - 1] = '=IFERROR(' + TL + R.adSalesTot + '/(' + TL + R.mkt + '-' + TL + web + '),"")';

  const width = TOTAL;
  rows.forEach(r => { for (let j = 0; j < width; j++) if (r[j] === undefined) r[j] = ''; });
  sh.getRange(1, 1, N, width).setValues(rows);

  sh.getRange(4, FIRST, N - 3, months.length + 1).setNumberFormat('#,##0');
  [R.cm2p, R.tacos].forEach(r => sh.getRange(r, FIRST, 1, months.length + 1).setNumberFormat('0.0%'));
  [R.roas, R.roasAd].forEach(r => sh.getRange(r, FIRST, 1, months.length + 1).setNumberFormat('0.00'));
  sh.getRange(2, FIRST, 1, months.length).setNumberFormat('mmm-yy');
  sh.getRange(2, 1, 1, width).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(12);
  [R.rev, R.cm1, R.cm2].forEach(r => sh.getRange(r, 1, 1, width).setFontWeight('bold'));
  [R.salesHead, R.spendHead, R.adSalesHead].forEach(r => sh.getRange(r, 1).setFontWeight('bold'));
  sh.setFrozenRows(2); sh.setFrozenColumns(3);
  sh.setColumnWidth(1, 250); sh.setColumnWidth(2, 160); sh.setColumnWidth(3, 110);

  return '[' + MM.TAB + '] BANA DIYA: ' + months.length + ' mahine (' + (months[0].getMonth() + 1) + '/' + months[0].getFullYear() +
    ' se ' + (months[months.length - 1].getMonth() + 1) + '/' + months[months.length - 1].getFullYear() + '). Sab formula hai, SpendData se live.\n' +
    'Sirf 5 portal (Amazon, Flipkart, Swiggy, Website, Blinkit) hain aur SKU / product split is data me nahi hota. ' +
    'COGS 30% aur Platform Margin 30% default formula hai.';
}

// new month -> new column before Total (copies the previous month column). Old columns are never changed.
function mmEnsureCurrentMonth_(ss) {
  const sh = ss.getSheetByName(MM.TAB);
  if (!sh || sh.getLastColumn() < 5 || sh.getLastRow() < 5) return;
  const n = new Date(), y = n.getFullYear(), m = n.getMonth(), hr = MM.HEADER_ROW;
  if (mmFindMonthCol_(sh, hr, y, m)) return;
  const hdr = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0];
  const tot = hdr.findIndex(h => String(h).trim().toLowerCase() === 'total') + 1;
  if (!tot || tot < 5) return;
  sh.insertColumnBefore(tot);
  sh.getRange(1, tot - 1, sh.getMaxRows(), 1).copyTo(sh.getRange(1, tot, sh.getMaxRows(), 1));
  sh.getRange(hr, tot).setValue(new Date(y, m, 1)).setNumberFormat('mmm-yy');
  mmLog_(ss, '[' + MM.TAB + '] naya mahina ' + (m + 1) + '/' + y + ' column ' + mmColLetter_(tot) + ' me jodd diya.');
}

function mmFindMonthCol_(sh, hr, y, m) {
  const hdr = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0];
  for (let i = 2; i < hdr.length; i++) {
    const v = hdr[i];
    if (v instanceof Date && !isNaN(v.getTime()) && v.getFullYear() === y && v.getMonth() === m) return i + 1;
  }
  return 0;
}

function mmColLetter_(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function mmLog_(ss, msg) {
  let sh = ss.getSheetByName(MM.LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(MM.LOG_TAB);
    sh.getRange(1, 1, 1, 2).setValues([['Time', 'Report']]).setFontWeight('bold');
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 700);
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 2).setValues([[new Date(), msg]]);
  sh.getRange(2, 2).setWrap(true);
  sh.getRange(2, 1, 1, 2).setVerticalAlignment('top');
}
