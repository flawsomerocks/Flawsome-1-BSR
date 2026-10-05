/** file 2 of 4 - hidden Data sheet (cleans Sales, never edits Sales) */
function buildData_(data, col) {
  const R = c => "'" + CFG.SOURCE + "'!" + c + '2:' + c;
  data.getRange('A1:E1').setValues([['Date', 'Portal', 'Universal SKU', 'Qty', 'Revenue']]);
  data.getRange('G1:K1').setValues([['F Date', 'F Portal', 'F SKU', 'F Qty', 'F Revenue']]);

  // real date or text date -> date; year before 2000 (e.g. 0206) -> blank
  data.getRange('A2').setFormula(
    '=ARRAYFORMULA(LET(x,IF(ISNUMBER(' + R(col.date) + '),' + R(col.date) +
    ',IFERROR(DATEVALUE(' + R(col.date) + '),0)),IF(x>=36526,INT(x),"")))');
  data.getRange('B2').setFormula('=ARRAYFORMULA(TRIM(' + R(col.portal) + '))');
  data.getRange('C2').setFormula('=ARRAYFORMULA(TRIM(' + R(col.sku) + '))');
  data.getRange('D2').setFormula('=ARRAYFORMULA(IFERROR(VALUE(' + R(col.qty) + '),0))');
  data.getRange('E2').setFormula('=ARRAYFORMULA(IFERROR(VALUE(' + R(col.rev) + '),0))');

  // rows inside Dashboard period + portal
  data.getRange('G2').setFormula(
    '=IFERROR(FILTER({A2:A,B2:B,C2:C,D2:D,E2:E},' +
    'A2:A<>"",A2:A>=Dashboard!$B$5,A2:A<=Dashboard!$B$6,' +
    '(Dashboard!$B$4="All")+(B2:B=Dashboard!$B$4)>0),"")');
  data.getRange('A2:A').setNumberFormat('dd-mmm-yyyy');
  data.getRange('G2:G').setNumberFormat('dd-mmm-yyyy');
  data.hideSheet();
}
