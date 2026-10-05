/** file 3 of 4 - Dashboard controls, KPIs and tables */
function buildDash_(d) {
  d.setHiddenGridlines(true);
  d.getRange('A1:K1').merge().setValue('FLAWSOME SALES DASHBOARD')
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');

  d.getRange('A3:A6').setValues([['Period'], ['Portal'], ['Start Date'], ['End Date']]).setFontWeight('bold');
  d.getRange('B3').setValue('MTD');
  d.getRange('B4').setValue('All');
  d.getRange('D3:D5').setValues([['Custom Start'], ['Custom End'], ['Latest Data Date']]).setFontWeight('bold');
  d.getRange('E3').setFormula('=E5');
  d.getRange('E4').setFormula('=E5');
  d.getRange('E5').setFormula('=IFERROR(MAX(Data!A2:A),"")');
  d.getRange('B5').setFormula('=IFERROR(SWITCH($B$3,"Last Day",$E$5,"Last 7 Days",$E$5-6,' +
    '"Last 30 Days",$E$5-29,"MTD",DATE(YEAR($E$5),MONTH($E$5),1),' +
    '"Last Month",DATE(YEAR($E$5),MONTH($E$5)-1,1),"All Time",MIN(Data!A2:A),' +
    '"Custom",$E$3,$E$5),"")');
  d.getRange('B6').setFormula('=IFERROR(SWITCH($B$3,"Last Month",EOMONTH($E$5,-1),"Custom",$E$4,$E$5),"")');

  d.getRange('G3:G4').setValues([['Prev Start'], ['Prev End']]).setFontWeight('bold');
  d.getRange('H3').setFormula('=IFERROR($B$5-($B$6-$B$5+1),"")');
  d.getRange('H4').setFormula('=IFERROR($B$5-1,"")');
  d.getRange('H3:H4').setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('left');

  d.getRange('Z1').setFormula('={"All";IFERROR(SORT(UNIQUE(FILTER(Data!B2:B,Data!B2:B<>""))),"")}');
  d.hideColumns(26);
  d.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Last Day', 'Last 7 Days', 'Last 30 Days', 'MTD', 'Last Month', 'All Time', 'Custom'], true)
    .setAllowInvalid(false).build());
  d.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(d.getRange('Z1:Z100'), true).setAllowInvalid(false).build());
  d.getRange('B3:B4').setBackground('#fff2cc').setHorizontalAlignment('left');
  d.getRange('E3:E4').setBackground('#fff2cc');
  d.getRange('B5:B6').setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('left');
  d.getRange('E3:E5').setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('left');

  d.getRange('A8:F8').setValues([['Total Revenue', 'Total Qty', 'Avg per Qty',
    'Last Day Revenue', 'Last Day Qty', 'Rows without Date']])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');
  const lastDay = c => '=IFERROR(SUM(FILTER(Data!' + c + '2:' + c + ',Data!A2:A=$E$5,' +
    '($B$4="All")+(Data!B2:B=$B$4)>0)),0)';
  d.getRange('A9').setFormula('=SUM(Data!K2:K)');
  d.getRange('B9').setFormula('=SUM(Data!J2:J)');
  d.getRange('C9').setFormula('=IFERROR(A9/B9,0)');
  d.getRange('D9').setFormula(lastDay('E'));
  d.getRange('E9').setFormula(lastDay('D'));
  d.getRange('F9').setFormula('=SUMPRODUCT((Data!A2:A="")*(Data!E2:E<>0))');
  d.getRange('A9:F9').setFontSize(16).setFontWeight('bold').setNumberFormat('#,##0').setHorizontalAlignment('right');

  const Q = 'Data!G2:K';
  const q = (sel, extra, order, lbl) => '=IFERROR(QUERY(' + Q + ',"select ' + sel +
    ' where Col1 is not null' + extra + ' ' + order + ' label ' + lbl + '",0),"")';
  d.getRange('A11').setValue('PORTAL WISE').setFontWeight('bold');
  d.getRange('E11').setValue('TOP 10 SKU').setFontWeight('bold');
  d.getRange('M51').setValue('DATE WISE').setFontWeight('bold');
  d.getRange('A12').setFormula(q('Col2, sum(Col4), sum(Col5)', ' group by Col2',
    'order by sum(Col5) desc', "Col2 'Portal', sum(Col4) 'Qty', sum(Col5) 'Revenue'"));
  d.getRange('E12').setFormula(q('Col3, sum(Col4), sum(Col5)', " and Col3 <> '' group by Col3",
    'order by sum(Col5) desc limit 10', "Col3 'Universal SKU', sum(Col4) 'Qty', sum(Col5) 'Revenue'"));
  d.getRange('M52').setFormula(q('Col1, sum(Col4), sum(Col5)', ' group by Col1',
    'order by Col1', "Col1 'Date', sum(Col4) 'Qty', sum(Col5) 'Revenue'"));

  d.getRange('M53:M').setNumberFormat('dd-mmm-yy').setHorizontalAlignment('left');
  d.getRangeList(['B13:C', 'F13:G', 'N53:O']).setNumberFormat('#,##0');
  d.getRangeList(['A12:C12', 'E12:G12', 'M52:O52']).setFontWeight('bold').setBackground('#d9eaf7');
  buildSkuDrop_(d);
  d.setColumnWidths(1, 11, 125);
  d.setColumnWidth(1, 190);
  d.setColumnWidth(5, 190);
  d.setFrozenRows(1);
}
