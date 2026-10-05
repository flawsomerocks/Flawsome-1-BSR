/** file 5 - SKU Drop: every SKU, previous period vs selected period (biggest drop on top) */
function buildSku_(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange('A1:I1').merge().setValue('SKU WISE - DROP / GROWTH')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e78').setFontColor('#ffffff');
  sh.getRange('A2').setFormula('="Selected: "&TEXT(Dashboard!B5,"dd-mmm-yy")&" to "&TEXT(Dashboard!B6,"dd-mmm-yy")' +
    '&"   |   Previous: "&TEXT(Dashboard!H3,"dd-mmm-yy")&" to "&TEXT(Dashboard!H4,"dd-mmm-yy")' +
    '&"   |   Portal: "&Dashboard!B4&"   (change Period / Portal on Dashboard tab)"');
  sh.getRange('A4:I4').setValues([['Universal SKU', 'Prev Qty', 'Curr Qty', 'Qty Change',
    'Prev Revenue', 'Curr Revenue', 'Revenue Change', 'Change %', 'Status']])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');

  sh.getRange('A5').setFormula(
    '=IFERROR(ARRAYFORMULA(LET(' +
    's,UNIQUE(FILTER({Data!I2:I;Data!O2:O},{Data!I2:I;Data!O2:O}<>"")),' +
    'pq,SUMIF(Data!O2:O,s,Data!P2:P),cq,SUMIF(Data!I2:I,s,Data!J2:J),' +
    'pr,SUMIF(Data!O2:O,s,Data!Q2:Q),cr,SUMIF(Data!I2:I,s,Data!K2:K),' +
    'SORT({s,pq,cq,cq-pq,pr,cr,cr-pr,IF(pr=0,"New",(cr-pr)/pr)},7,TRUE))),"")');
  sh.getRange('I5').setFormula(
    '=ARRAYFORMULA(IF(G5:G="",,IF(G5:G<0,"🔻 Drop",IF(G5:G>0,"▲ Growth","—"))))');

  sh.getRange('B5:G').setNumberFormat('#,##0;-#,##0;0');
  sh.getRange('H5:H').setNumberFormat('0.0%').setHorizontalAlignment('right');
  const red = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
    .setFontColor('#c00000').setRanges([sh.getRange('D5:D'), sh.getRange('G5:H')]).build();
  const green = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setFontColor('#1e7e34').setRanges([sh.getRange('D5:D'), sh.getRange('G5:H')]).build();
  const drop = SpreadsheetApp.newConditionalFormatRule().whenTextContains('Drop')
    .setBackground('#fde2e2').setRanges([sh.getRange('I5:I')]).build();
  sh.setConditionalFormatRules([red, green, drop]);
  sh.setColumnWidth(1, 200);
  sh.setColumnWidths(2, 8, 110);
  sh.setFrozenRows(4);
}
