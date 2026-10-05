/** file 5 - SKU drop / growth block on the Dashboard (row 25 onwards) */
function buildSkuDrop_(d) {
  d.getRange('A25').setValue('SKU WISE - DROP / GROWTH (selected period vs previous period)').setFontWeight('bold');
  d.getRange('A26').setFormula('="Selected: "&TEXT(B5,"dd-mmm-yy")&" to "&TEXT(B6,"dd-mmm-yy")' +
    '&"   |   Previous: "&TEXT(H3,"dd-mmm-yy")&" to "&TEXT(H4,"dd-mmm-yy")');
  d.getRange('A27:I27').setValues([['Universal SKU', 'Prev Qty', 'Curr Qty', 'Qty Change',
    'Prev Revenue', 'Curr Revenue', 'Revenue Change', 'Change %', 'Status']])
    .setFontWeight('bold').setBackground('#434343').setFontColor('#ffffff');

  d.getRange('A28').setFormula(
    '=IFERROR(ARRAYFORMULA(LET(' +
    's,UNIQUE(FILTER({Data!I2:I;Data!O2:O},{Data!I2:I;Data!O2:O}<>"")),' +
    'pq,SUMIF(Data!O2:O,s,Data!P2:P),cq,SUMIF(Data!I2:I,s,Data!J2:J),' +
    'pr,SUMIF(Data!O2:O,s,Data!Q2:Q),cr,SUMIF(Data!I2:I,s,Data!K2:K),' +
    'SORT({s,pq,cq,cq-pq,pr,cr,cr-pr,IF(pr=0,"New",(cr-pr)/pr)},7,TRUE))),"")');
  d.getRange('I28').setFormula(
    '=ARRAYFORMULA(IF(G28:G="",,IF(G28:G<0,"🔻 Drop",IF(G28:G>0,"▲ Growth","—"))))');

  d.getRange('B28:G').setNumberFormat('#,##0;-#,##0;0');
  d.getRange('H28:H').setNumberFormat('0.0%').setHorizontalAlignment('right');
  const red = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
    .setFontColor('#c00000').setRanges([d.getRange('D28:D'), d.getRange('G28:H')]).build();
  const green = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setFontColor('#1e7e34').setRanges([d.getRange('D28:D'), d.getRange('G28:H')]).build();
  const drop = SpreadsheetApp.newConditionalFormatRule().whenTextContains('Drop')
    .setBackground('#fde2e2').setRanges([d.getRange('I28:I')]).build();
  d.setConditionalFormatRules([red, green, drop]);
}
