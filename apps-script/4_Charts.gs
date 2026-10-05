/** file 4 of 4 - charts */
function buildCharts_(d) {
  const add = (type, r1, r2, title, row, h) => d.insertChart(d.newChart()
    .setChartType(type)
    .addRange(d.getRange(r1)).addRange(d.getRange(r2))
    .setNumHeaders(1)
    .setOption('title', title)
    .setOption('width', 520).setOption('height', h)
    .setPosition(row, 13, 0, 0).build());

  add(Charts.ChartType.COLUMN, 'M52:M171', 'O52:O171', 'Daily Revenue', 3, 260);
  add(Charts.ChartType.PIE, 'A12:A31', 'C12:C31', 'Portal-wise Revenue', 17, 280);
  add(Charts.ChartType.BAR, 'E12:E22', 'G12:G22', 'Top 10 SKU (Revenue)', 33, 300);
}
