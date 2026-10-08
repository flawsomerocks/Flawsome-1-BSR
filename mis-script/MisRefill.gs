// One-time helper: re-fills ALL months of "1 Total" and "ITEM PRICE" from the corrected DashData.
// Run REFILL_ALL; if the log says "Dobara run karo", run it again until it says done.
function REFILL_ALL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!dd) { misLog_(ss, 'ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const months = misMonthsFrom_(dd.getDataRange().getValues());
  const list = months.map(d => [d.getFullYear(), d.getMonth()]);
  const n = new Date();
  list.push([n.getFullYear(), n.getMonth()]);

  const props = PropertiesService.getScriptProperties();
  let i = Number(props.getProperty('REFILL_I') || 0);
  if (i >= list.length) i = 0;
  const t0 = Date.now();
  for (; i < list.length; i++) {
    if (Date.now() - t0 > 240000) break;
    fillMonths_([list[i]]);
  }
  if (i >= list.length) {
    props.deleteProperty('REFILL_I');
    misLog_(ss, 'REFILL_ALL: SAB DONE. ' + list.length + ' mahine dobara bhare (1 Total + ITEM PRICE).');
  } else {
    props.setProperty('REFILL_I', String(i));
    misLog_(ss, 'REFILL_ALL: ' + i + ' / ' + list.length + ' mahine ho gaye. Dobara REFILL_ALL run karo.');
  }
}
