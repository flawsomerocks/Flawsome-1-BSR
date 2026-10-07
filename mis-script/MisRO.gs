/**
 * "MIS as per RO"  -  SEPARATE NEW FILE (old MIS code is not touched)
 * Add this as a new file in the MIS DEEP ANALISIS Apps Script project, Save, run  RO_SETUP  once.
 *
 * Fills every channel / product / SKU / Units / Marketing-spend cell of the tab "MIS as per RO" from DashData
 * (Total column), exactly like "MIS as per 1 Total", BUT the Blinkit row is NEVER changed:
 * it belongs to the Invoice tool (SUMIFS on 'Invoice Log'). Existing Blinkit formulas are put back after every run.
 * New month -> a new column is added before Total (Blinkit cell of that new column stays empty until the
 * Invoice tool syncs). Old months are never touched.
 * Needs the old MIS file (MisAutoFill.gs) in the same project: this file uses its helper functions.
 */
const RO = {
  TAB_KEY: 'misasperro',     // tab name ignoring spaces and capitals
  MEASURE: 'total',
  SKIP_LABEL: 'blinkit',     // row that the Invoice tool owns
  HOUR: 11
};

// run once from the function dropdown: fills now + sets the daily trigger
function RO_SETUP() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'roDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('roDaily').timeBased().everyDays(1).atHour(RO.HOUR).create();
  roDaily();
}

// running month every day; previous month only until the 10th, then it is frozen
function roDaily() {
  const n = new Date();
  const list = [[n.getFullYear(), n.getMonth()]];
  if (n.getDate() <= 10) list.unshift([n.getFullYear(), n.getMonth() - 1]);
  roFill_(list);
}

function roFindTab_(ss) {
  return ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === RO.TAB_KEY) || null;
}

function roFill_(list) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = roFindTab_(ss);
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!tab) {
    const names = ss.getSheets().map(s => '"' + s.getName() + '"').join(', ');
    misLog_(ss, 'RO ERROR: tab "MIS as per RO" nahi mila. Tabs: ' + names);
    return;
  }
  if (!dd) { misLog_(ss, 'RO ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const values = dd.getDataRange().getValues();
  const msgs = [];
  list.forEach(ym => {
    const first = new Date(ym[0], ym[1], 1);
    const y = first.getFullYear(), m = first.getMonth();
    try { msgs.push(roFillMonth_(tab, values, y, m)); }
    catch (e) { msgs.push('RO ERROR ' + (m + 1) + '/' + y + ': ' + e.message); }
  });
  const msg = msgs.join('\n\n');
  Logger.log(msg);
  misLog_(ss, msg);
}

function roFillMonth_(tab, values, y, m) {
  const agg = misAggregate_(values, y, m, RO.MEASURE);
  if (!agg.rows || tab.getLastColumn() < 3 || tab.getLastRow() < 5) {
    return '[MIS as per RO] ' + (m + 1) + '/' + y + ': DashData me data nahi / tab khaali. Kuch nahi likha.';
  }
  const hr = misHeaderRow_(tab);
  // find the Blinkit row (label in columns A-D)
  const grid = tab.getRange(1, 1, tab.getLastRow(), 4).getValues();
  let row = 0;
  grid.forEach((r, i) => r.forEach(c => { if (!row && String(c).trim().toLowerCase() === RO.SKIP_LABEL) row = i + 1; }));

  // remember what the Invoice tool wrote in this month's Blinkit cell
  const col0 = misFindMonthCol_(tab, hr, y, m);
  let keep = null;
  if (row && col0) {
    const cell = tab.getRange(row, col0);
    keep = { f: cell.getFormula(), v: cell.getValue() };
  }

  const info = fillOneMis_(tab, { label: 'MIS as per RO', measure: RO.MEASURE }, agg, y, m);

  // put the Blinkit cell back
  const col1 = misFindMonthCol_(tab, hr, y, m);
  let note = '';
  if (row && col1) {
    const cell = tab.getRange(row, col1);
    if (keep) {
      if (keep.f) cell.setFormula(keep.f); else cell.setValue(keep.v);
      note = 'Blinkit row jyon ki tyon (invoice tool ki).';
    } else {
      cell.clearContent();
      note = 'Naya column bana: Blinkit cell khaali hai, invoice tool ke agle sync me bharega (tool me "Sync now").';
    }
  } else if (!row) {
    note = 'WARNING: "Blinkit" row nahi mili.';
  }
  return info + '\n' + note;
}
