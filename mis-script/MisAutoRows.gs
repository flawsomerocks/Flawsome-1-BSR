/**
 * AUTO CHANNEL ROWS  -  SEPARATE NEW FILE (old MIS code is not touched)   [add as "AutoRows.gs"]
 * Run  ROWS_SETUP  once. After that, every day ~8am (before the old 9am fill) it
 *  1. adds a CHANNEL row when DashData has a portal that has no row in MIS (e.g. Jio Mart, Smytten, any new portal),
 *     inside the Channel block so "Revenue = SUM(channels)" includes it, and back-fills that new row for all months;
 *  2. keeps those extra channel rows updated (running month; previous month until the 10th).
 * New SKUs are NOT added: they stay inside "Rest SKU's (Combo's)".
 * Works on: "MIS as per 1 Total", "MIS as per ITEM PRICE", "MIS as per RO". Old months are never overwritten.
 * Needs MisAutoFill.gs (helper functions) in the same project.
 */
const AR = {
  TABS: [
    { words: ['mis', 'total'], measure: 'total' },
    { words: ['mis', 'item'], measure: 'item price' },
    { key: 'misasperro', measure: 'total' }
  ],
  HOUR: 8,
  MIN_REV: 1,              // a portal needs at least this revenue in the month to get a row
  BACKFILL_CHANNELS: true  // fill the new channel row for all old months too (they were missing before)
};

// run once from the function dropdown
function ROWS_SETUP() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rowsDaily') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rowsDaily').timeBased().everyDays(1).atHour(AR.HOUR).create();
  arRun_(true);
  // then let the existing fills refresh the Revenue formulas right now
  if (typeof fillDaily === 'function') fillDaily();
  if (typeof roDaily === 'function') { try { roDaily(); } catch (e) { /* RO file not added: fine */ } }
}

function rowsDaily() { arRun_(false); }

function arTabs_(ss) {
  const out = [], seen = {};
  AR.TABS.forEach(spec => {
    const tab = spec.key
      ? ss.getSheets().find(s => s.getName().toLowerCase().replace(/\s+/g, '') === spec.key)
      : misFindSheet_(ss, spec.words, false);
    if (tab && !seen[tab.getName()]) { seen[tab.getName()] = 1; out.push({ tab: tab, measure: spec.measure }); }
  });
  return out;
}

function arRun_(full) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dd = misFindSheet_(ss, [MISC.DATA_TAB.toLowerCase()], true);
  if (!dd) { misLog_(ss, 'AUTO ROWS ERROR: "' + MISC.DATA_TAB + '" tab nahi mila.'); return; }
  const values = dd.getDataRange().getValues();
  const n = new Date();
  const win = [[n.getFullYear(), n.getMonth()]];
  if (n.getDate() <= 10) win.unshift([n.getFullYear(), n.getMonth() - 1]);
  const msgs = [];
  arTabs_(ss).forEach(t => {
    try { msgs.push(arSyncTab_(t.tab, t.measure, values, win, full)); }
    catch (e) { msgs.push('[' + t.tab.getName() + '] AUTO ROWS ERROR: ' + e.message); }
  });
  if (!msgs.length) msgs.push('AUTO ROWS: koi MIS tab nahi mila.');
  const msg = 'AUTO ROWS\n' + msgs.join('\n');
  Logger.log(msg);
  misLog_(ss, msg);
}

// original-case portal names from DashData (the aggregate works in lower case)
function arNames_(values) {
  const head = values[0].map(h => String(h).trim().toLowerCase());
  const iP = head.indexOf('portal');
  const portal = {};
  for (let r = 1; r < values.length; r++) {
    const p = String(values[r][iP]).trim();
    if (p && !portal[p.toLowerCase()]) portal[p.toLowerCase()] = p;
  }
  return { portal: portal };
}

function arReadGrid_(tab) {
  return tab.getRange(1, 1, tab.getLastRow(), 4).getValues();
}

// channel block: rows under the "Channel Wise Revenue Split" heading while column B has a label
function arChannelRows_(grid) {
  let start = -1;
  grid.forEach((r, i) => { if (start < 0 && String(r[0]).trim().toLowerCase() === 'channel wise revenue split') start = i; });
  const rows = [];
  if (start < 0) return rows;
  for (let i = start; i < grid.length; i++) {
    const label = String(grid[i][1]).trim();
    if (!label) break;
    rows.push({ row: i + 1, label: label });
  }
  return rows;
}

function arSyncTab_(tab, measure, values, win, full) {
  const name = tab.getName();
  if (tab.getLastColumn() < 4 || tab.getLastRow() < 20) return '[' + name + '] khaali tab, skip.';
  const hr = misHeaderRow_(tab);
  const hdr = tab.getRange(hr, 1, 1, tab.getLastColumn()).getValues()[0];
  const monthCols = [];
  let totCol = 0;
  hdr.forEach((h, i) => {
    if (String(h).trim().toLowerCase() === 'total') totCol = i + 1;
    else { const ym = misHeaderToYm_(h); if (ym && i >= 2) monthCols.push({ col: i + 1, y: ym.y, m: ym.m }); }
  });
  if (!monthCols.length) return '[' + name + '] month columns nahi mile, skip.';

  const aggCache = {};
  const aggOf = (y, m) => aggCache[y + '-' + m] || (aggCache[y + '-' + m] = misAggregate_(values, y, m, measure));
  const scan = full ? monthCols : monthCols.filter(c => win.some(w => new Date(w[0], w[1], 1).getFullYear() === c.y && new Date(w[0], w[1], 1).getMonth() === c.m));
  const names = arNames_(values);
  const log = [];

  // ---- what is missing? ----
  let grid = arReadGrid_(tab);
  const haveLabels = {}; arChannelRows_(grid).forEach(r => { haveLabels[r.label.toLowerCase()] = 1; });
  const newPortals = {};
  scan.forEach(c => {
    const a = aggOf(c.y, c.m);
    Object.keys(a.unmappedPortals).forEach(p => {
      if (p && p.charAt(0) !== '(' && a.unmappedPortals[p] >= AR.MIN_REV && !haveLabels[p]) newPortals[p] = 1;
    });
  });

  // ---- 1. new channel rows (inserted INSIDE the block, so Revenue = SUM(channels) includes them) ----
  Object.keys(newPortals).sort().forEach(p => {
    grid = arReadGrid_(tab);
    const rows = arChannelRows_(grid);
    if (rows.length < 2) return;
    const last = rows[rows.length - 1].row;               // insert before the last channel row => stays inside the SUM range
    tab.insertRowBefore(last);
    const r = last;
    arCopyRowStyle_(tab, r - 1, r);
    tab.getRange(r, 2).setValue(names.portal[p] || p);
    arCopyTotal_(tab, r - 1, r, totCol);
    const cols = (full || AR.BACKFILL_CHANNELS) ? monthCols : scan;
    cols.forEach(c => tab.getRange(r, c.col).setValue(Math.round((aggOf(c.y, c.m).unmappedPortals[p] || 0) * 100) / 100));
    log.push('naya CHANNEL row: ' + (names.portal[p] || p) + ' (row ' + r + (cols === monthCols ? ', saare mahine backfill' : '') + ')');
  });

  // ---- 2. keep extra channel rows updated (only the running window; old months stay frozen) ----
  grid = arReadGrid_(tab);
  const known = {}; Object.keys(MISC.CHANNELS).forEach(k => { known[k] = 1; });
  const extra = arChannelRows_(grid).filter(r => !known[r.label.toLowerCase()] && r.label.toLowerCase() !== 'blinkit');
  extra.forEach(r => {
    scan.forEach(c => {
      const v = aggOf(c.y, c.m).unmappedPortals[r.label.toLowerCase()];
      if (v !== undefined) tab.getRange(r.row, c.col).setValue(Math.round(v * 100) / 100);
    });
  });

  return '[' + name + '] ' + (log.length ? log.join('; ') : 'koi nayi row ki zarurat nahi') +
    (extra.length ? ' | extra channel rows update: ' + extra.map(r => r.label).join(', ') : '');
}

function arCopyRowStyle_(tab, from, to) {
  const w = tab.getLastColumn();
  tab.getRange(from, 1, 1, w).copyTo(tab.getRange(to, 1, 1, w), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

// the Total column formula of the row above (relative), so the new row is summed too
function arCopyTotal_(tab, from, to, totCol) {
  if (!totCol) return;
  const f = tab.getRange(from, totCol).getFormulaR1C1();
  if (f) tab.getRange(to, totCol).setFormulaR1C1(f);
}
