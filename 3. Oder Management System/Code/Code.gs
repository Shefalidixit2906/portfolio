// ============================================================
//  Company name. — ORDER MANAGEMENT SYSTEM
//  Google Apps Script — Code.gs  (v19 — FOCUS AREAS ADDED)
//
//  WHAT CHANGED FROM v18 (this revision):
//  ──────────────────────────────────────────────────────────
//  1. New getFocusAreaData(pdRows) builder identifies designs
//     (and rolls them up by category) that are producing PCS but
//     selling poorly relative to stock — i.e. the ones that need
//     a sales/marketing push. A design is flagged when:
//       • production PCS > 0 (something was actually made), AND
//       • booked / PCS  <  FOCUS_BOOKING_RATIO_THRESHOLD (40%), AND
//       • balance (PCS - booked) > 0 (there is unsold stock)
//     Each flagged design carries an "idleValue" = balance * MRP,
//     i.e. the rupee value of stock sitting unsold — this is what
//     both the design list and the category rollup are sorted by,
//     so the most business-critical items surface first.
//  2. buildAllDashboardDataLive_() now also returns `focusAreas:
//     { designs, categories, thresholdPct }` alongside the
//     existing keys.
//  3. Everything else is UNCHANGED from v18 (Agent/Distributor
//     column support etc.) — see prior header comments preserved
//     below for that history.
// ============================================================


// ─────────────────────────────────────────────────────────────
//  EXTERNAL SPREADSHEET IDs
// ─────────────────────────────────────────────────────────────

var PRODUCTION_SHEET_ID     = '1qmlioUHlLh9fA12o7BxZh2yzbeK0s8Owv3GUW0kxIaQ';
var DEALER_EXTERNAL_SHEET_ID = '1CQ10W-HsppHFVBk0CovD6j6UNyAvhvWBmRsBI64Q070';
var DEALER_EXTERNAL_TAB      = 'DEALER MASTER2';
var DESIGN_MASTER_SHEET_NAME = 'PRODUCTION-DESIGN_MASTER';


// ─────────────────────────────────────────────────────────────
//  CONFIGURATION
// ─────────────────────────────────────────────────────────────
var CONFIG = {
  SHEETS: {
    ORDER_MASTER:    'ORDER_MASTER',
    INTERNAL_ORDER:  'INTERNAL_ORDER_MASTER',
    DEALER_MASTER:   'DEALER_MASTER',
    PRODUCTION:      'PRODUCTION-DESIGN_MASTER'
  },
  OM_COLS: {
    SR_NO:1, DATE:2, DEALER:3, EXECUTIVE:4, AGENT:5, CITY:6, STATE:7,
    DESIGN_NO:8, QTY_28:9, QTY_30:10, QTY_32:11, QTY_34:12,
    QTY_36:13, QTY_38:14, QTY_40:15, QTY_42:16, QTY_44:17,
    QTY_46:18, TOTAL:19, MRP:20
  },
  IOM_COLS: {
    SR_NO:1, DATE:2, DEALER:3, STATE:4, CITY:5, AGENT:6, EXECUTIVE:7,
    DESIGN:8, CATEGORY:9, SUPPLIER:10, DESIGN_NO:11, DESIGN_CODE:12,
    QTY_28:13, QTY_30:14, QTY_32:15, QTY_34:16, QTY_36:17,
    QTY_38:18, QTY_40:19, QTY_42:20, QTY_44:21, QTY_46:22,
    TOTAL:23, AMOUNT:24, ORDER_STATUS:25
  },
  PDM_COLS: {
    DESIGN_NO:1, DESIGN_CODE:2, DESIGN_NAME:3, CATEGORY:4,
    SUPPLIER:5, FORM_MTR:6, MTR:7, READY_STOCK:8, PCS:9,
    MRP:10, BOOKED_QTY:11, BALANCE:12, STATUS:13,
    BOOKING_STATUS:14
  },
  DM_COLS: { DEALER:1, STATE:2, CITY:3, AGENT:4, EXECUTIVE:5, PHONE:6 },
  MTR_RATIO: {
    'SHIRTS':1.5, 'FORMAL SHIRTS':1.5, 'CASUAL SHIRT':1.5,
    'CASUAL SHIRT-CORE':1.5, 'FORMAL SHIRT':1.5,
    'JEANS':1.38, 'CASUAL TROUSER':1.35, 'COTTON TROUSER':1.35,
    'FORMAL TROUSER':1.23, 'TRAVEL PANTS':1.35
  },
  OVERBOOK_THRESHOLD: 0.20,
  ALERT_EMAIL: 'mis@nathjiagencies.in',
  HEADER_ROW: 1,
  DATA_START:  2
};

// v20: raised from 0.40 -> 0.60 per manager request — designs
// booking below 60% of their production (with unsold stock) now
// surface in Focus Areas, not just the ones below 40%.
var FOCUS_BOOKING_RATIO_THRESHOLD = 0.60; // 60%

var STATUS_COLORS = {
  'OVER BOOKED': '#FF0000', 'OUT OF STOCK': '#FFCCCC',
  'LOW STOCK':   '#FFD700', 'AVAILABLE':    '#C6EFCE',
  'PENDING':     '#FFFFFF', 'NOT SET':      '#F0F0F0'
};

var PROP_DEALER_MAP  = 'dealer_map_v1';
var PROP_DESIGN_MAP  = 'design_map_v1';
var CACHE_DASHBOARD  = 'dashboard_v1';
var CACHE_TTL_SEC    = 21600;

var DESIGN_TYPE_KEYWORDS = [
  'CHECKS','STRIPES','PRINT','PLAIN','TWILL','DOBBY','LINEN',
  'DENIM','DOT','SOLID','JACQUARD','MELANGE','HERRINGBONE','LYCRA'
];

const SOURCE_SHEETS = [
  {
    sheetName: "JEANS", headerRow: 1, category: "JEANS",
    designNoCol: "I", supplierCol: "D", designNameCol: "E",
    orderQtyCol: "J", productionMtrCol: "K", productionMtrCol2: null,
    mrpCol: "M"
  },
  {
    sheetName: "TRAVEL PANTS", headerRow: 1, category: "TRAVEL PANTS",
    designNoCol: "K", supplierCol: "D", designNameCol: "E",
    orderQtyCol: "I", productionMtrCol: "J", productionMtrCol2: null,
    mrpCol: "M"
  },
  {
    sheetName: "COTTON TROUSER", headerRow: 1, category: "COTTON TROUSER",
    designNoCol: "L", supplierCol: "D", designNameCol: "E",
    orderQtyCol: "I", productionMtrCol: "J", productionMtrCol2: "K",
    mrpCol: "N"
  },
  {
    sheetName: "FORMAL TROUSER", headerRow: 1, category: "FORMAL TROUSER",
    designNoCol: "L", supplierCol: "D", designNameCol: "E",
    orderQtyCol: "I", productionMtrCol: "J", productionMtrCol2: "K",
    mrpCol: "N"
  },
  {
    sheetName: "CASUAL SHIRT", headerRow: 1, category: "CASUAL SHIRT",
    designNoCol: "M", supplierCol: "D", designNameCol: "E",
    orderQtyCol: "J", productionMtrCol: "K", productionMtrCol2: "L",
    mrpCol: "P"
  },
  {
    sheetName: "FORMAL SHIRT", headerRow: 1, category: "FORMAL SHIRT",
    designNoCol: "J", supplierCol: "D", designNameCol: "E",
    orderQtyCol: null, productionMtrCol: "I", productionMtrCol2: null,
    mrpCol: "M"
  }
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Order System')
    .addItem('🔄 Full Resync (Orders → Internal)',   'syncOrderMasterToInternal')
    .addItem('📦 Update Production Master',           'updateProductionMaster')
    .addItem('🔁 Sync Production Qty from Source',    'updateOrderQty')
    .addItem('🏪 Sync Dealer Master from External',   'syncDealerMasterFromExternal')
    .addItem('🔧 Correct Existing Dealer Data',       'correctExistingDealerData')
    .addSeparator()
    .addItem('⚡ Refresh Property Cache',             'refreshPropertyCache')
    .addItem('📊 Refresh Dashboard Cache',            'buildDashboardCache')
    .addSeparator()
    .addItem('🚀 Open Dashboard',                     'showDashboard')
    .addToUi();
}

function normalizeDealerKey_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function setupAllTriggers() {
  try {
    SpreadsheetApp.getUi().alert(
      'ℹ️ No automatic triggers are configured.\n\n' +
      'Run all syncs manually from the "📊 Order System" menu:\n' +
      '• Sync Production Qty from Source\n' +
      '• Update Production Master\n' +
      '• Full Resync (Orders → Internal)\n' +
      '• Sync Dealer Master from External'
    );
  } catch(e) {}
}

function syncDealerMasterFromExternal() {
  if (!DEALER_EXTERNAL_SHEET_ID) {
    Logger.log('DEALER_EXTERNAL_SHEET_ID not set — skipping dealer sync.');
    return;
  }
  PropertiesService.getScriptProperties().deleteProperty(PROP_DEALER_MAP);
  CacheService.getScriptCache().remove('dealer_map_cache_v1');
  var map   = getDealerMapFast_();
  var count = Object.keys(map).length;
  Logger.log('Dealer sync complete: ' + count + ' dealers read directly from external Dealer Master.');
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('Dealer Master synced: ' + count + ' records (direct from external sheet).', '✅ Done', 4);
  } catch(e) {}
}

function correctExistingDealerData() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var omSheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER);
  if (!omSheet) { SpreadsheetApp.getUi().alert('ORDER_MASTER sheet not found.'); return; }

  PropertiesService.getScriptProperties().deleteProperty(PROP_DEALER_MAP);
  CacheService.getScriptCache().remove('dealer_map_cache_v1');
  var map = buildDealerMapRaw_(ss);
  if (Object.keys(map).length === 0) {
    SpreadsheetApp.getUi().alert('⚠️ Could not read DEALER MASTER2 — aborting correction. Check DEALER_EXTERNAL_SHEET_ID / tab name.');
    return;
  }

  var lastRow = omSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) {
    SpreadsheetApp.getUi().alert('No order rows found to correct.');
    return;
  }

  var numRows   = lastRow - CONFIG.DATA_START + 1;
  var dealerCol = omSheet.getRange(CONFIG.DATA_START, CONFIG.OM_COLS.DEALER,    numRows, 1).getValues();
  var execBlock = omSheet.getRange(CONFIG.DATA_START, CONFIG.OM_COLS.EXECUTIVE, numRows, 4).getValues();

  var correctedCount = 0, unchangedCount = 0;
  var unmatchedDealers = {};

  for (var i = 0; i < numRows; i++) {
    var dealerName = String(dealerCol[i][0] || '').trim();
    if (!dealerName) continue;

    var info = map[normalizeDealerKey_(dealerName)];
    if (!info) { unmatchedDealers[dealerName] = true; continue; }

    var newExec  = info.executive || '';
    var newAgent = info.agent     || '';
    var newCity  = info.city      || '';
    var newState = info.state     || '';
    var oldExec  = execBlock[i][0], oldAgent = execBlock[i][1],
        oldCity  = execBlock[i][2], oldState = execBlock[i][3];

    if (newExec !== oldExec || newAgent !== oldAgent || newCity !== oldCity || newState !== oldState) {
      execBlock[i] = [newExec, newAgent, newCity, newState];
      correctedCount++;
    } else {
      unchangedCount++;
    }
  }

  omSheet.getRange(CONFIG.DATA_START, CONFIG.OM_COLS.EXECUTIVE, numRows, 4).setValues(execBlock);

  syncOrderMasterToInternal();
  refreshPropertyCache();
  buildDashboardCache();

  var unmatchedList = Object.keys(unmatchedDealers);
  Logger.log(
    'correctExistingDealerData complete. Corrected: ' + correctedCount +
    ', Already correct: ' + unchangedCount +
    ', Unmatched dealers: ' + unmatchedList.length
  );
  if (unmatchedList.length > 0) Logger.log('Unmatched dealer names: ' + unmatchedList.join(', '));

  try {
    SpreadsheetApp.getUi().alert(
      'Dealer Data Correction Complete',
      'Rows corrected: '      + correctedCount +
      '\nAlready correct: '   + unchangedCount +
      '\nUnmatched dealers: ' + unmatchedList.length +
      (unmatchedList.length > 0
        ? '\n\nNot found in DEALER MASTER2 (' + unmatchedList.length + '):\n' +
          unmatchedList.slice(0, 20).join(', ') + (unmatchedList.length > 20 ? '…' : '')
        : '\n\nAll dealer names matched ✅'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {}
}

function buildDashboardCache() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var data = buildAllDashboardDataLive_(ss);
  var json = JSON.stringify(data);
  if (json.length < 95000) {
    CacheService.getScriptCache().put(CACHE_DASHBOARD, json, CACHE_TTL_SEC);
  } else {
    var chunks = chunkString_(json, 90000);
    var cache  = CacheService.getScriptCache();
    cache.put(CACHE_DASHBOARD + '_count', String(chunks.length), CACHE_TTL_SEC);
    chunks.forEach(function(chunk, i) {
      cache.put(CACHE_DASHBOARD + '_' + i, chunk, CACHE_TTL_SEC);
    });
    cache.remove(CACHE_DASHBOARD);
  }
  Logger.log('Dashboard cache built: ' + Math.round(json.length / 1024) + ' KB');
  try { SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard cache refreshed ✅', 'Cache', 3); } catch(e) {}
}

function getAllDashboardData() {
  var cache  = CacheService.getScriptCache();
  var single = cache.get(CACHE_DASHBOARD);
  if (single) return JSON.parse(single);
  var countStr = cache.get(CACHE_DASHBOARD + '_count');
  if (countStr) {
    var count  = parseInt(countStr, 10);
    var chunks = [];
    for (var i = 0; i < count; i++) {
      var chunk = cache.get(CACHE_DASHBOARD + '_' + i);
      if (!chunk) break;
      chunks.push(chunk);
    }
    if (chunks.length === count) return JSON.parse(chunks.join(''));
  }
  Logger.log('Cache miss — building live dashboard data');
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var data = buildAllDashboardDataLive_(ss);
  try {
    var json    = JSON.stringify(data);
    var chunks2 = chunkString_(json, 90000);
    if (chunks2.length === 1) {
      cache.put(CACHE_DASHBOARD, json, CACHE_TTL_SEC);
    } else {
      cache.put(CACHE_DASHBOARD + '_count', String(chunks2.length), CACHE_TTL_SEC);
      chunks2.forEach(function(c, i) { cache.put(CACHE_DASHBOARD + '_' + i, c, CACHE_TTL_SEC); });
    }
  } catch(cacheErr) { Logger.log('Cache store error: ' + cacheErr.message); }
  return data;
}

function getAllDashboardDataLive() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var data = buildAllDashboardDataLive_(ss);
  try { buildDashboardCache(); } catch(e) {}
  return data;
}

function chunkString_(str, size) {
  var chunks = [];
  for (var i = 0; i < str.length; i += size) chunks.push(str.substring(i, i + size));
  return chunks;
}

function getDealerMapFast_() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get('dealer_map_cache_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    Logger.log('Could not acquire lock for dealer map — falling back to direct read.');
    return buildDealerMapRaw_();
  }

  try {
    cached = cache.get('dealer_map_cache_v1');
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }
    var map = buildDealerMapRaw_();
    try {
      var json = JSON.stringify(map);
      if (json.length < 95000) {
        cache.put('dealer_map_cache_v1', json, 300);
      }
    } catch(e) {
      Logger.log('Dealer map too large to cache, will re-fetch each time.');
    }
    return map;
  } finally {
    lock.releaseLock();
  }
}

function getDesignMapFast_() {
  var props  = PropertiesService.getScriptProperties();
  var cached = props.getProperty(PROP_DESIGN_MAP);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var map = buildDesignMapRaw_(ss);
  try { props.setProperty(PROP_DESIGN_MAP, JSON.stringify(map)); }
  catch(e) { Logger.log('Design map too large for Properties — rebuilding each time'); }
  return map;
}

function refreshPropertyCache() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_DEALER_MAP);
  props.deleteProperty(PROP_DESIGN_MAP);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildDealerMapRaw_(ss);
  getDealerMapFast_();
  getDesignMapFast_();
  try { SpreadsheetApp.getActiveSpreadsheet().toast('Property cache refreshed ✅', 'Cache', 3); } catch(e) {}
}

function buildDealerMapRaw_(ss) {
  var map = {};
  if (!DEALER_EXTERNAL_SHEET_ID) return map;

  var externalSS = null;
  var lastError  = null;
  var maxRetries = 3;

  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      externalSS = SpreadsheetApp.openById(DEALER_EXTERNAL_SHEET_ID);
      break;
    } catch(err) {
      lastError = err;
      Logger.log('Attempt ' + attempt + ' failed to open external dealer sheet: ' + err.message);
      if (attempt < maxRetries) Utilities.sleep(500 * attempt);
    }
  }

  if (!externalSS) {
    Logger.log('ERROR: Could not open external dealer sheet after ' + maxRetries + ' attempts: ' + (lastError ? lastError.message : 'unknown'));
    try {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('⚠️ Could not open external Dealer Master after retries.', 'Dealer Sync Error', 8);
    } catch(e) {}
    return map;
  }

  var exSheet = externalSS.getSheetByName(DEALER_EXTERNAL_TAB);
  if (!exSheet) {
    Logger.log('ERROR: Tab "' + DEALER_EXTERNAL_TAB + '" not found in external sheet.');
    try {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast('⚠️ Tab "' + DEALER_EXTERNAL_TAB + '" not found in external Dealer Master.', 'Dealer Sync Error', 8);
    } catch(e) {}
    return map;
  }

  var lastRow = exSheet.getLastRow();
  if (lastRow < 2) return map;
  var data = exSheet.getRange(
    2, 1, lastRow - 1, Math.max(exSheet.getLastColumn(), CONFIG.DM_COLS.EXECUTIVE, CONFIG.DM_COLS.PHONE)
  ).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = normalizeDealerKey_(data[i][CONFIG.DM_COLS.DEALER - 1]);
    if (key) {
      map[key] = {
        executive: data[i][CONFIG.DM_COLS.EXECUTIVE - 1] || '',
        city:      data[i][CONFIG.DM_COLS.CITY      - 1] || '',
        state:     data[i][CONFIG.DM_COLS.STATE     - 1] || '',
        agent:     data[i][CONFIG.DM_COLS.AGENT     - 1] || ''
      };
    }
  }
  return map;
}

function buildDesignMapRaw_(ss) {
  var pdSheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTION);
  var map     = {};
  if (!pdSheet) return map;
  var pdData = pdSheet.getDataRange().getValues();
  for (var j = 1; j < pdData.length; j++) {
    var dKey = String(pdData[j][CONFIG.PDM_COLS.DESIGN_NO - 1]).trim();
    if (dKey) {
      map[dKey] = {
        designName: pdData[j][CONFIG.PDM_COLS.DESIGN_NAME - 1] || '',
        category:   pdData[j][CONFIG.PDM_COLS.CATEGORY    - 1] || '',
        supplier:   pdData[j][CONFIG.PDM_COLS.SUPPLIER    - 1] || '',
        designCode: pdData[j][CONFIG.PDM_COLS.DESIGN_CODE - 1] || '',
        // Read straight from the sheet — see getProductionRows_ for
        // why this must not be re-derived as pcs-booked.
        balance:    Number(pdData[j][CONFIG.PDM_COLS.BALANCE    - 1]) || 0,
        mrp:        Number(pdData[j][CONFIG.PDM_COLS.MRP        - 1]) || 0,
        pcs:        Number(pdData[j][CONFIG.PDM_COLS.PCS        - 1]) || 0,
        booked:     Number(pdData[j][CONFIG.PDM_COLS.BOOKED_QTY - 1]) || 0,
        status:     pdData[j][CONFIG.PDM_COLS.STATUS       - 1] || 'AVAILABLE'
      };
    }
  }
  return map;
}

function getDealerMap_(ss) { return getDealerMapFast_(); }
function getDesignMap_(ss) { return getDesignMapFast_(); }

function debugRetryBehavior() {
  for (var i = 1; i <= 5; i++) {
    var start = new Date().getTime();
    var map = buildDealerMapRaw_();
    var elapsed = new Date().getTime() - start;
    Logger.log('Run ' + i + ': got ' + Object.keys(map).length + ' dealers in ' + elapsed + 'ms');
  }
}

function onEdit(e) {
  onEditSimple_(e);
}

function onEditSimple_(e) {
  if (!e || !e.source || !e.range) return;
  var sheet     = e.source.getActiveSheet();
  var range     = e.range;
  var row       = range.getRow();
  var col       = range.getColumn();
  var sheetName = sheet.getName();

  if (sheetName === CONFIG.SHEETS.PRODUCTION && row > CONFIG.HEADER_ROW) {
    handleProductionEdit_(sheet, row, col);
    return;
  }
  if (sheetName !== CONFIG.SHEETS.ORDER_MASTER) return;
  if (row <= CONFIG.HEADER_ROW) return;

  if (col === CONFIG.OM_COLS.DESIGN_NO && range.getValue() !== '') {
    resolveRowColor_(sheet, row, String(range.getValue()).trim());
  }
  if (isOrderEditableColumn_(col)) {
    syncSingleRowToInternal_(sheet, row);
  }
}
function autoFillDealerInfoFast_(sheet, row, dealerName) {
  if (!dealerName) return;
  var map  = getDealerMapFast_();
  var info = map[normalizeDealerKey_(dealerName)];
  if (!info) {
    Logger.log('No dealer match found for: "' + dealerName + '"');
    return;
  }
  sheet.getRange(row, CONFIG.OM_COLS.EXECUTIVE, 1, 4).setValues([[
    info.executive || '', info.agent || '', info.city || '', info.state || ''
  ]]);
}

function onEditInstallable_(e) {
  if (!e || !e.source || !e.range) return;
  var sheet     = e.source.getActiveSheet();
  var range     = e.range;
  var row       = range.getRow();
  var col       = range.getColumn();
  var sheetName = sheet.getName();

  if (sheetName !== CONFIG.SHEETS.ORDER_MASTER) return;
  if (row <= CONFIG.HEADER_ROW) return;
  if (col !== CONFIG.OM_COLS.DEALER) return;

  var map = buildDealerMapRaw_();
  if (Object.keys(map).length === 0) {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('⚠️ Dealer Master could not be read — Executive/Agent/City/State not filled. Try again.', 'Dealer Sync Failed', 6);
  }

  var dealerValues = range.getValues();
  for (var rIdx = 0; rIdx < dealerValues.length; rIdx++) {
    var thisRow    = row + rIdx;
    var dealerName = String(dealerValues[rIdx][0] || '').trim();
    if (!dealerName) continue;
    autoFillSrAndDate_(sheet, thisRow);
    var info = map[normalizeDealerKey_(dealerName)];
    if (info) {
      sheet.getRange(thisRow, CONFIG.OM_COLS.EXECUTIVE, 1, 4).setValues([[
        info.executive || '', info.agent || '', info.city || '', info.state || ''
      ]]);
    } else {
      Logger.log('No dealer match found for: "' + dealerName + '"');
    }
  }
}


function installDealerAutoFillTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditInstallable_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditInstallable_')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Dealer auto-fill trigger installed ✅ — try typing a dealer name now.', 'Setup Complete', 5);
}


function isOrderEditableColumn_(col) {
  return col >= CONFIG.OM_COLS.DEALER && col <= CONFIG.OM_COLS.MRP;
}

function handleProductionEdit_(sheet, row, col) {
  var c = CONFIG.PDM_COLS;
  if (col !== c.FORM_MTR && col !== c.MTR && col !== c.READY_STOCK && col !== c.CATEGORY) return;
  var mtrVals    = sheet.getRange(row, c.FORM_MTR, 1, 3).getValues()[0];
  var formMtr    = mtrVals[0];
  var mtr        = mtrVals[1];
  var readyStock = mtrVals[2];
  var category   = sheet.getRange(row, c.CATEGORY).getValue();
  var pcs = calculatePCS(formMtr, mtr, readyStock, category);
  sheet.getRange(row, c.PCS).setValue(pcs === '' ? '' : pcs);
}

function syncSingleRowToInternal_(omSheet, omRow) {
  var ss            = omSheet.getParent();
  var internalSheet = ss.getSheetByName(CONFIG.SHEETS.INTERNAL_ORDER);
  if (!internalSheet) return;
  var dealerMap   = getDealerMapFast_();
  var designMap   = getDesignMapFast_();
  var lastCol     = Math.max(omSheet.getLastColumn(), CONFIG.OM_COLS.MRP);
  var omRowValues = omSheet.getRange(omRow, 1, 1, lastCol).getValues()[0];
  var built       = buildInternalRow_(omRowValues, dealerMap, designMap);
  if (built.isBlank) {
    internalSheet.getRange(omRow, 1, 1, built.row.length).clearContent();
    internalSheet.getRange(omRow, 1, 1, built.row.length).setBackground('#FFFFFF');
    return;
  }
  internalSheet.getRange(omRow, 1, 1, built.row.length).setValues([built.row]);
  internalSheet.getRange(omRow, 1, 1, built.row.length)
    .setBackground(STATUS_COLORS[built.status] || '#FFFFFF');
}

function syncOrderMasterToInternal() {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var omSheet       = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER);
  var internalSheet = ss.getSheetByName(CONFIG.SHEETS.INTERNAL_ORDER);
  if (!omSheet || !internalSheet) {
    SpreadsheetApp.getUi().alert('Could not find required sheets.'); return;
  }
  var dealerMap = buildDealerMapRaw_(ss);
  var designMap = buildDesignMapRaw_(ss);
  var omLastRow = omSheet.getLastRow();
  if (omLastRow < CONFIG.DATA_START) return;

  var omData = omSheet.getRange(
    CONFIG.DATA_START, 1, omLastRow - CONFIG.DATA_START + 1, omSheet.getLastColumn()
  ).getValues();

  var iomColCount = CONFIG.IOM_COLS.ORDER_STATUS;
  var blankRow    = new Array(iomColCount).fill('');
  var blankBg     = new Array(iomColCount).fill('#FFFFFF');

  var output      = [];
  var backgrounds = [];
  var writtenCount = 0;

  omData.forEach(function(rowValues) {
    var built = buildInternalRow_(rowValues, dealerMap, designMap);
    if (built.isBlank) {
      output.push(blankRow.slice());
      backgrounds.push(blankBg.slice());
    } else {
      output.push(built.row);
      backgrounds.push(new Array(built.row.length).fill(STATUS_COLORS[built.status] || '#FFFFFF'));
      writtenCount++;
    }
  });

  var intLastRow = internalSheet.getLastRow();
  if (intLastRow >= CONFIG.DATA_START) {
    internalSheet.getRange(
      CONFIG.DATA_START, 1,
      intLastRow - CONFIG.DATA_START + 1,
      internalSheet.getLastColumn()
    ).clearContent().setBackground('#FFFFFF');
  }

  if (output.length > 0) {
    var targetRange = internalSheet.getRange(CONFIG.DATA_START, 1, output.length, output[0].length);
    targetRange.setValues(output);
    targetRange.setBackgrounds(backgrounds);
  }

  refreshPropertyCache();
  buildDashboardCache();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Full resync complete: ' + writtenCount + ' orders synced across ' + output.length + ' rows.', '✅ Done', 4);
}

function updateProductionMaster() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var pdSheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTION);
  var omSheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER);
  if (!pdSheet || !omSheet) {
    SpreadsheetApp.getUi().alert('Could not find required sheets.'); return;
  }
  var pdLastRow = pdSheet.getLastRow();
  if (pdLastRow < CONFIG.DATA_START) return;

  var bookedMap = buildBookedMap_(omSheet);
  var c         = CONFIG.PDM_COLS;
  var pdRange   = pdSheet.getRange(CONFIG.DATA_START, 1, pdLastRow - 1, pdSheet.getLastColumn());
  var pdData    = pdRange.getValues();
  var pdBg      = pdRange.getBackgrounds();
  var overbookedDesigns = [];

  for (var i = 0; i < pdData.length; i++) {
    var row      = pdData[i];
    var designNo = String(row[c.DESIGN_NO - 1] || '').trim();
    var category = String(row[c.CATEGORY  - 1] || '').trim();
    if (!designNo) continue;

    var formMtr = row[c.FORM_MTR    - 1];
    var mtrQty  = row[c.MTR         - 1];
    var rs      = row[c.READY_STOCK - 1];

    var productionPCS = calculatePCS(formMtr, mtrQty, rs, category);
    var pcsIsBlank    = (productionPCS === '');
    if (pcsIsBlank) productionPCS = 0;
    pdData[i][c.PCS - 1] = pcsIsBlank ? '' : productionPCS;

    var booked  = bookedMap[designNo] || 0;
    var balance = pcsIsBlank ? '' : (productionPCS - booked);
    var status, bookingStatus, rowColor;

    if (pcsIsBlank) {
      status = 'NOT SET'; bookingStatus = '—'; rowColor = '#F0F0F0';
    } else if (productionPCS === 0 && booked > 0) {
      status = 'OUT OF STOCK'; bookingStatus = '❌ OUT'; rowColor = '#FFCCCC';
    } else if (productionPCS > 0 && booked > productionPCS * 1.20) {
      status = 'OVER BOOKED'; bookingStatus = '🔴 OVER BOOKED'; rowColor = '#FF0000';
      overbookedDesigns.push(designNo + ' (booked:' + booked + ', prod:' + productionPCS + ')');
    } else if (productionPCS > 0 && booked >= productionPCS * 0.80) {
      status = 'LOW STOCK'; bookingStatus = '⚠️ LOW'; rowColor = '#FFD700';
    } else if (productionPCS > 0) {
      status = 'AVAILABLE'; bookingStatus = '✅ OK';
      rowColor = (booked > 0) ? '#C6EFCE' : '#FFFFFF';
    } else {
      status = 'NOT SET'; bookingStatus = '—'; rowColor = '#F0F0F0';
    }

    pdData[i][c.BOOKED_QTY     - 1] = booked;
    pdData[i][c.BALANCE        - 1] = balance;
    pdData[i][c.STATUS         - 1] = status;
    pdData[i][c.BOOKING_STATUS - 1] = bookingStatus;
    for (var bgc = 0; bgc < pdBg[i].length; bgc++) pdBg[i][bgc] = rowColor;
  }

  pdRange.setValues(pdData);
  pdRange.setBackgrounds(pdBg);
  if (overbookedDesigns.length > 0) sendOverbookAlert_(overbookedDesigns);
  refreshPropertyCache();
  buildDashboardCache();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Production master updated! Over booked: ' + overbookedDesigns.length, '✅ Done', 4);
}

function calculatePCS(formMtr, mtr, readyStock, category) {
  var cat   = String(category || '').trim().toUpperCase();
  var ratio = CONFIG.MTR_RATIO[cat] || 1.5;

  var mtrVal     = (mtr        !== '' && mtr        !== null && !isNaN(Number(mtr)))        ? Number(mtr)        : 0;
  var formMtrVal = (formMtr    !== '' && formMtr    !== null && !isNaN(Number(formMtr)))    ? Number(formMtr)    : 0;
  var rsVal      = (readyStock !== '' && readyStock !== null && !isNaN(Number(readyStock))) ? Number(readyStock) : 0;

  var combinedMtr = formMtrVal + mtrVal;
  var hasCombined = combinedMtr > 0;
  var hasRS       = rsVal > 0;

  if (!hasCombined && !hasRS) return '';

  if (hasCombined) {
    var pcsFromMtr = Math.floor(combinedMtr / ratio);
    return hasRS ? (pcsFromMtr + rsVal) : pcsFromMtr;
  }
  return rsVal;
}

function updateOrderQty() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(DESIGN_MASTER_SHEET_NAME);
  if (!masterSheet) throw new Error('Sheet "' + DESIGN_MASTER_SHEET_NAME + '" not found.');
  if (!PRODUCTION_SHEET_ID) throw new Error('Please set PRODUCTION_SHEET_ID.');

  var productionSS   = SpreadsheetApp.openById(PRODUCTION_SHEET_ID);
  var codeToData      = {};
  var allSourceCodes  = [];

  SOURCE_SHEETS.forEach(function(cfg) {
    var sheet = productionSS.getSheetByName(cfg.sheetName);
    if (!sheet) {
      Logger.log('WARNING: Source sheet "' + cfg.sheetName + '" not found. Skipping.');
      return;
    }
    var lastRow      = sheet.getLastRow();
    var firstDataRow = cfg.headerRow + 1;
    if (lastRow < firstDataRow) {
      Logger.log('INFO: "' + cfg.sheetName + '" has no data rows. Skipping.');
      return;
    }

    var colsNeeded = [
      columnLetterToIndex(cfg.designNoCol),
      columnLetterToIndex(cfg.supplierCol),
      columnLetterToIndex(cfg.designNameCol),
      columnLetterToIndex(cfg.productionMtrCol),
      columnLetterToIndex(cfg.mrpCol)
    ];
    if (cfg.orderQtyCol) colsNeeded.push(columnLetterToIndex(cfg.orderQtyCol));
    if (cfg.productionMtrCol2) colsNeeded.push(columnLetterToIndex(cfg.productionMtrCol2));
    var maxCol  = Math.max.apply(null, colsNeeded);
    var numRows = lastRow - cfg.headerRow;

    var block = sheet.getRange(firstDataRow, 1, numRows, maxCol).getValues();

    var dIdx   = columnLetterToIndex(cfg.designNoCol)      - 1;
    var supIdx = columnLetterToIndex(cfg.supplierCol)       - 1;
    var nmIdx  = columnLetterToIndex(cfg.designNameCol)     - 1;
    var oqIdx  = cfg.orderQtyCol ? columnLetterToIndex(cfg.orderQtyCol) - 1 : null;
    var m1Idx  = columnLetterToIndex(cfg.productionMtrCol)  - 1;
    var m2Idx  = cfg.productionMtrCol2 ? columnLetterToIndex(cfg.productionMtrCol2) - 1 : null;
    var mrpIdx = columnLetterToIndex(cfg.mrpCol)            - 1;

    for (var i = 0; i < block.length; i++) {
      var r    = block[i];
      var code = normalizeCode(r[dIdx]);
      if (!code) continue;

      var m1Raw = r[m1Idx];
      var m1Num = (m1Raw !== '' && m1Raw !== null && !isNaN(Number(m1Raw))) ? Number(m1Raw) : 0;
      var productionMtr;

      if (m2Idx !== null) {
        var m2Raw = r[m2Idx];
        var m2Num = (m2Raw !== '' && m2Raw !== null && !isNaN(Number(m2Raw))) ? Number(m2Raw) : 0;
        productionMtr = (m1Num === 0 && m2Num === 0) ? '' : (m1Num + m2Num);
      } else {
        productionMtr = (m1Raw === '' || m1Raw === null || isNaN(Number(m1Raw))) ? '' : m1Raw;
      }

      var orderQty = '';
      if (oqIdx !== null) {
        var oqRaw = r[oqIdx];
        orderQty  = (oqRaw === '' || oqRaw === null || isNaN(Number(oqRaw))) ? '' : oqRaw;
      }

      var mrpRaw = r[mrpIdx];
      var mrp    = (mrpRaw === '' || mrpRaw === null || isNaN(Number(mrpRaw))) ? '' : mrpRaw;

      codeToData[code] = {
        orderQty:      orderQty,
        productionMtr: productionMtr,
        mrp:           mrp,
        supplier:      String(r[supIdx] || '').trim(),
        designName:    String(r[nmIdx]  || '').trim(),
        category:      cfg.category
      };
      allSourceCodes.push(code);
    }
    Logger.log('Read ' + block.length + ' rows from "' + cfg.sheetName + '"');
  });

  var c              = CONFIG.PDM_COLS;
  var lastMasterRow  = masterSheet.getLastRow();
  var firstDataRow2  = CONFIG.HEADER_ROW + 1;
  var masterMaxCol   = Math.max(
    c.DESIGN_NO, c.DESIGN_CODE, c.DESIGN_NAME, c.CATEGORY, c.SUPPLIER,
    c.FORM_MTR, c.MTR, c.READY_STOCK, c.PCS, c.MRP
  );

  var codeToRow = {};
  var masterBlock = [];

  if (lastMasterRow >= firstDataRow2) {
    var numMasterRows = lastMasterRow - CONFIG.HEADER_ROW;
    masterBlock = masterSheet.getRange(firstDataRow2, 1, numMasterRows, masterMaxCol).getValues();
    for (var mr = 0; mr < masterBlock.length; mr++) {
      var mCode = normalizeCode(masterBlock[mr][c.DESIGN_NO - 1]);
      if (mCode) codeToRow[mCode] = mr;
    }
  }

  var formMtrOut = masterBlock.map(function(r) { return [r[c.FORM_MTR    - 1]]; });
  var mtrOut     = masterBlock.map(function(r) { return [r[c.MTR         - 1]]; });
  var mrpOut     = masterBlock.map(function(r) { return [r[c.MRP         - 1]]; });
  var supOut     = masterBlock.map(function(r) { return [r[c.SUPPLIER    - 1]]; });
  var nameOut    = masterBlock.map(function(r) { return [r[c.DESIGN_NAME - 1]]; });
  var catOut     = masterBlock.map(function(r) { return [r[c.CATEGORY    - 1]]; });

  var matchedCount = 0, skippedRS = 0, matchedCodes = {};
  var newRows = [];

  var uniqueSourceCodes = allSourceCodes.filter(function(code, idx, arr) {
    return arr.indexOf(code) === idx;
  });

  uniqueSourceCodes.forEach(function(code) {
    var d = codeToData[code];

    if (codeToRow.hasOwnProperty(code)) {
      var mr = codeToRow[code];

      var formMtrVal = masterBlock[mr][c.FORM_MTR    - 1];
      var rsVal      = masterBlock[mr][c.READY_STOCK - 1];
      var formMtrNum = (formMtrVal !== '' && formMtrVal !== null && !isNaN(Number(formMtrVal))) ? Number(formMtrVal) : 0;
      var rsNum      = (rsVal !== '' && rsVal !== null && !isNaN(Number(rsVal))) ? Number(rsVal) : 0;

      var isReadyStockRow = (rsNum > 0 && formMtrNum === 0);
      if (isReadyStockRow) {
        skippedRS++;
        if (d.mrp        !== '') mrpOut[mr][0]  = d.mrp;
        if (d.supplier   !== '') supOut[mr][0]  = d.supplier;
        if (d.designName !== '') nameOut[mr][0] = d.designName;
        matchedCount++; matchedCodes[code] = true;
        return;
      }

      formMtrOut[mr][0] = d.orderQty;
      mtrOut[mr][0]     = d.productionMtr;
      if (d.mrp        !== '') mrpOut[mr][0]  = d.mrp;
      if (d.supplier   !== '') supOut[mr][0]  = d.supplier;
      if (d.designName !== '') nameOut[mr][0] = d.designName;
      matchedCount++; matchedCodes[code] = true;

    } else {
      var newRow = new Array(masterMaxCol).fill('');
      newRow[c.DESIGN_NO    - 1] = code;
      newRow[c.CATEGORY     - 1] = d.category;
      newRow[c.SUPPLIER     - 1] = d.supplier;
      newRow[c.DESIGN_NAME  - 1] = d.designName;
      newRow[c.FORM_MTR     - 1] = d.orderQty;
      newRow[c.MTR          - 1] = d.productionMtr;
      newRow[c.MRP          - 1] = d.mrp;
      newRows.push(newRow);
      matchedCodes[code] = true;
    }
  });

  if (masterBlock.length > 0) {
    masterSheet.getRange(firstDataRow2, c.FORM_MTR,    masterBlock.length, 1).setValues(formMtrOut);
    masterSheet.getRange(firstDataRow2, c.MTR,         masterBlock.length, 1).setValues(mtrOut);
    masterSheet.getRange(firstDataRow2, c.MRP,         masterBlock.length, 1).setValues(mrpOut);
    masterSheet.getRange(firstDataRow2, c.SUPPLIER,    masterBlock.length, 1).setValues(supOut);
    masterSheet.getRange(firstDataRow2, c.DESIGN_NAME, masterBlock.length, 1).setValues(nameOut);
    masterSheet.getRange(firstDataRow2, c.CATEGORY,    masterBlock.length, 1).setValues(catOut);
  }

  var addedCount = newRows.length;
  if (addedCount > 0) {
    var appendStartRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(appendStartRow, 1, newRows.length, masterMaxCol).setValues(newRows);
  }

  recalculateAllPCS_();

  refreshPropertyCache();
  buildDashboardCache();

  var missing = uniqueSourceCodes.filter(function(code) {
    return !matchedCodes[code];
  });

  Logger.log(
    'updateOrderQty complete.' +
    ' Updated existing: ' + (matchedCount - addedCount) +
    ', New rows added: '  + addedCount +
    ', READY STOCK protected: ' + skippedRS +
    ', Unmatched: ' + missing.length
  );
  if (missing.length > 0) Logger.log('Unmatched codes: ' + missing.join(', '));

  try {
    SpreadsheetApp.getUi().alert(
      'Production Qty Update Complete',
      'Existing rows updated: ' + (matchedCount - addedCount) +
      '\nNew rows added: '      + addedCount +
      '\nREADY STOCK rows protected: ' + skippedRS +
      (missing.length > 0
        ? '\n\nSource codes not processed (' + missing.length + '):\n' +
          missing.slice(0, 20).join(', ') + (missing.length > 20 ? '…' : '')
        : '\n\nAll source codes matched or added ✅'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(uiErr) {}
}

function recalculateAllPCS_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var pdSheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTION);
  if (!pdSheet) return;
  var lastRow = pdSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return;

  var c       = CONFIG.PDM_COLS;
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var mtrBlock = pdSheet.getRange(CONFIG.DATA_START, c.FORM_MTR, numRows, 3).getValues();
  var catBlock = pdSheet.getRange(CONFIG.DATA_START, c.CATEGORY, numRows, 1).getValues();
  var pcsOut   = [];

  for (var i = 0; i < numRows; i++) {
    var pcs = calculatePCS(mtrBlock[i][0], mtrBlock[i][1], mtrBlock[i][2], catBlock[i][0]);
    pcsOut.push([pcs === '' ? '' : pcs]);
  }

  pdSheet.getRange(CONFIG.DATA_START, c.PCS, numRows, 1).setValues(pcsOut);
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Company name — Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Company name — Manager Dashboard')
    .setWidth(1400);
  SpreadsheetApp.getUi()
    .showModelessDialog(html, '📊 Company name — Order Management Dashboard');
}

function refreshAllData() {
  updateProductionMaster();
  syncOrderMasterToInternal();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Dashboard Updated Successfully', '✅ Refresh Complete', 3);
}

function buildInternalRow_(omRowValues, dealerMap, designMap) {
  var srNo      = omRowValues[CONFIG.OM_COLS.SR_NO      - 1];
  var dateVal   = omRowValues[CONFIG.OM_COLS.DATE       - 1];
  var dealer    = String(omRowValues[CONFIG.OM_COLS.DEALER    - 1] || '').trim();
  var executive = String(omRowValues[CONFIG.OM_COLS.EXECUTIVE - 1] || '').trim();
  var agent     = String(omRowValues[CONFIG.OM_COLS.AGENT     - 1] || '').trim();
  var city      = String(omRowValues[CONFIG.OM_COLS.CITY      - 1] || '').trim();
  var state     = String(omRowValues[CONFIG.OM_COLS.STATE     - 1] || '').trim();
  var designNo  = String(omRowValues[CONFIG.OM_COLS.DESIGN_NO - 1] || '').trim();

  var dealerInfo = dealerMap[dealer] || {};
  if (!city)      city      = dealerInfo.city      || '';
  if (!state)     state     = dealerInfo.state     || '';
  if (!executive) executive = dealerInfo.executive || '';
  if (!agent)     agent     = dealerInfo.agent      || '';

  var qtys = [];
  for (var col = CONFIG.OM_COLS.QTY_28; col <= CONFIG.OM_COLS.QTY_46; col++) {
    qtys.push(Number(omRowValues[col - 1]) || 0);
  }

  var total = Number(omRowValues[CONFIG.OM_COLS.TOTAL - 1]) || 0;
  if (total === 0) total = qtys.reduce(function(a, b) { return a + b; }, 0);

  var designInfo        = designMap[designNo] || {};
  var prodPCS           = Number(designInfo.pcs    || 0);
  var totalBookedDesign = Number(designInfo.booked || 0);
  var pdmStatus         = String(designInfo.status || '').toUpperCase();
  var orderStatus;

  if (!designInfo.category) {
    orderStatus = 'PENDING';
  } else if (pdmStatus === 'NOT SET') {
    orderStatus = 'PENDING';
  } else if (prodPCS === 0 && total > 0) {
    orderStatus = 'OUT OF STOCK';
  } else if (prodPCS > 0 && totalBookedDesign > prodPCS * 1.20) {
    orderStatus = 'OVER BOOKED';
  } else if (prodPCS > 0 && totalBookedDesign >= prodPCS * 0.80) {
    orderStatus = 'LOW STOCK';
  } else {
    orderStatus = 'AVAILABLE';
  }

  var amount = (designInfo.mrp || 0) * total;
  var rowOut = [
    srNo, dateVal, dealer,
    state, city, agent, executive,
    designInfo.designName || '',
    designInfo.category   || '',
    designInfo.supplier   || '',
    designNo,
    designInfo.designCode || ''
  ].concat(qtys, [total, amount, orderStatus]);

  return { row: rowOut, status: orderStatus, isBlank: (!srNo && !dealer) };
}

function buildBookedMap_(omSheet) {
  var lastRow = omSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return {};
  var data = omSheet.getRange(
    CONFIG.DATA_START, 1, lastRow - 1, CONFIG.OM_COLS.TOTAL
  ).getValues();
  var map = {};
  data.forEach(function(row) {
    var designNo = String(row[CONFIG.OM_COLS.DESIGN_NO - 1] || '').trim();
    var total    = Number(row[CONFIG.OM_COLS.TOTAL     - 1]) || 0;
    if (total === 0) {
      for (var col = CONFIG.OM_COLS.QTY_28 - 1; col <= CONFIG.OM_COLS.QTY_46 - 1; col++) {
        total += Number(row[col]) || 0;
      }
    }
    if (designNo) map[designNo] = (map[designNo] || 0) + total;
  });
  return map;
}

function autoFillSrAndDate_(sheet, row) {
  var srCell = sheet.getRange(row, CONFIG.OM_COLS.SR_NO);
  if (srCell.getValue() === '' || srCell.getValue() === null) {
    var lastSr = 0;
    if (row > CONFIG.DATA_START) {
      var prev = sheet.getRange(row - 1, CONFIG.OM_COLS.SR_NO).getValue();
      if (!isNaN(prev) && prev !== '') lastSr = Number(prev);
    }
    srCell.setValue(lastSr + 1);
  }
  var dateCell = sheet.getRange(row, CONFIG.OM_COLS.DATE);
  if (dateCell.getValue() === '') dateCell.setValue(new Date());
}

function resolveRowColor_(sheet, editedRow, designNo) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var color   = null;
  var fontColor = null;

  if (designNo && lastRow >= CONFIG.DATA_START) {
    var designCol = sheet.getRange(
      CONFIG.DATA_START, CONFIG.OM_COLS.DESIGN_NO, lastRow - 1, 1
    ).getValues();
    for (var i = 0; i < designCol.length; i++) {
      var rowNo = i + CONFIG.DATA_START;
      if (rowNo === editedRow) continue;
      if (String(designCol[i][0]).trim() === designNo) {
        var bg = sheet.getRange(rowNo, 1).getBackground();
        if (bg && bg !== '#ffffff' && bg !== '#FFFFFF') {
          color     = bg;
          fontColor = sheet.getRange(rowNo, 1).getFontColor();
          break;
        }
      }
    }
  }
  if (!color) {
    var backgrounds = sheet.getRange(editedRow, 1, 1, lastCol).getBackgrounds()[0];
    for (var col = 0; col < backgrounds.length; col++) {
      var bgc = backgrounds[col];
      if (bgc && bgc !== '#ffffff' && bgc !== '#FFFFFF' && bgc !== 'white') {
        color = bgc; break;
      }
    }
  }
  if (!color) return;
  var targetRange = sheet.getRange(editedRow, 1, 1, lastCol);
  targetRange.setBackground(color);
  if (fontColor) targetRange.setFontColor(fontColor);
}

function buildAllDashboardDataLive_(ss) {
  var pdRows    = getProductionRows_(ss);
  var iomRows   = getInternalOrderRows_(ss);
  var dealerMap = buildDealerMapRaw_(ss);

  return {
    cards:          getCardsData(ss, pdRows, iomRows),
    dealers:        getDealerData(ss, iomRows),
    executives:     getExecutiveData(ss, iomRows, dealerMap),
    categories:     getCategoryData(ss, pdRows),
    suppliers:      getSupplierData(ss, pdRows),
    orders:         getOrderData(ss),
    internalOrders: getInternalOrderData(ss),
    production:     getProductionData(ss),
    dealerMaster:   getDealerMasterData(ss),
    focusAreas:     getFocusAreaData(pdRows)
  };
}

// ─────────────────────────────────────────────────────────────
//  ROOT-CAUSE FIX (v21): dashboard now matches the Google Sheet
//  exactly, verified against the real workbook.
//  ──────────────────────────────────────────────────────────
//  PRODUCTION-DESIGN_MASTER has ~330 rows where PCS is still
//  blank ("NOT SET" — nothing produced yet) but BOOKED QTY is
//  already non-zero (orders were placed ahead of production).
//  For those rows updateProductionMaster() deliberately leaves
//  BALANCE blank rather than writing a negative number — that's
//  the sheet's own convention for "not trackable yet".
//
//  A previous revision of this file DERIVED balance here as
//  (pcs - booked) for every row, which for those ~330 rows
//  produces a large NEGATIVE number instead of the sheet's blank
//  — that pulled the Overview/Production-Inventory KPI far below
//  what Category/Supplier (which summed the stored BALANCE
//  column) and the sheet itself show. Verified on the actual
//  workbook: summing PRODUCTION-DESIGN_MASTER's own BALANCE
//  column gives 56,967.20 (matches Supplier tab); the derive-live
//  formula gave 47,969.75 instead.
//
//  Fix: read BALANCE straight from the sheet again (no
//  re-derivation) so every section of the dashboard — Overview,
//  Production Inventory, Category, Supplier, Focus Areas — shows
//  the exact same number you'd get by summing that column in
//  Google Sheets yourself.
// ─────────────────────────────────────────────────────────────
function getProductionRows_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTION);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var c    = CONFIG.PDM_COLS;
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[c.DESIGN_NO - 1]) continue;
    rows.push({
      designNo:   String(row[c.DESIGN_NO    - 1] || '').trim(),
      designName: String(row[c.DESIGN_NAME  - 1] || '').trim(),
      category:   String(row[c.CATEGORY     - 1] || '').trim(),
      supplier:   String(row[c.SUPPLIER     - 1] || '').trim(),
      pcs:        Number(row[c.PCS          - 1]) || 0,
      mrp:        Number(row[c.MRP          - 1]) || 0,
      booked:     Number(row[c.BOOKED_QTY   - 1]) || 0,
      // Read straight from the sheet's own BALANCE column — this
      // is what makes the dashboard match the Google Sheet. See
      // the note above for why this must NOT be re-derived.
      balance:    Number(row[c.BALANCE      - 1]) || 0,
      status:     String(row[c.STATUS       - 1] || '').toUpperCase()
    });
  }
  return rows;
}

function getInternalOrderRows_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INTERNAL_ORDER);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var c    = CONFIG.IOM_COLS;
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[c.DEALER - 1]) continue;
    rows.push({
      dealer:    String(row[c.DEALER     - 1] || '').trim(),
      state:     String(row[c.STATE      - 1] || '').trim(),
      city:      String(row[c.CITY       - 1] || '').trim(),
      agent:     String(row[c.AGENT      - 1] || '').trim(),
      executive: String(row[c.EXECUTIVE  - 1] || '').trim(),
      category:  String(row[c.CATEGORY   - 1] || '').trim(),
      designNo:  String(row[c.DESIGN_NO  - 1] || '').trim(),
      total:     Number(row[c.TOTAL      - 1]) || 0,
      amount:    Number(row[c.AMOUNT     - 1]) || 0,
      status:    String(row[c.ORDER_STATUS - 1] || 'PENDING').trim()
    });
  }
  return rows;
}

function getCardsData(ss, pdRows, iomRows) {
  var cards = getDefaultCards();

  var avail = 0, outOfStock = 0, lowStock = 0, overbooked = 0;
  var totalProd = 0, totalBooked = 0, totalSales = 0, totalBalance = 0;
  pdRows.forEach(function(r) {
    totalProd += r.pcs; totalBooked += r.booked; totalSales += (r.booked * r.mrp);
    // Sum the sheet's own stored BALANCE (r.balance), NOT
    // (totalProd - totalBooked). Some rows have PCS still blank
    // ("NOT SET") but a non-zero BOOKED QTY from orders already
    // placed — the sheet leaves BALANCE blank/0 for those instead
    // of a negative number, so (pcs - booked) summed across every
    // row would under-count Balance Stock PCS vs. what the sheet
    // (and the Category/Supplier tabs) actually show.
    totalBalance += r.balance;
    if      (r.status === 'AVAILABLE')    avail++;
    else if (r.status === 'OUT OF STOCK') outOfStock++;
    else if (r.status === 'LOW STOCK')    lowStock++;
    else if (r.status === 'OVER BOOKED')  overbooked++;
  });
  cards.totalProductionPCS = totalProd;
  cards.totalBookedPCS     = totalBooked;
  cards.balanceStock       = totalBalance;
  cards.availableCount     = avail;
  cards.outOfStockCount    = outOfStock;
  cards.lowStockCount      = lowStock;
  cards.overbookedCount    = overbooked;
  cards.totalDesigns       = pdRows.length;
  cards.totalSales         = totalSales;

  var totalPCS = 0, dealers = {};
  iomRows.forEach(function(r) {
    totalPCS += r.total;
    if (r.dealer) dealers[r.dealer] = true;
  });
  cards.totalPCS     = totalPCS;
  cards.totalDealers = Object.keys(dealers).length;

  var typeTotals = {};
  pdRows.forEach(function(r) {
    var nameUpper = r.designName.toUpperCase();
    var matched   = null;
    for (var k = 0; k < DESIGN_TYPE_KEYWORDS.length; k++) {
      if (nameUpper.indexOf(DESIGN_TYPE_KEYWORDS[k]) !== -1) { matched = DESIGN_TYPE_KEYWORDS[k]; break; }
    }
    if (!matched) return;
    typeTotals[matched] = (typeTotals[matched] || 0) + (r.booked || r.pcs || 0);
  });
  var topDesigns = Object.entries(typeTotals)
    .sort(function(a, b) { return b[1] - a[1]; })
    .slice(0, 5)
    .map(function(e) { return { name: e[0], qty: Math.round(e[1]) }; });
  if (topDesigns.length > 0) cards.topDesigns = topDesigns;

  return cards;
}

function getDefaultCards() {
  return {
    totalDealers:0, totalPCS:0, totalSales:0,
    totalDesigns:0, totalSuppliers:13, totalCategories:4,
    totalProductionPCS:0, totalBookedPCS:0, balanceStock:0,
    availableCount:0, outOfStockCount:0, lowStockCount:0, overbookedCount:0,
    topDesigns:[{name:'CHECKS',qty:0},{name:'PRINT',qty:0},{name:'STRIPES',qty:0}]
  };
}

function getDealerData(ss, iomRows) {
  var map = {};
  iomRows.forEach(function(r) {
    if (!r.dealer) return;
    if (!map[r.dealer]) map[r.dealer] = { name: r.dealer, pcs: 0, designs: {}, shirts: 0, amount: 0, pending: 0 };
    var d = map[r.dealer];
    d.pcs    += r.total;
    d.amount += r.amount;
    if (r.designNo) d.designs[r.designNo] = true;
    if (r.category.toUpperCase().indexOf('SHIRT') !== -1) d.shirts += r.total;
    if (r.status === 'PENDING') d.pending++;
  });
  return Object.values(map).map(function(d) {
    return { name: d.name, pcs: Math.round(d.pcs), designs: Object.keys(d.designs).length,
             shirts: Math.round(d.shirts), amount: Math.round(d.amount), pending: d.pending };
  });
}

function getExecutiveData(ss, iomRows, dealerMap) {
  dealerMap = dealerMap || {};
  var map = {};
  iomRows.forEach(function(r) {
    var execName   = r.executive || 'Unassigned';
    var dealerInfo = dealerMap[normalizeDealerKey_(r.dealer)] || {};
    var agent      = r.agent || dealerInfo.agent || 'Unassigned';
    var key        = execName + '||' + agent;
    if (!map[key]) map[key] = { name: execName, agent: agent, states: {}, pcs: 0, sales: 0 };
    map[key].pcs   += r.total;
    map[key].sales += r.amount;
    if (r.state) map[key].states[r.state] = true;
  });
  return Object.values(map).map(function(e) {
    return {
      name:  e.name,
      agent: e.agent,
      state: Object.keys(e.states).sort().join(', '),
      pcs:   Math.round(e.pcs),
      sales: Math.round(e.sales)
    };
  });
}

function getCategoryData(ss, pdRows) {
  var map = {};
  pdRows.forEach(function(r) {
    var cat = r.category || 'UNCATEGORIZED';
    if (!map[cat]) map[cat] = { name: cat, designs: 0, pcs: 0, stock: 0 };
    map[cat].designs++;
    map[cat].pcs   += r.booked;
    map[cat].stock += r.balance;
  });
  return Object.values(map).sort(function(a, b) { return b.pcs - a.pcs; });
}

function getSupplierData(ss, pdRows) {
  var map = {};
  pdRows.forEach(function(r) {
    var sup = r.supplier || 'Unassigned';
    if (!map[sup]) map[sup] = { name: sup, designs: 0, production: 0, booked: 0, balance: 0 };
    map[sup].designs++;
    map[sup].production += r.pcs;
    map[sup].booked      += r.booked;
    map[sup].balance     += r.balance;
  });
  return Object.values(map);
}

function getFocusAreaData(pdRows) {
  var designs = [];

  pdRows.forEach(function(r) {
    if (r.pcs <= 0)      return;
    if (r.balance <= 0)  return;

    var ratio = r.booked / r.pcs;
    if (ratio >= FOCUS_BOOKING_RATIO_THRESHOLD) return;

    designs.push({
      designNo:     r.designNo,
      designName:   r.designName || '',
      category:     r.category   || 'UNCATEGORIZED',
      supplier:     r.supplier   || '',
      pcs:          r.pcs,
      booked:       r.booked,
      balance:      r.balance,
      bookingPct:   Math.round(ratio * 1000) / 10,
      mrp:          r.mrp,
      idleValue:    Math.round(r.balance * r.mrp)
    });
  });

  designs.sort(function(a, b) { return b.idleValue - a.idleValue; });

  var catMap = {};
  designs.forEach(function(d) {
    if (!catMap[d.category]) {
      catMap[d.category] = { category: d.category, designCount: 0, totalPcs: 0, totalBooked: 0, totalBalance: 0, totalIdleValue: 0 };
    }
    var c = catMap[d.category];
    c.designCount++;
    c.totalPcs        += d.pcs;
    c.totalBooked      += d.booked;
    c.totalBalance     += d.balance;
    c.totalIdleValue   += d.idleValue;
  });

  var categories = Object.values(catMap).map(function(c) {
    return {
      category:       c.category,
      designCount:    c.designCount,
      totalBalance:   Math.round(c.totalBalance),
      totalIdleValue: Math.round(c.totalIdleValue),
      avgBookingPct:  c.totalPcs > 0 ? Math.round((c.totalBooked / c.totalPcs) * 1000) / 10 : 0
    };
  }).sort(function(a, b) { return b.totalIdleValue - a.totalIdleValue; });

  return {
    designs:      designs,
    categories:   categories,
    thresholdPct: Math.round(FOCUS_BOOKING_RATIO_THRESHOLD * 100)
  };
}

function getOrderData(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER); if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), result = [], tz = Session.getScriptTimeZone();

  var iomSheet    = ss.getSheetByName(CONFIG.SHEETS.INTERNAL_ORDER);
  var statusByRow = {};
  if (iomSheet) {
    var iomLastRow = iomSheet.getLastRow();
    if (iomLastRow >= CONFIG.DATA_START) {
      var iomStatusCol = iomSheet.getRange(
        CONFIG.DATA_START, CONFIG.IOM_COLS.ORDER_STATUS, iomLastRow - CONFIG.DATA_START + 1, 1
      ).getValues();
      for (var s = 0; s < iomStatusCol.length; s++) {
        statusByRow[s + CONFIG.DATA_START] = String(iomStatusCol[s][0] || '').trim();
      }
    }
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !String(row[CONFIG.OM_COLS.DEALER - 1]).trim()) continue;
    var dateVal = row[CONFIG.OM_COLS.DATE - 1];
    var dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, tz, 'dd-MMM-yyyy')
      : String(dateVal).substring(0, 10);
    var dateTs = (dateVal instanceof Date) ? dateVal.getTime() : 0;
    result.push({
      sr:        row[CONFIG.OM_COLS.SR_NO    - 1],
      date:      dateStr,
      dateTs:    dateTs,
      dealer:    String(row[CONFIG.OM_COLS.DEALER    - 1]).trim(),
      city:      String(row[CONFIG.OM_COLS.CITY      - 1] || '').trim(),
      state:     String(row[CONFIG.OM_COLS.STATE     - 1] || '').trim(),
      executive: String(row[CONFIG.OM_COLS.EXECUTIVE - 1] || '').trim(),
      agent:     String(row[CONFIG.OM_COLS.AGENT      - 1] || '').trim(),
      design:    String(row[CONFIG.OM_COLS.DESIGN_NO - 1]).trim(),
      total:     Number(row[CONFIG.OM_COLS.TOTAL     - 1]) || 0,
      mrp:       Number(row[CONFIG.OM_COLS.MRP       - 1]) || 0,
      status:    statusByRow[i + 1] || 'PENDING'
    });
  }
  result.sort(function(a, b) { return b.sr - a.sr; });
  return result;
}
// new added (30 july)
// function getOrderData(ss) {
//   var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDER_MASTER); if (!sheet) return [];
//   var data = sheet.getDataRange().getValues(), result = [], tz = Session.getScriptTimeZone();
//   var designMap = getDesignMapFast_();   // ← add this

//   // ...unchanged status lookup block...

//   for (var i = 1; i < data.length; i++) {
//     var row = data[i];
//     if (!row[0] || !String(row[CONFIG.OM_COLS.DEALER - 1]).trim()) continue;
//     // ...unchanged date logic...
//     var designNo = String(row[CONFIG.OM_COLS.DESIGN_NO - 1]).trim();
//     var designInfo = designMap[designNo] || {};
//     result.push({
//       sr:        row[CONFIG.OM_COLS.SR_NO    - 1],
//       date:      dateStr,
//       dateTs:    dateTs,
//       dealer:    String(row[CONFIG.OM_COLS.DEALER    - 1]).trim(),
//       city:      String(row[CONFIG.OM_COLS.CITY      - 1] || '').trim(),
//       state:     String(row[CONFIG.OM_COLS.STATE     - 1] || '').trim(),
//       executive: String(row[CONFIG.OM_COLS.EXECUTIVE - 1] || '').trim(),
//       agent:     String(row[CONFIG.OM_COLS.AGENT      - 1] || '').trim(),
//       design:    designNo,
//       total:     Number(row[CONFIG.OM_COLS.TOTAL     - 1]) || 0,
//       mrp:       Number(designInfo.mrp) || 0,   // ← was reading the empty column, now looks it up
//       status:    statusByRow[i + 1] || 'PENDING'
//     });
//   }
//   result.sort(function(a, b) { return b.sr - a.sr; });
//   return result;
// }
function getInternalOrderData(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INTERNAL_ORDER);
  if (!sheet) return { statusCount:{}, designCount:{}, supplierCount:{} };
  var data = sheet.getDataRange().getValues();
  var statusCount = {'LOW STOCK':0,'AVAILABLE':0,'Pending':0,'OVER BOOKED':0};
  var designCount = {}, supplierCount = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i]; if (!row[0]) continue;
    var status   = String(row[CONFIG.IOM_COLS.ORDER_STATUS - 1] || 'Pending').trim();
    var design   = String(row[CONFIG.IOM_COLS.DESIGN       - 1] || '').trim();
    var supplier = String(row[CONFIG.IOM_COLS.SUPPLIER     - 1] || '').trim();
    statusCount[status] = (statusCount[status] || 0) + 1;
    if (design)   designCount[design]     = (designCount[design]     || 0) + 1;
    if (supplier) supplierCount[supplier] = (supplierCount[supplier] || 0) + 1;
  }
  return { statusCount:statusCount, designCount:designCount, supplierCount:supplierCount };
}

function getProductionData(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUCTION);
  if (!sheet) return { statusCount:{}, bookingCount:{}, supplierPCS:{} };
  var data = sheet.getDataRange().getValues();
  var statusCount = {}, bookingCount = {}, supplierPCS = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i]; if (!row[0]) continue;
    var status  = String(row[CONFIG.PDM_COLS.STATUS         - 1] || 'AVAILABLE').trim();
    var booking = String(row[CONFIG.PDM_COLS.BOOKING_STATUS - 1] || '').trim();
    var sup     = String(row[CONFIG.PDM_COLS.SUPPLIER       - 1] || '').trim();
    var pcs     = parseFloat(row[CONFIG.PDM_COLS.PCS        - 1]) || 0;
    statusCount[status] = (statusCount[status] || 0) + 1;
    if      (booking.indexOf('OK')  !== -1) bookingCount['OK']  = (bookingCount['OK']  || 0) + 1;
    else if (booking.indexOf('LOW') !== -1) bookingCount['LOW'] = (bookingCount['LOW'] || 0) + 1;
    if (sup) supplierPCS[sup] = (supplierPCS[sup] || 0) + pcs;
  }
  return { statusCount:statusCount, bookingCount:bookingCount, supplierPCS:supplierPCS };
}

function getDealerMasterData(ss) {
  if (!DEALER_EXTERNAL_SHEET_ID) return { total:0, states:[], execMap:{}, execAgentList:[] };
  var externalSS;
  try {
    externalSS = SpreadsheetApp.openById(DEALER_EXTERNAL_SHEET_ID);
  } catch(err) {
    Logger.log('ERROR opening external dealer sheet for dashboard: ' + err.message);
    return { total:0, states:[], execMap:{}, execAgentList:[] };
  }
  var exSheet = externalSS.getSheetByName(DEALER_EXTERNAL_TAB);
  if (!exSheet) return { total:0, states:[], execMap:{}, execAgentList:[] };
  var lastRow = exSheet.getLastRow();
  if (lastRow < 2) return { total:0, states:[], execMap:{}, execAgentList:[] };

  var data = exSheet.getRange(
    2, 1, lastRow - 1, Math.max(exSheet.getLastColumn(), CONFIG.DM_COLS.EXECUTIVE, CONFIG.DM_COLS.PHONE)
  ).getValues();

  var stateMap = {}, execMap = {}, execAgentMap = {}, total = 0;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dealerName = String(row[CONFIG.DM_COLS.DEALER - 1] || '').trim();
    if (!dealerName) continue;
    total++;
    var state = String(row[CONFIG.DM_COLS.STATE     - 1] || 'Unknown').trim() || 'Unknown';
    var exec  = String(row[CONFIG.DM_COLS.EXECUTIVE - 1] || 'Unknown').trim() || 'Unknown';
    var agent = String(row[CONFIG.DM_COLS.AGENT     - 1] || 'Unknown').trim() || 'Unknown';
    stateMap[state] = (stateMap[state] || 0) + 1;
    execMap[exec]   = (execMap[exec]   || 0) + 1;

    var eaKey = exec + '||' + agent;
    if (!execAgentMap[eaKey]) execAgentMap[eaKey] = { executive: exec, agent: agent, count: 0, states: {} };
    execAgentMap[eaKey].count++;
    if (state) execAgentMap[eaKey].states[state] = true;
  }
  var states = Object.entries(stateMap).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 10);
  var execAgentList = Object.values(execAgentMap)
    .map(function(e) {
      return {
        executive: e.executive,
        agent:     e.agent,
        state:     Object.keys(e.states).sort().join(', '),
        count:     e.count
      };
    })
    .sort(function(a, b) { return b.count - a.count; });
  return { total:total, states:states, execMap:execMap, execAgentList:execAgentList };
}

function sendOverbookAlert_(designs) {
  var subject = '🔴 OVER BOOKED ALERT — Company name Order System';
  var body    = 'Dear Manager,\n\nThe following design(s) are OVER BOOKED:\n\n';
  designs.forEach(function(d, i) { body += (i + 1) + '. ' + d + '\n'; });
  body += '\n\nGenerated: ' + new Date().toLocaleString();
  try { MailApp.sendEmail({ to: CONFIG.ALERT_EMAIL, subject: subject, body: body }); }
  catch(err) { Logger.log('Email send failed: ' + err.message); }
}

function normalizeCode(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function columnLetterToIndex(letter) {
  letter = String(letter).trim().toUpperCase();
  var col = 0;
  for (var i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64);
  return col;
}

function syncAllDesignColors() {
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDER_MASTER);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var lastCol  = sheet.getLastColumn();
  var data     = sheet.getRange(2, CONFIG.OM_COLS.DESIGN_NO, lastRow - 1, 1).getValues();
  var colorMap = {};
  for (var i = 0; i < data.length; i++) {
    var design = String(data[i][0]).trim(); if (!design) continue;
    var color  = sheet.getRange(i + 2, 1).getBackground();
    if (!colorMap[design] && color !== '#ffffff') colorMap[design] = color;
  }
  for (var j = 0; j < data.length; j++) {
    var design2 = String(data[j][0]).trim();
    if (colorMap[design2]) sheet.getRange(j + 2, 1, 1, lastCol).setBackground(colorMap[design2]);
  }
}

function debugDealerLookup() {
  var map = buildDealerMapRaw_();
  var count = Object.keys(map).length;
  Logger.log('Dealer count found: ' + count);
  Logger.log('First 5 keys: ' + JSON.stringify(Object.keys(map).slice(0, 5)));
  
  var testName = '24 CARAT';
  var testKey = normalizeDealerKey_(testName);
  Logger.log('Looking up "' + testName + '" as key "' + testKey + '"');
  Logger.log('Found: ' + JSON.stringify(map[testKey] || 'NOT FOUND'));
}