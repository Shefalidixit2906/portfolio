// ================================================================
// NATHJI AGENCIES — SALES DASHBOARD  (Code.gs)
// ================================================================
// SETUP:
//  1. Extensions → Apps Script in your Google Sheet
//  2. Replace ALL of Code.gs with this entire file
//  3. Make sure Index.html is also in place
//  4. Save both files (Ctrl+S)
//  5. Run → testDashboard first — check Execution Log for record count
//  6. Deploy → New Deployment → Web App
//     Execute as: Me  |  Who has access: Anyone
//  7. After any edit: Deploy → Manage Deployments → Edit → New Version
// ================================================================

// ── Sheet names ─────────────────────────────────────────────────
var SALES_SHEET  = "SALES_MASTER";
var DEALER_SHEET = "DEALER_MASTER";

// ── SALES_MASTER columns (1-based) ──────────────────────────────
//  1=Sr.No  2=Period  3=Category  4=Dealer Name  5=Executive  6=Quantity  7=Amount
var SM_PERIOD_COL = 2;
var SM_CAT_COL    = 3;
var SM_DEALER_COL = 4;
var SM_EXEC_COL   = 5;
var SM_QTY_COL    = 6;
var SM_AMT_COL    = 7;

// ── DEALER_MASTER columns (1-based) ─────────────────────────────
var DM_DEALER_COL = 1;
var DM_EXEC_COL   = 4;

// ── Month key/label lookup ───────────────────────────────────────
var MONTH_META = {
  1:  { key: "2025-01", label: "Jan 25" },
  2:  { key: "2025-02", label: "Feb 25" },
  3:  { key: "2025-03", label: "Mar 25" },
  4:  { key: "2025-04", label: "Apr 25" },
  5:  { key: "2025-05", label: "May 25" },
  6:  { key: "2025-06", label: "Jun 25" },
  7:  { key: "2025-07", label: "Jul 25" },
  8:  { key: "2025-08", label: "Aug 25" },
  9:  { key: "2025-09", label: "Sep 25" },
  10: { key: "2025-10", label: "Oct 25" },
  11: { key: "2025-11", label: "Nov 25" },
  12: { key: "2025-12", label: "Dec 25" }
};

// ── Text → month number map ──────────────────────────────────────
var TEXT_MONTH_MAP = {
  "JAN":1,"JANUARY":1,
  "FEB":2,"FEBRUARY":2,
  "MAR":3,"MARCH":3,
  "APR":4,"APRIL":4,
  "MAY":5,
  "JUN":6,"JUNE":6,
  "JUL":7,"JULY":7,
  "AUG":8,"AUGUST":8,
  "SEP":9,"SEPT":9,"SEPTEMBER":9,
  "OCT":10,"OCTOBER":10,
  "NOV":11,"NOVEMBER":11,
  "DEC":12,"DECEMBER":12
};

// ================================================================
// parsePeriod
// ================================================================
function parsePeriod(raw) {
  if (!raw) return null;

  if (raw instanceof Date) {
    var mo = raw.getMonth() + 1;
    return MONTH_META[mo] || null;
  }

  var s = raw.toString().trim();

  var isoMatch = s.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (isoMatch) {
    var mo = parseInt(isoMatch[2], 10);
    return MONTH_META[mo] || null;
  }

  var upper = s.toUpperCase().replace(/\s+/g, ' ').trim();
  var word  = upper.split(' ')[0];
  if (TEXT_MONTH_MAP[word] !== undefined) {
    return MONTH_META[TEXT_MONTH_MAP[word]] || null;
  }

  for (var key in TEXT_MONTH_MAP) {
    if (upper.indexOf(key) === 0) {
      return MONTH_META[TEXT_MONTH_MAP[key]] || null;
    }
  }

  return null;
}

// ================================================================
// parseAmount
// ================================================================
function parseAmount(raw) {
  if (!raw) return 0;
  if (typeof raw === 'number') return raw;
  var cleaned = raw.toString().replace(/[₹,\s]/g, '');
  return parseFloat(cleaned) || 0;
}

// ================================================================
function doGet() {
  var template = HtmlService.createTemplateFromFile("Index");
  try {
    var data = getDashboardData();
    template.data = JSON.stringify(data);
  } catch (e) {
    template.data = JSON.stringify({
      records: [], execs: [], cats: [], months: [], ml: {},
      error: e.message
    });
  }
  return template
    .evaluate()
    .setTitle("Nathji Agencies — Sales Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================================================================
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Build dealer → exec map from DEALER_MASTER ────────
  var dealerExecMap = {};
  var dmSh = ss.getSheetByName(DEALER_SHEET);
  if (dmSh) {
    var dmData = dmSh.getDataRange().getValues();
    for (var i = 1; i < dmData.length; i++) {
      var dName = dmData[i][DM_DEALER_COL - 1];
      var dExec = dmData[i][DM_EXEC_COL  - 1];
      if (dName && dName.toString().trim()) {
        var dKey = dName.toString().trim().toUpperCase();
        if (dExec && dExec.toString().trim()) {
          dealerExecMap[dKey] = dExec.toString().trim().toUpperCase().replace(/\s+/g, ' ');
        }
      }
    }
    Logger.log("✅ DEALER_MASTER loaded: " + Object.keys(dealerExecMap).length + " dealers");
  } else {
    Logger.log("⚠️ Sheet not found: [" + DEALER_SHEET + "]");
  }

  // ── Step 2: Read all records from SALES_MASTER ────────────────
  var smSh = ss.getSheetByName(SALES_SHEET);
  if (!smSh) {
    throw new Error("Sheet not found: [" + SALES_SHEET + "]. Check the sheet tab name.");
  }

  var smData = smSh.getDataRange().getValues();
  Logger.log("📋 SALES_MASTER rows (incl. header): " + smData.length);

  var records  = [];
  var execSet  = {};
  var catSet   = {};
  var monthSet = {};
  var ml       = {};
  var skipped  = 0;

  for (var row = 1; row < smData.length; row++) {
    var r       = smData[row];
    var period  = r[SM_PERIOD_COL  - 1];
    var cat     = r[SM_CAT_COL     - 1];
    var dealer  = r[SM_DEALER_COL  - 1];
    var execRaw = r[SM_EXEC_COL    - 1];
    var qty     = r[SM_QTY_COL     - 1];
    var amt     = r[SM_AMT_COL     - 1];

    // Skip blank / header-repeat rows
    if (!dealer || dealer.toString().trim() === "")     { skipped++; continue; }
    if (!cat    || cat.toString().trim()    === "")     { skipped++; continue; }
    if (cat.toString().trim().toUpperCase() === "CATEGORY NAME") { skipped++; continue; }
    if (cat.toString().trim().toUpperCase() === "CATEGORY") { skipped++; continue; }

    var periodInfo = parsePeriod(period);
    if (!periodInfo) {
      Logger.log("⚠️ Row " + (row + 1) + ": unrecognised period [" + period + "] — skipped");
      skipped++;
      continue;
    }

    var dealerKey = dealer.toString().trim().toUpperCase();

    // ✅ FIX: .trim() prevents " T-SHIRTS" and "TRAVEL PANTS " being counted separately
    var catClean  = cat.toString().trim().toUpperCase().replace(/\s+/g, ' ');

    var qtyVal    = parseFloat(qty) || 0;
    var amtVal    = parseAmount(amt);

    var execClean;
    if (execRaw && execRaw.toString().trim()) {
      execClean = execRaw.toString().trim().toUpperCase().replace(/\s+/g, ' ');
    } else {
      execClean = dealerExecMap[dealerKey] || "UNKNOWN";
    }

    monthSet[periodInfo.key] = true;
    ml[periodInfo.key]       = periodInfo.label;
    catSet[catClean]         = true;
    execSet[execClean]       = true;

    records.push({
      m: periodInfo.key,
      d: dealerKey,
      e: execClean,
      q: qtyVal,
      a: amtVal,
      c: catClean
    });
  }

  Logger.log("✅ Records loaded: " + records.length + " | Skipped: " + skipped);

  var sortedMonths = Object.keys(monthSet).sort();
  var execList = Object.keys(execSet)
    .filter(function(e) { return e !== "UNKNOWN"; })
    .sort();

  return {
    records : records,
    execs   : execList,
    cats    : Object.keys(catSet).sort(),
    months  : sortedMonths,
    ml      : ml
  };
}

// ================================================================
// getDataJson — called by google.script.run from the browser
// Returns the dashboard data as a JSON string (not an object)
// ================================================================
function getDataJson() {
  return JSON.stringify(getDashboardData());
}

// ================================================================
// TEST FUNCTION
// ================================================================
function testDashboard() {
  var data = getDashboardData();
  Logger.log("=================================");
  Logger.log("✅ Total records: "  + data.records.length);
  Logger.log("📅 Months found: "   + JSON.stringify(data.months));
  Logger.log("📅 Month labels: "   + JSON.stringify(data.ml));
  Logger.log("👤 Executives: "     + JSON.stringify(data.execs));
  Logger.log("🏷️  Categories: "    + JSON.stringify(data.cats));
  Logger.log("=================================");

  // Print qty per category to verify counts
  var catTotals = {};
  data.records.forEach(function(r) {
    catTotals[r.c] = (catTotals[r.c] || 0) + r.q;
  });
  Logger.log("📦 Qty by category: " + JSON.stringify(catTotals));

  if (data.records.length === 0) {
    Logger.log("❌ NO DATA — Possible reasons:");
    Logger.log("   1. Sheet tab name doesn't match: [" + SALES_SHEET + "]");
    Logger.log("   2. Period column values are unrecognised");
    Logger.log("   3. Dealer or Category columns are blank");
  } else {
    Logger.log("🔢 Sample [0]:   " + JSON.stringify(data.records[0]));
    Logger.log("✅ Data looks good — safe to deploy!");
  }
}

// ================================================================
// UTILITY: List all sheet names
// ================================================================
function listSheetNames() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(s) {
    Logger.log("[" + s.getName() + "] — " + s.getName().length + " chars");
  });
}