// ================================================================
// NATHJI AGENCIES — OPTIMIZED VERSION
// ================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 WhatsApp Tool')
    .addItem('Open Generator', 'showSidebar')
    .addItem('📊 Open Outstanding Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Sync Dealers from Tally', 'setupDealerMaster')
    .addItem('🔄 Sync Ageing Data', 'updateAndProtectAgeing')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('WhatsApp Generator')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function sortMonthHeaders(headers) {
  const monthMap = { Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11 };
  return headers.sort((a, b) => {
    const [am, ay] = a.toString().trim().split(' ');
    const [bm, by] = b.toString().trim().split(' ');
    return new Date(Number("20"+ay), monthMap[am], 1) - new Date(Number("20"+by), monthMap[bm], 1);
  });
}

function setupDealerMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tally = ss.getSheetByName('Tally_Import');
  const master = ss.getSheetByName('DEALER_MASTER');
  if (!tally || !master) return;

  const data = tally.getRange(2, 1, tally.getLastRow() - 1, 8).getValues();
  let tallyMap = {};
  data.forEach(r => {
    if (!r[2]) return;
    let key = r[2].toString().toLowerCase().trim();
    tallyMap[key] = { name: r[2], executive: r[7] };
  });

  const masterData = master.getRange(2, 1, master.getLastRow() - 1, 1).getValues();
  let masterMap = {};
  masterData.forEach((r, i) => {
    if (r[0]) masterMap[r[0].toString().toLowerCase().trim()] = i + 2;
  });

  let newRows = [];
  Object.keys(tallyMap).forEach(key => {
    const dealer = tallyMap[key];
    if (masterMap[key]) {
      if (dealer.executive) master.getRange(masterMap[key], 4).setValue(dealer.executive);
    } else {
      newRows.push([dealer.name, "", "", dealer.executive || "", "", "", "", ""]);
    }
  });

  if (newRows.length) {
    master.getRange(master.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
  }
}

function updateAndProtectAgeing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tally = ss.getSheetByName("Tally_Import2");
  const ageing = ss.getSheetByName("AGEING_ANALYSIS");
  const dealerMaster = ss.getSheetByName("DEALER_MASTER");

  if (!tally || !ageing || !dealerMaster) {
    SpreadsheetApp.getUi().alert("Required sheets missing!");
    return;
  }

  const tallyData = tally.getDataRange().getValues();
  if (tallyData.length < 2) return;

  const headers = tallyData[1];
  let rawMonths = headers.slice(2).filter(h => h && h.toString().includes(" ") && h !== "Grand Total");
  rawMonths = sortMonthHeaders(rawMonths);

  // ✅ FIX: Build index map ONCE using original header positions
  let monthIndexMap = {};
  headers.forEach((h, i) => { monthIndexMap[h] = i; });

  const masterData = dealerMaster.getDataRange().getValues();
  let dealerMap = {};
  for (let i = 1; i < masterData.length; i++) {
    let row = masterData[i];
    if (!row[0]) continue;
    let cleanKey = row[0].toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim();
    dealerMap[cleanKey] = {
      executive: row[3] || "",
      phone: row[6] ? row[6].toString().replace(/\D/g, '') : "",
      city: row[2] || ""
    };
  }

  ageing.clearContents();

  const finalHeader = ["Dealer Name", "Executive", "Phone", "City", "Last Update", ...rawMonths, "Status"];
  ageing.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);

  let finalData = [];
  const today = new Date();

  for (let i = 2; i < tallyData.length; i++) {
    const row = tallyData[i];
    const dealerName = row[0];
    if (!dealerName) continue;

    const key = dealerName.toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim();
    const master = dealerMap[key] || {};
    let newRow = [dealerName, master.executive || "", master.phone || "", master.city || "", today];

    rawMonths.forEach(month => {
      newRow.push(row[monthIndexMap[month]] || 0);
    });

    newRow.push("");
    finalData.push(newRow);
  }

  if (finalData.length) {
    ageing.getRange(2, 1, finalData.length, finalData[0].length).setValues(finalData);
  }

  ageing.autoResizeColumns(1, finalHeader.length);
  SpreadsheetApp.getUi().alert("✅ AGEING_ANALYSIS Updated Successfully");
}

// ================= GET DATA — OPTIMIZED =================
function getData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AGEING_ANALYSIS");
  if (!sheet) return { dealers: [], executives: [], dynamicMonths: [] };

  const data = sheet.getDataRange().getValues(); // ✅ Single batch read
  if (data.length < 2) return { dealers: [], executives: [], dynamicMonths: [] };

  const headers = data[0];

  // ✅ FIX: Detect month start dynamically instead of hardcoding index 5
  const fixedCols = ["Dealer Name", "Executive", "Phone", "City", "Last Update"];
  const monthStartIndex = fixedCols.length; // = 5, but now safely anchored

  // ✅ FIX: Build a position map for month columns BEFORE sorting
  // so row[colIndex] lookups use actual column positions, not assumed order
  let monthColMap = {}; // month label → actual column index in sheet
  for (let c = monthStartIndex; c < headers.length; c++) {
    const h = headers[c];
    if (h && h !== "Status") monthColMap[h] = c;
  }

  // Sort month labels chronologically
  let monthHeaders = sortMonthHeaders(Object.keys(monthColMap));

  let dealers = [];
  let execSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const exec = row[1] || "";
    if (exec) execSet.add(exec);

    let months = {};
    monthHeaders.forEach(m => {
      // ✅ FIX: Use actual column index, not assumed sorted position
      months[m] = row[monthColMap[m]] || 0;
    });

    dealers.push({
      name: row[0],
      executive: exec,
      // ✅ FIX: Clean phone once here, not in sidebar
      phone: row[2] ? row[2].toString().replace(/\D/g, '') : "",
      months: months
    });
  }

  return {
    dealers: dealers,
    executives: [...execSet].sort(),
    dynamicMonths: monthHeaders
  };
}

// ================= SEND WHATSAPP — OPTIMIZED =================
function sendWhatsAppFromSidebar(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AGEING_ANALYSIS");

  const dealer = data.dealer;
  const executiveName = data.executive;

  const executivePhones = {
    "ABHISHEK":          "919648896777",
    "AJAY":              "918601696777",
    "AMIT":              "919648896777",
    "UMAKANT":           "919871550046",   // ✅ added country code
    "SHEERY":            "919837789007",   // ✅ added country code
    "MAHIPAL SHIKHAWAT": "919887513856"    // ✅ added country code
  };

  const execPhone = executivePhones[(executiveName || "").toUpperCase().trim()] || "";
  const time = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "hh:mm a");

  if (execPhone) {
    const notifyMsg = `CRM contacted your dealer ${dealer} at ${time}`;
    const notifyLink = `https://wa.me/${execPhone}?text=` + encodeURIComponent(notifyMsg);
    const html = HtmlService.createHtmlOutput(
      `<script>window.open("${notifyLink}", "_blank");</script>`
    );
    SpreadsheetApp.getUi().showModelessDialog(html, "...");
  } else {
    // ✅ Warn if executive not mapped — helps catch missing numbers
    SpreadsheetApp.getUi().alert(`⚠️ No phone mapped for executive: "${executiveName}"`);
  }

  // ✅ Find dealer row
  const finder = sheet.createTextFinder(dealer).matchEntireCell(true);
  const found = finder.findNext();

  if (found) {
    // ✅ Dynamically find the "Status" column from headers — no hardcoding
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusColIndex = headers.indexOf("Status") + 1; // +1 = 1-based

    if (statusColIndex > 0) {
      sheet.getRange(found.getRow(), statusColIndex).setValue(`Sent by CRM at ${time}`);
    } else {
      Logger.log("⚠️ 'Status' column not found in headers.");
    }
  } else {
    Logger.log(`⚠️ Dealer not found in sheet: ${dealer}`);
  }
}

// ================================================================
// OUTSTANDING DASHBOARD
// ================================================================

// Opens the dashboard as a modal dialog from the Sheet's menu.
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1280)
    .setHeight(840);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Outstanding Dashboard');
}

// Lets the same page be deployed as a standalone web app
// (Deploy > New deployment > Web app), so it can be opened in a
// browser without going through the Sheet — handy on a phone.
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Outstanding Dashboard — Nathji Agencies')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// "Mon YY" / "Mon YYYY" -> Date at the 1st of that month.
function monthLabelToDate_(label) {
  const monthMap = { Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11 };
  const parts = label.toString().trim().split(' ');
  if (parts.length !== 2) return null;
  const mon = monthMap[parts[0]];
  if (mon === undefined) return null;
  const yr = parts[1].length === 2 ? 2000 + Number(parts[1]) : Number(parts[1]);
  return new Date(yr, mon, 1);
}

// Buckets a month column against "today" — current month is 0-30,
// one month back is 31-60, two back is 61-90, older is 90+.
function ageingBucket_(monthDate, today) {
  if (!monthDate) return 'b0_30';
  const diff = (today.getFullYear() - monthDate.getFullYear()) * 12 +
               (today.getMonth() - monthDate.getMonth());
  if (diff <= 0) return 'b0_30';
  if (diff === 1) return 'b31_60';
  if (diff === 2) return 'b61_90';
  return 'b90plus';
}

// Aggregates AGEING_ANALYSIS into everything the dashboard needs:
// KPI totals, ageing buckets, executive/city rollups, a monthly
// trend, and the full dealer list with a per-dealer ageing split.
function getDashboardData() {
  const base = getData(); // reuse the existing dealers/executives/months reader
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AGEING_ANALYSIS");
  const raw = sheet ? sheet.getDataRange().getValues() : [];
  const headers = raw[0] || [];
  const cityCol = headers.indexOf("City");
  const statusCol = headers.indexOf("Status");

  // Dealer name (lowercased) -> city / status, since getData() doesn't carry them.
  let extra = {};
  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row[0]) continue;
    extra[row[0].toString().toLowerCase().trim()] = {
      city: cityCol > -1 ? (row[cityCol] || "") : "",
      status: statusCol > -1 ? (row[statusCol] || "") : ""
    };
  }

  const today = new Date();
  const monthDates = {};
  base.dynamicMonths.forEach(m => { monthDates[m] = monthLabelToDate_(m); });

  let totalOutstanding = 0;
  let ageingTotals = { b0_30: 0, b31_60: 0, b61_90: 0, b90plus: 0 };
  let execMap = {};
  let cityMap = {};
  let monthlyTotals = {};
  base.dynamicMonths.forEach(m => { monthlyTotals[m] = 0; });

  let dealers = base.dealers.map(d => {
    const key = d.name.toString().toLowerCase().trim();
    const info = extra[key] || {};
    let total = 0;
    let buckets = { b0_30: 0, b31_60: 0, b61_90: 0, b90plus: 0 };

    base.dynamicMonths.forEach(m => {
      const val = Number(d.months[m]) || 0;
      total += val;
      monthlyTotals[m] += val;
      if (val !== 0) {
        buckets[ageingBucket_(monthDates[m], today)] += val;
      }
    });

    totalOutstanding += total;
    ageingTotals.b0_30 += buckets.b0_30;
    ageingTotals.b31_60 += buckets.b31_60;
    ageingTotals.b61_90 += buckets.b61_90;
    ageingTotals.b90plus += buckets.b90plus;

    const execName = d.executive || "Unassigned";
    if (!execMap[execName]) execMap[execName] = { name: execName, total: 0, count: 0 };
    execMap[execName].total += total;
    if (total > 0) execMap[execName].count += 1;

    const cityName = info.city || "Unknown";
    if (!cityMap[cityName]) cityMap[cityName] = { name: cityName, total: 0, count: 0 };
    cityMap[cityName].total += total;
    if (total > 0) cityMap[cityName].count += 1;

    return {
      name: d.name,
      executive: d.executive || "",
      phone: d.phone,
      city: info.city || "",
      status: info.status || "",
      total: total,
      buckets: buckets
    };
  });

  dealers.sort((a, b) => b.total - a.total);

  const byExecutive = Object.keys(execMap)
    .map(k => execMap[k])
    .sort((a, b) => b.total - a.total)
    .map(e => ({ name: e.name, total: e.total, count: e.count,
                 pct: totalOutstanding ? (e.total / totalOutstanding * 100) : 0 }));

  const byCity = Object.keys(cityMap)
    .map(k => cityMap[k])
    .sort((a, b) => b.total - a.total);

  const monthlyTotalsArr = base.dynamicMonths.map(m => ({ month: m, total: monthlyTotals[m] }));
  const topOverdue = dealers.filter(d => d.total > 0).slice(0, 15);

  return {
    generatedAt: today.toISOString(),
    totalOutstanding: totalOutstanding,
    dealerCount: dealers.filter(d => d.total > 0).length,
    totalDealers: dealers.length,
    ageingTotals: ageingTotals,
    byExecutive: byExecutive,
    byCity: byCity,
    topOverdue: topOverdue,
    dealers: dealers,
    monthlyTotals: monthlyTotalsArr,
    months: base.dynamicMonths
  };
}
