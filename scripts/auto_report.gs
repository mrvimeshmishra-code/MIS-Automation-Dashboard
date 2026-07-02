// ============================================================
// MIS Auto Report Generator
// Author  : Vimesh Kumar Mishra
// Email   : mr.vimeshmishra@gmail.com
// GitHub  : https://github.com/mrvimeshmishra-code
// Version : 1.0
// Description: Automates daily MIS report generation and
//              sends email alerts when KPIs exceed thresholds
// ============================================================

/**
 * SETUP INSTRUCTIONS:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code
 * 4. Run setDailyTrigger() once to schedule automation
 * 5. Reports will auto-generate every day at 9 AM
 *
 * SHEET STRUCTURE REQUIRED:
 * Sheet 1 Name: "Raw_Data"
 *   Columns: Date | Agent_Name | Revenue | Status | Region
 * Sheet 2 Name: "MIS_Report"
 *   (This will be auto-created and updated)
 */

// ---- CONFIGURATION ----
const CONFIG = {
  rawDataSheet: "Raw_Data",
  reportSheet:  "MIS_Report",
  alertEmail:   "mr.vimeshmishra@gmail.com",
  exceptionStatus: "Exception",
  reportHour:   9,   // 9 AM daily trigger
  headerColor:  "#1a3c6e",
  exceptionHighlightColor: "#FFE0E0"
};

// ---- MAIN FUNCTION ----
function generateDailyMISReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet   = ss.getSheetByName(CONFIG.rawDataSheet);
  const reportSheet = ss.getSheetByName(CONFIG.reportSheet)
                     || ss.insertSheet(CONFIG.reportSheet);

  if (!dataSheet) {
    Logger.log("ERROR: Sheet '" + CONFIG.rawDataSheet + "' not found!");
    return;
  }

  // Step 1: Clear previous report
  reportSheet.clearContents();
  reportSheet.clearFormats();

  // Step 2: Read raw data
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);

  // Step 3: Calculate KPIs
  let totalRevenue   = 0;
  let totalRecords   = rows.length;
  let exceptionCount = 0;
  let regionMap      = {};

  rows.forEach(function(row) {
    const revenue = parseFloat(row[2]) || 0;
    const status  = row[3];
    const region  = row[4] || "Unknown";

    totalRevenue += revenue;
    if (status === CONFIG.exceptionStatus) exceptionCount++;
    regionMap[region] = (regionMap[region] || 0) + revenue;
  });

  // Step 4: Write Summary Header
  const today = new Date().toLocaleDateString("en-IN");
  reportSheet.getRange(1, 1, 1, 5).setValues([[
    "📊 MIS Report — " + today, "", "", "", ""
  ]]);
  reportSheet.getRange(1, 1, 1, 5)
    .merge()
    .setBackground(CONFIG.headerColor)
    .setFontColor("white")
    .setFontSize(13)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Step 5: Write KPI Summary
  const kpiHeaders = [["Metric", "Value"]];
  const kpiData = [
    ["Report Date",          today],
    ["Total Records",        totalRecords],
    ["Total Revenue (₹)",    "₹" + totalRevenue.toLocaleString("en-IN")],
    ["Total Exceptions",     exceptionCount],
    ["Exception Rate (%)",   ((exceptionCount / totalRecords) * 100).toFixed(2) + "%"]
  ];

  reportSheet.getRange(3, 1, 1, 2).setValues(kpiHeaders)
    .setBackground("#2E4057").setFontColor("white").setFontWeight("bold");
  reportSheet.getRange(4, 1, kpiData.length, 2).setValues(kpiData);

  // Highlight exception row if high
  if (exceptionCount > 0) {
    reportSheet.getRange(7, 1, 1, 2).setBackground(CONFIG.exceptionHighlightColor);
  }

  // Step 6: Region-wise Revenue Breakdown
  const regionStart = 4 + kpiData.length + 2;
  reportSheet.getRange(regionStart, 1, 1, 2).setValues([["Region", "Revenue (₹)"]])
    .setBackground("#2E4057").setFontColor("white").setFontWeight("bold");

  let r = regionStart + 1;
  Object.keys(regionMap).forEach(function(region) {
    reportSheet.getRange(r, 1, 1, 2).setValues([[region, "₹" + regionMap[region].toLocaleString("en-IN")]]);
    r++;
  });

  // Step 7: Auto-resize columns
  reportSheet.autoResizeColumns(1, 2);

  Logger.log("✅ MIS Report Generated Successfully: " + today);

  // Step 8: Send alert email if exceptions found
  if (exceptionCount > 0) {
    sendExceptionAlert(exceptionCount, totalRevenue, today);
  }
}

// ---- EMAIL ALERT FUNCTION ----
function sendExceptionAlert(exceptions, revenue, date) {
  const exceptionRate = "N/A";
  const subject = "⚠️ MIS Alert — " + exceptions + " Exceptions Found | " + date;
  const htmlBody =
    "<h2 style='color:#cc0000;'>⚠️ MIS Exception Alert</h2>" +
    "<table border='1' cellpadding='8' cellspacing='0'>" +
    "<tr><th>Metric</th><th>Value</th></tr>" +
    "<tr><td>Report Date</td><td>" + date + "</td></tr>" +
    "<tr><td>Total Revenue</td><td>₹" + revenue.toLocaleString("en-IN") + "</td></tr>" +
    "<tr><td style='color:red;'>Total Exceptions</td><td style='color:red;font-weight:bold;'>" + exceptions + "</td></tr>" +
    "</table><br/>" +
    "<p>Please review the MIS Dashboard immediately.</p>" +
    "<p><i>— Auto MIS System | Vimesh Kumar Mishra</i></p>";

  MailApp.sendEmail({
    to:       CONFIG.alertEmail,
    subject:  subject,
    htmlBody: htmlBody
  });

  Logger.log("📧 Exception alert email sent to: " + CONFIG.alertEmail);
}

// ---- TRIGGER SETUP (Run once manually) ----
function setDailyTrigger() {
  // Delete existing triggers first
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  // Create new daily trigger at 9 AM
  ScriptApp.newTrigger("generateDailyMISReport")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.reportHour)
    .create();

  Logger.log("✅ Daily trigger set for " + CONFIG.reportHour + ":00 AM");
}
