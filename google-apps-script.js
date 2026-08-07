/**
 * Google Apps Script — บธว. Daily Revenue & Expense Report Backend
 * 
 * SETUP INSTRUCTIONS:
 * ──────────────────────────────────────────────────────────────────
 * 1. Create a new Google Spreadsheet (name it e.g. "บธว. Daily Diary Report").
 * 2. Go to Extensions > Apps Script.
 * 3. Delete any default code and paste this entire script into Code.gs.
 * 4. Click Deploy > New deployment.
 *    - Select type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - Click Deploy, authorize permissions, copy the Web App URL.
 * 5. Paste the Web App URL into the frontend Settings panel (ตั้งค่า API URL).
 *
 * SHEET STRUCTURE (auto-created on first run):
 * ──────────────────────────────────────────────────────────────────
 * Sheet "DailyReports"   : date | status | savedAt | createdBy
 * Sheet "ReportLines"    : date | itemId | group | coopRev | coopExp | djRev | djExp | note
 * Sheet "Debtors"        : date | name | coopRev | coopExp
 * Sheet "PaymentChannels": date | channelId | coopAmount | djAmount
 * Sheet "Config"         : key | value
 */

// ================================================================
// SHEET NAMES
// ================================================================
var SHEET_DAILY_REPORTS    = "DailyReports";
var SHEET_REPORT_LINES     = "ReportLines";
var SHEET_DEBTORS          = "Debtors";
var SHEET_PAYMENT_CHANNELS = "PaymentChannels";
var SHEET_CONFIG           = "Config";

// ================================================================
// HTTP HANDLERS
// ================================================================

/**
 * Handle GET requests
 * 
 * Supported actions:
 *   ?action=getReport&date=YYYY-MM-DD       → Get single report by date
 *   ?action=listReports                      → List all report headers
 *   ?action=queryReports&from=...&to=...     → Query reports in date range
 *   ?action=getConfig                        → Get categories/settings
 *   ?action=ping                             → Health check & return Sheet URL
 *   (default)                                → List all reports
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "listReports";
  var response = {};

  try {
    ensureSheetsExist();
    var ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    switch (action) {
      case "ping":
        response = { 
          status: "success", 
          message: "บธว. Diary API is ready.", 
          spreadsheetUrl: ssUrl,
          timestamp: new Date().toISOString() 
        };
        break;

      case "getReport":
        var date = e.parameter.date;
        if (!date) throw new Error("Missing 'date' parameter.");
        response = { status: "success", data: getReportByDate(date), spreadsheetUrl: ssUrl };
        break;

      case "listReports":
        response = { status: "success", data: listAllReports(), spreadsheetUrl: ssUrl };
        break;

      case "queryReports":
        var from = e.parameter.from || "";
        var to   = e.parameter.to   || "";
        var unit = e.parameter.unit || "ALL";
        response = { status: "success", data: queryReports(from, to, unit), spreadsheetUrl: ssUrl };
        break;

      case "getConfig":
        response = { status: "success", data: getConfig(), spreadsheetUrl: ssUrl };
        break;

      default:
        response = { status: "success", data: listAllReports(), spreadsheetUrl: ssUrl };
    }
  } catch (err) {
    response = { status: "error", message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle POST requests
 * 
 * Supported actions (in JSON body):
 *   { action: "saveReport",   report: {...} }   → Save/update a daily report
 *   { action: "deleteReport", date: "..." }     → Delete a report by date
 *   { action: "saveConfig",   config: {...} }   → Save categories config
 */
function doPost(e) {
  var response = {};

  try {
    ensureSheetsExist();
    var ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;

    switch (action) {
      case "saveReport":
        var result = saveReport(postData.report);
        response = { status: "success", message: result.message, date: result.date, spreadsheetUrl: ssUrl };
        break;

      case "deleteReport":
        deleteReport(postData.date);
        response = { status: "success", message: "ลบรายงานวันที่ " + postData.date + " เรียบร้อย", spreadsheetUrl: ssUrl };
        break;

      case "saveConfig":
        saveConfig(postData.config);
        response = { status: "success", message: "บันทึกการตั้งค่าเรียบร้อย", spreadsheetUrl: ssUrl };
        break;

      default:
        throw new Error("Unknown action: " + action);
    }
  } catch (err) {
    response = { status: "error", message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
// SHEET INITIALIZATION
// ================================================================

function ensureSheetsExist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_DAILY_REPORTS)) {
    var s = ss.insertSheet(SHEET_DAILY_REPORTS);
    s.appendRow(["date", "status", "savedAt", "createdBy"]);
    s.getRange(1, 1, 1, 4).setFontWeight("bold");
    s.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_REPORT_LINES)) {
    var s = ss.insertSheet(SHEET_REPORT_LINES);
    s.appendRow(["date", "itemId", "group", "coopRev", "coopExp", "djRev", "djExp", "note"]);
    s.getRange(1, 1, 1, 8).setFontWeight("bold");
    s.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_DEBTORS)) {
    var s = ss.insertSheet(SHEET_DEBTORS);
    s.appendRow(["date", "name", "coopRev", "coopExp"]);
    s.getRange(1, 1, 1, 4).setFontWeight("bold");
    s.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_PAYMENT_CHANNELS)) {
    var s = ss.insertSheet(SHEET_PAYMENT_CHANNELS);
    s.appendRow(["date", "channelId", "coopAmount", "djAmount"]);
    s.getRange(1, 1, 1, 4).setFontWeight("bold");
    s.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_CONFIG)) {
    var s = ss.insertSheet(SHEET_CONFIG);
    s.appendRow(["key", "value"]);
    s.getRange(1, 1, 1, 2).setFontWeight("bold");
    s.setFrozenRows(1);
  }
}

// ================================================================
// SAVE REPORT
// ================================================================

function saveReport(report) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = report.date;

  if (!date) throw new Error("Missing report date.");

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var existingRow = -1;

  for (var i = 1; i < headerData.length; i++) {
    if (String(headerData[i][0]).trim() === date) {
      existingRow = i + 1;
      break;
    }
  }

  var now = new Date().toISOString();
  if (existingRow > 0) {
    headerSheet.getRange(existingRow, 2).setValue(report.status || "DRAFT");
    headerSheet.getRange(existingRow, 3).setValue(now);
  } else {
    headerSheet.appendRow([date, report.status || "DRAFT", now, report.createdBy || ""]);
  }

  clearRowsByDate(SHEET_REPORT_LINES, date);
  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  var data = report.data || {};

  var lineItems = {};
  var paymentItems = {};

  Object.keys(data).forEach(function(key) {
    var val = data[key];
    if (val === 0 || val === "" || val === null) return;

    if (key.indexOf("PAY_") === 0) {
      if (key.endsWith("_COOP")) {
        var channelId = key.replace(/_COOP$/, "");
        if (!paymentItems[channelId]) paymentItems[channelId] = { coopAmount: 0, djAmount: 0 };
        paymentItems[channelId].coopAmount = val;
      } else if (key.endsWith("_DJ")) {
        var channelId = key.replace(/_DJ$/, "");
        if (!paymentItems[channelId]) paymentItems[channelId] = { coopAmount: 0, djAmount: 0 };
        paymentItems[channelId].djAmount = val;
      }
    } else if (key.indexOf("DEBTOR_") === 0) {
      // Skip debtors here
    } else {
      var suffixes = ["_COOP_REV", "_COOP_EXP", "_DJ_REV", "_DJ_EXP"];
      for (var si = 0; si < suffixes.length; si++) {
        var suffix = suffixes[si];
        if (key.endsWith(suffix)) {
          var itemId = key.replace(suffix, "");
          if (!lineItems[itemId]) lineItems[itemId] = { coopRev: 0, coopExp: 0, djRev: 0, djExp: 0, group: "" };
          if (suffix === "_COOP_REV") lineItems[itemId].coopRev = val;
          else if (suffix === "_COOP_EXP") lineItems[itemId].coopExp = val;
          else if (suffix === "_DJ_REV")   lineItems[itemId].djRev = val;
          else if (suffix === "_DJ_EXP")   lineItems[itemId].djExp = val;

          if (itemId.indexOf("REV_") === 0)     lineItems[itemId].group = "revenue";
          else if (itemId.indexOf("EXP_") === 0) lineItems[itemId].group = "opex";
          else if (itemId.indexOf("COGS_") === 0) lineItems[itemId].group = "cogs";
          else if (itemId.indexOf("PAYABLE_") === 0) lineItems[itemId].group = "payable";
          break;
        }
      }
    }
  });

  var lineRows = [];
  Object.keys(lineItems).forEach(function(itemId) {
    var item = lineItems[itemId];
    lineRows.push([date, itemId, item.group, item.coopRev, item.coopExp, item.djRev, item.djExp, ""]);
  });
  if (lineRows.length > 0) {
    linesSheet.getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, 8).setValues(lineRows);
  }

  clearRowsByDate(SHEET_DEBTORS, date);
  var debtorSheet = ss.getSheetByName(SHEET_DEBTORS);
  var debtors = report.debtors || [];
  var debtorRows = [];

  debtors.forEach(function(d) {
    if (d.name || d.coopRev || d.coopExp) {
      debtorRows.push([date, d.name || "", d.coopRev || 0, d.coopExp || 0]);
    }
  });
  if (debtorRows.length > 0) {
    debtorSheet.getRange(debtorSheet.getLastRow() + 1, 1, debtorRows.length, 4).setValues(debtorRows);
  }

  clearRowsByDate(SHEET_PAYMENT_CHANNELS, date);
  var paySheet = ss.getSheetByName(SHEET_PAYMENT_CHANNELS);
  var payRows = [];

  Object.keys(paymentItems).forEach(function(channelId) {
    var ch = paymentItems[channelId];
    if (ch.coopAmount || ch.djAmount) {
      payRows.push([date, channelId, ch.coopAmount, ch.djAmount]);
    }
  });
  if (payRows.length > 0) {
    paySheet.getRange(paySheet.getLastRow() + 1, 1, payRows.length, 4).setValues(payRows);
  }

  return { message: "บันทึกรายงานวันที่ " + date + " สำเร็จ", date: date };
}

// ================================================================
// GET REPORT BY DATE
// ================================================================

function getReportByDate(date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var report = null;

  for (var i = 1; i < headerData.length; i++) {
    if (String(headerData[i][0]).trim() === date) {
      report = {
        date: date,
        status: headerData[i][1],
        savedAt: headerData[i][2],
        createdBy: headerData[i][3],
        data: {},
        debtors: []
      };
      break;
    }
  }

  if (!report) return null;

  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  var linesData = linesSheet.getDataRange().getValues();
  for (var i = 1; i < linesData.length; i++) {
    if (String(linesData[i][0]).trim() === date) {
      var itemId = linesData[i][1];
      var group = linesData[i][2];
      var coopRev = Number(linesData[i][3]) || 0;
      var coopExp = Number(linesData[i][4]) || 0;
      var djRev   = Number(linesData[i][5]) || 0;
      var djExp   = Number(linesData[i][6]) || 0;

      if (group === "revenue") {
        if (coopRev) report.data[itemId + "_COOP_REV"] = coopRev;
        if (djRev)   report.data[itemId + "_DJ_REV"] = djRev;
      } else {
        if (coopExp) report.data[itemId + "_COOP_EXP"] = coopExp;
        if (djExp)   report.data[itemId + "_DJ_EXP"] = djExp;
      }
    }
  }

  var debtorSheet = ss.getSheetByName(SHEET_DEBTORS);
  var debtorData = debtorSheet.getDataRange().getValues();
  for (var i = 1; i < debtorData.length; i++) {
    if (String(debtorData[i][0]).trim() === date) {
      report.debtors.push({
        name: debtorData[i][1],
        coopRev: Number(debtorData[i][2]) || 0,
        coopExp: Number(debtorData[i][3]) || 0
      });
    }
  }

  var paySheet = ss.getSheetByName(SHEET_PAYMENT_CHANNELS);
  var payData = paySheet.getDataRange().getValues();
  for (var i = 1; i < payData.length; i++) {
    if (String(payData[i][0]).trim() === date) {
      var channelId = payData[i][1];
      var coopAmt = Number(payData[i][2]) || 0;
      var djAmt   = Number(payData[i][3]) || 0;
      if (coopAmt) report.data[channelId + "_COOP"] = coopAmt;
      if (djAmt)   report.data[channelId + "_DJ"] = djAmt;
    }
  }

  return report;
}

// ================================================================
// LIST ALL REPORTS
// ================================================================

function listAllReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var data = sheet.getDataRange().getValues();
  var reports = [];

  for (var i = 1; i < data.length; i++) {
    var date = String(data[i][0]).trim();
    if (!date) continue;
    reports.push({
      date: date,
      status: data[i][1] || "DRAFT",
      savedAt: data[i][2] || "",
      createdBy: data[i][3] || ""
    });
  }

  reports.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return reports;
}

// ================================================================
// QUERY REPORTS
// ================================================================

function queryReports(from, to, unit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var datesInRange = {};

  for (var i = 1; i < headerData.length; i++) {
    var date = String(headerData[i][0]).trim();
    if (!date) continue;
    if (from && date < from) continue;
    if (to && date > to) continue;
    datesInRange[date] = {
      date: date,
      status: headerData[i][1] || "DRAFT",
      coopRev: 0, coopExp: 0, djRev: 0, djExp: 0
    };
  }

  if (Object.keys(datesInRange).length === 0) return [];

  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  var linesData = linesSheet.getDataRange().getValues();
  for (var i = 1; i < linesData.length; i++) {
    var date = String(linesData[i][0]).trim();
    if (!datesInRange[date]) continue;
    var group = linesData[i][2];

    if (group === "revenue") {
      datesInRange[date].coopRev += Number(linesData[i][3]) || 0;
      datesInRange[date].djRev   += Number(linesData[i][5]) || 0;
    } else {
      datesInRange[date].coopExp += Number(linesData[i][4]) || 0;
      datesInRange[date].djExp   += Number(linesData[i][6]) || 0;
    }
  }

  var results = Object.keys(datesInRange).map(function(d) { return datesInRange[d]; });
  results.sort(function(a, b) { return a.date.localeCompare(b.date); });
  return results;
}

// ================================================================
// DELETE REPORT
// ================================================================

function deleteReport(date) {
  if (!date) throw new Error("Missing date to delete.");
  clearRowsByDate(SHEET_DAILY_REPORTS, date);
  clearRowsByDate(SHEET_REPORT_LINES, date);
  clearRowsByDate(SHEET_DEBTORS, date);
  clearRowsByDate(SHEET_PAYMENT_CHANNELS, date);
}

// ================================================================
// CONFIG (Categories)
// ================================================================

function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "categories") {
      try {
        return JSON.parse(data[i][1]);
      } catch (e) {
        return null;
      }
    }
  }
  return null;
}

function saveConfig(config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  var data = sheet.getDataRange().getValues();
  var configJson = JSON.stringify(config);
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "categories") {
      sheet.getRange(i + 1, 2).setValue(configJson);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow(["categories", configJson]);
  }
}

function clearRowsByDate(sheetName, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === date) {
      sheet.deleteRow(i + 1);
    }
  }
}

function testSetup() {
  ensureSheetsExist();
  Logger.log("✅ All sheets created successfully!");
  Logger.log("Spreadsheet URL: " + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}
