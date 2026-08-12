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
// DATE NORMALIZER HELPER
// ================================================================

/**
 * Normalizes any date value (Date object, "14/07/2026", "2026-07-14", "Tue Jul 14...")
 * into standard ISO "YYYY-MM-DD" string format.
 */
function normalizeDateStr(val) {
  if (!val) return "";

  // 1. If it's a native Apps Script Date object
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    var yyyy = val.getFullYear();
    var mm = String(val.getMonth() + 1); if (mm.length < 2) mm = "0" + mm;
    var dd = String(val.getDate());       if (dd.length < 2) dd = "0" + dd;
    return yyyy + "-" + mm + "-" + dd;
  }

  var str = String(val).trim();
  if (!str) return "";

  // 2. If already YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

  // 3. If DD/MM/YYYY or DD-MM-YYYY
  var dmyMatch = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (dmyMatch) {
    var d = String(dmyMatch[1]); if (d.length < 2) d = "0" + d;
    var m = String(dmyMatch[2]); if (m.length < 2) m = "0" + m;
    var y = Number(dmyMatch[3]);
    if (y > 2400) y = y - 543; // Convert Thai BE to AD
    return y + "-" + m + "-" + d;
  }

  // 4. Try parsing general date string
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    var yyyy = parsed.getFullYear();
    var mm = String(parsed.getMonth() + 1); if (mm.length < 2) mm = "0" + mm;
    var dd = String(parsed.getDate());       if (dd.length < 2) dd = "0" + dd;
    return yyyy + "-" + mm + "-" + dd;
  }

  return str;
}

// ================================================================
// HTTP HANDLERS
// ================================================================

/**
 * Handle GET requests
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
        var date = normalizeDateStr(e.parameter.date);
        if (!date) throw new Error("Missing or invalid 'date' parameter.");
        response = { status: "success", data: getReportByDate(date), spreadsheetUrl: ssUrl };
        break;

      case "listReports":
        response = { status: "success", data: listAllReports(), spreadsheetUrl: ssUrl };
        break;

      case "queryReports":
        var from = normalizeDateStr(e.parameter.from || "");
        var to   = normalizeDateStr(e.parameter.to   || "");
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
        var targetDate = normalizeDateStr(postData.date);
        deleteReport(targetDate);
        response = { status: "success", message: "ลบรายงานวันที่ " + targetDate + " เรียบร้อย", spreadsheetUrl: ssUrl };
        break;

      case "saveConfig":
        saveConfig(postData.config);
        response = { status: "success", message: "บันทึกการตั้งค่าเรียบร้อย", spreadsheetUrl: ssUrl };
        break;

      case "importSilomRevenue":
        var result = importSilomRevenue(postData.date, postData.revenueData);
        response = { status: "success", message: result.message, date: result.date, spreadsheetUrl: ssUrl };
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
  var date = normalizeDateStr(report.date);

  if (!date) throw new Error("Missing or invalid report date.");

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var existingRow = -1;

  for (var i = 1; i < headerData.length; i++) {
    var d = normalizeDateStr(headerData[i][0]);
    if (d === date) {
      existingRow = i + 1;
      break;
    }
  }

  var now = new Date().toISOString();
  if (existingRow > 0) {
    headerSheet.getRange(existingRow, 1).setValue(date);
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

function getReportByDate(targetDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var date = normalizeDateStr(targetDateStr);
  if (!date) return null;

  var report = {
    date: date,
    status: "DRAFT",
    savedAt: "",
    createdBy: "",
    data: {},
    debtors: []
  };
  var foundAny = false;

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  if (headerSheet) {
    var headerData = headerSheet.getDataRange().getValues();
    for (var i = 1; i < headerData.length; i++) {
      var d = normalizeDateStr(headerData[i][0]);
      if (d === date) {
        report.status = headerData[i][1] || "DRAFT";
        report.savedAt = headerData[i][2] || "";
        report.createdBy = headerData[i][3] || "";
        foundAny = true;
        break;
      }
    }
  }

  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  if (linesSheet) {
    var linesData = linesSheet.getDataRange().getValues();
    for (var i = 1; i < linesData.length; i++) {
      var d = normalizeDateStr(linesData[i][0]);
      if (d === date) {
        foundAny = true;
        var itemId = String(linesData[i][1]).trim();
        var group = String(linesData[i][2]).trim();
        var coopRev = parseNum(linesData[i][3]);
        var coopExp = parseNum(linesData[i][4]);
        var djRev   = parseNum(linesData[i][5]);
        var djExp   = parseNum(linesData[i][6]);

        if (group === "revenue") {
          var keyCoop = itemId + "_COOP_REV";
          if (coopRev !== 0 || report.data[keyCoop] === undefined) {
            report.data[keyCoop] = coopRev;
          }
          var keyDj = itemId + "_DJ_REV";
          if (djRev !== 0 || report.data[keyDj] === undefined) {
            report.data[keyDj] = djRev;
          }
        } else {
          var keyCoop = itemId + "_COOP_EXP";
          if (coopExp !== 0 || report.data[keyCoop] === undefined) {
            report.data[keyCoop] = coopExp;
          }
          var keyDj = itemId + "_DJ_EXP";
          if (djExp !== 0 || report.data[keyDj] === undefined) {
            report.data[keyDj] = djExp;
          }
        }
      }
    }
  }

  var debtorSheet = ss.getSheetByName(SHEET_DEBTORS);
  if (debtorSheet) {
    var debtorData = debtorSheet.getDataRange().getValues();
    for (var i = 1; i < debtorData.length; i++) {
      var d = normalizeDateStr(debtorData[i][0]);
      if (d === date) {
        foundAny = true;
        report.debtors.push({
          name: String(debtorData[i][1]).trim(),
          coopRev: Number(debtorData[i][2]) || 0,
          coopExp: Number(debtorData[i][3]) || 0
        });
      }
    }
  }

  var paySheet = ss.getSheetByName(SHEET_PAYMENT_CHANNELS);
  if (paySheet) {
    var payData = paySheet.getDataRange().getValues();
    for (var i = 1; i < payData.length; i++) {
      var d = normalizeDateStr(payData[i][0]);
      if (d === date) {
        foundAny = true;
        var channelId = String(payData[i][1]).trim();
        var coopAmt = Number(payData[i][2]) || 0;
        var djAmt   = Number(payData[i][3]) || 0;
        if (coopAmt) report.data[channelId + "_COOP"] = coopAmt;
        if (djAmt)   report.data[channelId + "_DJ"] = djAmt;
      }
    }
  }

  return foundAny ? report : null;
}

// ================================================================
// LIST ALL REPORTS
// ================================================================

function listAllReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportsMap = {};

  var sheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var date = normalizeDateStr(data[i][0]);
      if (!date) continue;
      reportsMap[date] = {
        date: date,
        status: data[i][1] || "DRAFT",
        savedAt: data[i][2] || "",
        createdBy: data[i][3] || ""
      };
    }
  }

  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  if (linesSheet) {
    var linesData = linesSheet.getDataRange().getValues();
    for (var i = 1; i < linesData.length; i++) {
      var date = normalizeDateStr(linesData[i][0]);
      if (!date) continue;
      if (!reportsMap[date]) {
        reportsMap[date] = {
          date: date,
          status: "DRAFT",
          savedAt: "",
          createdBy: "Silom POS Auto-Import"
        };
      }
    }
  }

  var reports = [];
  Object.keys(reportsMap).forEach(function(d) {
    reports.push(reportsMap[d]);
  });
  reports.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return reports;
}

// ================================================================
// QUERY REPORTS
// ================================================================

function queryReports(from, to, unit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var normFrom = normalizeDateStr(from);
  var normTo   = normalizeDateStr(to);

  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var datesInRange = {};

  for (var i = 1; i < headerData.length; i++) {
    var date = normalizeDateStr(headerData[i][0]);
    if (!date) continue;
    if (normFrom && date < normFrom) continue;
    if (normTo   && date > normTo) continue;
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
    var date = normalizeDateStr(linesData[i][0]);
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

function deleteReport(targetDateStr) {
  var date = normalizeDateStr(targetDateStr);
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

function clearRowsByDate(sheetName, targetDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var date = normalizeDateStr(targetDateStr);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = normalizeDateStr(data[i][0]);
    if (rowDate === date) {
      sheet.deleteRow(i + 1);
    }
  }
}

function testSetup() {
  ensureSheetsExist();
  Logger.log("✅ All sheets created successfully!");
  Logger.log("Spreadsheet URL: " + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}

/**
 * Imports/merges Silom POS daily revenue figures into SbacDiaryRPT without overwriting existing expenses or debtors.
 */
function importSilomRevenue(rawDate, revenueData) {
  var date = normalizeDateStr(rawDate);
  if (!date) throw new Error("Invalid date for Silom import");

  ensureSheetsExist();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ensure DailyReports entry exists
  var headerSheet = ss.getSheetByName(SHEET_DAILY_REPORTS);
  var headerData = headerSheet.getDataRange().getValues();
  var existingRow = -1;
  for (var i = 1; i < headerData.length; i++) {
    if (normalizeDateStr(headerData[i][0]) === date) {
      existingRow = i + 1;
      break;
    }
  }

  var now = new Date().toISOString();
  if (existingRow === -1) {
    headerSheet.appendRow([date, "DRAFT", now, "Silom POS Auto-Import"]);
  }

  // 2. Merge into ReportLines sheet
  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  var linesData = linesSheet.getDataRange().getValues();

  var itemRowIndices = {};
  for (var r = 1; r < linesData.length; r++) {
    if (normalizeDateStr(linesData[r][0]) === date) {
      var itemId = String(linesData[r][1]).trim();
      if (!itemRowIndices[itemId]) itemRowIndices[itemId] = [];
      itemRowIndices[itemId].push(r + 1);
    }
  }

  var silomItems = ["REV_COOP_SALES", "REV_CANTEEN_RICE", "REV_BAKERY_CONSIGN", "REV_CONSIGNMENT", "REV_UNIFORM"];

  silomItems.forEach(function(itemId) {
    var coopRevVal = parseNum(revenueData[itemId]);

    if (itemRowIndices[itemId] && itemRowIndices[itemId].length > 0) {
      itemRowIndices[itemId].forEach(function(rowIndex) {
        linesSheet.getRange(rowIndex, 4).setValue(coopRevVal);
        linesSheet.getRange(rowIndex, 8).setValue("Silom POS Auto-Import");
      });
    } else {
      linesSheet.appendRow([date, itemId, "revenue", coopRevVal, 0, 0, 0, "Silom POS Auto-Import"]);
    }
  });

  return { status: "success", message: "Silom revenue imported for " + date, date: date };
}

function parseNum(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  var str = String(val).replace(/,/g, "").trim();
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function testApiGetReport() {
  var report = getReportByDate("2026-08-11");
  Logger.log("Report 2026-08-11: " + JSON.stringify(report));
}

function debugInspectSheetDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var linesSheet = ss.getSheetByName(SHEET_REPORT_LINES);
  if (linesSheet) {
    var linesData = linesSheet.getDataRange().getValues();
    Logger.log("--- ReportLines Total Rows: " + linesData.length + " ---");
    for (var i = 1; i < linesData.length; i++) {
      var rawCell = linesData[i][0];
      var norm = normalizeDateStr(rawCell);
      if (norm.indexOf("2026-08") !== -1) {
        Logger.log("Row " + (i+1) + " | Date: '" + norm + "' | Item: " + linesData[i][1] + " | CoopRev: " + linesData[i][3]);
      }
    }
  } else {
    Logger.log("No ReportLines sheet found!");
  }
}
