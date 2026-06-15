// ============================================================
// MY FINANCE TRACKER — Google Apps Script
// Paste this entire file into your Apps Script editor
// then Deploy as a Web App (Execute as: Me, Access: Anyone)
// ============================================================

const SHEET_NAME = "Transactions";
const HEADERS = ["ID", "Date", "Time", "Type", "Account", "Category", "Description", "Amount", "Synced At"];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "addEntry") {
      return addEntry(data.entry);
    }
    if (data.action === "syncAll") {
      return syncAll(data.entries);
    }
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "Finance tracker is running!" })
  ).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Add headers with formatting
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#6c63ff");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 110);
    sheet.setColumnWidth(6, 130);
    sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(8, 90);
    sheet.setColumnWidth(9, 160);
  }
  return sheet;
}

function addEntry(entry) {
  const sheet = getOrCreateSheet();
  const now = new Date().toISOString();
  const row = [
    entry.id || now,
    entry.date || "",
    entry.time || "",
    entry.type || "",
    entry.account || "",
    entry.category || "",
    entry.description || "",
    entry.amount || 0,
    now
  ];

  // Check for duplicate ID
  const existingIds = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  if (existingIds.includes(row[0])) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "duplicate", message: "Entry already exists" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  sheet.appendRow(row);

  // Color rows by type
  const lastRow = sheet.getLastRow();
  if (entry.type === "income") {
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#e8f5e9");
  } else {
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#ffffff");
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: "success", message: "Entry added", row: lastRow })
  ).setMimeType(ContentService.MimeType.JSON);
}

function syncAll(entries) {
  if (!entries || !entries.length) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "success", message: "Nothing to sync" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getOrCreateSheet();
  const existingIds = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
    : [];

  let added = 0;
  const now = new Date().toISOString();

  entries.forEach(entry => {
    if (!existingIds.includes(entry.id)) {
      const row = [
        entry.id || now,
        entry.date || "",
        entry.time || "",
        entry.type || "",
        entry.account || "",
        entry.category || "",
        entry.description || "",
        entry.amount || 0,
        now
      ];
      sheet.appendRow(row);
      const lastRow = sheet.getLastRow();
      if (entry.type === "income") {
        sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#e8f5e9");
      }
      added++;
    }
  });

  return ContentService.createTextOutput(
    JSON.stringify({ status: "success", message: `Synced ${added} new entries`, added })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// OPTIONAL: Run this function manually once to create a
// "Monthly Summary" sheet with automatic totals
// ============================================================
function createSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Monthly Summary");
  if (!sheet) {
    sheet = ss.insertSheet("Monthly Summary");
  }
  sheet.clearContents();
  const headers = ["Month", "Personal Income", "Personal Expenses", "Personal Net", "CEK Income", "CEK Expenses", "CEK Net"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#6c63ff").setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("Monthly Summary sheet created! It will auto-populate as you add entries.");
}
