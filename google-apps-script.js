// ============================================================
// MY FINANCE TRACKER — Google Apps Script (Two-Way Sync v2)
// Paste this entire file into your Apps Script editor,
// replacing everything that was there before.
// Then Deploy as a Web App (Execute as: Me, Access: Anyone)
// ============================================================
//
// HOW THIS WORKS:
// - The full app state (entries, projects, suppliers, budgets,
//   gold/silver purchases, everything) is stored as ONE JSON
//   blob in a hidden "AppState" sheet, tagged with a timestamp.
// - Every device that opens the app PULLS this blob on load and
//   MERGES it with whatever is stored locally, using each
//   record's unique ID — so nothing is lost or duplicated.
// - A human-readable "Transactions" sheet is also kept in sync
//   automatically, purely so you can browse/export your data
//   normally in Google Sheets. It is NOT the source of truth —
//   the AppState blob is.
// ============================================================

const STATE_SHEET = "AppState";
const LOG_SHEET = "Transactions";
const LOG_HEADERS = ["ID", "Date", "Time", "Type", "Account", "Category", "Project", "Supplier/Client", "Description", "Amount"];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "pushState") return pushState(data.state, data.deviceId);
    if (data.action === "pullState") return pullState();
    return jsonOut({ status: "error", message: "Unknown action" });
  } catch (err) {
    return jsonOut({ status: "error", message: err.message });
  }
}

function doGet(e) {
  // Allow pulling via GET too (some environments restrict POST)
  if (e && e.parameter && e.parameter.action === "pullState") {
    return pullState();
  }
  return jsonOut({ status: "ok", message: "Finance tracker sync is running!" });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getStateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STATE_SHEET);
    sheet.getRange(1,1,1,3).setValues([["UpdatedAt","DeviceId","StateJSON"]]);
    sheet.hideSheet();
  }
  return sheet;
}

// ── PULL: any device calls this on load to get the latest known state ──
function pullState() {
  const sheet = getStateSheet();
  if (sheet.getLastRow() < 2) {
    return jsonOut({ status: "empty" });
  }
  const row = sheet.getRange(sheet.getLastRow(), 1, 1, 3).getValues()[0];
  return jsonOut({ status: "ok", updatedAt: row[0], deviceId: row[1], state: JSON.parse(row[2]) });
}

// ── PUSH: a device sends its full (already-merged) state up ──
function pushState(state, deviceId) {
  const sheet = getStateSheet();
  const now = new Date().toISOString();
  // Always overwrite row 2 — we only ever keep ONE current snapshot
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([now, deviceId||"unknown", JSON.stringify(state)]);
  } else {
    sheet.getRange(2,1,1,3).setValues([[now, deviceId||"unknown", JSON.stringify(state)]]);
  }
  // Also refresh the human-readable Transactions log for browsing in Sheets
  try { refreshTransactionLog(state); } catch(err) {}
  return jsonOut({ status: "success", updatedAt: now });
}

// ── Rebuilds the readable Transactions sheet from the latest entries ──
function refreshTransactionLog(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
  }
  sheet.clearContents();
  const headerRange = sheet.getRange(1,1,1,LOG_HEADERS.length);
  headerRange.setValues([LOG_HEADERS]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#6c63ff");
  headerRange.setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  const entries = (state.entries||[]).slice().sort((a,b)=> (a.date<b.date?1:-1));
  if (entries.length===0) return;

  const rows = entries.map(en => [
    en.id || "",
    en.date || "",
    en.time || "",
    en.type || "",
    en.account || "",
    en.category || "",
    en.project || "",
    en.supplier || en.client || "",
    en.description || "",
    en.amount || 0
  ]);
  sheet.getRange(2,1,rows.length,LOG_HEADERS.length).setValues(rows);

  // Color income rows light green for quick scanning
  for (let i=0;i<entries.length;i++) {
    if (entries[i].type==='income') {
      sheet.getRange(i+2,1,1,LOG_HEADERS.length).setBackground("#e8f5e9");
    }
  }
  sheet.autoResizeColumns(1, LOG_HEADERS.length);
}
