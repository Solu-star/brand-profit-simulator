const SPREADSHEET_ID = "1ss7Agd7xUP8HW8-JL8hJikLY3QXTKKwJEni3ENb3ZfM";
const SHEET_NAME = "利益計算 ";
const TOKEN = "CHANGE_ME_TO_YOUR_SECRET";
const MAX_ROWS_PER_REQUEST = 300;
const COLUMN_COUNT = 13;

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");
    if (payload.token !== TOKEN) {
      return jsonResponse({ ok: false, error: "invalid token" });
    }

    const rows = normalizeRows(payload.rows);
    if (rows.length === 0) {
      return jsonResponse({ ok: false, error: "no rows" });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
      if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

      const startRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(startRow, 1, rows.length, COLUMN_COUNT).setValues(rows);
      SpreadsheetApp.flush();
      return jsonResponse({ ok: true, appended: rows.length, startRow });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function normalizeRows(inputRows) {
  if (!Array.isArray(inputRows)) return [];
  return inputRows
    .slice(0, MAX_ROWS_PER_REQUEST)
    .map((row) => Array.isArray(row) ? row.slice(0, COLUMN_COUNT) : [])
    .filter((row) => row.length === COLUMN_COUNT && row[0] !== "")
    .map((row) => row.map((value, index) => {
      if (index === 0 || index === 12) return String(value || "");
      const number = Number(value);
      return Number.isFinite(number) ? number : "";
    }));
}

function jsonResponse(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
