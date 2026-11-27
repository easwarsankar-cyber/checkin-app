// Google Apps Script backend for the Session Check-In app.
// Setup:
//   1. Create a Google Sheet. Rename the first tab to "CheckIns".
//   2. Extensions > Apps Script. Replace the default code with this file.
//   3. Deploy > New deployment > type: Web app
//        Execute as: Me
//        Who has access: Anyone
//   4. Copy the Web app URL into CONFIG.APPS_SCRIPT_URL in index.html.

const SHEET_NAME = "CheckIns";
const HEADERS = [
  "Received At", "Check-in Time (UTC)", "First Name", "Last Name", "Title",
  "Company", "Email", "Phone", "Message", "Session Location", "Session Date"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Write the header row once if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }

    sheet.appendRow([
      new Date(),
      data.timestamp || "",
      data.firstName || "",
      data.lastName || "",
      data.title || "",
      data.company || "",
      data.email || "",
      data.phone || "",
      data.message || "",
      data.sessionLocation || "",
      data.sessionDate || ""
    ]);

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// Handy for checking the deployment works by opening the URL in a browser
function doGet() {
  return jsonResponse({ ok: true, message: "Check-in endpoint is running." });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
