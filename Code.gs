// Google Apps Script backend for the Session Check-In app.
// Setup:
//   1. Create a Google Sheet. Rename the first tab to "CheckIns".
//   2. Extensions > Apps Script. Replace the default code with this file.
//   3. Deploy > New deployment > type: Web app
//        Execute as: Me
//        Who has access: Anyone
//   4. Copy the Web app URL into CONFIG.APPS_SCRIPT_URL in index.html.

const SHEET_NAME = "CheckIns";
const PROFILE_SHEET_NAME = "Profiles";

const HEADERS = [
  "Received At", "Check-in Time (UTC)", "First Name", "Last Name", "Title",
  "Company", "Email", "Phone", "Message", "Session Location", "Session Date"
];

const PROFILE_HEADERS = [
  "Last Updated", "First Name", "Last Name", "Preferred Location", "Email", "Phone"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "profile") {
      return saveProfile(data);
    }

    // ----- Check-in (unchanged logic) -----
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

// ----- Profile (one row per email; updates in place on re-save) -----
function saveProfile(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROFILE_SHEET_NAME) || ss.insertSheet(PROFILE_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(PROFILE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PROFILE_HEADERS.length).setFontWeight("bold");
  }

  const row = [
    new Date(),
    data.firstName || "",
    data.lastName || "",
    data.location || "",
    data.email || "",
    data.phone || ""
  ];

  // Find an existing row with the same email (column E)
  const email = String(data.email || "").trim().toLowerCase();
  let rowNumber = -1;
  if (email && sheet.getLastRow() > 1) {
    const emails = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues();
    const idx = emails.findIndex(r => String(r[0]).trim().toLowerCase() === email);
    if (idx >= 0) rowNumber = idx + 2;
  }

  if (rowNumber > 0) {
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return jsonResponse({ ok: true, action: "profile" });
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
