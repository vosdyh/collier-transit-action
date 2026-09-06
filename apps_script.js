// Google Apps Script to accept POST requests and append to a Google Sheet
// To use this:
// 1. Create a new Google Sheet.
// 2. Add headers to the first row: Timestamp, Email, Zip Code, Persona/Role
// 3. Go to Extensions > Apps Script.
// 4. Paste this code, save, and click Deploy > New Deployment.
// 5. Select "Web app".
// 6. Execute as: "Me"
// 7. Who has access: "Anyone"
// 8. Deploy and copy the Web App URL.

const SHEET_NAME = "Sheet1";

function doOptions(e) {
  // Handle CORS preflight requests
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Sheet not found" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*");
  }

  try {
    let payload;

    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "No payload found" }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader("Access-Control-Allow-Origin", "*");
    }

    const email = payload.email || "";
    const zip = payload.zip || "";
    const role = payload.role || "";
    const timestamp = new Date();

    // Headers must match: Timestamp, Email, Zip Code, Persona/Role
    sheet.appendRow([timestamp, email, zip, role]);

    return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*");

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*");
  }
}
