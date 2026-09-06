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

function doGet(e) {
  return ContentService.createTextOutput("Collier Transit Action backend is active.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Sheet not found" }))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader("Access-Control-Allow-Origin", "*");
  }

  try {
    var payload = {};
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      // fallback if payload is parsed via parameter
      payload = e.parameter || {};
    }

    // Honeypot check
    if (payload.website) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader("Access-Control-Allow-Origin", "*");
    }

    // Server-side validations
    const zipRegex = /^341(0[1-9]|1[0-9]|20|3[7-9]|4[0-3]|45|46)$/;
    if (!payload.zip || !zipRegex.test(payload.zip)) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Invalid input" }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader("Access-Control-Allow-Origin", "*");
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!payload.email || !emailRegex.test(payload.email)) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Invalid input" }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader("Access-Control-Allow-Origin", "*");
    }

    const validRoles = ["Hospitality / Service Worker", "Local Rideshare / Taxi Driver", "Concerned Taxpayer", "Family of Senior / Transit Dependent"];
    if (!payload.role || !validRoles.includes(payload.role)) {
      return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": "Invalid input" }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader("Access-Control-Allow-Origin", "*");
    }

    const sanitizeInput = (str) => {
      if (!str) return "";
      const stringified = String(str);
      // Prevent formula/CSV injection
      if (/^[=+\-@]/.test(stringified)) {
        return "'" + stringified;
      }
      return stringified;
    };

    const email = sanitizeInput(payload.email || "");
    const zip = sanitizeInput(payload.zip || "");
    const role = sanitizeInput(payload.role || "");
    const timestamp = new Date().toISOString();

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
