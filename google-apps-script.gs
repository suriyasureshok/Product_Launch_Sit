/**
 * Kitronix – Contact Form → Google Sheets integration
 * ─────────────────────────────────────────────────────
 * SETUP STEPS
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Delete any existing code and paste this entire file
 *  3. Save the project (Ctrl+S), then click Deploy → New Deployment
 *  4. Choose type: Web App
 *     • Execute as:  Me
 *     • Who has access: Anyone
 *  5. Click Deploy → copy the web-app URL
 *  6. In index.html replace 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL'
 *     with the copied URL
 * ─────────────────────────────────────────────────────
 * The sheet will be created automatically with these columns:
 *   Timestamp | Name | Email | Company | Message
 */

var SHEET_NAME = 'Contact Submissions';  // Change if you prefer a different tab name

function doPost(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Create the sheet + header row the first time
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Company', 'Message']);

      // Style the header row
      var headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#FFE401');          // Kitronix yellow
      headerRange.setFontColor('#000000');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 160);
      sheet.setColumnWidth(3, 220);
      sheet.setColumnWidth(4, 180);
      sheet.setColumnWidth(5, 400);
    }

    var params = e.parameter;

    sheet.appendRow([
      params.date    || new Date().toLocaleString(),
      params.name    || '',
      params.email   || '',
      params.company || '',
      params.message || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Quick test — run this function manually in the Apps Script editor to verify the sheet is created
function testInsert() {
  doPost({
    parameter: {
      name   : 'Test User',
      email  : 'test@example.com',
      company: 'Test Co.',
      message: 'This is a manual test entry.',
      date   : new Date().toLocaleString()
    }
  });
}
