// Google Apps Script — deploy as Web App (Anyone, even anonymous)
// Paste this in https://script.google.com → New Project

const SHEET_ID = '1Yyrk-S1NG4qHyQXzU3oKjKT-SdX3a8ECLRm2GfaW_3M';

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    // Add headers on first run
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Subject', 'Message']);
    }

    sheet.appendRow([
      new Date(),
      e.parameter.name,
      e.parameter.email,
      e.parameter.subject,
      e.parameter.message
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
