function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Company Memos & Announcements')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMemosFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Memos"); // change if needed
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header row

  const memos = data
    .filter(row => row[5]?.toString().toLowerCase() === "yes")
    .map(row => ({
      dateSubmitted: formatDate(row[0]), // Date Submitted
      weeklyTidbit: sanitize(row[1]),    // Weekly Tidbit Information
      pointOfContact: row[2],            // Point of Contact
      bu: row[3],                         // BU
      postBy: formatDate(row[4]),        // Post no later than date
      published: row[5]                  // Published
    }));

  return memos;
}

function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "MMM dd, yyyy");
}

function sanitize(input) {
  return typeof input === 'string'
    ? input.replace(/</g, '&lt;').replace(/>/g, '&gt;')
    : '';
}
