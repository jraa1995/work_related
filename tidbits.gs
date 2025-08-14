function getMemosFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const memos = data.map(row => {
    return {
      dateSubmitted: row[0],
      weeklyTidbit: row[1],
      pointOfContact: row[2],
      bu: row[3],
      postBy: row[4],
      published: row[5]
    };
  }).filter(memo => memo.published.toString().toLowerCase() === "yes");
  return memos;
}
