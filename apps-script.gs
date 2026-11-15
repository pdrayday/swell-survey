/**
 * Swell survey → Google Sheet
 *
 * 1. Make a new Google Sheet (name it "Swell Survey Responses").
 * 2. Extensions → Apps Script. Delete what's there, paste this whole file, Save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me      Who has access: Anyone
 *    Click Deploy, approve access, copy the Web app URL (ends in /exec).
 * 4. Paste that URL into index.html:  window.SWELL_ENDPOINT = "https://script.google.com/.../exec";
 *
 * Every submission becomes one row. New questions add new columns automatically.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var lastCol = sheet.getLastColumn();
    var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(row).forEach(function (k) {
      if (headers.indexOf(k) === -1) { headers.push(k); sheet.getRange(1, headers.length).setValue(k); }
    });
    sheet.appendRow(headers.map(function (h) { return row[h] === undefined ? "" : row[h]; }));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Swell survey endpoint is live.");
}
