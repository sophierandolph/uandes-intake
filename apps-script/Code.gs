/**
 * UANDES intake survey receiver.
 * Bind this script to the responses Google Sheet (Extensions > Apps Script), then
 * Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * Paste the web app URL into `endpoint` in index.html (and part-2.html later).
 *
 * Each POST is one JSON object. Rows go to a tab named "Part 1" or "Part 2".
 * New fields get new columns automatically, so editing questions never breaks the sheet.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const part = String(data.part || '1').replace(/[^0-9]/g, '') || '1';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = 'Part ' + part;
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);

    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) headers.push(k);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Prefix a quote on anything that starts like a formula, so answers stay plain text.
    const row = headers.map(function (h) {
      const v = data[h] == null ? '' : String(data[h]);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    });
    sheet.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** Run once from the editor to check the sheet wiring. Adds a row tagged TEST. */
function testWrite() {
  const fake = { postData: { contents: JSON.stringify({ part: 1, submitted_at: new Date().toISOString(), full_name: 'TEST' }) } };
  Logger.log(doPost(fake).getContent());
}
