/**
 * UANDES intake survey receiver.
 * Bound to the responses Google Sheet (Extensions > Apps Script), deployed as a
 * Web app: Execute as "Me", Who has access "Anyone".
 * After editing: Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy,
 * so the URL in index.html / part-2.html stays the same.
 *
 * One row per student in the "Responses" tab. Each POST carries a response_id.
 * Part 1 adds the row; Part 2 (same browser, same response_id) fills in the rest of
 * that row. If no matching row is found (e.g. Part 2 on another phone), it adds a new row.
 * New fields get new columns automatically, so editing questions never breaks the sheet.
 */
const TAB = 'Responses';
const FIRST_COLS = ['response_id', 'full_name', 'part1_submitted_at', 'part2_submitted_at'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB) || ss.insertSheet(TAB, 0);

    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];
    if (headers.length === 0) headers = FIRST_COLS.slice();
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) headers.push(k);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Prefix a quote on anything that starts like a formula, so answers stay plain text.
    const clean = function (v) {
      v = v == null ? '' : String(v);
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    };

    // Find this student's existing row by response_id.
    let rowNum = -1;
    const idCol = headers.indexOf('response_id') + 1;
    if (data.response_id && idCol > 0 && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
      for (let r = 0; r < ids.length; r++) {
        if (String(ids[r][0]) === String(data.response_id)) { rowNum = r + 2; break; }
      }
    }

    if (rowNum === -1) {
      sheet.appendRow(headers.map(function (h) { return clean(data[h]); }));
    } else {
      // Fill in only the fields this POST sent; leave the other part's answers alone.
      const range = sheet.getRange(rowNum, 1, 1, headers.length);
      const current = range.getValues()[0];
      const updated = headers.map(function (h, c) {
        return Object.prototype.hasOwnProperty.call(data, h) ? clean(data[h]) : current[c];
      });
      range.setValues([updated]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/** Run once from the editor to check the wiring. Adds a TEST row, then fills it in as Part 2. */
function testWrite() {
  const id = 'TEST-' + Date.now();
  Logger.log(doPost({ postData: { contents: JSON.stringify({ response_id: id, part1_submitted_at: new Date().toISOString(), full_name: 'TEST' }) } }).getContent());
  Logger.log(doPost({ postData: { contents: JSON.stringify({ response_id: id, part2_submitted_at: new Date().toISOString(), hype_song: 'TEST' }) } }).getContent());
}
