/**
 * ANVI Fashion Jewellery - Quick Message Google Apps Script
 *
 * 1. Open the target Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1GMZ3wRnzdSiE1r9qX--i6_VcRx9CXOjgah13YBu5MaA/edit
 * 2. Extensions -> Apps Script
 * 3. Replace the default code with this file and Save.
 * 4. Deploy -> New deployment -> Web app
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Copy the Web app URL ending in /exec.
 * 6. Paste it into CONTACT_FORM_ENDPOINT in the website HTML.
 */
const SPREADSHEET_ID = '1GMZ3wRnzdSiE1r9qX--i6_VcRx9CXOjgah13YBu5MaA';
const SHEET_NAME = ''; // Leave blank to use the first tab, or enter the exact tab name.
const OUTPUT_HEADERS = ['Timestamp', 'Name', 'Mobile Number', 'Message', 'Page URL'];

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(body);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
    if (!sheet) throw new Error('Target sheet not found.');

    ensureHeaders_(sheet);
    sheet.appendRow([
      new Date(),
      String(data.name || '').trim(),
      String(data.mobile || '').trim(),
      String(data.message || '').trim(),
      String(data.page || '').trim()
    ]);

    return json_({ok:true});
  } catch (err) {
    console.error(err);
    return json_({ok:false, error:String(err)});
  }
}

function doGet() {
  return json_({ok:true, service:'ANVI contact form'});
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), OUTPUT_HEADERS.length);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const hasAnyHeader = firstRow.some(v => String(v || '').trim() !== '');
  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    return;
  }
  // If the first row already has headers, append missing form columns to the right.
  const existing = firstRow.map(v => String(v || '').trim().toLowerCase());
  let col = existing.length;
  OUTPUT_HEADERS.forEach(header => {
    if (!existing.includes(header.toLowerCase())) {
      col += 1;
      sheet.getRange(1, col).setValue(header);
    }
  });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
