/**
 * Digital Business Bootcamp – Business Understanding Form
 * Google Sheets receiver (Google Apps Script Web App)
 *
 * SETUP
 * 1. Create a new Google Sheet (e.g. "Bootcamp – Business Understanding Responses").
 * 2. In the sheet: Extensions → Apps Script. Delete the sample code, paste this file, click Save.
 * 3. Click Deploy → New deployment → gear icon → "Web app".
 *      Description:     Form receiver
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Click Deploy and allow the permissions.
 * 4. Copy the "Web app URL" (it ends with /exec) and paste it into
 *    CONFIG.SHEET_ENDPOINT in index.html.
 *
 * If you change this script later: Deploy → Manage deployments → pencil icon →
 * Version: "New version" → Deploy. The URL stays the same.
 *
 * HOW PARTIAL RESPONSES WORK
 * The form sends the full set of answers after every question. Each visitor has a
 * Session ID, so their row is updated in place (never duplicated). The Status column
 * reads "Partial" until the last question is submitted, then "Completed".
 */

const SHEET_NAME = 'Responses';
const META_HEADERS = ['Session ID', 'Status', 'Progress', 'Last Answered', 'Started At', 'Last Updated', 'Device'];
const SEQ_HEADER = 'Sync #';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }

  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.sessionId || !Array.isArray(data.fields)) {
      return json_({ ok: false, error: 'invalid payload' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    const headers = ensureHeaders_(sheet, data.fields.map(f => String(f.label)));
    const col = name => headers.indexOf(name);

    // Find this visitor's existing row (search from the bottom — recent rows first)
    const lastRow = sheet.getLastRow();
    let rowIndex = -1;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === data.sessionId) { rowIndex = i + 2; break; }
      }
    }

    let row;
    if (rowIndex > 0) {
      row = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      // Ignore requests that arrive out of order (older than what we already stored)
      if (Number(data.seq) <= Number(row[col(SEQ_HEADER)] || 0)) {
        return json_({ ok: true, skipped: 'stale' });
      }
    } else {
      rowIndex = lastRow + 1;
      row = new Array(headers.length).fill('');
    }

    const tz = ss.getSpreadsheetTimeZone();
    const fmt = d => Utilities.formatDate(d, tz, 'dd MMM yyyy, hh:mm a');

    row[col('Session ID')] = data.sessionId;
    row[col('Status')] = data.status || 'Partial';
    row[col('Progress')] = data.progress || '';
    row[col('Last Answered')] = data.lastAnswered || '';
    row[col('Started At')] = data.startedAt ? fmt(new Date(data.startedAt)) : '';
    row[col('Last Updated')] = fmt(new Date());
    row[col('Device')] = data.device || '';
    row[col(SEQ_HEADER)] = String(Number(data.seq) || 0);
    data.fields.forEach(f => {
      row[col(String(f.label))] = f.value == null ? '' : String(f.value);
    });

    // Store everything as plain text so phone numbers keep their + and leading zeros
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    range.setNumberFormat('@').setValues([row]).setVerticalAlignment('top');
    sheet.getRange(rowIndex, col('Status') + 1)
      .setBackground(data.status === 'Completed' ? '#D8F3E4' : '#FFF1D0')
      .setFontWeight('bold');

    return json_({ ok: true, row: rowIndex });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Visit the Web app URL in a browser to confirm it is live.
function doGet() {
  return ContentService.createTextOutput('Bootcamp form endpoint is live ✅');
}

function ensureHeaders_(sheet, fieldLabels) {
  const wanted = META_HEADERS.concat(fieldLabels, [SEQ_HEADER]);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, wanted.length)
      .setValues([wanted])
      .setFontWeight('bold')
      .setBackground('#0A1433')
      .setFontColor('#FFFFFF')
      .setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    return wanted;
  }

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const missing = wanted.filter(h => headers.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length)
      .setValues([missing])
      .setFontWeight('bold')
      .setBackground('#0A1433')
      .setFontColor('#FFFFFF');
    headers = headers.concat(missing);
  }
  return headers;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
