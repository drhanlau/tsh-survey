/**
 * TSH AI Readiness Survey — Google Sheets endpoint
 *
 * Receives submissions from tsh-ai-readiness-survey.html and appends them
 * as rows to this spreadsheet (tab with gid 1880581065):
 *   https://docs.google.com/spreadsheets/d/13G8FbhECV0YBrNvRYton1eihgBZIZ8IjuGfyd_LvlA0/edit
 *
 * ONE-TIME SETUP (~3 minutes):
 *   1. Open the spreadsheet above → Extensions → Apps Script.
 *   2. Replace the default Code.gs content with this entire file, then Save.
 *   3. Select the `setup` function in the toolbar and click Run — grant the
 *      requested permissions. This writes the header row into the survey tab.
 *   4. Deploy → New deployment → gear icon → "Web app":
 *        - Execute as:     Me
 *        - Who has access: Anyone   (respondents submit without signing in)
 *      Click Deploy, authorize if asked, and copy the Web app URL (…/exec).
 *   5. Paste that URL into SCRIPT_URL near the top of the <script> block in
 *      tsh-ai-readiness-survey.html.
 *
 * You can open the …/exec URL in a browser to check it is alive — it replies
 * with a small JSON status. Only doPost writes data; the sheet is never
 * exposed for reading through this endpoint.
 */

var SPREADSHEET_ID = '13G8FbhECV0YBrNvRYton1eihgBZIZ8IjuGfyd_LvlA0';
var SHEET_GID = 1880581065;

var HEADERS = [
  'timestamp',
  'name',
  'role',
  'q1_tools_used',
  'q2_understanding_rating',
  'q3_concerns',
  'q4_tasks_wanted',
  'q5_repetitive_task',
  'q6_workshop_goal'
];

/** Run once from the editor to create the header row. */
function setup() {
  var sheet = getTargetSheet_();
  ensureHeaders_(sheet);
}

function doGet() {
  return json_({
    result: 'ok',
    message: 'TSH Pre-Training Survey endpoint is live. Submissions arrive via POST.'
  });
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return json_({ result: 'error', message: 'Empty request body.' });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ result: 'error', message: 'Busy, please retry.' });
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getTargetSheet_();
    ensureHeaders_(sheet);
    sheet.appendRow(HEADERS.map(function (h) {
      var v = data[h];
      if (v === undefined || v === null) return '';
      return Array.isArray(v) ? v.join(', ') : String(v);
    }));
    return json_({ result: 'ok' });
  } catch (err) {
    return json_({ result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---- helpers -------------------------------------------------------------

function getTargetSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  throw new Error('No tab with gid ' + SHEET_GID + ' in spreadsheet ' + SPREADSHEET_ID);
}

function ensureHeaders_(sheet) {
  var firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var hasAny = firstRow.some(function (v) { return v !== ''; });
  if (!hasAny) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
