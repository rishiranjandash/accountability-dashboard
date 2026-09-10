/**
 * Warehouse Leaders Accountability Dashboard — Apps Script backend
 *
 * Reads/writes the "Accountability Playbook — Poll Options & Follow-ups"
 * Google Sheet and exposes a small JSON API that the static dashboard
 * (hosted on GitHub Pages) calls over plain GET requests.
 *
 * Deploy as: Extensions > Apps Script (bound to the Sheet, or standalone
 * with SPREADSHEET_ID set below) > Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone with the link
 * Copy the resulting /exec URL into docs/index.html (APPS_SCRIPT_URL).
 *
 * Everything here only ever touches this one spreadsheet — no other
 * Google data, no other scopes needed.
 */

// If this script is bound to the spreadsheet (created via Extensions > Apps
// Script from inside the Sheet), leave this blank and SpreadsheetApp.getActive()
// is used. If deployed standalone, paste the Sheet ID here instead.
const SPREADSHEET_ID = '';

// Simple shared-secret so random internet visitors can't write rows.
// Set your own value in Project Settings > Script properties as WRITE_TOKEN,
// then put the same value in docs/index.html (WRITE_TOKEN). Reads (GET
// without action=markDone) do not require it.
function getWriteToken_() {
  return PropertiesService.getScriptProperties().getProperty('WRITE_TOKEN') || '';
}

const TAB_TASKS = 'Poll Options & Follow-ups';
const TAB_LEAVE = 'Leave Tracker';
const TAB_POINTS = 'Points';
const TAB_SLIPPAGE = 'Slippage Report';
const TAB_LOG = 'Daily Log';

function getSpreadsheet_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheetToObjects_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c === null)) continue; // skip blank rows
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

function todayStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

/** Merge task definitions + leave tracker + today's Daily Log rows into one board. */
function buildTodayBoard_() {
  const tasks = sheetToObjects_(TAB_TASKS);
  const leave = sheetToObjects_(TAB_LEAVE);
  const log = sheetToObjects_(TAB_LOG);
  const today = todayStr_();

  const leaveByName = {};
  leave.forEach(r => { leaveByName[String(r['Name']).trim()] = r; });

  // index today's log rows by "Task / Option" (last write wins)
  const logByTask = {};
  log.forEach(r => {
    const d = r['Date'];
    const dStr = d instanceof Date ? Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd') : String(d);
    if (dStr === today) {
      logByTask[String(r['Task / Option']).trim()] = r;
    }
  });

  return tasks
    .filter(t => String(t['Active (Y/N)']).trim().toUpperCase() === 'Y')
    .map(t => {
      const owner = String(t['Assigned To'] || '').trim();
      // "Suraj/ Ajay" style multi-owner cells — match leave on the first name
      const firstOwner = owner.split('/')[0].trim();
      const leaveRow = leaveByName[firstOwner];
      const onLeave = leaveRow && String(leaveRow['On Leave Today (Y/N)']).trim().toUpperCase() === 'Y';
      const logRow = logByTask[String(t['Task / Option']).trim()];
      const status = logRow ? String(logRow['Poll Status'] || 'Done') : 'Pending';
      return {
        cutoff: t['Cutoff Time'],
        group: t['Group'],
        task: t['Task / Option'],
        assignedTo: owner,
        onLeave: !!onLeave,
        coveredBy: onLeave ? (leaveRow['Covered By'] || '') : '',
        leaveNote: onLeave ? (leaveRow['Notes'] || '') : '',
        status: status,
        markedAt: logRow ? logRow['Marked At'] : '',
        followUp: t['Follow-up Question If Missed']
      };
    });
}

function doGet(e) {
  const action = (e.parameter.action || 'tasks').toLowerCase();
  let payload;
  try {
    switch (action) {
      case 'tasks':
        payload = { date: todayStr_(), tasks: buildTodayBoard_() };
        break;
      case 'leaderboard':
        payload = { rows: sheetToObjects_(TAB_POINTS) };
        break;
      case 'leave':
        payload = { rows: sheetToObjects_(TAB_LEAVE) };
        break;
      case 'slippage':
        payload = { rows: sheetToObjects_(TAB_SLIPPAGE) };
        break;
      case 'log':
        payload = { rows: sheetToObjects_(TAB_LOG) };
        break;
      case 'markdone':
        return handleMarkDone_(e);
      default:
        payload = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    payload = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Mark a task done for today, appended to Daily Log.
 * GET params: action=markDone, task=<Task / Option text>, by=<name>, token=<WRITE_TOKEN>
 * (uses GET, not POST, so the static frontend can call it without hitting CORS preflight)
 */
function handleMarkDone_(e) {
  const token = getWriteToken_();
  if (token && e.parameter.token !== token) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'invalid token' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const taskName = e.parameter.task;
  const markedBy = e.parameter.by || '';
  if (!taskName) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'missing task param' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const tasks = sheetToObjects_(TAB_TASKS);
  const def = tasks.find(t => String(t['Task / Option']).trim() === String(taskName).trim());
  const sheet = getSpreadsheet_().getSheetByName(TAB_LOG);
  const now = new Date();
  sheet.appendRow([
    todayStr_(),
    def ? def['Group'] : '',
    taskName,
    (def ? def['Assigned To'] : '') + (markedBy ? ' (marked by ' + markedBy + ')' : ''),
    def ? def['Cutoff Time'] : '',
    'Done',
    now,
    '', '', '', '', ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
