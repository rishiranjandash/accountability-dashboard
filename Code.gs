/**
 * Warehouse Leaders Accountability Dashboard — Apps Script backend
 *
 * Reads/writes the "Accountability Playbook — Poll Options & Follow-ups"
 * Google Sheet and exposes a small JSON API that the static dashboard
 * (hosted on GitHub Pages) calls over plain GET/POST requests.
 *
 * Deploy as: Extensions > Apps Script (bound to the Sheet, or standalone
 * with SPREADSHEET_ID set below) > Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone with the link
 * Copy the resulting /exec URL into docs/index.html (APPS_SCRIPT_URL).
 *
 * Everything here only ever touches this one spreadsheet, plus (as of the
 * webapp-voting feature) a single Drive folder used to store vote-proof
 * screenshots. No other Google data, no other scopes needed.
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

// OAuth 2.0 Web Client ID from Google Cloud Console ("Sign in with Google"),
// used to check the `aud` claim on every ID token this backend verifies, so
// a token minted for some other app can't be replayed here. Leave blank
// while the client hasn't been created yet — this only *skips* the aud
// check, it does not disable signature/expiry verification (still done via
// Google's own tokeninfo endpoint). Fill in once created and redeploy. The
// SAME value must be set as GOOGLE_CLIENT_ID in docs/index.html.
const OAUTH_CLIENT_ID = '';

// NOTE: the sheet tab holding the poll options/tasks data is still named
// "Untitled" internally (never renamed after creation) even though the
// spreadsheet FILE is called "Accountability Playbook — Poll Options &
// Follow-ups" and the columns match that schema. Using the real tab name
// here rather than renaming the live tab, to avoid touching anything the
// existing WhatsApp automation might depend on.
const TAB_TASKS = 'Untitled';
const TAB_LEAVE = 'Leave Tracker';
const TAB_POINTS = 'Points';
const TAB_SLIPPAGE = 'Slippage Report';
const TAB_LOG = 'Daily Log';
const TAB_RULES = 'Point Rules';
const TAB_TEAM = 'Team';

// Drive folder (created on first use) that holds vote-proof screenshots.
const PROOF_FOLDER_NAME = 'Accountability Dashboard — Vote Proofs';

// A task is auto-marked "Missed" this many minutes after its cutoff time if
// nobody has voted it Done and its owner isn't on leave. Matches the
// existing WhatsApp-automation grace window (1.5h).
const AUTO_MISS_GRACE_MINUTES = 90;

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

// Sheets stores a "time of day" cell (e.g. Cutoff Time) as a Date on the
// 1899-12-30 epoch. Left as a raw Date object it serializes via JSON.stringify
// into an ugly/misleading ISO string like "1899-12-30T16:30:00.000Z". Format
// it back into the time-of-day the sheet actually displays instead.
function formatTimeOfDay_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'h:mm a');
  }
  return val || '';
}

// Same idea for a real date+time value (e.g. Daily Log "Marked At").
function formatDateTime_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'd MMM, h:mm a');
  }
  return val || '';
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
        cutoff: formatTimeOfDay_(t['Cutoff Time']),
        group: t['Group'],
        task: t['Task / Option'],
        assignedTo: owner,
        onLeave: !!onLeave,
        coveredBy: onLeave ? (leaveRow['Covered By'] || '') : '',
        leaveNote: onLeave ? (leaveRow['Notes'] || '') : '',
        status: status,
        markedAt: logRow ? formatDateTime_(logRow['Marked At']) : '',
        followUp: t['Follow-up Question If Missed']
      };
    });
}

/**
 * Reads the editable point formula from the "Point Rules" tab (Setting | Value | Notes).
 * Falls back to the original default (+1 done / -1 missed, net score) if the tab is
 * missing or a value is blank/non-numeric, so the sheet can never break the leaderboard.
 * Run setupPointRulesTab_() once (Apps Script editor > select it > Run) to create the
 * tab with these defaults pre-filled, or to reset it if it gets mangled.
 */
function getPointRules_() {
  const rules = { done: 1, missed: -1 };
  const sheet = getSpreadsheet_().getSheetByName(TAB_RULES);
  if (!sheet) return rules;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const label = String(values[i][0] || '').trim().toLowerCase();
    const val = parseFloat(values[i][1]);
    if (isNaN(val)) continue;
    if (label === 'points for done') rules.done = val;
    else if (label === 'points for missed') rules.missed = val;
  }
  return rules;
}

/** One-time (or reset-anytime) setup: creates/repopulates the "Point Rules" tab. */
function setupPointRulesTab_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(TAB_RULES);
  if (!sheet) sheet = ss.insertSheet(TAB_RULES);
  sheet.clear();
  sheet.getRange(1, 1, 6, 3).setValues([
    ['Setting', 'Value', 'Notes'],
    ['Points for Done', 1, "Points added to a person's Leaderboard total for each task logged as Done."],
    ['Points for Missed', -1, 'Points added for each task logged as Missed. Use a negative number so misses subtract from the total (set to 0 to stop penalizing misses).'],
    ['', '', ''],
    ['How it works', '', 'Total Points = (Tasks Done x Points for Done) + (Tasks Missed x Points for Missed). Edit the Value column above and the Leaderboard picks it up automatically next time it loads — no redeploy needed.'],
    ['Applies to', '', 'Leaderboard tab only. Slippage Report status thresholds (On Track / Needs Attention / At Risk) are separate and not affected by these settings.']
  ]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 560);
  sheet.setFrozenRows(1);
}

/**
 * One-time (or reset-anytime) setup: creates/repopulates the "Team" tab —
 * the self-service roster that maps each person's Gmail address to the
 * display name used everywhere else in the sheet (Assigned To, Leave
 * Tracker, etc). Add, remove, or edit rows here any time to change who can
 * sign in and vote on the webapp — no code change or redeploy needed.
 * Seeded with the names currently in use; emails left blank on purpose —
 * fill each person's Gmail address in before they can vote.
 */
function setupTeamTab_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(TAB_TEAM);
  if (!sheet) sheet = ss.insertSheet(TAB_TEAM);
  if (sheet.getLastRow() > 0) return; // don't clobber an existing roster on re-run
  sheet.getRange(1, 1, 8, 4).setValues([
    ['Name', 'Email', 'Active (Y/N)', 'Notes'],
    ['Chandu', '', 'Y', ''],
    ['Imran', '', 'Y', ''],
    ['Vinod', '', 'Y', ''],
    ['Manoj', '', 'Y', ''],
    ['Ajay', '', 'Y', ''],
    ['Suraj', '', 'Y', ''],
    ['', '', '', 'Add a row per person. Name must exactly match the name used in the "Assigned To" column of the task list (and Leave Tracker). Email is the Gmail address they sign in with on the webapp — required before they can vote. Set Active to N (instead of deleting the row) to temporarily disable someone without losing history.']
  ]);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 560);
  sheet.setFrozenRows(1);
}

/**
 * Points leaderboard, computed live from Daily Log.
 * Formula: Total Points = (Tasks Done x Points for Done) + (Tasks Missed x Points for Missed),
 * read from the "Point Rules" tab via getPointRules_() — defaults to the original
 * net-score formula (+1 done / -1 missed) if that tab is missing or blank.
 * A task logged as "On Leave"/covered does not count against the assigned owner
 * (it is simply excluded from both done and missed for that owner).
 */
function computePoints_() {
  const tasks = sheetToObjects_(TAB_TASKS);
  const log = sheetToObjects_(TAB_LOG);
  const rules = getPointRules_();

  const owners = [...new Set(
    tasks.map(t => String(t['Assigned To'] || '').trim()).filter(Boolean)
  )];
  const byOwner = {};
  owners.forEach(o => { byOwner[o] = { done: 0, missed: 0, lastActivity: null }; });

  log.forEach(r => {
    const taskName = String(r['Task / Option'] || '').trim();
    const def = tasks.find(t => String(t['Task / Option']).trim() === taskName);
    const owner = def ? String(def['Assigned To'] || '').trim() : String(r['Assigned To'] || '').trim();
    if (!owner) return;
    if (!byOwner[owner]) byOwner[owner] = { done: 0, missed: 0, lastActivity: null };

    const status = String(r['Poll Status'] || '').trim();
    if (status === 'Done') {
      byOwner[owner].done++;
    } else if (status === 'Missed') {
      byOwner[owner].missed++;
    }
    // "On Leave" / "Covered" (or any other status) neither helps nor hurts the owner.

    const markedAt = r['Marked At'];
    if (markedAt instanceof Date && (!byOwner[owner].lastActivity || markedAt > byOwner[owner].lastActivity)) {
      byOwner[owner].lastActivity = markedAt;
    }
  });

  return owners.map(o => {
    const s = byOwner[o];
    const total = s.done + s.missed;
    return {
      'Name': o,
      'Group': 'Warehouse Leaders',
      'Total Points': s.done * rules.done + s.missed * rules.missed,
      'Tasks Done': s.done,
      'Tasks Missed': s.missed,
      'Completion Rate': total ? Math.round((s.done / total) * 100) + '%' : 'N/A',
      'Last Activity Date': s.lastActivity ? formatDateTime_(s.lastActivity) : 'N/A'
    };
  }).sort((a, b) => b['Total Points'] - a['Total Points']);
}

/**
 * Slippage report, computed live from Daily Log.
 * "Missed (Last 3 Days)" counts Missed rows dated within the trailing 3 days
 * (today inclusive); "Missed (All-Time)" counts every Missed row ever logged.
 * Status: On Track (0 misses in last 3 days), Needs Attention (1-2), At Risk (3+).
 */
function computeSlippage_() {
  const tasks = sheetToObjects_(TAB_TASKS);
  const log = sheetToObjects_(TAB_LOG);

  const owners = [...new Set(
    tasks.map(t => String(t['Assigned To'] || '').trim()).filter(Boolean)
  )];
  const now = new Date();
  const threeDaysAgo = new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000);

  const byOwner = {};
  owners.forEach(o => { byOwner[o] = { missed3: 0, missedAll: 0, done: 0, missed: 0 }; });

  log.forEach(r => {
    const taskName = String(r['Task / Option'] || '').trim();
    const def = tasks.find(t => String(t['Task / Option']).trim() === taskName);
    const owner = def ? String(def['Assigned To'] || '').trim() : String(r['Assigned To'] || '').trim();
    if (!owner) return;
    if (!byOwner[owner]) byOwner[owner] = { missed3: 0, missedAll: 0, done: 0, missed: 0 };

    const status = String(r['Poll Status'] || '').trim();
    const dateVal = r['Date'];
    const rowDate = dateVal instanceof Date ? dateVal : new Date(String(dateVal));

    if (status === 'Missed') {
      byOwner[owner].missedAll++;
      byOwner[owner].missed++;
      if (!isNaN(rowDate.getTime()) && rowDate >= threeDaysAgo) {
        byOwner[owner].missed3++;
      }
    } else if (status === 'Done') {
      byOwner[owner].done++;
    }
  });

  return owners.map(o => {
    const s = byOwner[o];
    const total = s.done + s.missed;
    const rate = total ? Math.round((s.done / total) * 100) : null;
    let status;
    if (s.missed3 === 0) status = 'On Track';
    else if (s.missed3 <= 2) status = 'Needs Attention';
    else status = 'At Risk';
    return {
      'Name': o,
      'Group': 'Warehouse Leaders',
      'Missed (Last 3 Days)': s.missed3,
      'Missed (All-Time)': s.missedAll,
      'Completion Rate': rate !== null ? rate + '%' : 'N/A',
      'Status': status,
      'Last Updated': formatDateTime_(now)
    };
  }).sort((a, b) => b['Missed (Last 3 Days)'] - a['Missed (Last 3 Days)']);
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
        payload = { rows: computePoints_() };
        break;
      case 'leave':
        payload = { rows: sheetToObjects_(TAB_LEAVE) };
        break;
      case 'slippage':
        payload = { rows: computeSlippage_() };
        break;
      case 'log':
        payload = { rows: sheetToObjects_(TAB_LOG) };
        break;
      case 'markdone':
        return handleMarkDone_(e);
      case 'logresult':
        return handleLogResult_(e);
      case 'setupteam':
        payload = handleAdminSetup_(e, setupTeamTab_);
        break;
      case 'setuptrigger':
        payload = handleAdminSetup_(e, installAutoMissTrigger_);
        break;
      case 'checkautomiss':
        // Lets you trigger a manual sweep on demand (e.g. right after
        // editing a cutoff time) without waiting for the next 15-min tick.
        payload = handleAdminSetup_(e, checkAutoMiss_);
        break;
      default:
        payload = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    payload = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Shared guard for the one-off admin/setup GET actions above. */
function handleAdminSetup_(e, fn) {
  const token = getWriteToken_();
  if (token && e.parameter.token !== token) {
    return { error: 'invalid token' };
  }
  fn();
  return { ok: true };
}

function doPost(e) {
  const action = (e.parameter.action || '').toLowerCase();
  let payload;
  try {
    const body = (e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    switch (action) {
      case 'whoami':
        payload = handleWhoAmI_(body);
        break;
      case 'vote':
        payload = handleVote_(body);
        break;
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
 * Verifies a Google Sign-In ID token via Google's own tokeninfo endpoint
 * (no crypto library needed — Google checks the signature/expiry for us
 * and just tells us the result). Throws on anything invalid.
 */
function verifyGoogleIdToken_(idToken) {
  if (!idToken) throw new Error('not signed in');
  const resp = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  const body = JSON.parse(resp.getContentText() || '{}');
  if (resp.getResponseCode() !== 200 || body.error) {
    throw new Error('sign-in expired or invalid — please sign in again' + (body.error_description ? ' (' + body.error_description + ')' : ''));
  }
  if (OAUTH_CLIENT_ID && body.aud !== OAUTH_CLIENT_ID) {
    throw new Error('sign-in token was not issued for this app');
  }
  if (body.email_verified !== 'true' && body.email_verified !== true) {
    throw new Error('Google account email is not verified');
  }
  return { email: String(body.email || '').trim().toLowerCase(), name: body.name || '', picture: body.picture || '' };
}

/** Resolves a verified email address to the Team-tab display Name (or null). */
function resolveTeamName_(email) {
  const rows = sheetToObjects_(TAB_TEAM);
  const match = rows.find(r =>
    String(r['Email'] || '').trim().toLowerCase() === email &&
    String(r['Active (Y/N)'] || '').trim().toUpperCase() !== 'N'
  );
  return match ? String(match['Name'] || '').trim() : null;
}

/** "Suraj/ Ajay" style multi-owner cells -> ['Suraj','Ajay']. */
function taskOwners_(task) {
  return String(task['Assigned To'] || '').split('/').map(s => s.trim()).filter(Boolean);
}

/**
 * POST action=whoami {idToken}. Lets the frontend find out (a) that the
 * signed-in Google account is recognized at all, and (b) which sheet Name
 * it maps to, so it can highlight that person's tasks. Read-only, but still
 * goes through full token verification since it echoes back a Name.
 */
function handleWhoAmI_(body) {
  const identity = verifyGoogleIdToken_(body.idToken);
  const name = resolveTeamName_(identity.email);
  if (!name) {
    return { ok: true, recognized: false, email: identity.email,
      message: "Your Google account (" + identity.email + ") isn't on the Team tab yet — ask Rishi to add it." };
  }
  return { ok: true, recognized: true, name: name, email: identity.email };
}

/** Get-or-create the Drive folder that holds vote-proof screenshots. */
function getProofFolder_() {
  const it = DriveApp.getFoldersByName(PROOF_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(PROOF_FOLDER_NAME);
}

/**
 * POST action=vote {idToken, task, proofBase64, proofMimeType, proofFilename}.
 * This is the webapp's own voting flow (replaces posting a WhatsApp poll):
 * verifies the signed-in Google account, resolves it to a Team-tab Name,
 * confirms that Name actually owns the task, requires a screenshot as
 * proof (same "proof before voting" rule the WhatsApp process already
 * uses), saves it to Drive, and appends a Done row to Daily Log.
 */
function handleVote_(body) {
  const identity = verifyGoogleIdToken_(body.idToken);
  const name = resolveTeamName_(identity.email);
  if (!name) {
    return { error: "Your Google account (" + identity.email + ") isn't on the Team tab yet — ask Rishi to add it before you can vote." };
  }

  const taskName = String(body.task || '').trim();
  if (!taskName) return { error: 'missing task' };

  const tasks = sheetToObjects_(TAB_TASKS);
  const def = tasks.find(t =>
    String(t['Task / Option']).trim() === taskName &&
    String(t['Active (Y/N)']).trim().toUpperCase() === 'Y'
  );
  if (!def) return { error: 'unknown or inactive task' };

  const owners = taskOwners_(def).map(o => o.toLowerCase());
  if (owners.indexOf(name.toLowerCase()) === -1) {
    return { error: 'That task is assigned to ' + def['Assigned To'] + ', not you (' + name + ').' };
  }

  // Idempotent: if today's outcome is already logged, don't double-append.
  const today = todayStr_();
  const existing = sheetToObjects_(TAB_LOG).find(r => {
    const d = r['Date'];
    const dStr = d instanceof Date ? Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd') : String(d);
    return dStr === today && String(r['Task / Option']).trim() === taskName;
  });
  if (existing) {
    return { ok: true, alreadyLogged: true, status: existing['Poll Status'] || 'Done' };
  }

  if (!body.proofBase64) {
    return { error: 'Please attach a screenshot as proof before marking this done.' };
  }

  let proofUrl = '';
  try {
    const mimeType = body.proofMimeType || 'image/png';
    const filename = (today + '_' + taskName + '_' + name).replace(/[^\w\-. ]+/g, '_') +
      (body.proofFilename ? '_' + body.proofFilename : '');
    const bytes = Utilities.base64Decode(body.proofBase64);
    const blob = Utilities.newBlob(bytes, mimeType, filename);
    const file = getProofFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    proofUrl = file.getUrl();
  } catch (err) {
    return { error: 'Could not save the proof screenshot: ' + err };
  }

  const sheet = getSpreadsheet_().getSheetByName(TAB_LOG);
  const now = new Date();
  sheet.appendRow([
    today,
    def['Group'],
    taskName,
    name,
    def['Cutoff Time'],
    'Done',
    now,
    '', '', '',
    'Proof: ' + proofUrl + ' (marked via webapp by ' + name + ')',
    ''
  ]);
  return { ok: true, name: name, task: taskName, status: 'Done', proofUrl: proofUrl };
}

/**
 * Scans today's active tasks and auto-logs "Missed" for anything past its
 * cutoff + AUTO_MISS_GRACE_MINUTES with no vote yet, unless the owner is on
 * leave (leave/coverage is left for a human to reconcile, same as today).
 * Installed as a time-driven trigger via installAutoMissTrigger_() — run
 * that once (via ?action=setuptrigger&token=... or the Apps Script editor)
 * after deploying.
 */
function checkAutoMiss_() {
  const tasks = sheetToObjects_(TAB_TASKS).filter(t => String(t['Active (Y/N)']).trim().toUpperCase() === 'Y');
  const leave = sheetToObjects_(TAB_LEAVE);
  const leaveByName = {};
  leave.forEach(r => { leaveByName[String(r['Name']).trim()] = r; });

  const today = todayStr_();
  const loggedToday = new Set();
  sheetToObjects_(TAB_LOG).forEach(r => {
    const d = r['Date'];
    const dStr = d instanceof Date ? Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd') : String(d);
    if (dStr === today) loggedToday.add(String(r['Task / Option']).trim());
  });

  const now = new Date();
  const sheet = getSpreadsheet_().getSheetByName(TAB_LOG);

  tasks.forEach(t => {
    const taskName = String(t['Task / Option']).trim();
    if (loggedToday.has(taskName)) return;

    const cutoff = t['Cutoff Time'];
    if (!(cutoff instanceof Date)) return; // can't compute a grace deadline without a real time
    const hm = Utilities.formatDate(cutoff, 'Asia/Kolkata', 'HH:mm').split(':');
    const todayDateStr = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd');
    const cutoffToday = new Date(todayDateStr + 'T' + hm[0] + ':' + hm[1] + ':00+05:30');
    const deadline = new Date(cutoffToday.getTime() + AUTO_MISS_GRACE_MINUTES * 60 * 1000);
    if (now.getTime() < deadline.getTime()) return; // still within grace

    const owner = String(t['Assigned To'] || '').trim();
    const firstOwner = owner.split('/')[0].trim();
    const leaveRow = leaveByName[firstOwner];
    const onLeave = leaveRow && String(leaveRow['On Leave Today (Y/N)']).trim().toUpperCase() === 'Y';
    if (onLeave) return; // leave/coverage is reconciled by a human, not auto-missed

    sheet.appendRow([
      today, t['Group'], taskName, owner, t['Cutoff Time'], 'Missed', now,
      '', '', '', 'Auto-marked Missed: cutoff + ' + AUTO_MISS_GRACE_MINUTES + 'min grace passed with no vote.', ''
    ]);
    loggedToday.add(taskName); // guard against double-appending within this same run
  });
}

/** Installs (idempotently) the recurring trigger that drives checkAutoMiss_(). */
function installAutoMissTrigger_() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (tr.getHandlerFunction() === 'checkAutoMiss_') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('checkAutoMiss_').timeBased().everyMinutes(15).create();
}

/**
 * Mark a task done for today, appended to Daily Log.
 * GET params: action=markDone, task=<Task / Option text>, by=<name>, token=<WRITE_TOKEN>
 * (uses GET, not POST, so the static frontend can call it without hitting CORS preflight)
 * Kept for backward compatibility / manual/admin use; the webapp's own voting
 * flow now goes through POST action=vote (handleVote_) instead.
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

/**
 * Log a task's outcome for today (Done or Missed), appended to Daily Log.
 * Originally built for the WhatsApp checkpoint automation to call once it
 * finalized each task's status; kept for that use (and any other manual/
 * admin logging) even though the webapp's own flow now covers Done (POST
 * action=vote) and Missed (automatic, via checkAutoMiss_) on its own.
 * GET params: action=logresult, task=<Task / Option text>, status=Done|Missed, token=<WRITE_TOKEN>
 */
function handleLogResult_(e) {
  const token = getWriteToken_();
  if (token && e.parameter.token !== token) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'invalid token' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const taskName = e.parameter.task;
  const status = e.parameter.status;
  if (!taskName || !status) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'missing task or status param' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (status !== 'Done' && status !== 'Missed') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'status must be Done or Missed' }))
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
    def ? def['Assigned To'] : '',
    def ? def['Cutoff Time'] : '',
    status,
    now,
    '', '', '', '', ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
