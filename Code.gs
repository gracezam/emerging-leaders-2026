// ─────────────────────────────────────────────
// CRD2026 — Google Apps Script Backend
// Paste this entire file into Apps Script, deploy as web app
//
// SETUP STEPS:
// 1. Open your Google Sheet
// 2. Extensions → Apps Script
// 3. Delete all existing code, paste this entire file
// 4. Click Save (floppy disk icon)
// 5. Click Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Click Deploy → copy the Web app URL
// 7. Paste that URL into config.js as script_url
// ─────────────────────────────────────────────

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Tab names
const TABS = {
  users:         'users',
  conversations: 'conversations',
  coffee:        'coffee_consults',
  views:         'profile_views',
  survey:        'survey',
  dashboard:     'dashboard'
};

// ── ROUTE INCOMING REQUESTS ───────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if      (action === 'identify')      handleIdentify(data);
    else if (action === 'log_convo')     handleConversation(data);
    else if (action === 'undo_convo')    handleUndoConversation(data);
    else if (action === 'coffee')        handleCoffee(data);
    else if (action === 'view_profile')  handleProfileView(data);
    else if (action === 'survey')        handleSurvey(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET: health check OR data export for admin dashboard
function doGet(e) {
  const action = e.parameter ? e.parameter.action : null;

  if (action === 'export') {
    // Return all tracking data for the admin dashboard
    const data = {
      users:  getSheetData(TABS.users),
      convos: getSheetData(TABS.conversations),
      coffee: getSheetData(TABS.coffee),
      views:  getSheetData(TABS.views),
      survey: getSheetData(TABS.survey),
    };
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'CRD2026 backend running' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HANDLERS ─────────────────────────────────

function handleIdentify(data) {
  // Track unique app users by session ID
  const sheet = getOrCreateTab(TABS.users, [
    'session_id', 'name', 'role', 'program', 'timestamp', 'date'
  ]);

  // Check if session already exists — update if so, insert if new
  const rows = sheet.getDataRange().getValues();
  const sessionCol = 0;
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][sessionCol] === data.session_id) {
      // Update existing row (user may have changed role)
      sheet.getRange(i + 1, 1, 1, 6).setValues([[
        data.session_id,
        data.name || '',
        data.role || '',
        data.program || '',
        new Date(),
        formatDate(new Date())
      ]]);
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([
      data.session_id,
      data.name || '',
      data.role || '',
      data.program || '',
      new Date(),
      formatDate(new Date())
    ]);
  }
  refreshDashboard();
}

function handleConversation(data) {
  const sheet = getOrCreateTab(TABS.conversations, [
    'session_id', 'viewer_name', 'viewer_role', 'viewer_program',
    'participant_id', 'participant_name', 'participant_role', 'participant_program',
    'timestamp', 'date', 'status'
  ]);
  sheet.appendRow([
    data.session_id || '',
    data.viewer_name || '',
    data.viewer_role || '',
    data.viewer_program || '',
    data.participant_id || '',
    data.participant_name || '',
    data.participant_role || '',
    data.participant_program || '',
    new Date(),
    formatDate(new Date()),
    'logged'
  ]);
  refreshDashboard();
}

function handleUndoConversation(data) {
  // Mark as undone rather than delete — preserves audit trail
  const sheet = getOrCreateTab(TABS.conversations, []);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === data.session_id && rows[i][4] === data.participant_id) {
      sheet.getRange(i + 1, 11).setValue('undone');
      break;
    }
  }
  refreshDashboard();
}

function handleCoffee(data) {
  const sheet = getOrCreateTab(TABS.coffee, [
    'session_id', 'requester_name', 'requester_role', 'requester_program',
    'participant_id', 'participant_name', 'participant_role', 'participant_program',
    'action', 'timestamp', 'date'
  ]);
  sheet.appendRow([
    data.session_id || '',
    data.requester_name || '',
    data.requester_role || '',
    data.requester_program || '',
    data.participant_id || '',
    data.participant_name || '',
    data.participant_role || '',
    data.participant_program || '',
    data.action || 'selected',   // 'selected' or 'removed'
    new Date(),
    formatDate(new Date())
  ]);
  refreshDashboard();
}

function handleProfileView(data) {
  const sheet = getOrCreateTab(TABS.views, [
    'session_id', 'viewer_role', 'viewer_program',
    'participant_id', 'participant_name', 'participant_role', 'participant_program',
    'timestamp', 'date'
  ]);
  sheet.appendRow([
    data.session_id || '',
    data.viewer_role || '',
    data.viewer_program || '',
    data.participant_id || '',
    data.participant_name || '',
    data.participant_role || '',
    data.participant_program || '',
    new Date(),
    formatDate(new Date())
  ]);
  // Don't refresh dashboard on every view — too frequent, use scheduled refresh instead
}

function handleSurvey(data) {
  const sheet = getOrCreateTab(TABS.survey, [
    'session_id', 'name', 'role', 'program',
    'meeting_happened', 'meeting_useful', 'continue_collaboration',
    'most_useful', 'timestamp', 'date'
  ]);
  sheet.appendRow([
    data.session_id || '',
    data.name || '',
    data.role || '',
    data.program || '',
    data.meeting_happened || '',
    data.meeting_useful || '',
    data.continue_collaboration || '',
    data.most_useful || '',
    new Date(),
    formatDate(new Date())
  ]);
}

// ── DASHBOARD REFRESH ─────────────────────────

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dash = ss.getSheetByName(TABS.dashboard);
  if (!dash) {
    dash = ss.insertSheet(TABS.dashboard);
    // Move dashboard to first position
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  }
  dash.clearContents();

  const users     = getSheetData(TABS.users);
  const convos    = getSheetData(TABS.conversations);
  const coffee    = getSheetData(TABS.coffee);
  const views     = getSheetData(TABS.views);

  // Filter active convos (not undone)
  const activeConvos = convos.filter(r => r.status !== 'undone');
  // Active coffee selections
  const coffeeSelected = coffee.filter(r => r.action === 'selected');
  const coffeeRemoved  = coffee.filter(r => r.action === 'removed');
  // Net coffee selections: selected minus subsequently removed
  const netCoffee = coffeeSelected.filter(sel =>
    !coffeeRemoved.some(rem =>
      rem.session_id === sel.session_id && rem.participant_id === sel.participant_id
      && new Date(rem.timestamp) > new Date(sel.timestamp)
    )
  );

  // ── SECTION 1: EVENT OVERVIEW ──
  const now = new Date();
  writeSection(dash, 1, 1, 'EVENT OVERVIEW', [
    ['Cancer Research Day 2026', 'October 14, 2026'],
    ['Last updated', now.toLocaleString()],
    [],
    ['METRIC', 'COUNT'],
    ['Unique app users', users.length],
    ['Total conversations logged', activeConvos.length],
    ['Coffee Consult selections (net)', netCoffee.length],
    ['Profile views', views.length],
  ]);

  // ── SECTION 2: USERS BY ROLE ──
  const byRole = groupBy(users, 'role');
  const roleRows = [['ROLE', 'COUNT']].concat(
    Object.entries(byRole)
      .sort((a, b) => b[1].length - a[1].length)
      .map(([role, list]) => [role, list.length])
  );
  writeSection(dash, 1, 12, 'APP USERS BY ROLE', roleRows);

  // ── SECTION 3: USERS BY PROGRAM ──
  const byProgram = groupBy(users.filter(u => u.program), 'program');
  const programRows = [['PROGRAM / DEPARTMENT', 'COUNT']].concat(
    Object.entries(byProgram)
      .sort((a, b) => b[1].length - a[1].length)
      .map(([prog, list]) => [prog, list.length])
  );
  writeSection(dash, 1, 26, 'APP USERS BY PROGRAM', programRows);

  // ── SECTION 4: CONVERSATIONS ──
  const convoByCBG = activeConvos.filter(c =>
    c.participant_role === 'PhD Student' &&
    (c.participant_program || '').toLowerCase().includes('cancer biology')
  );
  const convoByFellow = activeConvos.filter(c =>
    c.participant_role === 'Clinical Fellow / Resident'
  );
  writeSection(dash, 5, 1, 'CONVERSATIONS', [
    ['METRIC', 'COUNT'],
    ['Total conversations logged', activeConvos.length],
    ['With CBG PhD students', convoByCBG.length],
    ['With Clinical Fellows', convoByFellow.length],
    [],
    ['MOST-VISITED POSTERS', 'VIEWS'],
    ...getMostVisited(views, 10)
  ]);

  // ── SECTION 5: COFFEE CONSULTS ──
  // Unique matched pairs (both sides selected each other)
  const mutualPairs = findMutualMatches(netCoffee);
  writeSection(dash, 5, 12, 'COFFEE CONSULTS', [
    ['METRIC', 'COUNT'],
    ['Net Coffee Consult selections', netCoffee.length],
    ['Mutual matches (both sides)', mutualPairs.length],
    [],
    ['SELECTIONS', 'REQUESTER → PARTICIPANT'],
    ...netCoffee.slice(0, 15).map(c => [
      c.requester_name + ' (' + c.requester_role + ')',
      c.participant_name + ' (' + c.participant_role + ')'
    ])
  ]);

  // ── SECTION 6: PROFILE VIEWS ──
  const viewsByRole = groupBy(views, 'viewer_role');
  writeSection(dash, 5, 26, 'PROFILE VIEWS BY VIEWER ROLE', [
    ['VIEWER ROLE', 'VIEWS'],
    ...Object.entries(viewsByRole)
      .sort((a, b) => b[1].length - a[1].length)
      .map(([role, list]) => [role, list.length])
  ]);

  // Style the dashboard
  styleDashboard(dash);
}

// ── HELPERS ───────────────────────────────────

function getOrCreateTab(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#990011')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getSheetData(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function groupBy(arr, key) {
  return arr.reduce((acc, item) => {
    const k = item[key] || 'Unknown';
    if (!acc[k]) acc[k] = [];
    acc[k].push(item);
    return acc;
  }, {});
}

function getMostVisited(views, n) {
  const counts = {};
  views.forEach(v => {
    const key = v.participant_name + ' — ' + v.participant_id;
    counts[key] = (counts[key] || 0) + 1;
  });
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .map(([name, count]) => [name, count]);
}

function findMutualMatches(coffeeList) {
  const pairs = [];
  coffeeList.forEach(a => {
    const mutual = coffeeList.find(b =>
      b.session_id === a.participant_id &&
      b.participant_id === a.session_id
    );
    if (mutual && !pairs.find(p =>
      (p[0] === a.requester_name && p[1] === mutual.requester_name) ||
      (p[1] === a.requester_name && p[0] === mutual.requester_name)
    )) {
      pairs.push([a.requester_name, mutual.requester_name]);
    }
  });
  return pairs;
}

function writeSection(sheet, col, row, title, rows) {
  sheet.getRange(row, col).setValue(title)
    .setBackground('#990011')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11);

  rows.forEach((r, i) => {
    if (r.length === 0) return;
    const range = sheet.getRange(row + 1 + i, col, 1, r.length);
    range.setValues([r]);
    // Style header rows (all caps first cell)
    if (typeof r[0] === 'string' && r[0] === r[0].toUpperCase() && r[0].length > 1) {
      range.setBackground('#F4F4F4').setFontWeight('bold');
    }
  });
}

function styleDashboard(dash) {
  // Auto-resize columns
  dash.autoResizeColumns(1, 30);
  // Freeze nothing — dashboard is read-only reference
  dash.setTabColor('#990011');
  // Style all data tabs
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [TABS.users, TABS.conversations, TABS.coffee, TABS.views, TABS.survey].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) s.autoResizeColumns(1, s.getLastColumn());
  });
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

// ── SCHEDULED REFRESH ─────────────────────────
// Set this to run every 30 minutes via Triggers (Clock trigger)
// Keeps dashboard current even if no new actions come in

function scheduledRefresh() {
  refreshDashboard();
}
