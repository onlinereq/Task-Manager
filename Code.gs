/**
 * Google Apps Script Web App Backend for Team Task Manager
 * Timezone: Asia/Phnom_Penh
 * Deployed as Web App (Supports both doGet and doPost)
 * All endpoints require server-side credential verification.
 */

// ============================================================================
// CONFIGURATION & CONSTANTS
// ============================================================================
var CONFIG = {
  TIMEZONE: 'Asia/Phnom_Penh',
  DATETIME_FORMAT: 'yyyy-MM-dd HH:mm:ss',
  DATE_FORMAT: 'yyyy-MM-dd',
  SHEETS: {
    USERS: 'Users',
    TASKS: 'Tasks',
    HISTORY: 'History',
    SETTINGS: 'Settings'
  },
  DEFAULT_STATUSES: ['Register', 'In Progress', 'Review', 'Done'],
  VALID_STATES: ['Active', 'Completed', 'Archived'],
  DATA_VERSION: 4,            // Increment integer to force instant client-side cache refresh
  MAX_FAILED_ATTEMPTS: 5,     // 5 failed attempts per username triggers lockout
  LOCKOUT_SECONDS: 900,       // 15 minutes lockout duration (900 seconds)
  FAILURE_DELAY_MS: 1200,     // Delay applied only on authentication failure
  CACHE_USERS_SEC: 300,       // Cache user hashes in CacheService for 5 minutes
  CACHE_SNAPSHOT_SEC: 60,     // 60 seconds
  LOCK_TIMEOUT_MS: 15000      // 15 seconds for write locks
};

// ============================================================================
// SCRIPT PROPERTIES & DATABASE ACCESS
// ============================================================================

/**
 * Retrieves the spreadsheet instance.
 * Reads SPREADSHEET_ID from ScriptProperties if configured,
 * otherwise falls back to the bound ActiveSpreadsheet.
 */
function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SPREADSHEET_ID');
  if (sheetId && sheetId.trim()) {
    try {
      return SpreadsheetApp.openById(sheetId.trim());
    } catch (e) {
      // Fallback to active spreadsheet if openById fails
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Retrieves the cryptographic salt used for PIN hashing.
 * Persisted in ScriptProperties so hashes remain stable across sessions.
 */
function getPinSalt() {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('PIN_SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty('PIN_SALT', salt);
  }
  return salt;
}

// ============================================================================
// CRYPTOGRAPHY, HASHING & MIGRATION
// ============================================================================

/**
 * Computes a salted SHA-256 hex digest for a PIN.
 */
function hashPin(pin, salt) {
  if (pin === null || pin === undefined) return '';
  var raw = String(pin).trim() + salt;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < digest.length; i++) {
    var b = digest[i];
    if (b < 0) b += 256;
    var byteStr = b.toString(16);
    if (byteStr.length === 1) byteStr = '0' + byteStr;
    hex += byteStr;
  }
  return hex;
}

/**
 * One-time migration function to convert existing plain-text PINs in the
 * Users sheet into salted SHA-256 hashes. Run once from Apps Script editor.
 */
function migrateUserPinsToHashes() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (!sheet) {
    throw new Error('Users sheet not found. Run setupSheet() first.');
  }

  var salt = getPinSalt();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 'No users found to migrate.';

  var migratedCount = 0;
  for (var i = 1; i < data.length; i++) {
    var rawPin = String(data[i][1] || '').trim();
    if (!rawPin) continue;

    // Check if already a 64-character hex hash
    var isAlreadyHashed = /^[0-9a-fA-F]{64}$/.test(rawPin);
    if (!isAlreadyHashed) {
      var hashed = hashPin(rawPin, salt);
      var cell = sheet.getRange(i + 1, 2);
      cell.setNumberFormat('@');
      cell.setValue(hashed);
      migratedCount++;
    }
  }

  // Invalidate users cache so fresh hashes are immediately loaded
  CacheService.getScriptCache().remove('users_list');
  SpreadsheetApp.flush();

  return 'Migration complete: ' + migratedCount + ' PIN(s) converted to salted SHA-256 hashes.';
}

// ============================================================================
// FORMULA INJECTION & SANITIZATION HELPERS
// ============================================================================

/**
 * Prevents formula injection by prepending an apostrophe to strings
 * starting with '=', '+', '-', or '@'.
 */
function sanitizeCellValue(val) {
  if (val === null || val === undefined) return '';
  var str = String(val);
  if (str.length > 0 && /^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

/**
 * Strips leading escape apostrophe if it was prepended to escape formulas.
 */
function unescapeCellValue(val) {
  if (val === null || val === undefined) return '';
  var str = String(val);
  if (str.length > 1 && str.charAt(0) === "'" && /^[=+\-@]/.test(str.substring(1))) {
    return str.substring(1);
  }
  return str;
}

// ============================================================================
// ONE-TIME SETUP FUNCTION (SAFE FOR REPEATED EXECUTION)
// ============================================================================

/**
 * This function is safe to run multiple times. It never deletes existing data.
 * Initializes missing database tabs, verifies expected header columns, and applies formatting.
 */
function setupSheet() {
  var ss = getSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  var summaryLogs = [];

  // 1. Setup Users Tab
  var existedUsers = !!ss.getSheetByName(CONFIG.SHEETS.USERS);
  var usersSheet = getOrCreateSheet(ss, CONFIG.SHEETS.USERS);
  var usersResult = setupUsersSheet(usersSheet);
  if (!existedUsers || usersResult.isBrandNew) {
    summaryLogs.push('Users sheet: created new with 1 starter user (Alex Chen).');
  } else {
    var addedNote = usersResult.addedHeaders.length > 0 ? ' (added missing header(s): ' + usersResult.addedHeaders.join(', ') + ')' : '';
    summaryLogs.push('Users sheet: already existed, ' + usersResult.rowCount + ' row(s) kept' + addedNote + '.');
  }

  // 2. Setup Tasks Tab
  var existedTasks = !!ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var tasksSheet = getOrCreateSheet(ss, CONFIG.SHEETS.TASKS);
  var tasksResult = setupTasksSheet(tasksSheet);
  if (!existedTasks || tasksResult.isBrandNew) {
    summaryLogs.push('Tasks sheet: created new.');
  } else {
    var addedTaskNote = tasksResult.addedHeaders.length > 0 ? ' (added missing header(s): ' + tasksResult.addedHeaders.join(', ') + ')' : '';
    summaryLogs.push('Tasks sheet: already existed, ' + tasksResult.rowCount + ' task(s) kept' + addedTaskNote + '.');
  }

  // 3. Setup History Tab
  var existedHistory = !!ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  var historySheet = getOrCreateSheet(ss, CONFIG.SHEETS.HISTORY);
  var historyResult = setupHistorySheet(historySheet);
  if (!existedHistory || historyResult.isBrandNew) {
    summaryLogs.push('History sheet: created new.');
  } else {
    var addedHistNote = historyResult.addedHeaders.length > 0 ? ' (added missing header(s): ' + historyResult.addedHeaders.join(', ') + ')' : '';
    summaryLogs.push('History sheet: already existed, ' + historyResult.rowCount + ' history row(s) kept' + addedHistNote + '.');
  }

  // 4. Setup Settings Tab
  var existedSettings = !!ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  var settingsSheet = getOrCreateSheet(ss, CONFIG.SHEETS.SETTINGS);
  var settingsResult = setupSettingsSheet(settingsSheet);
  if (!existedSettings || settingsResult.isBrandNew) {
    summaryLogs.push('Settings sheet: created new with default statuses.');
  } else {
    summaryLogs.push('Settings sheet: already existed, ' + settingsResult.rowCount + ' status(es) kept.');
  }

  // 5. Clean default Sheet1 only if completely empty and unused
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    if (sheet.getName().toLowerCase() === 'sheet1' && allSheets.length > 1) {
      if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
        ss.deleteSheet(sheet);
        summaryLogs.push('Sheet1: removed (empty default tab).');
      }
      break;
    }
  }

  // Clear caches
  var cache = CacheService.getScriptCache();
  cache.remove('users_list');
  cache.remove('data_snapshot');
  cache.remove('all_tasks_raw');

  SpreadsheetApp.flush();

  var finalSummary = summaryLogs.join('\n');
  Logger.log('setupSheet() Summary:\n' + finalSummary);
  return finalSummary;
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Ensures all expected headers exist in row 1 of the sheet.
 * Never touches or overwrites existing data rows (rows 2..N).
 * If expected headers are missing, appends them to row 1.
 * Returns an array of newly added header names.
 */
function ensureHeaderColumns(sheet, expectedHeaders, columnWidths) {
  var addedHeaders = [];
  var lastCol = sheet.getLastColumn();
  var existingHeaders = [];

  if (lastCol > 0 && sheet.getLastRow() > 0) {
    var headerRowValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headerRowValues.length; i++) {
      existingHeaders.push(String(headerRowValues[i] || '').trim().toLowerCase());
    }
  }

  // If sheet was empty (no row 1 at all)
  if (existingHeaders.length === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    formatHeaderRow(sheet, expectedHeaders.length);
    sheet.setFrozenRows(1);
    if (columnWidths) {
      for (var w = 0; w < columnWidths.length; w++) {
        sheet.setColumnWidth(w + 1, columnWidths[w]);
      }
    }
    return expectedHeaders.slice();
  }

  // If headers exist, check which expected headers are missing
  var currentColCount = lastCol;
  for (var h = 0; h < expectedHeaders.length; h++) {
    var expectedName = expectedHeaders[h];
    var normExpected = expectedName.toLowerCase();

    // Check if this header already exists anywhere in row 1
    var exists = existingHeaders.indexOf(normExpected) !== -1;
    if (!exists) {
      currentColCount++;
      var newCell = sheet.getRange(1, currentColCount);
      newCell.setValue(expectedName);
      formatHeaderCell(newCell);
      if (columnWidths && columnWidths[h]) {
        sheet.setColumnWidth(currentColCount, columnWidths[h]);
      }
      addedHeaders.push(expectedName);
      existingHeaders.push(normExpected);
    }
  }

  // Ensure row 1 is frozen
  try {
    if (sheet.getFrozenRows() < 1) {
      sheet.setFrozenRows(1);
    }
  } catch (e) {}

  return addedHeaders;
}

function setupUsersSheet(sheet) {
  var headers = ['Name', 'PIN', 'Team', 'Status'];
  var widths = [180, 280, 160, 120];
  var lastRow = sheet.getLastRow();
  var isBrandNew = (lastRow === 0 && sheet.getLastColumn() === 0);

  var addedHeaders = ensureHeaderColumns(sheet, headers, widths);

  if (isBrandNew) {
    // Starter user with salted SHA-256 hash (default PIN: '1234')
    var salt = getPinSalt();
    var starterHash = hashPin('1234', salt);
    sheet.appendRow(['Alex Chen', starterHash, 'Core Team', 'Active']);
  }

  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  applyAlternatingColors(sheet, 1, headers.length);

  return {
    isBrandNew: isBrandNew,
    rowCount: Math.max(sheet.getLastRow() - 1, 0),
    addedHeaders: addedHeaders
  };
}

function setupTasksSheet(sheet) {
  var headers = [
    'Task ID',
    'Task Name',
    'Description',
    'Assigned To',
    'Deadline',
    'Status',
    'State',
    'Created By',
    'CreatedAt',
    'Last Updated By',
    'Last Updated At',
    'IsPrivate'
  ];
  var widths = [120, 220, 280, 140, 110, 110, 100, 120, 150, 120, 150, 100];
  var lastRow = sheet.getLastRow();
  var isBrandNew = (lastRow === 0 && sheet.getLastColumn() === 0);

  var addedHeaders = ensureHeaderColumns(sheet, headers, widths);

  sheet.getRange(2, 3, Math.max(sheet.getMaxRows() - 1, 1), 1).setWrap(true);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  applyAlternatingColors(sheet, 1, headers.length);
  applyStatusConditionalFormatting(sheet, 6);

  return {
    isBrandNew: isBrandNew,
    rowCount: Math.max(sheet.getLastRow() - 1, 0),
    addedHeaders: addedHeaders
  };
}

function setupHistorySheet(sheet) {
  var headers = [
    'History ID',
    'Task ID',
    'Task Name',
    'Status',
    'Note',
    'Updated By',
    'Timestamp'
  ];
  var widths = [140, 120, 220, 110, 300, 130, 160];
  var lastRow = sheet.getLastRow();
  var isBrandNew = (lastRow === 0 && sheet.getLastColumn() === 0);

  var addedHeaders = ensureHeaderColumns(sheet, headers, widths);

  sheet.getRange(2, 5, Math.max(sheet.getMaxRows() - 1, 1), 1).setWrap(true);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  applyAlternatingColors(sheet, 1, headers.length);

  return {
    isBrandNew: isBrandNew,
    rowCount: Math.max(sheet.getLastRow() - 1, 0),
    addedHeaders: addedHeaders
  };
}

function setupSettingsSheet(sheet) {
  var headers = ['Status List'];
  var widths = [180];
  var lastRow = sheet.getLastRow();
  var isBrandNew = (lastRow === 0 && sheet.getLastColumn() === 0);

  var addedHeaders = ensureHeaderColumns(sheet, headers, widths);

  var statusCount = Math.max(sheet.getLastRow() - 1, 0);
  if (statusCount === 0) {
    for (var i = 0; i < CONFIG.DEFAULT_STATUSES.length; i++) {
      sheet.appendRow([CONFIG.DEFAULT_STATUSES[i]]);
    }
    statusCount = CONFIG.DEFAULT_STATUSES.length;
  }

  applyAlternatingColors(sheet, 1, 1);

  return {
    isBrandNew: isBrandNew,
    rowCount: statusCount,
    addedHeaders: addedHeaders
  };
}

function formatHeaderCell(cell) {
  cell
    .setBackground('#F2F4F7')
    .setFontColor('#1D1D1F')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#D0D5DD', SpreadsheetApp.BorderStyle.SOLID);
}

function formatHeaderRow(sheet, numColumns) {
  if (!numColumns || numColumns <= 0) return;
  var headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange
    .setBackground('#F2F4F7')
    .setFontColor('#1D1D1F')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#D0D5DD', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 34);
}

function applyAlternatingColors(sheet, startCol, endCol) {
  try {
    var existingBandings = sheet.getBandings();
    if (!existingBandings || existingBandings.length === 0) {
      var range = sheet.getRange(1, startCol, Math.max(sheet.getMaxRows(), 2), endCol - startCol + 1);
      range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    }
  } catch (e) {}
}

function applyStatusConditionalFormatting(sheet, statusColIndex) {
  try {
    var range = sheet.getRange(2, statusColIndex, Math.max(sheet.getMaxRows() - 1, 1), 1);
    var rules = sheet.getConditionalFormatRules() || [];

    var existingRulesText = '';
    for (var r = 0; r < rules.length; r++) {
      var crit = rules[r].getBooleanCriteria();
      if (crit) {
        existingRulesText += crit.getCriteriaValues().join(',') + ';';
      }
    }

    var colorMap = [
      { text: 'Register', bg: '#F2F4F7', font: '#475467' },
      { text: 'In Progress', bg: '#EBF3FF', font: '#0071E3' },
      { text: 'Review', bg: '#FEF3E6', font: '#B45309' },
      { text: 'Done', bg: '#EDFBF0', font: '#1B7A34' }
    ];

    var needsUpdate = false;
    colorMap.forEach(function (cfg) {
      if (existingRulesText.indexOf(cfg.text) === -1) {
        var rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(cfg.text)
          .setBackground(cfg.bg)
          .setFontColor(cfg.font)
          .setRanges([range])
          .build();
        rules.push(rule);
        needsUpdate = true;
      }
    });

    if (needsUpdate) {
      sheet.setConditionalFormatRules(rules);
    }
  } catch (e) {}
}

// ============================================================================
// SERVER-SIDE CREDENTIAL VERIFICATION & LOCKOUT PROTECTION
// ============================================================================

/**
 * Authenticates user credentials with server-side lockout and rate limiting.
 * - Lockout: 5 failed attempts per name blocks for 15 minutes in CacheService.
 * - Generic error returned ({ok:false, error:"unauthorized"}), never reveals if name or PIN was wrong.
 * - Adds a small delay only on failure.
 * - Caches user hashes in CacheService for 5 minutes.
 * - Never returns PINs or raw hash data to callers.
 */
function verifyCredentials(ss, name, pin) {
  var rawName = String(name || '').trim();
  var rawPin = (pin !== undefined && pin !== null) ? String(pin).trim() : '';

  // Empty credentials fail immediately with unauthorized
  if (!rawName || !rawPin) {
    Utilities.sleep(CONFIG.FAILURE_DELAY_MS);
    return { ok: false, error: 'unauthorized' };
  }

  var normName = rawName.toLowerCase();
  var cache = CacheService.getScriptCache();
  var lockoutKey = 'lockout_' + normName;
  var failKey = 'fails_' + normName;

  // 1. Check if user is locked out
  var isLocked = cache.get(lockoutKey);
  if (isLocked) {
    return { ok: false, error: 'locked' };
  }

  // 2. Load cached user list (names, hashed PINs, team, status)
  var users = getCachedUsers(ss);
  var matchedUser = null;

  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (String(u.name || '').trim().toLowerCase() === normName) {
      matchedUser = u;
      break;
    }
  }

  // Helper to record a failure and enforce 5-failure lockout
  function handleAuthFailure() {
    var currentFails = Number(cache.get(failKey) || 0) + 1;
    if (currentFails >= CONFIG.MAX_FAILED_ATTEMPTS) {
      cache.put(lockoutKey, '1', CONFIG.LOCKOUT_SECONDS);
      cache.remove(failKey);
      Utilities.sleep(CONFIG.FAILURE_DELAY_MS);
      return { ok: false, error: 'locked' };
    } else {
      cache.put(failKey, String(currentFails), CONFIG.LOCKOUT_SECONDS);
      Utilities.sleep(CONFIG.FAILURE_DELAY_MS);
      return { ok: false, error: 'unauthorized' };
    }
  }

  // User does not exist or is inactive
  if (!matchedUser || String(matchedUser.status || '').trim().toLowerCase() !== 'active') {
    return handleAuthFailure();
  }

  // 3. Verify PIN against stored hash
  var salt = getPinSalt();
  var inputHash = hashPin(rawPin, salt);
  var storedHash = String(matchedUser.hash || '').trim();

  var isMatch = false;
  if (/^[0-9a-fA-F]{64}$/.test(storedHash)) {
    isMatch = (inputHash.toLowerCase() === storedHash.toLowerCase());
  } else {
    // Unmigrated legacy plain PIN fallback
    isMatch = (rawPin === storedHash);
  }

  if (!isMatch) {
    return handleAuthFailure();
  }

  // 4. Successful verification: clear failed attempt counters
  cache.remove(failKey);
  cache.remove(lockoutKey);

  return {
    ok: true,
    user: {
      name: matchedUser.name,
      team: matchedUser.team,
      status: matchedUser.status
    }
  };
}

/**
 * Loads user list from Users sheet, caching hashed credentials for 5 minutes.
 * Never stores plain PINs.
 */
function getCachedUsers(ss) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('users_list');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var userName = String(row[0] || '').trim();
    if (userName) {
      users.push({
        name: userName,
        hash: String(row[1] || '').trim(), // Salted SHA-256 hash
        team: String(row[2] || '').trim(),
        status: String(row[3] || '').trim()
      });
    }
  }

  try {
    cache.put('users_list', JSON.stringify(users), CONFIG.CACHE_USERS_SEC);
  } catch (e) {}

  return users;
}

function invalidateDataCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove('data_snapshot');
    cache.remove('all_tasks_raw');
  } catch (e) {}
}

// ============================================================================
// MANUAL LOCKOUT RESET & MANAGEMENT
// ============================================================================

/**
 * Resets a user's 15-minute lockout and clears their failed-attempt counter.
 * CacheService keys cleared:
 *  - 'lockout_' + lowerName (lockout flag)
 *  - 'fails_' + lowerName   (failed attempts counter)
 *
 * @param {string} [name] - Target user name to unlock. If omitted, uses DEFAULT_EDITOR_TARGET.
 * @param {string} [source] - Diagnostic label for the reset audit log.
 */
function clearLockout(name, source) {
  var DEFAULT_EDITOR_TARGET = 'Alex Chen'; // <-- Edit name here if running clearLockout directly from Apps Script editor
  var rawName = String(name || DEFAULT_EDITOR_TARGET).trim();
  if (!rawName) {
    return { ok: false, error: 'invalid_name', message: 'User name is required.' };
  }

  var normName = rawName.toLowerCase();
  var cache = CacheService.getScriptCache();
  var lockoutKey = 'lockout_' + normName;
  var failKey = 'fails_' + normName;

  cache.remove(lockoutKey);
  cache.remove(failKey);

  logLockoutReset(rawName, source || 'manual');

  var msg = 'Lockout and failed attempts successfully cleared for: ' + rawName;
  Logger.log(msg);

  return {
    ok: true,
    success: true,
    message: msg,
    user: rawName,
    keysCleared: [lockoutKey, failKey]
  };
}

/**
 * Direct execution from Apps Script Editor dropdown.
 * Edit TARGET_NAME below to the name you want to unlock, then click Run.
 */
function runClearLockout() {
  var TARGET_NAME = 'Alex Chen'; // <-- Edit this name to unlock a user from the Apps Script editor
  var result = clearLockout(TARGET_NAME, 'Apps Script Editor: runClearLockout');
  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * Handles action: "clearLockout" from the web app API.
 * Requires a separate admin password configured in Script Properties ('ADMIN_PASSWORD').
 */
function handleClearLockoutAction(targetName, adminPassword) {
  var props = PropertiesService.getScriptProperties();
  var configuredAdminPass = props.getProperty('ADMIN_PASSWORD');

  if (!configuredAdminPass) {
    return {
      ok: false,
      error: 'unauthorized',
      message: 'ADMIN_PASSWORD is not configured in Script Properties.'
    };
  }

  if (!adminPassword || String(adminPassword).trim() !== configuredAdminPass.trim()) {
    Utilities.sleep(CONFIG.FAILURE_DELAY_MS);
    return {
      ok: false,
      error: 'unauthorized',
      message: 'Invalid admin credentials.'
    };
  }

  return clearLockout(targetName, 'API action (admin password verified)');
}

/**
 * Records manual lockout resets with timestamp and source in ScriptProperties.
 */
function logLockoutReset(name, source) {
  try {
    var timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);
    var props = PropertiesService.getScriptProperties();
    var rawLogs = props.getProperty('LOCKOUT_RESET_LOGS');
    var logs = [];
    if (rawLogs) {
      try { logs = JSON.parse(rawLogs); } catch (e) {}
    }
    logs.unshift({
      name: name,
      timestamp: timestamp,
      source: source || 'manual'
    });
    if (logs.length > 50) logs = logs.slice(0, 50);
    props.setProperty('LOCKOUT_RESET_LOGS', JSON.stringify(logs));
  } catch (e) {
    Logger.log('Could not write reset log: ' + e);
  }
}

// ============================================================================
// WEB APP API DISPATCHER (doGet & doPost)
// ============================================================================

/**
 * Handles GET requests. Credentials and parameters passed via query string.
 * Every request must include name and pin, verified before returning data.
 */
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  return executeApiRequest(params.action || 'read', params.name, params.pin, params);
}

/**
 * Handles POST requests. Payload passed as text/plain JSON.
 * Every request must include name and pin, verified before returning data or changing state.
 */
function doPost(e) {
  var payload = {};
  if (e && e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return createJsonResponse({ ok: false, error: 'unauthorized' });
    }
  } else {
    return createJsonResponse({ ok: false, error: 'unauthorized' });
  }

  var action = payload.action || 'read';
  var name = payload.name;
  var pin = payload.pin;

  return executeApiRequest(action, name, pin, payload);
}

/**
 * Unified request processor:
 * 1. Authenticates name & PIN against Users tab with lockout.
 * 2. Enforces authorization rules:
 *    - Anyone authenticated can update task status.
 *    - Only the task owner can edit, change state, or delete their task.
 * 3. Sanitizes all inputs and writes safely with formula injection prevention.
 */
function executeApiRequest(action, name, pin, payload) {
  try {
    var ss = getSpreadsheet();

    // 0. Protected admin action: clearLockout (authenticated via ADMIN_PASSWORD)
    if (action === 'clearLockout') {
      var targetName = payload.targetName || name || '';
      var adminPassword = payload.adminPassword || pin || '';
      return createJsonResponse(handleClearLockoutAction(targetName, adminPassword));
    }

    // 1. Mandatory server-side authentication for ALL requests and actions
    var authResult = verifyCredentials(ss, name, pin);
    if (!authResult.ok) {
      return createJsonResponse({
        ok: false,
        error: authResult.error // strictly "unauthorized" or "locked"
      });
    }

    var currentUser = authResult.user;

    // 2. Dispatch action
    // READ / SYNC / EXPORT ACTIONS: No write lock needed
    if (action === 'start' || action === 'login' || action === 'sync' || action === 'read') {
      var readData = handleStart(ss, payload, currentUser);
      return createJsonResponse(readData);
    }

    if (action === 'export' || action === 'exportData') {
      var exportData = handleExport(ss, currentUser);
      return createJsonResponse(exportData);
    }

    // WRITE ACTIONS: Wrapped with LockService
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    } catch (lockError) {
      return createJsonResponse({
        ok: false,
        success: false,
        error: 'busy',
        message: 'Server is currently busy with another write. Please retry.'
      });
    }

    try {
      var writeResponse;
      switch (action) {
        case 'createTask':
        case 'add':
          writeResponse = handleCreateTask(ss, payload, currentUser);
          break;

        case 'updateTask':
        case 'update':
          writeResponse = handleUpdateTask(ss, payload, currentUser);
          break;

        case 'setState':
        case 'edit':
          writeResponse = handleSetState(ss, payload, currentUser);
          break;

        case 'deleteTask':
        case 'delete':
          writeResponse = handleDeleteTask(ss, payload, currentUser);
          break;

        case 'setPrivacy':
        case 'togglePrivacy':
          writeResponse = handleSetPrivacy(ss, payload, currentUser);
          break;

        case 'addStatus':
          writeResponse = handleAddStatus(ss, payload, currentUser);
          break;

        case 'editStatus':
          writeResponse = handleEditStatus(ss, payload, currentUser);
          break;

        case 'deleteStatus':
          writeResponse = handleDeleteStatus(ss, payload, currentUser);
          break;

        case 'saveStatuses':
          writeResponse = handleSaveStatuses(ss, payload, currentUser);
          break;

        default:
          writeResponse = {
            ok: false,
            success: false,
            error: 'invalid_action',
            message: 'Unknown action: ' + action
          };
      }

      if (writeResponse && writeResponse.ok !== false && writeResponse.success !== false) {
        invalidateDataCache();
      }

      return createJsonResponse(writeResponse);
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return createJsonResponse({
      ok: false,
      success: false,
      error: 'server_error',
      message: 'Server error: ' + (err.message || err.toString())
    });
  }
}

// ============================================================================
// ACTION HANDLERS
// ============================================================================

/**
 * Action: start / login / sync / read
 * Returns user info, tasks, history, statuses, and sanitized active users.
 * Never returns PINs or password hashes.
 */
function handleStart(ss, payload, currentUser) {
  var action = String(payload.action || '').trim();
  var lastSync = payload.lastSync ? String(payload.lastSync).trim() : '';
  var forceRefresh = payload.forceRefresh === true || action === 'start' || action === 'login' || action === 'clearCache';
  var serverTime = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);

  if (forceRefresh) {
    invalidateDataCache();
  }

  var statuses = fetchStatuses(ss);
  var activeUsers = fetchActiveUsers(ss); // PINs stripped
  var rawTasks = fetchTasks(ss);
  var rawHistory = fetchHistory(ss);

  // Filter tasks: include if NOT private OR created by currentUser
  var currentUserName = (currentUser && currentUser.name ? currentUser.name : '').trim().toLowerCase();
  var allTasks = rawTasks.filter(function (t) {
    if (!t.isPrivate) return true;
    return (t.createdBy || '').trim().toLowerCase() === currentUserName;
  });

  // Allowed task IDs map for this user
  var allowedTaskIds = {};
  allTasks.forEach(function (t) {
    allowedTaskIds[String(t.taskId).toLowerCase()] = true;
  });

  // Filter history: only include history for tasks visible to this user
  var allHistory = rawHistory.filter(function (h) {
    return !!allowedTaskIds[String(h.taskId).toLowerCase()];
  });

  var serverTaskIds = allTasks.map(function (t) { return t.taskId; });

  if (forceRefresh || !lastSync || allTasks.length === 0) {
    return {
      ok: true,
      success: true,
      user: currentUser,
      data: {
        tasks: allTasks,
        history: allHistory,
        serverTaskIds: serverTaskIds,
        statuses: statuses,
        users: activeUsers,
        serverTime: serverTime,
        dataVersion: CONFIG.DATA_VERSION,
        delta: false
      }
    };
  }

  // Delta sync for quiet background updates
  var lastSyncEpoch = parseToEpoch(lastSync);
  var deltaTasks = allTasks.filter(function (t) {
    var tEpoch = parseToEpoch(t.lastUpdatedAt || t.createdAt || '');
    return tEpoch >= lastSyncEpoch;
  });

  var deltaHistory = allHistory.filter(function (h) {
    var hEpoch = parseToEpoch(h.timestamp || '');
    return hEpoch >= lastSyncEpoch;
  });

  return {
    ok: true,
    success: true,
    user: currentUser,
    data: {
      tasks: deltaTasks,
      history: deltaHistory,
      serverTaskIds: serverTaskIds,
      statuses: statuses,
      users: activeUsers,
      serverTime: serverTime,
      dataVersion: CONFIG.DATA_VERSION,
      delta: true
    }
  };
}

/**
 * Action: export / exportData
 * Secure export of tasks and history without leaking user PINs or system hashes.
 */
function handleExport(ss, currentUser) {
  var serverTime = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);
  var rawTasks = fetchTasks(ss);
  var rawHistory = fetchHistory(ss);
  var currentUserName = (currentUser && currentUser.name ? currentUser.name : '').trim().toLowerCase();

  var tasks = rawTasks.filter(function (t) {
    if (!t.isPrivate) return true;
    return (t.createdBy || '').trim().toLowerCase() === currentUserName;
  });

  var allowedTaskIds = {};
  tasks.forEach(function (t) {
    allowedTaskIds[String(t.taskId).toLowerCase()] = true;
  });

  var history = rawHistory.filter(function (h) {
    return !!allowedTaskIds[String(h.taskId).toLowerCase()];
  });

  return {
    ok: true,
    success: true,
    user: currentUser,
    exportedAt: serverTime,
    tasks: tasks,
    history: history,
    statuses: fetchStatuses(ss)
  };
}

/**
 * Action: createTask / add
 * Validates inputs, prevents formula injection, inserts task row and initial history row.
 */
function handleCreateTask(ss, payload, currentUser) {
  // Validate and sanitize inputs; ignore unexpected fields
  var taskName = String(payload.taskName || '').trim();
  var description = String(payload.description || '').trim();
  var assignedTo = String(payload.assignedTo || '').trim();
  var deadline = String(payload.deadline || '').trim();
  var clientTaskId = payload.taskId ? String(payload.taskId).trim() : '';
  var clientHistoryId = payload.historyId ? String(payload.historyId).trim() : '';

  if (!taskName || taskName.length > 200) {
    return { ok: false, success: false, message: 'Valid task name (1-200 characters) is required.' };
  }
  if (!assignedTo || assignedTo.length > 100) {
    return { ok: false, success: false, message: 'Valid Assigned To name is required.' };
  }
  if (!deadline || !/^\d{4}-\d{2}-\d{2}$/.test(deadline)) {
    return { ok: false, success: false, message: 'Valid deadline (YYYY-MM-DD) is required.' };
  }
  if (description.length > 5000) {
    description = description.substring(0, 5000);
  }

  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);

  if (!tasksSheet || !historySheet) {
    return { ok: false, success: false, message: 'Required database sheets missing. Run setupSheet().' };
  }

  var now = new Date();
  var formattedNow = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);
  var taskId = clientTaskId || ('TSK-' + now.getTime().toString(36).toUpperCase());
  var initialStatus = 'Register';
  var initialState = 'Active';
  var isPrivate = (payload.isPrivate === true || String(payload.isPrivate).trim().toLowerCase() === 'true' || payload.isPrivate === 1);

  // Apply formula injection prevention
  tasksSheet.appendRow([
    sanitizeCellValue(taskId),
    sanitizeCellValue(taskName),
    sanitizeCellValue(description),
    sanitizeCellValue(assignedTo),
    sanitizeCellValue(deadline),
    sanitizeCellValue(initialStatus),
    sanitizeCellValue(initialState),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow),
    isPrivate ? true : false
  ]);

  // First history log
  var historyId = clientHistoryId || ('HIST-' + now.getTime().toString(36).toUpperCase() + '-1');
  historySheet.appendRow([
    sanitizeCellValue(historyId),
    sanitizeCellValue(taskId),
    sanitizeCellValue(taskName),
    sanitizeCellValue(initialStatus),
    sanitizeCellValue('Task created'),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow)
  ]);

  return {
    ok: true,
    success: true,
    task: {
      taskId: taskId,
      taskName: taskName,
      description: description,
      assignedTo: assignedTo,
      deadline: deadline,
      status: initialStatus,
      state: initialState,
      createdBy: currentUser.name,
      createdAt: formattedNow,
      lastUpdatedBy: currentUser.name,
      lastUpdatedAt: formattedNow,
      isPrivate: isPrivate
    },
    historyEntry: {
      historyId: historyId,
      taskId: taskId,
      taskName: taskName,
      status: initialStatus,
      note: 'Task created',
      updatedBy: currentUser.name,
      timestamp: formattedNow
    }
  };
}

/**
 * Action: updateTask / update
 * Authorization rule: ANY authenticated user can update task status and append a history note.
 * Validates allowed status values.
 */
function handleUpdateTask(ss, payload, currentUser) {
  var taskId = String(payload.taskId || '').trim();
  var newStatus = String(payload.status || '').trim();
  var note = String(payload.note || '').trim();
  var clientHistoryId = payload.historyId ? String(payload.historyId).trim() : '';

  if (!taskId) {
    return { ok: false, success: false, message: 'Task ID is required.' };
  }
  if (!newStatus) {
    return { ok: false, success: false, message: 'Status is required.' };
  }

  // Validate allowed status value
  var validStatuses = fetchStatuses(ss);
  var isValidStatus = validStatuses.some(function (st) {
    return st.toLowerCase() === newStatus.toLowerCase();
  });
  if (!isValidStatus) {
    return { ok: false, success: false, message: 'Status "' + newStatus + '" is not in allowed status list.' };
  }

  if (note.length > 2000) {
    note = note.substring(0, 2000);
  }

  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);

  var taskRowIndex = findRowIndexByColumnValue(tasksSheet, 1, taskId);
  if (taskRowIndex === -1) {
    return { ok: false, success: false, message: 'Task ID not found.' };
  }

  // Check privacy authorization: private tasks can only be updated by their creator
  var rawPrivate = tasksSheet.getRange(taskRowIndex, 12).getValue();
  var isPrivate = (rawPrivate === true || String(rawPrivate).trim().toLowerCase() === 'true' || rawPrivate === 1 || String(rawPrivate).trim() === '1');
  var createdBy = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 8).getValue() || '');
  if (isPrivate && createdBy.trim().toLowerCase() !== currentUser.name.trim().toLowerCase()) {
    return {
      ok: false,
      success: false,
      error: 'forbidden',
      message: 'Private tasks can only be updated by their creator.'
    };
  }

  var now = new Date();
  var formattedNow = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);

  var canonicalTaskId = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 1).getValue() || taskId);
  var taskName = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 2).getValue() || '');

  // Update Status (col 6), Last Updated By (col 10), Last Updated At (col 11)
  tasksSheet.getRange(taskRowIndex, 6).setValue(sanitizeCellValue(newStatus));
  tasksSheet.getRange(taskRowIndex, 10).setValue(sanitizeCellValue(currentUser.name));
  tasksSheet.getRange(taskRowIndex, 11).setValue(sanitizeCellValue(formattedNow));

  // Append history row
  var historyId = clientHistoryId || ('HIST-' + now.getTime().toString(36).toUpperCase());
  historySheet.appendRow([
    sanitizeCellValue(historyId),
    sanitizeCellValue(canonicalTaskId),
    sanitizeCellValue(taskName),
    sanitizeCellValue(newStatus),
    sanitizeCellValue(note || ''),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow)
  ]);

  var updatedRowValues = tasksSheet.getRange(taskRowIndex, 1, 1, 12).getValues()[0];
  var updatedTask = mapRowToTask(updatedRowValues);

  return {
    ok: true,
    success: true,
    task: updatedTask,
    historyEntry: {
      historyId: historyId,
      taskId: canonicalTaskId,
      taskName: taskName,
      status: newStatus,
      note: note || '',
      updatedBy: currentUser.name,
      timestamp: formattedNow
    }
  };
}

/**
 * Action: setState / edit
 * Authorization rule: ONLY THE OWNER (createdBy) can edit or change state of their tasks.
 */
function handleSetState(ss, payload, currentUser) {
  var taskId = String(payload.taskId || '').trim();
  var newState = String(payload.state || '').trim();
  var note = String(payload.note || '').trim();
  var clientHistoryId = payload.historyId ? String(payload.historyId).trim() : '';

  if (!taskId) {
    return { ok: false, success: false, message: 'Task ID is required.' };
  }
  if (CONFIG.VALID_STATES.indexOf(newState) === -1) {
    return { ok: false, success: false, message: 'Invalid state. Must be one of: ' + CONFIG.VALID_STATES.join(', ') };
  }

  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);

  var taskRowIndex = findRowIndexByColumnValue(tasksSheet, 1, taskId);
  if (taskRowIndex === -1) {
    return { ok: false, success: false, message: 'Task ID not found.' };
  }

  // Enforce server-side authorization: only owner can edit/change task state
  var createdBy = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 8).getValue() || '');
  if (createdBy.trim().toLowerCase() !== currentUser.name.trim().toLowerCase()) {
    return {
      ok: false,
      success: false,
      error: 'forbidden',
      message: 'Only the task owner (' + createdBy + ') can edit or change the state of this task.'
    };
  }

  var now = new Date();
  var formattedNow = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);

  var canonicalTaskId = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 1).getValue() || taskId);
  var taskName = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 2).getValue() || '');
  var currentStatus = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 6).getValue() || '');

  // Update State (col 7), Last Updated By (col 10), Last Updated At (col 11)
  tasksSheet.getRange(taskRowIndex, 7).setValue(sanitizeCellValue(newState));
  tasksSheet.getRange(taskRowIndex, 10).setValue(sanitizeCellValue(currentUser.name));
  tasksSheet.getRange(taskRowIndex, 11).setValue(sanitizeCellValue(formattedNow));

  var historyId = clientHistoryId || ('HIST-' + now.getTime().toString(36).toUpperCase());
  var logNote = 'State set to ' + newState + (note ? ': ' + note : '');

  historySheet.appendRow([
    sanitizeCellValue(historyId),
    sanitizeCellValue(canonicalTaskId),
    sanitizeCellValue(taskName),
    sanitizeCellValue(currentStatus),
    sanitizeCellValue(logNote),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow)
  ]);

  var updatedRowValues = tasksSheet.getRange(taskRowIndex, 1, 1, 12).getValues()[0];
  var updatedTask = mapRowToTask(updatedRowValues);

  return {
    ok: true,
    success: true,
    task: updatedTask,
    historyEntry: {
      historyId: historyId,
      taskId: canonicalTaskId,
      taskName: taskName,
      status: currentStatus,
      note: logNote,
      updatedBy: currentUser.name,
      timestamp: formattedNow
    }
  };
}

/**
 * Action: deleteTask / delete
 * Authorization rule: ONLY THE OWNER (createdBy) can delete their tasks.
 */
function handleDeleteTask(ss, payload, currentUser) {
  var taskId = String(payload.taskId || '').trim();
  if (!taskId) {
    return { ok: false, success: false, message: 'Task ID is required for deletion.' };
  }

  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);

  var taskRowIndex = findRowIndexByColumnValue(tasksSheet, 1, taskId);
  if (taskRowIndex === -1) {
    return { ok: false, success: false, message: 'Task ID not found.' };
  }

  // Enforce server-side authorization: only owner can delete
  var createdBy = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 8).getValue() || '');
  if (createdBy.trim().toLowerCase() !== currentUser.name.trim().toLowerCase()) {
    return {
      ok: false,
      success: false,
      error: 'forbidden',
      message: 'Only the task owner (' + createdBy + ') can delete this task.'
    };
  }

  // Delete the task row
  tasksSheet.deleteRow(taskRowIndex);

  // Prune history rows associated with this task
  if (historySheet) {
    var histData = historySheet.getDataRange().getValues();
    for (var i = histData.length - 1; i >= 1; i--) {
      var hTaskId = unescapeCellValue(histData[i][1]);
      if (hTaskId.toLowerCase() === taskId.toLowerCase()) {
        historySheet.deleteRow(i + 1);
      }
    }
  }

  return {
    ok: true,
    success: true,
    taskId: taskId
  };
}

/**
 * Action: setPrivacy / togglePrivacy
 * Authorization rule: ONLY THE CREATOR (createdBy) can change task privacy.
 */
function handleSetPrivacy(ss, payload, currentUser) {
  var taskId = String(payload.taskId || '').trim();
  var newIsPrivate = (payload.isPrivate === true || String(payload.isPrivate).trim().toLowerCase() === 'true' || payload.isPrivate === 1);
  var clientHistoryId = payload.historyId ? String(payload.historyId).trim() : '';

  if (!taskId) {
    return { ok: false, success: false, message: 'Task ID is required.' };
  }

  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  var historySheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);

  var taskRowIndex = findRowIndexByColumnValue(tasksSheet, 1, taskId);
  if (taskRowIndex === -1) {
    return { ok: false, success: false, message: 'Task ID not found.' };
  }

  // Authorization: Only the creator can change privacy
  var createdBy = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 8).getValue() || '');
  if (createdBy.trim().toLowerCase() !== currentUser.name.trim().toLowerCase()) {
    return {
      ok: false,
      success: false,
      error: 'forbidden',
      message: 'Only the task creator (' + createdBy + ') can change privacy.'
    };
  }

  var now = new Date();
  var formattedNow = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT);

  var canonicalTaskId = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 1).getValue() || taskId);
  var taskName = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 2).getValue() || '');
  var currentStatus = unescapeCellValue(tasksSheet.getRange(taskRowIndex, 6).getValue() || '');

  // Update IsPrivate (col 12), Last Updated By (col 10), Last Updated At (col 11)
  tasksSheet.getRange(taskRowIndex, 12).setValue(newIsPrivate ? true : false);
  tasksSheet.getRange(taskRowIndex, 10).setValue(sanitizeCellValue(currentUser.name));
  tasksSheet.getRange(taskRowIndex, 11).setValue(sanitizeCellValue(formattedNow));

  var historyId = clientHistoryId || ('HIST-' + now.getTime().toString(36).toUpperCase());
  var privacyLabel = newIsPrivate ? 'Private (Only Me)' : 'Public (Shared with team)';
  var logNote = 'Privacy changed to ' + privacyLabel;

  historySheet.appendRow([
    sanitizeCellValue(historyId),
    sanitizeCellValue(canonicalTaskId),
    sanitizeCellValue(taskName),
    sanitizeCellValue(currentStatus),
    sanitizeCellValue(logNote),
    sanitizeCellValue(currentUser.name),
    sanitizeCellValue(formattedNow)
  ]);

  var updatedRowValues = tasksSheet.getRange(taskRowIndex, 1, 1, 12).getValues()[0];
  var updatedTask = mapRowToTask(updatedRowValues);

  return {
    ok: true,
    success: true,
    task: updatedTask,
    historyEntry: {
      historyId: historyId,
      taskId: canonicalTaskId,
      taskName: taskName,
      status: currentStatus,
      note: logNote,
      updatedBy: currentUser.name,
      timestamp: formattedNow
    }
  };
}

/**
 * Action: addStatus
 */
function handleAddStatus(ss, payload, currentUser) {
  var status = String(payload.status || '').trim();
  if (!status || status.length > 50) {
    return { ok: false, success: false, message: 'Valid status name (1-50 chars) is required.' };
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    sheet.appendRow(['Status List']);
    formatHeaderRow(sheet, 1);
  }

  var existing = fetchStatuses(ss);
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].toLowerCase() === status.toLowerCase()) {
      return { ok: false, success: false, message: 'Status already exists.' };
    }
  }

  sheet.appendRow([sanitizeCellValue(status)]);
  return {
    ok: true,
    success: true,
    statuses: fetchStatuses(ss)
  };
}

/**
 * Action: editStatus
 */
function handleEditStatus(ss, payload, currentUser) {
  var oldStatus = String(payload.oldStatus || '').trim();
  var newStatus = String(payload.newStatus || '').trim();

  if (!oldStatus || !newStatus) {
    return { ok: false, success: false, message: 'Both current and new status names are required.' };
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    return { ok: false, success: false, message: 'Settings sheet not found.' };
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (unescapeCellValue(data[i][0]).toLowerCase() === oldStatus.toLowerCase()) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    return { ok: false, success: false, message: 'Status "' + oldStatus + '" not found.' };
  }

  sheet.getRange(foundRow, 1).setValue(sanitizeCellValue(newStatus));

  // Rename across existing tasks
  var tasksSheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  if (tasksSheet) {
    var taskData = tasksSheet.getDataRange().getValues();
    for (var t = 1; t < taskData.length; t++) {
      if (unescapeCellValue(taskData[t][5]).toLowerCase() === oldStatus.toLowerCase()) {
        tasksSheet.getRange(t + 1, 6).setValue(sanitizeCellValue(newStatus));
        tasksSheet.getRange(t + 1, 10).setValue(sanitizeCellValue(currentUser.name));
        tasksSheet.getRange(t + 1, 11).setValue(sanitizeCellValue(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATETIME_FORMAT)));
      }
    }
  }

  return {
    ok: true,
    success: true,
    statuses: fetchStatuses(ss)
  };
}

/**
 * Action: deleteStatus
 */
function handleDeleteStatus(ss, payload, currentUser) {
  var status = String(payload.status || '').trim();
  if (!status) {
    return { ok: false, success: false, message: 'Status name is required.' };
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    return { ok: false, success: false, message: 'Settings sheet not found.' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 2) {
    return { ok: false, success: false, message: 'Cannot delete the only remaining status.' };
  }

  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (unescapeCellValue(data[i][0]).toLowerCase() === status.toLowerCase()) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    return { ok: false, success: false, message: 'Status "' + status + '" not found.' };
  }

  sheet.deleteRow(foundRow);

  return {
    ok: true,
    success: true,
    statuses: fetchStatuses(ss)
  };
}

/**
 * Action: saveStatuses
 */
function handleSaveStatuses(ss, payload, currentUser) {
  var statuses = payload.statuses;
  if (!Array.isArray(statuses) || statuses.length === 0) {
    return { ok: false, success: false, message: 'At least one status is required.' };
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
  }

  // Ensure header in A1
  sheet.getRange(1, 1).setValue('Status List');
  formatHeaderRow(sheet, 1);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);

  // Clear existing status rows in column A (rows 2 and downwards) without wiping the sheet
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }

  var validStatuses = [];
  for (var i = 0; i < statuses.length; i++) {
    var s = String(statuses[i] || '').trim();
    if (s && validStatuses.indexOf(s) === -1) {
      validStatuses.push([sanitizeCellValue(s)]);
    }
  }

  if (validStatuses.length > 0) {
    sheet.getRange(2, 1, validStatuses.length, 1).setValues(validStatuses);
  }

  return {
    ok: true,
    success: true,
    statuses: fetchStatuses(ss)
  };
}

// ============================================================================
// DATA FETCHING & MAPPING HELPERS
// ============================================================================

function fetchTasks(ss) {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get('all_tasks_raw');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {}

  var sheet = ss.getSheetByName(CONFIG.SHEETS.TASKS);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var tasks = [];
  for (var i = 1; i < data.length; i++) {
    var task = mapRowToTask(data[i]);
    if (task.taskId) {
      tasks.push(task);
    }
  }

  try {
    var jsonStr = JSON.stringify(tasks);
    if (jsonStr.length < 95000) {
      cache.put('all_tasks_raw', jsonStr, CONFIG.CACHE_SNAPSHOT_SEC);
    }
  } catch (e) {}

  return tasks;
}

function mapRowToTask(row) {
  var rawPrivate = row[11];
  var isPrivate = (rawPrivate === true || String(rawPrivate).trim().toLowerCase() === 'true' || rawPrivate === 1 || String(rawPrivate).trim() === '1');
  return {
    taskId: unescapeCellValue(row[0]),
    taskName: unescapeCellValue(row[1]),
    description: unescapeCellValue(row[2]),
    assignedTo: unescapeCellValue(row[3]),
    deadline: formatCellDate(row[4], CONFIG.DATE_FORMAT),
    status: unescapeCellValue(row[5]),
    state: unescapeCellValue(row[6]) || 'Active',
    createdBy: unescapeCellValue(row[7]),
    createdAt: formatCellDate(row[8], CONFIG.DATETIME_FORMAT),
    lastUpdatedBy: unescapeCellValue(row[9]),
    lastUpdatedAt: formatCellDate(row[10], CONFIG.DATETIME_FORMAT),
    isPrivate: isPrivate
  };
}

function fetchHistory(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.HISTORY);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var history = [];
  var seen = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var taskId = unescapeCellValue(row[1]).trim();
    if (!taskId) continue;

    var historyId = unescapeCellValue(row[0]).trim() || ('HIST-ROW-' + (i + 1));
    var taskName = unescapeCellValue(row[2]).trim();
    var status = unescapeCellValue(row[3]).trim();
    var note = unescapeCellValue(row[4]).trim();
    var updatedBy = unescapeCellValue(row[5]).trim();
    var timestamp = formatCellDate(row[6], CONFIG.DATETIME_FORMAT);
    var epochSec = Math.floor(parseToEpoch(timestamp) / 1000);

    var dedupeKey = historyId + '|' + taskId.toLowerCase() + '|' + status.toLowerCase() + '|' + note.toLowerCase() + '|' + Math.floor(epochSec / 60);
    if (seen[dedupeKey]) continue;
    seen[dedupeKey] = true;

    history.push({
      historyId: historyId,
      taskId: taskId,
      taskName: taskName,
      status: status,
      note: note,
      updatedBy: updatedBy,
      timestamp: timestamp
    });
  }
  return history;
}

function fetchStatuses(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!sheet) return CONFIG.DEFAULT_STATUSES;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return CONFIG.DEFAULT_STATUSES;

  var statuses = [];
  for (var i = 1; i < data.length; i++) {
    var val = unescapeCellValue(data[i][0]).trim();
    if (val && statuses.indexOf(val) === -1) {
      statuses.push(val);
    }
  }

  return statuses.length > 0 ? statuses : CONFIG.DEFAULT_STATUSES;
}

/**
 * Returns active user list with PINs completely removed.
 */
function fetchActiveUsers(ss) {
  var users = getCachedUsers(ss);
  var active = [];
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (String(u.status || '').trim().toLowerCase() === 'active') {
      active.push({
        name: u.name,
        team: u.team,
        status: u.status
      });
    }
  }
  return active;
}

function findRowIndexByColumnValue(sheet, colIndex, targetValue) {
  var data = sheet.getDataRange().getValues();
  var target = String(targetValue || '').trim().toLowerCase();
  if (!target) return -1;
  for (var i = 1; i < data.length; i++) {
    if (unescapeCellValue(data[i][colIndex - 1]).toLowerCase() === target) {
      return i + 1;
    }
  }
  return -1;
}

function parseToEpoch(dateStr) {
  if (!dateStr) return 0;
  if (dateStr instanceof Date) return dateStr.getTime();
  var str = String(dateStr).trim();
  if (!str) return 0;
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(str)) {
    str = str.replace(' ', 'T');
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    str = str + 'T00:00:00';
  }
  var parsed = new Date(str).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

function formatCellDate(cellValue, pattern) {
  if (!cellValue) return '';
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, CONFIG.TIMEZONE, pattern);
  }
  return unescapeCellValue(cellValue).trim();
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
