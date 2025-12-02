/****************************************************
 * IDEMPOTENT OPERATIONS WRAPPER
 * Phase 6 - Medium Priority
 *
 * Makes operations safe to retry without creating duplicates
 * Tracks operation state to ensure idempotency
 *
 * Based on ADDITIONAL_ENHANCEMENTS.md Phase 3
 ****************************************************/

/**
 * Make a function idempotent using caching
 * Operations can be safely retried without side effects
 *
 * @param {Function} operation - The operation to make idempotent
 * @param {Function} keyGenerator - Function to generate unique key from arguments
 * @param {number} ttl - Cache time-to-live in seconds (default: 600 = 10 minutes)
 * @returns {Function} Idempotent version of the operation
 */
function makeIdempotent(operation, keyGenerator, ttl = 600) {
  return function(...args) {
    const key = keyGenerator(...args);
    const cache = CacheService.getScriptCache();

    const lockKey = `lock_${key}`;
    const resultKey = `result_${key}`;

    // Check if operation already in progress
    const inProgress = cache.get(lockKey);
    if (inProgress) {
      throw new Error(
        `Operation "${key}" is already in progress. ` +
        `Please wait for it to complete before retrying.`
      );
    }

    // Check if operation already completed
    const cachedResult = cache.get(resultKey);
    if (cachedResult) {
      Logger.log(`Returning cached result for operation: ${key}`);

      try {
        return JSON.parse(cachedResult);
      } catch (e) {
        // If result can't be parsed, re-execute
        Logger.log(`Failed to parse cached result, will re-execute`);
      }
    }

    try {
      // Set lock (5 minute max duration)
      cache.put(lockKey, 'true', 300);
      Logger.log(`Lock acquired for operation: ${key}`);

      // Execute operation
      const result = operation.apply(this, args);

      // Cache result
      try {
        cache.put(resultKey, JSON.stringify(result), ttl);
        Logger.log(`Result cached for operation: ${key} (TTL: ${ttl}s)`);
      } catch (e) {
        // Result might not be serializable, log but don't fail
        Logger.log(`Could not cache result: ${e.message}`);
      }

      return result;

    } finally {
      // Always remove lock
      cache.remove(lockKey);
      Logger.log(`Lock released for operation: ${key}`);
    }
  };
}

/**
 * Idempotent member addition
 * Safe to retry - won't create duplicates
 */
var addMemberIdempotent = makeIdempotent(
  function(memberData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      throw new Error('Member Directory sheet not found');
    }

    const memberID = memberData[0];

    // Check if member already exists
    const data = memberSheet.getDataRange().getValues();
    const existingRow = data.findIndex(row => row[0] === memberID);

    if (existingRow > 0) {
      // Update existing member
      memberSheet.getRange(existingRow + 1, 1, 1, memberData.length).setValues([memberData]);
      Logger.log(`Updated existing member: ${memberID}`);

      return {
        action: 'updated',
        memberID: memberID,
        row: existingRow + 1
      };
    } else {
      // Add new member
      memberSheet.appendRow(memberData);
      const newRow = memberSheet.getLastRow();
      Logger.log(`Added new member: ${memberID}`);

      return {
        action: 'added',
        memberID: memberID,
        row: newRow
      };
    }
  },
  (memberData) => `add_member_${memberData[0]}` // Use Member ID as key
);

/**
 * Idempotent grievance addition
 */
var addGrievanceIdempotent = makeIdempotent(
  function(grievanceData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet) {
      throw new Error('Grievance Log sheet not found');
    }

    const grievanceID = grievanceData[0];

    // Check if grievance already exists
    const data = grievanceSheet.getDataRange().getValues();
    const existingRow = data.findIndex(row => row[0] === grievanceID);

    if (existingRow > 0) {
      // Update existing grievance
      grievanceSheet.getRange(existingRow + 1, 1, 1, grievanceData.length).setValues([grievanceData]);
      Logger.log(`Updated existing grievance: ${grievanceID}`);

      return {
        action: 'updated',
        grievanceID: grievanceID,
        row: existingRow + 1
      };
    } else {
      // Add new grievance
      grievanceSheet.appendRow(grievanceData);
      const newRow = grievanceSheet.getLastRow();
      Logger.log(`Added new grievance: ${grievanceID}`);

      return {
        action: 'added',
        grievanceID: grievanceID,
        row: newRow
      };
    }
  },
  (grievanceData) => `add_grievance_${grievanceData[0]}` // Use Grievance ID as key
);

/**
 * Idempotent backup creation
 * Won't create duplicate backups if retried within TTL
 */
var createBackupIdempotent = makeIdempotent(
  function() {
    if (typeof createIncrementalBackup === 'function') {
      return createIncrementalBackup();
    }

    throw new Error('createIncrementalBackup function not available');
  },
  () => {
    // Key is based on current hour - allows one backup per hour
    const now = new Date();
    const hourKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd-HH');
    return `backup_${hourKey}`;
  },
  3600 // 1 hour TTL
);

/**
 * Idempotent recalculation wrapper
 * Prevents multiple simultaneous recalculations
 */
var recalcAllIdempotent = makeIdempotent(
  function() {
    // Run both member and grievance recalculations
    const results = {};

    if (typeof recalcAllMembers === 'function') {
      results.members = recalcAllMembers();
    }

    if (typeof recalcAllGrievancesBatched === 'function') {
      results.grievances = recalcAllGrievancesBatched();
    }

    return results;
  },
  () => {
    // Key is based on current 5-minute window
    const now = new Date();
    const minutes = Math.floor(now.getMinutes() / 5) * 5;
    const timeKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd-HH') + `-${minutes}`;
    return `recalc_all_${timeKey}`;
  },
  300 // 5 minute TTL
);

/**
 * Idempotent dashboard rebuild
 */
var rebuildDashboardIdempotent = makeIdempotent(
  function() {
    if (typeof rebuildDashboardOptimized === 'function') {
      return rebuildDashboardOptimized();
    } else if (typeof rebuildDashboard === 'function') {
      return rebuildDashboard();
    }

    throw new Error('No dashboard rebuild function available');
  },
  () => {
    // Key is based on current 5-minute window
    const now = new Date();
    const minutes = Math.floor(now.getMinutes() / 5) * 5;
    const timeKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd-HH') + `-${minutes}`;
    return `rebuild_dashboard_${timeKey}`;
  },
  300 // 5 minute TTL
);

/**
 * Clear all idempotent caches
 * Use when you need to force re-execution of operations
 */
function clearIdempotentCache() {
  const cache = CacheService.getScriptCache();

  // Remove all keys with common prefixes
  const prefixes = [
    'lock_',
    'result_',
    'add_member_',
    'add_grievance_',
    'backup_',
    'recalc_all_',
    'rebuild_dashboard_'
  ];

  // Note: CacheService doesn't have a way to list all keys
  // So we can only remove specific known patterns
  // Store a list of active keys in properties for tracking

  const props = PropertiesService.getScriptProperties();
  const activeKeys = props.getProperty('IDEMPOTENT_KEYS');

  if (activeKeys) {
    const keys = JSON.parse(activeKeys);

    for (const key of keys) {
      cache.remove(key);
      cache.remove(`lock_${key}`);
      cache.remove(`result_${key}`);
    }

    props.deleteProperty('IDEMPOTENT_KEYS');
    Logger.log(`Cleared ${keys.length} idempotent cache entries`);

  } else {
    Logger.log('No active idempotent cache entries found');
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Idempotent Cache Cleared',
    'All cached operation results have been cleared.\n\n' +
    'Operations will re-execute on next call.',
    ui.ButtonSet.OK
  );
}

/**
 * Show idempotent operation status
 */
function showIdempotentStatus() {
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const activeKeys = props.getProperty('IDEMPOTENT_KEYS');

  const ui = SpreadsheetApp.getUi();

  if (!activeKeys) {
    ui.alert(
      'Idempotent Operations',
      'No cached operations found.\n\n' +
      'All operations are fresh and can be executed.',
      ui.ButtonSet.OK
    );
    return;
  }

  const keys = JSON.parse(activeKeys);

  let message = `Active Idempotent Operations: ${keys.length}\n\n`;

  for (const key of keys) {
    const hasResult = cache.get(`result_${key}`) !== null;
    const hasLock = cache.get(`lock_${key}`) !== null;

    message += `• ${key}\n`;
    if (hasLock) {
      message += `  Status: In Progress\n`;
    } else if (hasResult) {
      message += `  Status: Completed (Cached)\n`;
    } else {
      message += `  Status: Unknown\n`;
    }
  }

  ui.alert('Idempotent Operations', message, ui.ButtonSet.OK);
}

/**
 * Track active idempotent keys
 * @param {string} key - Operation key to track
 */
function trackIdempotentKey(key) {
  const props = PropertiesService.getScriptProperties();
  const activeKeys = props.getProperty('IDEMPOTENT_KEYS');

  let keys = activeKeys ? JSON.parse(activeKeys) : [];

  if (!keys.includes(key)) {
    keys.push(key);
    props.setProperty('IDEMPOTENT_KEYS', JSON.stringify(keys));
  }
}

/**
 * Menu functions
 */
function CLEAR_IDEMPOTENT_CACHE() {
  clearIdempotentCache();
}

function SHOW_IDEMPOTENT_STATUS() {
  showIdempotentStatus();
}
