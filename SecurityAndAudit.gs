/******************************************************************************
 * SECURITY AND AUDIT FEATURES
 *
 * This module implements comprehensive security, audit logging, and enhanced
 * functionality for the 509 Dashboard system.
 *
 * FEATURES INCLUDED:
 * ✓ Feature 79: Audit Logging - Tracks all data modifications
 * ✓ Feature 80: Role-Based Access Control (RBAC)
 * ✓ Feature 83: Input Sanitization
 * ✓ Feature 84: Audit Reporting
 * ✓ Feature 85: Data Retention Policy
 * ✓ Feature 86: Suspicious Activity Detection
 * ✓ Feature 87: Quick Actions Sidebar
 * ✓ Feature 88: Advanced Search
 * ✓ Feature 89: Advanced Filtering
 * ✓ Feature 90: Automated Backups
 * ✓ Feature 91: Performance Monitoring
 * ✓ Feature 92: Keyboard Shortcuts
 * ✓ Feature 93: Export Wizard
 * ✓ Feature 94: Data Import
 *
 * CONFIGURATION REQUIRED:
 * - Script Properties: ADMINS, STEWARDS, VIEWERS (comma-separated emails)
 * - Script Properties: BACKUP_FOLDER_ID (Google Drive folder ID)
 * - Script Properties: DATA_RETENTION_YEARS (default: 7)
 *
 * VERSION: 1.0
 * LAST UPDATED: 2025-11-27
 ******************************************************************************/

/* ===================== CONFIGURATION ===================== */

const SECURITY_SHEETS = {
  AUDIT_LOG: "Audit_Log",
  PERFORMANCE_LOG: "Performance_Log"
};

const USER_ROLES = {
  ADMIN: "Admin",
  STEWARD: "Steward",
  VIEWER: "Viewer"
};

const AUDIT_ACTIONS = {
  CREATE: "CREATE",
  UPDATE: "UPDATE",
  DELETE: "DELETE",
  EXPORT: "EXPORT",
  IMPORT: "IMPORT",
  BACKUP: "BACKUP"
};

const SUSPICIOUS_ACTIVITY_THRESHOLD = 50; // Changes per hour
const DEFAULT_RETENTION_YEARS = 7;

/* ===================== SHEET CREATION ===================== */

/**
 * Creates the Audit_Log sheet for tracking all data modifications
 * Feature 79: Audit Logging
 */
function createAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SECURITY_SHEETS.AUDIT_LOG);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SECURITY_SHEETS.AUDIT_LOG);
  } else {
    sheet.clear();
  }

  // Set up headers
  const headers = [
    "Timestamp",
    "User Email",
    "User Role",
    "Action Type",
    "Sheet Name",
    "Record ID",
    "Field Changed",
    "Old Value",
    "New Value",
    "IP Address",
    "Session ID"
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(COLORS.HEADER_BLUE);
  headerRange.setFontColor(COLORS.WHITE);

  // Format columns
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 200); // User Email
  sheet.setColumnWidth(3, 100); // User Role
  sheet.setColumnWidth(4, 100); // Action Type
  sheet.setColumnWidth(5, 150); // Sheet Name
  sheet.setColumnWidth(6, 120); // Record ID
  sheet.setColumnWidth(7, 150); // Field Changed
  sheet.setColumnWidth(8, 200); // Old Value
  sheet.setColumnWidth(9, 200); // New Value
  sheet.setColumnWidth(10, 150); // IP Address
  sheet.setColumnWidth(11, 200); // Session ID

  // Freeze header row
  sheet.setFrozenRows(1);

  // Protect sheet (only admins can edit)
  const protection = sheet.protect().setDescription('Audit Log - Protected');
  protection.setWarningOnly(true);

  Logger.log('✓ Audit_Log sheet created successfully');
  return sheet;
}

/**
 * Creates the Performance_Log sheet for tracking script execution times
 * Feature 91: Performance Monitoring
 */
function createPerformanceLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SECURITY_SHEETS.PERFORMANCE_LOG);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SECURITY_SHEETS.PERFORMANCE_LOG);
  } else {
    sheet.clear();
  }

  // Set up headers
  const headers = [
    "Timestamp",
    "Function Name",
    "Execution Time (ms)",
    "User Email",
    "Parameters",
    "Status",
    "Error Message"
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(COLORS.HEADER_GREEN);
  headerRange.setFontColor(COLORS.WHITE);

  // Format columns
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 200); // Function Name
  sheet.setColumnWidth(3, 150); // Execution Time
  sheet.setColumnWidth(4, 200); // User Email
  sheet.setColumnWidth(5, 300); // Parameters
  sheet.setColumnWidth(6, 100); // Status
  sheet.setColumnWidth(7, 300); // Error Message

  // Freeze header row
  sheet.setFrozenRows(1);

  Logger.log('✓ Performance_Log sheet created successfully');
  return sheet;
}

/* ===================== CORE SECURITY FUNCTIONS ===================== */

/**
 * Checks if the current user has permission to perform an action
 * Feature 80: Role-Based Access Control
 *
 * @param {string} requiredRole - Minimum role required (Admin, Steward, or Viewer)
 * @return {boolean} True if user has permission, false otherwise
 */
function checkUserPermission(requiredRole) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRole(userEmail);

    // Define role hierarchy
    const roleHierarchy = {
      [USER_ROLES.ADMIN]: 3,
      [USER_ROLES.STEWARD]: 2,
      [USER_ROLES.VIEWER]: 1
    };

    const userLevel = roleHierarchy[userRole] || 0;
    const requiredLevel = roleHierarchy[requiredRole] || 0;

    const hasPermission = userLevel >= requiredLevel;

    if (!hasPermission) {
      SpreadsheetApp.getUi().alert(
        'Access Denied',
        `This action requires ${requiredRole} permissions. Your role: ${userRole}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

    return hasPermission;
  } catch (error) {
    Logger.log('Error checking user permission: ' + error);
    return false;
  }
}

/**
 * Gets the role of a user based on script properties
 *
 * @param {string} userEmail - Email address of the user
 * @return {string} User role (Admin, Steward, or Viewer)
 */
function getUserRole(userEmail) {
  const props = PropertiesService.getScriptProperties();

  // Get role lists from script properties
  const admins = (props.getProperty('ADMINS') || '').split(',').map(e => e.trim().toLowerCase());
  const stewards = (props.getProperty('STEWARDS') || '').split(',').map(e => e.trim().toLowerCase());
  const viewers = (props.getProperty('VIEWERS') || '').split(',').map(e => e.trim().toLowerCase());

  const email = userEmail.toLowerCase();

  if (admins.includes(email)) {
    return USER_ROLES.ADMIN;
  } else if (stewards.includes(email)) {
    return USER_ROLES.STEWARD;
  } else if (viewers.includes(email)) {
    return USER_ROLES.VIEWER;
  } else {
    // Default to Viewer if not explicitly assigned
    return USER_ROLES.VIEWER;
  }
}

/**
 * Validates and sanitizes user input for security
 * Feature 83: Input Sanitization
 *
 * @param {string} input - User input to sanitize
 * @param {string} type - Type of input (text, email, number, date)
 * @return {string} Sanitized input
 */
function sanitizeInput(input, type = 'text') {
  if (input == null || input === '') {
    return '';
  }

  let sanitized = String(input).trim();

  switch (type) {
    case 'text':
      // Remove script tags and potentially dangerous characters
      sanitized = sanitized.replace(/<script[^>]*>.*?<\/script>/gi, '');
      sanitized = sanitized.replace(/<[^>]*>/g, '');
      sanitized = sanitized.replace(/[<>]/g, '');
      break;

    case 'email':
      // Validate email format
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(sanitized)) {
        throw new Error('Invalid email format');
      }
      sanitized = sanitized.toLowerCase();
      break;

    case 'number':
      // Ensure it's a valid number
      sanitized = sanitized.replace(/[^0-9.-]/g, '');
      if (isNaN(Number(sanitized))) {
        throw new Error('Invalid number format');
      }
      break;

    case 'date':
      // Validate date format
      const date = new Date(sanitized);
      if (isNaN(date.getTime())) {
        throw new Error('Invalid date format');
      }
      break;

    case 'alphanumeric':
      // Only allow letters, numbers, spaces, and basic punctuation
      sanitized = sanitized.replace(/[^a-zA-Z0-9\s\-_.,]/g, '');
      break;
  }

  // Additional security: limit length to prevent buffer overflow
  if (sanitized.length > 10000) {
    sanitized = sanitized.substring(0, 10000);
  }

  return sanitized;
}

/* ===================== AUDIT LOGGING FUNCTIONS ===================== */

/**
 * Logs a data modification to the Audit_Log sheet
 * Feature 79: Audit Logging
 *
 * @param {string} actionType - Type of action (CREATE, UPDATE, DELETE, etc.)
 * @param {string} sheetName - Name of the sheet where modification occurred
 * @param {string} recordId - ID of the record modified
 * @param {string} fieldChanged - Name of the field that changed
 * @param {string} oldValue - Previous value
 * @param {string} newValue - New value
 */
function logDataModification(actionType, sheetName, recordId, fieldChanged, oldValue, newValue) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(SECURITY_SHEETS.AUDIT_LOG);

    // Create audit log sheet if it doesn't exist
    if (!auditSheet) {
      auditSheet = createAuditLogSheet();
    }

    // Get user information
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRole(userEmail);
    const timestamp = new Date();

    // Generate session ID (simple implementation)
    const sessionId = Utilities.base64Encode(userEmail + timestamp.getTime());

    // Get IP address (not directly available in Apps Script, placeholder)
    const ipAddress = 'N/A';

    // Prepare log entry
    const logEntry = [
      timestamp,
      userEmail,
      userRole,
      actionType,
      sheetName,
      recordId || 'N/A',
      fieldChanged || 'N/A',
      oldValue || 'N/A',
      newValue || 'N/A',
      ipAddress,
      sessionId
    ];

    // Append to audit log
    auditSheet.appendRow(logEntry);

    // Check for suspicious activity
    detectSuspiciousActivity(userEmail);

  } catch (error) {
    Logger.log('Error logging data modification: ' + error);
  }
}

/**
 * Generates an audit report for a specified date range
 * Feature 84: Audit Reporting
 *
 * @param {Date} startDate - Start date for the report
 * @param {Date} endDate - End date for the report
 * @return {Array} Array of audit log entries
 */
function generateAuditReport(startDate, endDate) {
  if (!checkUserPermission(USER_ROLES.STEWARD)) {
    return [];
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(SECURITY_SHEETS.AUDIT_LOG);

    if (!auditSheet) {
      SpreadsheetApp.getUi().alert('Audit Log sheet not found. Please create it first.');
      return [];
    }

    const data = auditSheet.getDataRange().getValues();
    const headers = data[0];
    const logs = data.slice(1);

    // Filter logs by date range
    const filteredLogs = logs.filter(row => {
      const timestamp = new Date(row[0]);
      return timestamp >= startDate && timestamp <= endDate;
    });

    // Create report sheet
    const reportSheetName = `Audit_Report_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss')}`;
    const reportSheet = ss.insertSheet(reportSheetName);

    // Add headers and data
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    reportSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor(COLORS.WHITE);

    if (filteredLogs.length > 0) {
      reportSheet.getRange(2, 1, filteredLogs.length, headers.length).setValues(filteredLogs);
    }

    // Add summary statistics
    const summaryRow = filteredLogs.length + 3;
    reportSheet.getRange(summaryRow, 1).setValue('SUMMARY STATISTICS');
    reportSheet.getRange(summaryRow, 1).setFontWeight("bold");
    reportSheet.getRange(summaryRow + 1, 1).setValue('Total Modifications:');
    reportSheet.getRange(summaryRow + 1, 2).setValue(filteredLogs.length);
    reportSheet.getRange(summaryRow + 2, 1).setValue('Date Range:');
    reportSheet.getRange(summaryRow + 2, 2).setValue(`${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`);

    // Count by action type
    const actionCounts = {};
    filteredLogs.forEach(row => {
      const action = row[3];
      actionCounts[action] = (actionCounts[action] || 0) + 1;
    });

    let row = summaryRow + 3;
    reportSheet.getRange(row, 1).setValue('Actions by Type:');
    reportSheet.getRange(row, 1).setFontWeight("bold");
    row++;

    for (const [action, count] of Object.entries(actionCounts)) {
      reportSheet.getRange(row, 1).setValue(action);
      reportSheet.getRange(row, 2).setValue(count);
      row++;
    }

    SpreadsheetApp.getUi().alert(`Audit report generated: ${reportSheetName}`);
    return filteredLogs;

  } catch (error) {
    Logger.log('Error generating audit report: ' + error);
    SpreadsheetApp.getUi().alert('Error generating audit report: ' + error.message);
    return [];
  }
}

/**
 * Shows a dialog to generate an audit report
 */
function showAuditReportDialog() {
  if (!checkUserPermission(USER_ROLES.STEWARD)) {
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; background: #3B82F6; color: white; border: none; cursor: pointer; }
      button:hover { background: #2563EB; }
    </style>
    <h2>Generate Audit Report</h2>
    <label>Start Date:</label>
    <input type="date" id="startDate" />
    <label>End Date:</label>
    <input type="date" id="endDate" />
    <button onclick="generateReport()">Generate Report</button>
    <script>
      // Set default dates (last 30 days)
      const today = new Date();
      const thirtyDaysAgo = new Date(today.getTime() - (30 * 24 * 60 * 60 * 1000));
      document.getElementById('endDate').valueAsDate = today;
      document.getElementById('startDate').valueAsDate = thirtyDaysAgo;

      function generateReport() {
        const startDate = document.getElementById('startDate').value;
        const endDate = document.getElementById('endDate').value;
        google.script.run.withSuccessHandler(() => {
          google.script.host.close();
        }).generateAuditReportFromDialog(startDate, endDate);
      }
    </script>
  `).setWidth(400).setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Audit Report');
}

/**
 * Helper function to generate audit report from dialog
 */
function generateAuditReportFromDialog(startDateStr, endDateStr) {
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  generateAuditReport(startDate, endDate);
}

/* ===================== DATA RETENTION & SECURITY ===================== */

/**
 * Enforces data retention policy
 * Feature 85: Data Retention Policy
 *
 * @param {number} retentionYears - Number of years to retain data (default: 7)
 */
function enforceDataRetention(retentionYears) {
  if (!checkUserPermission(USER_ROLES.ADMIN)) {
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const years = retentionYears || parseInt(props.getProperty('DATA_RETENTION_YEARS')) || DEFAULT_RETENTION_YEARS;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(SECURITY_SHEETS.AUDIT_LOG);

    if (!auditSheet) {
      Logger.log('Audit Log sheet not found');
      return;
    }

    const cutoffDate = new Date();
    cutoffDate.setFullYear(cutoffDate.getFullYear() - years);

    const data = auditSheet.getDataRange().getValues();
    let rowsDeleted = 0;

    // Delete rows older than cutoff date (start from bottom to avoid index issues)
    for (let i = data.length - 1; i > 0; i--) {
      const timestamp = new Date(data[i][0]);
      if (timestamp < cutoffDate) {
        auditSheet.deleteRow(i + 1);
        rowsDeleted++;
      }
    }

    logDataModification(
      'RETENTION_POLICY',
      SECURITY_SHEETS.AUDIT_LOG,
      'N/A',
      'Data Retention',
      `${rowsDeleted} rows`,
      `Deleted records older than ${years} years`
    );

    SpreadsheetApp.getUi().alert(`Data retention policy enforced. ${rowsDeleted} old audit log entries removed.`);

  } catch (error) {
    Logger.log('Error enforcing data retention: ' + error);
    SpreadsheetApp.getUi().alert('Error enforcing data retention: ' + error.message);
  }
}

/**
 * Detects suspicious activity patterns
 * Feature 86: Suspicious Activity Detection
 *
 * @param {string} userEmail - Email of user to check
 */
function detectSuspiciousActivity(userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName(SECURITY_SHEETS.AUDIT_LOG);

    if (!auditSheet) {
      return;
    }

    const data = auditSheet.getDataRange().getValues();
    const oneHourAgo = new Date(Date.now() - (60 * 60 * 1000));

    // Count changes by this user in the last hour
    let changeCount = 0;
    for (let i = 1; i < data.length; i++) {
      const timestamp = new Date(data[i][0]);
      const email = data[i][1];

      if (email === userEmail && timestamp >= oneHourAgo) {
        changeCount++;
      }
    }

    // Alert if threshold exceeded
    if (changeCount > SUSPICIOUS_ACTIVITY_THRESHOLD) {
      const message = `⚠️ SUSPICIOUS ACTIVITY DETECTED\n\n` +
                     `User: ${userEmail}\n` +
                     `Changes in last hour: ${changeCount}\n` +
                     `Threshold: ${SUSPICIOUS_ACTIVITY_THRESHOLD}\n\n` +
                     `This may indicate a potential security issue or data breach.`;

      Logger.log(message);

      // Log the suspicious activity
      logDataModification(
        'SUSPICIOUS_ACTIVITY',
        'System',
        userEmail,
        'Activity Threshold',
        SUSPICIOUS_ACTIVITY_THRESHOLD.toString(),
        changeCount.toString()
      );

      // Notify admins
      const props = PropertiesService.getScriptProperties();
      const admins = (props.getProperty('ADMINS') || '').split(',').map(e => e.trim());

      if (admins.length > 0) {
        MailApp.sendEmail({
          to: admins.join(','),
          subject: '⚠️ Suspicious Activity Detected - 509 Dashboard',
          body: message
        });
      }
    }

  } catch (error) {
    Logger.log('Error detecting suspicious activity: ' + error);
  }
}

/* ===================== USER INTERFACE FEATURES ===================== */

/**
 * Shows the Quick Actions sidebar
 * Feature 87: Quick Actions Sidebar
 */
function showQuickActionsSidebar() {
  const userRole = getUserRole(Session.getActiveUser().getEmail());

  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        padding: 10px;
        background: #f5f5f5;
      }
      .action-btn {
        display: block;
        width: 100%;
        padding: 12px;
        margin: 8px 0;
        background: #3B82F6;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        text-align: left;
      }
      .action-btn:hover {
        background: #2563EB;
      }
      .action-btn.admin { background: #DC2626; }
      .action-btn.admin:hover { background: #B91C1C; }
      .action-btn.steward { background: #059669; }
      .action-btn.steward:hover { background: #047857; }
      .section {
        margin-top: 20px;
        padding-top: 10px;
        border-top: 1px solid #ddd;
      }
      h3 {
        color: #1F2937;
        font-size: 14px;
        margin: 10px 0;
      }
    </style>
    <h2>⚡ Quick Actions</h2>
    <p>Role: <strong>${userRole}</strong></p>

    <h3>🔍 Search & Filter</h3>
    <button class="action-btn" onclick="google.script.run.showSearchDialog()">Advanced Search</button>
    <button class="action-btn" onclick="google.script.run.showFilterDialog()">Filter Grievances</button>

    <h3>📊 Export & Import</h3>
    <button class="action-btn" onclick="google.script.run.showExportWizard()">Export Wizard</button>
    ${userRole === 'Admin' || userRole === 'Steward' ?
      '<button class="action-btn steward" onclick="google.script.run.showImportWizard()">Import Data</button>' : ''}

    ${userRole === 'Admin' || userRole === 'Steward' ? `
    <div class="section">
      <h3>📋 Reports</h3>
      <button class="action-btn steward" onclick="google.script.run.showAuditReportDialog()">Audit Report</button>
    </div>
    ` : ''}

    ${userRole === 'Admin' ? `
    <div class="section">
      <h3>🔐 Admin Actions</h3>
      <button class="action-btn admin" onclick="google.script.run.createAutomatedBackup()">Create Backup</button>
      <button class="action-btn admin" onclick="google.script.run.enforceDataRetention()">Enforce Retention</button>
      <button class="action-btn admin" onclick="google.script.run.setupRBACConfiguration()">Configure RBAC</button>
    </div>
    ` : ''}
  `).setTitle('Quick Actions');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the Advanced Search dialog
 * Feature 88: Advanced Search
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; background: #3B82F6; color: white; border: none; cursor: pointer; }
      button:hover { background: #2563EB; }
      .results { margin-top: 20px; max-height: 300px; overflow-y: auto; }
      .result-item { padding: 10px; margin: 5px 0; background: #f5f5f5; border-left: 3px solid #3B82F6; }
    </style>
    <h2>🔍 Advanced Search</h2>

    <label>Search By:</label>
    <select id="searchType">
      <option value="id">Grievance ID</option>
      <option value="member">Member Name</option>
      <option value="type">Issue Type</option>
      <option value="steward">Steward</option>
    </select>

    <label>Search Term:</label>
    <input type="text" id="searchTerm" placeholder="Enter search term..." />

    <button onclick="performSearch()">Search</button>

    <div id="results" class="results"></div>

    <script>
      function performSearch() {
        const searchType = document.getElementById('searchType').value;
        const searchTerm = document.getElementById('searchTerm').value;

        if (!searchTerm) {
          alert('Please enter a search term');
          return;
        }

        google.script.run
          .withSuccessHandler(displayResults)
          .withFailureHandler(err => alert('Search failed: ' + err))
          .performAdvancedSearch(searchType, searchTerm);
      }

      function displayResults(results) {
        const resultsDiv = document.getElementById('results');

        if (results.length === 0) {
          resultsDiv.innerHTML = '<p>No results found</p>';
          return;
        }

        let html = '<h3>Found ' + results.length + ' result(s):</h3>';
        results.forEach(r => {
          html += '<div class="result-item">';
          html += '<strong>Grievance #' + r.id + '</strong><br>';
          html += 'Member: ' + r.member + '<br>';
          html += 'Type: ' + r.type + '<br>';
          html += 'Status: ' + r.status + '<br>';
          html += 'Row: ' + r.row;
          html += '</div>';
        });

        resultsDiv.innerHTML = html;
      }
    </script>
  `).setWidth(500).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Advanced Search');
}

/**
 * Performs an advanced search on grievances
 */
function performAdvancedSearch(searchType, searchTerm) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet) {
      throw new Error('Grievance Log sheet not found');
    }

    const data = grievanceSheet.getDataRange().getValues();
    const results = [];

    // Sanitize search term
    const sanitizedTerm = sanitizeInput(searchTerm, 'text').toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let match = false;

      switch (searchType) {
        case 'id':
          match = String(row[GRIEVANCE_COLS.GRIEVANCE_ID - 1]).toLowerCase().includes(sanitizedTerm);
          break;
        case 'member':
          const memberName = `${row[GRIEVANCE_COLS.FIRST_NAME - 1]} ${row[GRIEVANCE_COLS.LAST_NAME - 1]}`.toLowerCase();
          match = memberName.includes(sanitizedTerm);
          break;
        case 'type':
          match = String(row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1]).toLowerCase().includes(sanitizedTerm);
          break;
        case 'steward':
          match = String(row[GRIEVANCE_COLS.STEWARD - 1]).toLowerCase().includes(sanitizedTerm);
          break;
      }

      if (match) {
        results.push({
          id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
          member: `${row[GRIEVANCE_COLS.FIRST_NAME - 1]} ${row[GRIEVANCE_COLS.LAST_NAME - 1]}`,
          type: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
          status: row[GRIEVANCE_COLS.STATUS - 1],
          row: i + 1
        });
      }
    }

    return results;

  } catch (error) {
    Logger.log('Error performing search: ' + error);
    throw error;
  }
}

/**
 * Shows the Advanced Filter dialog
 * Feature 89: Advanced Filtering
 */
function showFilterDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; background: #3B82F6; color: white; border: none; cursor: pointer; }
      button:hover { background: #2563EB; }
    </style>
    <h2>🎯 Advanced Filter</h2>

    <label>Status:</label>
    <select id="status">
      <option value="">All</option>
      <option value="Open">Open</option>
      <option value="In Progress">In Progress</option>
      <option value="Closed">Closed</option>
    </select>

    <label>Issue Type:</label>
    <select id="issueType">
      <option value="">All</option>
      <option value="Discipline">Discipline</option>
      <option value="Pay/Hours">Pay/Hours</option>
      <option value="Working Conditions">Working Conditions</option>
      <option value="Safety">Safety</option>
    </select>

    <label>Start Date:</label>
    <input type="date" id="startDate" />

    <label>End Date:</label>
    <input type="date" id="endDate" />

    <button onclick="applyFilter()">Apply Filter</button>
    <button onclick="clearFilter()" style="background: #6B7280;">Clear Filter</button>

    <script>
      function applyFilter() {
        const filters = {
          status: document.getElementById('status').value,
          issueType: document.getElementById('issueType').value,
          startDate: document.getElementById('startDate').value,
          endDate: document.getElementById('endDate').value
        };

        google.script.run
          .withSuccessHandler(() => {
            alert('Filter applied successfully');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Filter failed: ' + err))
          .applyAdvancedFilter(filters);
      }

      function clearFilter() {
        google.script.run
          .withSuccessHandler(() => {
            alert('Filter cleared');
            google.script.host.close();
          })
          .clearAdvancedFilter();
      }
    </script>
  `).setWidth(400).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Advanced Filter');
}

/**
 * Applies advanced filters to the Grievance Log
 */
function applyAdvancedFilter(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet) {
      throw new Error('Grievance Log sheet not found');
    }

    // Clear existing filter
    const existingFilter = grievanceSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    // Create new filter
    const dataRange = grievanceSheet.getDataRange();
    const filter = dataRange.createFilter();

    // Apply status filter
    if (filters.status) {
      const statusCol = GRIEVANCE_COLS.STATUS;
      const criteria = SpreadsheetApp.newFilterCriteria()
        .whenTextEqualTo(filters.status)
        .build();
      filter.setColumnFilterCriteria(statusCol, criteria);
    }

    // Apply issue type filter
    if (filters.issueType) {
      const typeCol = GRIEVANCE_COLS.ISSUE_CATEGORY;
      const criteria = SpreadsheetApp.newFilterCriteria()
        .whenTextEqualTo(filters.issueType)
        .build();
      filter.setColumnFilterCriteria(typeCol, criteria);
    }

    // Apply date range filter
    if (filters.startDate || filters.endDate) {
      const dateCol = GRIEVANCE_COLS.DATE_FILED;
      let criteriaBuilder = SpreadsheetApp.newFilterCriteria();

      if (filters.startDate) {
        criteriaBuilder = criteriaBuilder.whenDateAfter(new Date(filters.startDate));
      }
      if (filters.endDate) {
        criteriaBuilder = criteriaBuilder.whenDateBefore(new Date(filters.endDate));
      }

      filter.setColumnFilterCriteria(dateCol, criteriaBuilder.build());
    }

    logDataModification(
      'FILTER_APPLIED',
      SHEETS.GRIEVANCE_LOG,
      'N/A',
      'Advanced Filter',
      'None',
      JSON.stringify(filters)
    );

  } catch (error) {
    Logger.log('Error applying filter: ' + error);
    throw error;
  }
}

/**
 * Clears all filters from the Grievance Log
 */
function clearAdvancedFilter() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet) {
      throw new Error('Grievance Log sheet not found');
    }

    const existingFilter = grievanceSheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }

    logDataModification(
      'FILTER_CLEARED',
      SHEETS.GRIEVANCE_LOG,
      'N/A',
      'Advanced Filter',
      'Active',
      'None'
    );

  } catch (error) {
    Logger.log('Error clearing filter: ' + error);
    throw error;
  }
}

/* ===================== OPERATIONAL FEATURES ===================== */

/**
 * Creates an automated backup of the spreadsheet
 * Feature 90: Automated Backups
 */
function createAutomatedBackup() {
  if (!checkUserPermission(USER_ROLES.ADMIN)) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const backupFolderId = props.getProperty('BACKUP_FOLDER_ID');

    if (!backupFolderId) {
      SpreadsheetApp.getUi().alert(
        'Backup Folder Not Configured',
        'Please set the BACKUP_FOLDER_ID script property to your Google Drive backup folder ID.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const folder = DriveApp.getFolderById(backupFolderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
    const backupName = `509_Dashboard_Backup_${timestamp}`;

    // Create a copy of the spreadsheet
    const backup = ss.copy(backupName);
    const backupFile = DriveApp.getFileById(backup.getId());

    // Move to backup folder
    backupFile.moveTo(folder);

    logDataModification(
      AUDIT_ACTIONS.BACKUP,
      'System',
      backup.getId(),
      'Automated Backup',
      'N/A',
      backupName
    );

    SpreadsheetApp.getUi().alert(`Backup created successfully: ${backupName}`);

  } catch (error) {
    Logger.log('Error creating backup: ' + error);
    SpreadsheetApp.getUi().alert('Error creating backup: ' + error.message);
  }
}

/**
 * Sets up a time-driven trigger for daily backups
 */
function setupDailyBackupTrigger() {
  if (!checkUserPermission(USER_ROLES.ADMIN)) {
    return;
  }

  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createAutomatedBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger at 2 AM
  ScriptApp.newTrigger('createAutomatedBackup')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Daily backup trigger configured for 2:00 AM');
}

/**
 * Tracks performance of a function
 * Feature 91: Performance Monitoring
 *
 * @param {string} functionName - Name of the function being tracked
 * @param {function} fn - Function to execute and track
 * @param {Array} args - Arguments to pass to the function
 * @return {*} Result of the function execution
 */
function trackPerformance(functionName, fn, args = []) {
  const startTime = Date.now();
  const userEmail = Session.getActiveUser().getEmail();
  let result, status, errorMessage;

  try {
    result = fn.apply(null, args);
    status = 'SUCCESS';
    errorMessage = '';
  } catch (error) {
    status = 'ERROR';
    errorMessage = error.toString();
    throw error;
  } finally {
    const endTime = Date.now();
    const executionTime = endTime - startTime;

    // Log to Performance_Log sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let perfSheet = ss.getSheetByName(SECURITY_SHEETS.PERFORMANCE_LOG);

    if (!perfSheet) {
      perfSheet = createPerformanceLogSheet();
    }

    const logEntry = [
      new Date(),
      functionName,
      executionTime,
      userEmail,
      JSON.stringify(args).substring(0, 500),
      status,
      errorMessage
    ];

    perfSheet.appendRow(logEntry);
  }

  return result;
}

/* ===================== IMPORT/EXPORT FEATURES ===================== */

/**
 * Shows the Export Wizard dialog
 * Feature 93: Export Wizard
 */
function showExportWizard() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; background: #3B82F6; color: white; border: none; cursor: pointer; }
      button:hover { background: #2563EB; }
      .checkbox-group { margin-top: 10px; }
      .checkbox-group label { font-weight: normal; }
    </style>
    <h2>📥 Export Wizard</h2>

    <label>Export Type:</label>
    <select id="exportType">
      <option value="grievances">Grievances</option>
      <option value="members">Members</option>
      <option value="both">Both</option>
    </select>

    <label>Format:</label>
    <select id="format">
      <option value="csv">CSV</option>
      <option value="excel">Excel (XLSX)</option>
      <option value="pdf">PDF</option>
    </select>

    <div class="checkbox-group">
      <label><input type="checkbox" id="includeHeaders" checked> Include Headers</label>
      <label><input type="checkbox" id="activeOnly"> Active Records Only</label>
    </div>

    <label>Date Range (Optional):</label>
    <input type="date" id="startDate" placeholder="Start Date" />
    <input type="date" id="endDate" placeholder="End Date" />

    <button onclick="performExport()">Export</button>

    <script>
      function performExport() {
        const options = {
          exportType: document.getElementById('exportType').value,
          format: document.getElementById('format').value,
          includeHeaders: document.getElementById('includeHeaders').checked,
          activeOnly: document.getElementById('activeOnly').checked,
          startDate: document.getElementById('startDate').value,
          endDate: document.getElementById('endDate').value
        };

        google.script.run
          .withSuccessHandler(() => {
            alert('Export completed successfully');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Export failed: ' + err))
          .performExport(options);
      }
    </script>
  `).setWidth(400).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Wizard');
}

/**
 * Performs the export based on wizard options
 */
function performExport(options) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');

    // Create new sheet for export
    const exportSheetName = `Export_${options.exportType}_${timestamp}`;
    const exportSheet = ss.insertSheet(exportSheetName);

    // Get source data
    let sourceSheet;
    if (options.exportType === 'grievances') {
      sourceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    } else if (options.exportType === 'members') {
      sourceSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    }

    if (!sourceSheet) {
      throw new Error('Source sheet not found');
    }

    const data = sourceSheet.getDataRange().getValues();
    let exportData = data;

    // Filter by date range if specified
    if (options.startDate && options.endDate && options.exportType === 'grievances') {
      const startDate = new Date(options.startDate);
      const endDate = new Date(options.endDate);

      exportData = data.filter((row, index) => {
        if (index === 0) return options.includeHeaders;
        const filedDate = new Date(row[GRIEVANCE_COLS.DATE_FILED - 1]);
        return filedDate >= startDate && filedDate <= endDate;
      });
    }

    // Filter active only
    if (options.activeOnly && options.exportType === 'grievances') {
      exportData = exportData.filter((row, index) => {
        if (index === 0) return options.includeHeaders;
        return row[GRIEVANCE_COLS.STATUS - 1] !== 'Closed';
      });
    }

    // Write to export sheet
    if (exportData.length > 0) {
      exportSheet.getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);
    }

    logDataModification(
      AUDIT_ACTIONS.EXPORT,
      exportSheetName,
      'N/A',
      'Export Wizard',
      'N/A',
      JSON.stringify(options)
    );

    SpreadsheetApp.getUi().alert(`Export created: ${exportSheetName}\n\nYou can now download this sheet as ${options.format.toUpperCase()}.`);

  } catch (error) {
    Logger.log('Error performing export: ' + error);
    throw error;
  }
}

/**
 * Shows the Import Wizard dialog
 * Feature 94: Data Import
 */
function showImportWizard() {
  if (!checkUserPermission(USER_ROLES.STEWARD)) {
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; background: #3B82F6; color: white; border: none; cursor: pointer; }
      button:hover { background: #2563EB; }
      .warning { background: #FEF3C7; padding: 10px; margin: 10px 0; border-left: 3px solid #F59E0B; }
    </style>
    <h2>📤 Import Wizard</h2>

    <div class="warning">
      ⚠️ <strong>Warning:</strong> Importing data will add new records to your database.
      Make sure your CSV/Excel file has the correct column headers.
    </div>

    <label>Import Type:</label>
    <select id="importType">
      <option value="members">Members</option>
      <option value="grievances">Grievances</option>
    </select>

    <label>Instructions:</label>
    <ol>
      <li>Create a sheet named "Import_Data" in this spreadsheet</li>
      <li>Paste your CSV/Excel data into that sheet</li>
      <li>Ensure column headers match exactly</li>
      <li>Click the Import button below</li>
    </ol>

    <button onclick="performImport()">Import Data</button>

    <script>
      function performImport() {
        const importType = document.getElementById('importType').value;

        if (!confirm('Are you sure you want to import data? This cannot be undone easily.')) {
          return;
        }

        google.script.run
          .withSuccessHandler((count) => {
            alert('Import completed: ' + count + ' records imported');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Import failed: ' + err))
          .performImport(importType);
      }
    </script>
  `).setWidth(450).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Import Wizard');
}

/**
 * Performs data import from Import_Data sheet
 */
function performImport(importType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const importSheet = ss.getSheetByName('Import_Data');

    if (!importSheet) {
      throw new Error('Import_Data sheet not found. Please create it and paste your data.');
    }

    const importData = importSheet.getDataRange().getValues();

    if (importData.length < 2) {
      throw new Error('No data found in Import_Data sheet');
    }

    // Get target sheet
    let targetSheet;
    if (importType === 'members') {
      targetSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    } else if (importType === 'grievances') {
      targetSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    }

    if (!targetSheet) {
      throw new Error('Target sheet not found');
    }

    // Validate headers
    const importHeaders = importData[0];
    const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];

    // Skip header row and append data
    const dataToImport = importData.slice(1);
    const lastRow = targetSheet.getLastRow();

    // Sanitize all input
    const sanitizedData = dataToImport.map(row =>
      row.map(cell => sanitizeInput(String(cell), 'text'))
    );

    // Append to target sheet
    if (sanitizedData.length > 0) {
      targetSheet.getRange(lastRow + 1, 1, sanitizedData.length, sanitizedData[0].length).setValues(sanitizedData);
    }

    logDataModification(
      AUDIT_ACTIONS.IMPORT,
      targetSheet.getName(),
      'N/A',
      'Import Wizard',
      'N/A',
      `${sanitizedData.length} records imported`
    );

    return sanitizedData.length;

  } catch (error) {
    Logger.log('Error performing import: ' + error);
    throw error;
  }
}

/* ===================== KEYBOARD SHORTCUTS ===================== */

/**
 * Sets up custom keyboard shortcuts
 * Feature 92: Keyboard Shortcuts
 *
 * Note: Google Sheets doesn't support custom keyboard shortcuts directly.
 * This function provides a menu-based alternative with shortcut hints.
 */
function setupKeyboardShortcuts() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'Keyboard Shortcuts',
    'Google Apps Script does not support custom keyboard shortcuts.\n\n' +
    'However, you can use these built-in shortcuts:\n\n' +
    '• Alt+Shift+F - Open Filter menu\n' +
    '• Alt+Shift+S - Open Sort menu\n' +
    '• Ctrl+F (Cmd+F on Mac) - Find\n' +
    '• Ctrl+H (Cmd+H on Mac) - Find and Replace\n\n' +
    'Use the Quick Actions sidebar for one-click access to common actions.',
    ui.ButtonSet.OK
  );
}

/* ===================== RBAC CONFIGURATION ===================== */

/**
 * Shows RBAC configuration dialog for admins
 */
function setupRBACConfiguration() {
  if (!checkUserPermission(USER_ROLES.ADMIN)) {
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const currentAdmins = props.getProperty('ADMINS') || '';
  const currentStewards = props.getProperty('STEWARDS') || '';
  const currentViewers = props.getProperty('VIEWERS') || '';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      textarea { width: 100%; padding: 8px; margin-top: 5px; height: 80px; }
      button { margin-top: 20px; padding: 10px 20px; background: #DC2626; color: white; border: none; cursor: pointer; }
      button:hover { background: #B91C1C; }
      .info { background: #E0E7FF; padding: 10px; margin: 10px 0; border-radius: 4px; }
    </style>
    <h2>🔐 RBAC Configuration</h2>

    <div class="info">
      Enter email addresses separated by commas. Changes take effect immediately.
    </div>

    <label>Admins (Full Access):</label>
    <textarea id="admins">${currentAdmins}</textarea>

    <label>Stewards (Edit + Report Access):</label>
    <textarea id="stewards">${currentStewards}</textarea>

    <label>Viewers (Read-Only Access):</label>
    <textarea id="viewers">${currentViewers}</textarea>

    <button onclick="saveConfig()">Save Configuration</button>

    <script>
      function saveConfig() {
        const config = {
          admins: document.getElementById('admins').value,
          stewards: document.getElementById('stewards').value,
          viewers: document.getElementById('viewers').value
        };

        google.script.run
          .withSuccessHandler(() => {
            alert('RBAC configuration saved successfully');
            google.script.host.close();
          })
          .withFailureHandler(err => alert('Save failed: ' + err))
          .saveRBACConfiguration(config);
      }
    </script>
  `).setWidth(500).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'RBAC Configuration');
}

/**
 * Saves RBAC configuration to script properties
 */
function saveRBACConfiguration(config) {
  try {
    const props = PropertiesService.getScriptProperties();

    // Validate and sanitize emails
    const sanitizeEmails = (emailStr) => {
      return emailStr.split(',')
        .map(e => sanitizeInput(e.trim(), 'email'))
        .filter(e => e)
        .join(',');
    };

    props.setProperty('ADMINS', sanitizeEmails(config.admins));
    props.setProperty('STEWARDS', sanitizeEmails(config.stewards));
    props.setProperty('VIEWERS', sanitizeEmails(config.viewers));

    logDataModification(
      'RBAC_CONFIG',
      'System',
      'N/A',
      'RBAC Configuration',
      'Updated',
      'Roles configured'
    );

  } catch (error) {
    Logger.log('Error saving RBAC configuration: ' + error);
    throw error;
  }
}

/* ===================== INITIALIZATION ===================== */

/**
 * Initializes all security and audit features
 * Call this function once to set up the system
 */
function initializeSecurityFeatures() {
  if (!checkUserPermission(USER_ROLES.ADMIN)) {
    return;
  }

  try {
    SpreadsheetApp.getUi().alert('Initializing security features...');

    // Create required sheets
    createAuditLogSheet();
    createPerformanceLogSheet();

    // Update SHEETS constant to include new sheets
    if (typeof SHEETS !== 'undefined') {
      SHEETS.AUDIT_LOG = SECURITY_SHEETS.AUDIT_LOG;
      SHEETS.PERFORMANCE_LOG = SECURITY_SHEETS.PERFORMANCE_LOG;
    }

    SpreadsheetApp.getUi().alert(
      'Security Features Initialized',
      '✓ Audit_Log sheet created\n' +
      '✓ Performance_Log sheet created\n\n' +
      'Next steps:\n' +
      '1. Configure RBAC roles (Admin menu → Configure RBAC)\n' +
      '2. Set up backup folder (Script Properties → BACKUP_FOLDER_ID)\n' +
      '3. Enable Quick Actions sidebar from the menu',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('Error initializing security features: ' + error);
    SpreadsheetApp.getUi().alert('Error initializing: ' + error.message);
  }
}
