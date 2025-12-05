/**
 * ------------------------------------------------------------------------====
 * SECURITY UTILITIES - Core Security Functions
 * ------------------------------------------------------------------------====
 *
 * Provides essential security functions for the 509 Dashboard:
 * - HTML sanitization to prevent XSS attacks
 * - Role-based access control (RBAC)
 * - Input validation
 * - Audit logging
 * - Email security
 *
 * @module SecurityUtils
 * @version 2.0.0
 * @author SEIU Local 509 Tech Team
 * ------------------------------------------------------------------------====
 */

/* --------------------= CONFIGURATION --------------------= */

/**
 * Role definitions for access control
 * @const {Object}
 */
const SECURITY_ROLES = {
  ADMIN: 'ADMIN',
  STEWARD: 'STEWARD',
  MEMBER: 'MEMBER',
  GUEST: 'GUEST'
};

/**
 * Gets admin email addresses from Config sheet
 * Falls back to ADMIN_CONFIG.FALLBACK_ADMINS if Config sheet not available
 * @returns {string[]} Array of admin email addresses
 */
function getAdminEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEETS.CONFIG);

    if (!configSheet) {
      Logger.log('Config sheet not found, using fallback admin emails');
      return ADMIN_CONFIG.FALLBACK_ADMINS;
    }

    // Read admin emails from Config sheet (Column S, starting from row 2)
    const lastRow = configSheet.getLastRow();
    if (lastRow < 2) {
      return ADMIN_CONFIG.FALLBACK_ADMINS;
    }

    const adminEmailsRange = configSheet.getRange(2, ADMIN_CONFIG.CONFIG_COLUMN, lastRow - 1, 1);
    const adminEmailsData = adminEmailsRange.getValues();

    // Filter out empty rows and validate emails
    const adminEmails = adminEmailsData
      .flat()
      .filter(email => email && typeof email === 'string' && email.trim().length > 0)
      .map(email => email.trim());

    // If no admin emails found in Config, use fallback
    if (adminEmails.length === 0) {
      Logger.log('No admin emails found in Config sheet, using fallback');
      return ADMIN_CONFIG.FALLBACK_ADMINS;
    }

    return adminEmails;

  } catch (error) {
    Logger.log('Error reading admin emails from Config: ' + error.message);
    return ADMIN_CONFIG.FALLBACK_ADMINS;
  }
}

// Audit log configuration moved to Constants.gs (AUDIT_LOG_CONFIG)

/**
 * Rate limiting configuration
 * @const {Object}
 */
const RATE_LIMITS = {
  EMAIL_MIN_INTERVAL: 5000,        // 5 seconds between emails
  API_CALLS_PER_MINUTE: 60,        // Max API calls per minute
  EXPORT_MIN_INTERVAL: 30000,      // 30 seconds between exports
  BULK_OPERATION_MIN_INTERVAL: 60000 // 1 minute between bulk operations
};

/* --------------------= HTML SANITIZATION --------------------= */

/**
 * Sanitizes HTML string to prevent XSS attacks
 * Escapes dangerous characters that could execute scripts
 *
 * @param {string|number|null|undefined} input - Input to sanitize
 * @returns {string} Sanitized safe string
 *
 * @example
 * sanitizeHTML('<script>alert("XSS")</script>')
 * // Returns: '&lt;script&gt;alert(&quot;XSS&quot;)&lt;/script&gt;'
 */
function sanitizeHTML(input) {
  if (input === null || input === undefined) {
    return '';
  }

  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;');
}

/**
 * Sanitizes array of values
 *
 * @param {Array} arr - Array to sanitize
 * @returns {Array} Array with sanitized values
 */
function sanitizeArray(arr) {
  if (!Array.isArray(arr)) {
    return [];
  }
  return arr.map(function(item) { return sanitizeHTML(item); });
}

/**
 * Sanitizes object properties recursively
 *
 * @param {Object} obj - Object to sanitize
 * @returns {Object} Object with sanitized values
 */
function sanitizeObject(obj) {
  if (typeof obj !== 'object' || obj === null) {
    return sanitizeHTML(obj);
  }

  const sanitized = {};
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        sanitized[key] = sanitizeObject(obj[key]);
      } else {
        sanitized[key] = sanitizeHTML(obj[key]);
      }
    }
  }
  return sanitized;
}

/* --------------------= ACCESS CONTROL --------------------= */

/**
 * Gets the current user's email address
 *
 * @returns {string} User's email address
 */
function getCurrentUserEmail() {
  try {
    return Session.getEffectiveUser().getEmail();
  } catch (e) {
    Logger.log('Error getting user email: ' + e.message);
    return '';
  }
}

/**
 * Checks if current user is an administrator
 * Reads admin emails from Config sheet
 *
 * @returns {boolean} True if user is admin
 */
function isAdmin() {
  const userEmail = getCurrentUserEmail();
  const adminEmails = getAdminEmails();
  return adminEmails.includes(userEmail);
}

/**
 * Checks if current user is a steward
 * Queries Member Directory for Is Steward = "Yes"
 *
 * @returns {boolean} True if user is a steward
 */
function isSteward() {
  const userEmail = getCurrentUserEmail();

  if (!userEmail) {
    return false;
  }

  // Admins are also considered stewards
  if (isAdmin()) {
    return true;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      return false;
    }

    const data = memberSheet.getDataRange().getValues();

    // Check if user's email is in Member Directory with Is Steward = "Yes"
    for (let i = 1; i < data.length; i++) {
      const email = data[i][MEMBER_COLS.EMAIL - 1];
      const isStewardFlag = data[i][MEMBER_COLS.IS_STEWARD - 1];

      if (email === userEmail && isStewardFlag === 'Yes') {
        return true;
      }
    }

    return false;
  } catch (e) {
    Logger.log('Error checking steward status: ' + e.message);
    return false;
  }
}

/**
 * Checks if current user is a registered member
 *
 * @returns {boolean} True if user is in Member Directory
 */
function isMember() {
  const userEmail = getCurrentUserEmail();

  if (!userEmail) {
    return false;
  }

  // Admins and stewards are also members
  if (isAdmin() || isSteward()) {
    return true;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      return false;
    }

    const data = memberSheet.getDataRange().getValues();

    // Check if user's email is in Member Directory
    for (let i = 1; i < data.length; i++) {
      const email = data[i][MEMBER_COLS.EMAIL - 1];

      if (email === userEmail) {
        return true;
      }
    }

    return false;
  } catch (e) {
    Logger.log('Error checking member status: ' + e.message);
    return false;
  }
}

/**
 * Gets the current user's role
 *
 * @returns {string} User's role (ADMIN, STEWARD, MEMBER, or GUEST)
 */
function getUserRole() {
  if (isAdmin()) {
    return SECURITY_ROLES.ADMIN;
  }
  if (isSteward()) {
    return SECURITY_ROLES.STEWARD;
  }
  if (isMember()) {
    return SECURITY_ROLES.MEMBER;
  }
  return SECURITY_ROLES.GUEST;
}

/**
 * Checks if user has required role
 *
 * @param {string} requiredRole - Required role (ADMIN, STEWARD, or MEMBER)
 * @returns {boolean} True if user has required role or higher
 */
function hasRole(requiredRole) {
  const userRole = getUserRole();

  // Role hierarchy: ADMIN > STEWARD > MEMBER > GUEST
  const roleHierarchy = {
    ADMIN: 4,
    STEWARD: 3,
    MEMBER: 2,
    GUEST: 1
  };

  return roleHierarchy[userRole] >= roleHierarchy[requiredRole];
}

/**
 * Enforces role-based access control
 * Throws error if user doesn't have required role
 *
 * @param {string} requiredRole - Required role (ADMIN, STEWARD, or MEMBER)
 * @param {string} operationName - Name of operation for error message
 * @throws {Error} If user doesn't have required role
 *
 * @example
 * requireRole('STEWARD', 'Start Grievance');
 */
function requireRole(requiredRole, operationName) {
  if (!hasRole(requiredRole)) {
    const userRole = getUserRole();
    const errorMsg = `⛔ Access Denied: ${operationName} requires ${requiredRole} role. You have ${userRole} role.`;

    logAuditEvent('ACCESS_DENIED', {
      operation: operationName,
      requiredRole: requiredRole,
      userRole: userRole
    }, 'WARNING');

    throw new Error(errorMsg);
  }
}

/**
 * Shows access denied dialog to user
 *
 * @param {string} operationName - Name of operation
 * @param {string} requiredRole - Required role
 */
function showAccessDeniedDialog(operationName, requiredRole) {
  const ui = SpreadsheetApp.getUi();
  const userRole = getUserRole();

  ui.alert(
    '⛔ Access Denied',
    `${operationName} requires ${requiredRole} privileges.\n\n` +
    `Your current role: ${userRole}\n\n` +
    `Please contact an administrator if you need access.`,
    ui.ButtonSet.OK
  );
}

/* --------------------= INPUT VALIDATION --------------------= */

/**
 * Validates email address format
 *
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }

  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates phone number format
 * Accepts formats: (555) 123-4567, 555-123-4567, 5551234567
 *
 * @param {string} phone - Phone number to validate
 * @returns {boolean} True if valid phone format
 */
function isValidPhone(phone) {
  if (!phone || typeof phone !== 'string') {
    return false;
  }

  const phoneRegex = /^\(?([0-9]{3})\)?[-. ]?([0-9]{3})[-. ]?([0-9]{4})$/;
  return phoneRegex.test(phone);
}

/**
 * Validates member ID format
 * Expected format: M000001, M000002, etc.
 *
 * @param {string} memberId - Member ID to validate
 * @returns {boolean} True if valid member ID format
 */
function isValidMemberId(memberId) {
  if (!memberId || typeof memberId !== 'string') {
    return false;
  }

  // Allow M followed by 6 digits, or just digits
  const memberIdRegex = /^M?\d{6}$/;
  return memberIdRegex.test(memberId) && memberId.length <= 20;
}

/**
 * Validates grievance ID format
 * Expected format: G-000001, G-000002, etc.
 *
 * @param {string} grievanceId - Grievance ID to validate
 * @returns {boolean} True if valid grievance ID format
 */
function isValidGrievanceId(grievanceId) {
  if (!grievanceId || typeof grievanceId !== 'string') {
    return false;
  }

  const grievanceIdRegex = /^G-\d{6}$/;
  return grievanceIdRegex.test(grievanceId) && grievanceId.length <= 20;
}

/**
 * Validates date is not in the future (unless allowed)
 *
 * @param {Date} date - Date to validate
 * @param {boolean} allowFuture - Whether to allow future dates
 * @returns {boolean} True if valid date
 */
function isValidDate(date, allowFuture = false) {
  if (!(date instanceof Date) || isNaN(date)) {
    return false;
  }

  if (!allowFuture) {
    const now = new Date();
    return date <= now;
  }

  return true;
}

/**
 * Sanitizes and validates user input
 *
 * @param {string} input - Input to validate
 * @param {string} type - Type of input (email, phone, memberId, text)
 * @param {number} maxLength - Maximum length
 * @returns {Object} {valid: boolean, sanitized: string, error: string}
 */
function validateInput(input, type, maxLength = 255) {
  const result = {
    valid: false,
    sanitized: '',
    error: ''
  };

  // Sanitize first
  const sanitized = sanitizeHTML(input);
  result.sanitized = sanitized;

  // Check length
  if (sanitized.length > maxLength) {
    result.error = `Input exceeds maximum length of ${maxLength} characters`;
    return result;
  }

  // Type-specific validation
  switch (type) {
    case 'email':
      if (!isValidEmail(sanitized)) {
        result.error = 'Invalid email address format';
        return result;
      }
      break;

    case 'phone':
      if (!isValidPhone(sanitized)) {
        result.error = 'Invalid phone number format';
        return result;
      }
      break;

    case 'memberId':
      if (!isValidMemberId(sanitized)) {
        result.error = 'Invalid member ID format';
        return result;
      }
      break;

    case 'grievanceId':
      if (!isValidGrievanceId(sanitized)) {
        result.error = 'Invalid grievance ID format';
        return result;
      }
      break;

    case 'text':
      // Already sanitized, just check it's not empty
      if (!sanitized || sanitized.trim().length === 0) {
        result.error = 'Input cannot be empty';
        return result;
      }
      break;

    default:
      result.error = 'Unknown validation type';
      return result;
  }

  result.valid = true;
  return result;
}

/* --------------------= AUDIT LOGGING --------------------= */

/**
 * Logs security and audit events
 *
 * @param {string} action - Action name (e.g., 'GRIEVANCE_VIEWED', 'EMAIL_SENT')
 * @param {Object} details - Additional details about the action
 * @param {string} level - Log level (INFO, WARNING, ERROR, CRITICAL)
 *
 * @example
 * logAuditEvent('GRIEVANCE_VIEWED', {grievanceId: 'G-001234'}, 'INFO');
 */
function logAuditEvent(action, details = {}, level = 'INFO') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditLog = ss.getSheetByName(AUDIT_LOG_CONFIG.LOG_SHEET_NAME);

    // Create audit log sheet if it doesn't exist
    if (!auditLog) {
      auditLog = ss.insertSheet(AUDIT_LOG_CONFIG.LOG_SHEET_NAME);

      // Set up headers
      auditLog.appendRow([
        'Timestamp',
        'User Email',
        'User Role',
        'Action',
        'Level',
        'Details',
        'IP Address (N/A in Apps Script)'
      ]);

      // Format headers
      const headerRange = auditLog.getRange(1, 1, 1, 7);
      headerRange.setBackground('#1a73e8');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');

      // Protect sheet (warning only, so admins can still access)
      const protection = auditLog.protect();
      protection.setWarningOnly(true);
      protection.setDescription('Audit log - modifications are tracked');

      // Hide sheet from regular users
      auditLog.hideSheet();
    }

    // Add audit entry
    const userEmail = getCurrentUserEmail() || 'UNKNOWN';
    const userRole = getUserRole();
    const timestamp = new Date();
    const detailsJson = JSON.stringify(details);

    auditLog.appendRow([
      timestamp,
      userEmail,
      userRole,
      action,
      level,
      detailsJson,
      'N/A' // Apps Script doesn't provide IP addresses
    ]);

    // Trim old entries if needed (configured in AUDIT_LOG_CONFIG)
    if (AUDIT_LOG_CONFIG.AUTO_TRIM_ENABLED) {
      const lastRow = auditLog.getLastRow();
      const maxRows = AUDIT_LOG_CONFIG.MAX_ENTRIES + AUDIT_LOG_CONFIG.HEADER_ROWS;

      if (lastRow > maxRows) {
        const rowsToDelete = lastRow - maxRows;
        auditLog.deleteRows(2, rowsToDelete); // Delete from row 2 (after header)
      }
    }

  } catch (e) {
    // Don't let audit logging failures break the app
    Logger.log('Failed to log audit event: ' + e.message);
  }
}

/**
 * Gets recent audit log entries
 *
 * @param {number} limit - Maximum number of entries to return
 * @param {string} action - Filter by action (optional)
 * @returns {Array<Object>} Array of audit log entries
 */
function getAuditLog(limit = 100, action = null) {
  requireRole('ADMIN', 'View Audit Log');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditLog = ss.getSheetByName(AUDIT_LOG_CONFIG.LOG_SHEET_NAME);

    if (!auditLog) {
      return [];
    }

    const data = auditLog.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).reverse(); // Most recent first

    const filtered = action ? rows.filter(function(row) { return row[AUDIT_LOG_COLS.ACTION - 1] === action; }) : rows;

    const result = filtered.slice(0, limit).map(function(row) { return {
      timestamp: row[AUDIT_LOG_COLS.TIMESTAMP - 1],
      userEmail: row[AUDIT_LOG_COLS.USER_EMAIL - 1],
      userRole: row[AUDIT_LOG_COLS.USER_ROLE - 1],
      action: row[AUDIT_LOG_COLS.ACTION - 1],
      level: row[AUDIT_LOG_COLS.LEVEL - 1],
      details: row[AUDIT_LOG_COLS.DETAILS - 1],
      ipAddress: row[AUDIT_LOG_COLS.IP_ADDRESS - 1]
    };});

    return result;
  } catch (e) {
    Logger.log('Error getting audit log: ' + e.message);
    return [];
  }
}

/* --------------------= RATE LIMITING --------------------= */

/**
 * Checks if rate limit has been exceeded
 *
 * @param {string} operation - Operation name (e.g., 'EMAIL_SEND', 'EXPORT')
 * @param {number} minInterval - Minimum interval in milliseconds
 * @returns {Object} {allowed: boolean, waitTime: number}
 */
function checkRateLimit(operation, minInterval) {
  const props = PropertiesService.getUserProperties();
  const key = `RATE_LIMIT_${operation}`;
  const lastTime = props.getProperty(key);

  if (!lastTime) {
    return { allowed: true, waitTime: 0 };
  }

  const now = new Date().getTime();
  const elapsed = now - parseInt(lastTime);

  if (elapsed < minInterval) {
    const waitTime = minInterval - elapsed;
    return { allowed: false, waitTime: waitTime };
  }

  return { allowed: true, waitTime: 0 };
}

/**
 * Updates rate limit timestamp
 *
 * @param {string} operation - Operation name
 */
function updateRateLimit(operation) {
  const props = PropertiesService.getUserProperties();
  const key = `RATE_LIMIT_${operation}`;
  const now = new Date().getTime();
  props.setProperty(key, now.toString());
}

/**
 * Enforces rate limit for operation
 *
 * @param {string} operation - Operation name
 * @param {number} minInterval - Minimum interval in milliseconds
 * @throws {Error} If rate limit exceeded
 */
function enforceRateLimit(operation, minInterval) {
  const rateCheck = checkRateLimit(operation, minInterval);

  if (!rateCheck.allowed) {
    const waitSeconds = Math.ceil(rateCheck.waitTime / 1000);
    throw new Error(
      `⏱️ Rate limit exceeded. Please wait ${waitSeconds} seconds before trying again.`
    );
  }

  updateRateLimit(operation);
}

/* --------------------= EMAIL SECURITY --------------------= */

/**
 * Validates email recipient is a registered member or steward
 *
 * @param {string} email - Email address to validate
 * @returns {boolean} True if email is in Member Directory
 */
function isRegisteredEmail(email) {
  if (!isValidEmail(email)) {
    return false;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      return false;
    }

    const data = memberSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const memberEmail = data[i][MEMBER_COLS.EMAIL - 1];
      if (memberEmail === email) {
        return true;
      }
    }

    return false;
  } catch (e) {
    Logger.log('Error checking registered email: ' + e.message);
    return false;
  }
}

/**
 * Gets list of all registered member emails
 *
 * @returns {string[]} Array of email addresses
 */
function getAllMemberEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      return [];
    }

    const data = memberSheet.getDataRange().getValues();
    const emails = [];

    for (let i = 1; i < data.length; i++) {
      const email = data[i][MEMBER_COLS.EMAIL - 1];
      if (email && isValidEmail(email)) {
        emails.push(email);
      }
    }

    return emails;
  } catch (e) {
    Logger.log('Error getting member emails: ' + e.message);
    return [];
  }
}

/* --------------------= SECURITY AUDIT --------------------= */

/**
 * Runs security audit and returns report
 * Only accessible by admins
 *
 * @returns {Object} Security audit report
 */
function runSecurityAudit() {
  requireRole('ADMIN', 'Run Security Audit');

  const report = {
    timestamp: new Date(),
    auditedBy: getCurrentUserEmail(),
    results: {
      totalUsers: 0,
      adminCount: 0,
      stewardCount: 0,
      memberCount: 0,
      recentAccessDenied: 0,
      recentEmailsSent: 0,
      auditLogSize: 0
    },
    recommendations: []
  };

  // Count users by role
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (memberSheet) {
    const data = memberSheet.getDataRange().getValues();
    report.results.totalUsers = data.length - 1; // Exclude header

    for (let i = 1; i < data.length; i++) {
      const email = data[i][MEMBER_COLS.EMAIL - 1];
      const isStewardFlag = data[i][MEMBER_COLS.IS_STEWARD - 1];

      if (getAdminEmails().includes(email)) {
        report.results.adminCount++;
      } else if (isStewardFlag === 'Yes') {
        report.results.stewardCount++;
      } else {
        report.results.memberCount++;
      }
    }
  }

  // Get recent security events
  const auditLog = ss.getSheetByName(AUDIT_LOG_CONFIG.LOG_SHEET_NAME);
  if (auditLog) {
    const data = auditLog.getDataRange().getValues();
    report.results.auditLogSize = data.length - 1;

    // Count recent events (last 7 days)
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);

    for (let i = 1; i < data.length; i++) {
      const timestamp = data[i][AUDIT_LOG_COLS.TIMESTAMP - 1];
      const action = data[i][AUDIT_LOG_COLS.ACTION - 1];

      if (timestamp >= weekAgo) {
        if (action === 'ACCESS_DENIED') {
          report.results.recentAccessDenied++;
        } else if (action === 'EMAIL_SENT') {
          report.results.recentEmailsSent++;
        }
      }
    }
  }

  // Generate recommendations
  if (report.results.adminCount === 0) {
    report.recommendations.push('⚠️ No admins configured. Add admin emails to ADMIN_EMAILS constant.');
  }

  if (report.results.stewardCount === 0) {
    report.recommendations.push('⚠️ No stewards found. Ensure Member Directory has Is Steward = "Yes" for stewards.');
  }

  if (report.results.recentAccessDenied > 10) {
    report.recommendations.push('⚠️ High number of access denied events. Review user roles and permissions.');
  }

  if (report.results.auditLogSize > AUDIT_LOG_CONFIG.MAX_ENTRIES * 0.9) {
    report.recommendations.push(`ℹ️ Audit log approaching limit (${AUDIT_LOG_CONFIG.MAX_ENTRIES} entries). Old entries will be auto-deleted.`);
  }

  logAuditEvent('SECURITY_AUDIT', report.results, 'INFO');

  return report;
}

/**
 * Shows security audit report to admin
 */
function showSecurityAudit() {
  try {
    const report = runSecurityAudit();
    const ui = SpreadsheetApp.getUi();

    let message = '🔒 SECURITY AUDIT REPORT\n\n';
    message += `Total Users: ${report.results.totalUsers}\n`;
    message += `Admins: ${report.results.adminCount}\n`;
    message += `Stewards: ${report.results.stewardCount}\n`;
    message += `Members: ${report.results.memberCount}\n\n`;
    message += `Recent Events (7 days):\n`;
    message += `- Access Denied: ${report.results.recentAccessDenied}\n`;
    message += `- Emails Sent: ${report.results.recentEmailsSent}\n\n`;
    message += `Audit Log Entries: ${report.results.auditLogSize}\n\n`;

    if (report.recommendations.length > 0) {
      message += 'RECOMMENDATIONS:\n';
      report.recommendations.forEach(function(rec) {
        message += `${rec}\n`;
      });
    } else {
      message += '✅ No security issues detected.';
    }

    ui.alert('Security Audit Report', message, ui.ButtonSet.OK);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', 'Security audit failed: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
