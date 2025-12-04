/**
 * ------------------------------------------------------------------------====
 * SECURITY SERVICE - RBAC and Audit Logging
 * ------------------------------------------------------------------------====
 *
 * Provides role-based access control and comprehensive audit logging for
 * the 509 Dashboard to ensure security and compliance.
 *
 * Features:
 * - Role-based permissions (Admin, Steward, Member, Viewer)
 * - Function-level access control
 * - Comprehensive audit logging
 * - User activity tracking
 * - Change history
 *
 * ------------------------------------------------------------------------====
 */

/* --------------------= ROLE DEFINITIONS --------------------= */

/**
 * Role definitions with permissions
 * @type {Object}
 */
const ROLES = {
  ADMIN: {
    name: 'Admin',
    permissions: [
      'view', 'edit', 'delete', 'export', 'configure',
      'seed_data', 'clear_data', 'manage_users', 'view_audit_log'
    ],
    description: 'Full system access including configuration and user management'
  },
  STEWARD: {
    name: 'Steward',
    permissions: [
      'view', 'edit', 'comment', 'export', 'create_grievance',
      'update_grievance', 'view_assigned_grievances'
    ],
    description: 'Can view and edit grievances, manage assigned cases'
  },
  COORDINATOR: {
    name: 'Grievance Coordinator',
    permissions: [
      'view', 'edit', 'comment', 'export', 'create_grievance',
      'update_grievance', 'view_all_grievances', 'assign_stewards'
    ],
    description: 'Can manage all grievances and assign stewards'
  },
  MEMBER: {
    name: 'Member',
    permissions: [
      'view_own', 'view_own_grievances', 'comment'
    ],
    description: 'Can view own member data and grievances only'
  },
  VIEWER: {
    name: 'Viewer',
    permissions: [
      'view'
    ],
    description: 'Read-only access to anonymized data'
  }
};

/* --------------------= USER ROLE MANAGEMENT --------------------= */

/**
 * Gets the role for a user
 * @param {string} userEmail - User's email address (defaults to current user)
 * @returns {string} Role name (defaults to 'VIEWER')
 */
function getUserRole(userEmail) {
  if (!userEmail) {
    userEmail = Session.getActiveUser().getEmail();
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userRolesSheet = ss.getSheetByName('User Roles');

    if (!userRolesSheet) {
      // If sheet doesn't exist, create it and assign current user as admin
      userRolesSheet = createUserRolesSheet();
      assignRole(userEmail, 'ADMIN');
      return 'ADMIN';
    }

    // Look up user's role
    const data = userRolesSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === userEmail.toLowerCase()) {
        return data[i][1] || 'VIEWER';
      }
    }

    // User not found, default to VIEWER
    return 'VIEWER';
  } catch (error) {
    handleError(error, 'getUserRole');
    return 'VIEWER';
  }
}

/**
 * Assigns a role to a user
 * @param {string} userEmail - User's email
 * @param {string} role - Role to assign (ADMIN, STEWARD, COORDINATOR, MEMBER, VIEWER)
 * @returns {boolean} Success status
 */
function assignRole(userEmail, role) {
  try {
    if (!ROLES[role]) {
      throw new Error(`Invalid role: ${role}`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userRolesSheet = ss.getSheetByName('User Roles');

    if (!userRolesSheet) {
      userRolesSheet = createUserRolesSheet();
    }

    // Check if user already exists
    const data = userRolesSheet.getDataRange().getValues();
    let userRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toLowerCase() === userEmail.toLowerCase()) {
        userRow = i + 1;
        break;
      }
    }

    const timestamp = new Date();
    const assignedBy = Session.getActiveUser().getEmail();

    if (userRow > 0) {
      // Update existing user
      userRolesSheet.getRange(userRow, 2, 1, 3).setValues([[role, timestamp, assignedBy]]);
    } else {
      // Add new user
      userRolesSheet.appendRow([userEmail, role, timestamp, assignedBy]);
    }

    // Log the role assignment
    logAudit('ROLE_ASSIGNMENT', `Assigned role ${role} to ${userEmail}`, {
      userEmail: userEmail,
      role: role,
      assignedBy: assignedBy
    });

    return true;
  } catch (error) {
    handleError(error, 'assignRole');
    return false;
  }
}

/**
 * Creates the User Roles sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The created sheet
 */
function createUserRolesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('User Roles');

  const headers = ['Email', 'Role', 'Assigned Date', 'Assigned By'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4A5568')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.hideSheet(); // Hide from regular users

  return sheet;
}

/* --------------------= PERMISSION CHECKING --------------------= */

/**
 * Checks if current user has a specific permission
 * @param {string} permission - Permission to check
 * @param {string} userEmail - User email (optional, defaults to current user)
 * @returns {boolean} True if user has permission
 */
function hasPermission(permission, userEmail) {
  try {
    const role = getUserRole(userEmail);
    const roleConfig = ROLES[role];

    if (!roleConfig) {
      return false;
    }

    return roleConfig.permissions.includes(permission);
  } catch (error) {
    handleError(error, 'hasPermission');
    return false;
  }
}

/**
 * Checks permission and throws error if not authorized
 * @param {string} permission - Required permission
 * @param {string} action - Action being attempted (for error message)
 * @throws {Error} If user lacks permission
 */
function requirePermission(permission, action) {
  if (!hasPermission(permission)) {
    const userEmail = Session.getActiveUser().getEmail();
    const role = getUserRole(userEmail);

    logAudit('ACCESS_DENIED', `User ${userEmail} attempted ${action} without permission ${permission}`, {
      userEmail: userEmail,
      role: role,
      permission: permission,
      action: action
    });

    throw new Error(`Access Denied: You need ${permission} permission to ${action}. Your role: ${role}`);
  }
}

/**
 * Wraps a function with permission check
 * @param {Function} fn - Function to wrap
 * @param {string} requiredPermission - Permission required
 * @param {string} actionDescription - Description of action
 * @returns {Function} Wrapped function
 */
function withPermission(fn, requiredPermission, actionDescription) {
  return function(...args) {
    requirePermission(requiredPermission, actionDescription);
    return fn.apply(this, args);
  };
}

/* --------------------= AUDIT LOGGING --------------------= */

/**
 * Logs an audit event
 * @param {string} eventType - Type of event (LOGIN, DATA_CHANGE, ACCESS_DENIED, etc.)
 * @param {string} description - Human-readable description
 * @param {Object} metadata - Additional metadata
 */
function logAudit(eventType, description, metadata) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName('Audit Log');

    if (!auditSheet) {
      auditSheet = createAuditLogSheet();
    }

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail();
    const ipAddress = ''; // Apps Script doesn't provide IP directly
    const metadataJson = metadata ? JSON.stringify(metadata) : '';

    auditSheet.appendRow([
      timestamp,
      userEmail,
      eventType,
      description,
      metadataJson,
      ipAddress
    ]);

    // Keep only last 10,000 rows for performance
    const lastRow = auditSheet.getLastRow();
    if (lastRow > 10000) {
      auditSheet.deleteRows(2, lastRow - 10000);
    }
  } catch (error) {
    // Don't throw error in logging to avoid breaking functionality
    Logger.log('Error in logAudit: ' + error.message);
  }
}

/**
 * Creates the Audit Log sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The created sheet
 */
function createAuditLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('Audit Log');

  const headers = ['Timestamp', 'User Email', 'Event Type', 'Description', 'Metadata', 'IP Address'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}

/**
 * Logs data change event
 * @param {string} sheetName - Name of sheet changed
 * @param {string} operation - Operation (INSERT, UPDATE, DELETE)
 * @param {Object} details - Details about the change
 */
function logDataChange(sheetName, operation, details) {
  logAudit('DATA_CHANGE', `${operation} in ${sheetName}`, {
    sheetName: sheetName,
    operation: operation,
    ...details
  });
}

/**
 * Logs user login/access
 */
function logUserAccess() {
  const userEmail = Session.getActiveUser().getEmail();
  const role = getUserRole(userEmail);

  logAudit('ACCESS', `User accessed dashboard`, {
    userEmail: userEmail,
    role: role
  });
}

/* --------------------= DATA ACCESS CONTROL --------------------= */

/**
 * Filters member data based on user permissions
 * @param {Array<Array>} memberData - Full member data
 * @param {string} userEmail - User email (optional)
 * @returns {Array<Array>} Filtered data
 */
function filterMemberDataByPermission(memberData, userEmail) {
  if (!userEmail) {
    userEmail = Session.getActiveUser().getEmail();
  }

  const role = getUserRole(userEmail);

  // Admin and Coordinator can see all
  if (role === 'ADMIN' || role === 'COORDINATOR') {
    return memberData;
  }

  // Stewards can see all (but with limited edit permissions)
  if (role === 'STEWARD') {
    return memberData;
  }

  // Members can only see their own data
  if (role === 'MEMBER') {
    return memberData.filter(function(row, index) {
      if (index === 0) return true; // Keep header
      return row[7] && row[7].toLowerCase() === userEmail.toLowerCase(); // Email column
    });
  }

  // Viewers see anonymized data (remove PII)
  if (role === 'VIEWER') {
    return memberData.map(function(row, index) {
      if (index === 0) return row; // Keep header

      // Anonymize PII
      return row.map(function(cell, colIndex) {
        // Hide email (column 8) and phone (column 9)
        if (colIndex === 7 || colIndex === 8) {
          return '[REDACTED]';
        }
        return cell;
      });
    });
  }

  return [memberData[0]]; // Return only header for unknown roles
}

/**
 * Filters grievance data based on user permissions
 * @param {Array<Array>} grievanceData - Full grievance data
 * @param {string} userEmail - User email (optional)
 * @returns {Array<Array>} Filtered data
 */
function filterGrievanceDataByPermission(grievanceData, userEmail) {
  if (!userEmail) {
    userEmail = Session.getActiveUser().getEmail();
  }

  const role = getUserRole(userEmail);

  // Admin and Coordinator can see all
  if (role === 'ADMIN' || role === 'COORDINATOR') {
    return grievanceData;
  }

  // Stewards can only see assigned grievances
  if (role === 'STEWARD') {
    // First, get steward name from member directory
    const memberData = getSheetDataSafely(SHEETS.MEMBER_DIR);
    let stewardName = '';

    if (memberData) {
      const stewardRow = memberData.find(function(row) {
        return row[7] && row[7].toLowerCase() === userEmail.toLowerCase();
      });
      if (stewardRow) {
        stewardName = `${stewardRow[1]} ${stewardRow[2]}`; // First + Last name
      }
    }

    return grievanceData.filter(function(row, index) {
      if (index === 0) return true; // Keep header
      const assignedSteward = row[26]; // Assigned Steward column
      return assignedSteward && assignedSteward.includes(stewardName);
    });
  }

  // Members can only see their own grievances
  if (role === 'MEMBER') {
    return grievanceData.filter(function(row, index) {
      if (index === 0) return true; // Keep header
      const memberEmail = row[23]; // Member Email column
      return memberEmail && memberEmail.toLowerCase() === userEmail.toLowerCase();
    });
  }

  // Viewers see anonymized data
  if (role === 'VIEWER') {
    return grievanceData.map(function(row, index) {
      if (index === 0) return row; // Keep header

      // Anonymize PII
      return row.map(function(cell, colIndex) {
        // Hide names (2,3), email (23)
        if ([2, 3, 23].includes(colIndex)) {
          return '[REDACTED]';
        }
        return cell;
      });
    });
  }

  return [grievanceData[0]]; // Return only header for unknown roles
}

/* --------------------= PROTECTED OPERATIONS --------------------= */

/**
 * Protected version of SEED_MEMBERS - requires admin permission
 */
function protectedSeedMembers() {
  requirePermission('seed_data', 'seed member data');
  logAudit('SEED_DATA', 'Started seeding member data');

  // Call original function
  return SEED_20K_MEMBERS();
}

/**
 * Protected version of CLEAR_ALL_DATA - requires admin permission
 */
function protectedClearAllData() {
  requirePermission('clear_data', 'clear all data');

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ DANGER ZONE ⚠️',
    'This will permanently delete all member and grievance data!\n\n' +
    'Are you absolutely sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  logAudit('CLEAR_DATA', 'User confirmed: Clearing all data');

  // Call original clear function
  return CLEAR_ALL_DATA();
}

/* --------------------= ADMIN FUNCTIONS --------------------= */

/**
 * Shows the user management dialog (Admin only)
 */
function showUserManagement() {
  requirePermission('manage_users', 'manage user roles');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userRolesSheet = ss.getSheetByName('User Roles');

  if (!userRolesSheet) {
    userRolesSheet = createUserRolesSheet();
  }

  userRolesSheet.showSheet();
  userRolesSheet.activate();

  SpreadsheetApp.getUi().alert(
    '👥 User Management',
    'You can now view and edit user roles.\n\n' +
    'Columns:\n' +
    '• Email: User email address\n' +
    '• Role: ADMIN, COORDINATOR, STEWARD, MEMBER, or VIEWER\n' +
    '• Assigned Date: When role was assigned\n' +
    '• Assigned By: Who assigned the role\n\n' +
    'Remember to hide this sheet when done!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows the audit log (Admin only)
 */
function showAuditLog() {
  requirePermission('view_audit_log', 'view audit log');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit Log');

  if (!auditSheet) {
    SpreadsheetApp.getUi().alert('No audit log found.');
    return;
  }

  auditSheet.showSheet();
  auditSheet.activate();

  SpreadsheetApp.getUi().alert(
    '📋 Audit Log',
    'Showing all system access and change events.\n\n' +
    'This log tracks:\n' +
    '• User access\n' +
    '• Data changes\n' +
    '• Permission checks\n' +
    '• Security events\n\n' +
    'Log is limited to last 10,000 events.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Exports audit log to CSV (Admin only)
 */
function exportAuditLog() {
  requirePermission('view_audit_log', 'export audit log');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit Log');

  if (!auditSheet) {
    SpreadsheetApp.getUi().alert('No audit log found.');
    return;
  }

  // This would create a CSV file in Google Drive
  const data = auditSheet.getDataRange().getValues();
  const csv = data.map(function(row) { return row.join(','); }).join('\n');

  const folder = DriveApp.getRootFolder();
  const file = folder.createFile(
    `Audit_Log_${new Date().toISOString().slice(0, 10)}.csv`,
    csv,
    MimeType.CSV
  );

  logAudit('EXPORT_AUDIT_LOG', 'Exported audit log to CSV');

  SpreadsheetApp.getUi().alert(
    '✅ Export Complete',
    `Audit log exported to:\n${file.getName()}\n\nFile ID: ${file.getId()}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
