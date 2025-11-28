/**
 * ============================================================================
 * ARCHIVE IMPLEMENTED FEATURES
 * ============================================================================
 *
 * Adds Features 79-94 to the Feedback & Development sheet as "Completed"
 * This documents what has been implemented and removes them from active development
 *
 * ============================================================================
 */

/**
 * Archives all implemented features (79-94) to Feedback & Development sheet
 */
function archiveImplementedFeatures() {
  const ss = SpreadsheetApp.getActive();
  const feedbackSheet = ss.getSheetByName("Feedback & Development");

  if (!feedbackSheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'Feedback & Development sheet not found. Please create it first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const implementedFeatures = [
    // [Type, Submitted/Started, Submitted By, Priority, Title, Description, Status, Progress %, Complexity, Target Completion, Assigned To, Blockers, Resolution/Notes, Last Updated]
    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 79: Audit Logging",
     "Tracks all data modifications with comprehensive details including user, timestamp, action type, sheet name, record ID, field changes, and session ID",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Fully implemented in SecurityAndAudit.gs - logDataModification() function with automatic suspicious activity detection", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 80: Role-Based Access Control (RBAC)",
     "Three-tier permission system (Admin, Steward, Viewer) with role checking and configuration dialog",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Fully implemented with checkUserPermission() and setupRBACConfiguration() - configurable via Script Properties", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 83: Input Sanitization",
     "Validates and sanitizes user input to prevent security vulnerabilities (script tags, HTML, validates emails, numbers, dates)",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented sanitizeInput() with support for text, email, number, date, alphanumeric types", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Medium", "Feature 84: Audit Reporting",
     "Generates comprehensive audit trail reports for specified date ranges with summary statistics and action counts",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented generateAuditReport() with interactive date range dialog - requires Steward or Admin role", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Medium", "Feature 85: Data Retention Policy",
     "Enforces configurable data retention policy (default: 7 years) with automatic deletion of old audit log entries",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented enforceDataRetention() - configurable via DATA_RETENTION_YEARS script property", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 86: Suspicious Activity Detection",
     "Automatically detects unusual activity patterns (>50 changes/hour) and sends email alerts to admins",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented detectSuspiciousActivity() - runs automatically on each audit log entry", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Medium", "Feature 87: Quick Actions Sidebar",
     "Interactive sidebar with one-click access to common actions, role-based menu items, and admin tools",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented showQuickActionsSidebar() with HTML-based interactive UI and role-based filtering", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Medium", "Feature 88: Advanced Search",
     "Interactive grievance search with multiple criteria: Grievance ID, Member Name, Issue Type, Steward",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented showSearchDialog() with live results display - accessible from menu and sidebar", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Medium", "Feature 89: Advanced Filtering",
     "Filter grievances by status (Open/In Progress/Closed), issue type, and date range with clear filters option",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented showFilterDialog() using Google Sheets native filters - accessible from menu and sidebar", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 90: Automated Backups",
     "Creates automated backups to Google Drive with optional daily scheduled backups (2:00 AM)",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented createAutomatedBackup() and setupDailyBackupTrigger() - requires BACKUP_FOLDER_ID script property", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Low", "Feature 91: Performance Monitoring",
     "Tracks script execution times and performance with execution time, parameters, and error status logging",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented trackPerformance() - creates Performance_Log sheet automatically", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "Low", "Feature 92: Keyboard Shortcuts",
     "Provides guidance on keyboard shortcuts (Google Sheets doesn't support custom shortcuts)",
     "Completed", "100%", "Simple", new Date("2025-11-27"), "Claude AI",
     "", "Implemented setupKeyboardShortcuts() with documentation of built-in Google Sheets shortcuts", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 93: Export Wizard",
     "Guided export with multiple options: type selection, format (CSV/Excel/PDF), headers, date filtering",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented showExportWizard() with interactive wizard and input sanitization", new Date()],

    ["Completed", new Date("2025-11-27"), "System", "High", "Feature 94: Data Import",
     "Bulk data import capability for members and grievances from CSV/Excel with validation and error handling",
     "Completed", "100%", "Moderate", new Date("2025-11-27"), "Claude AI",
     "", "Implemented showImportWizard() with input sanitization - requires Steward or Admin role", new Date()]
  ];

  // Find the last row with data
  const lastRow = feedbackSheet.getLastRow();

  // Add all implemented features
  feedbackSheet.getRange(lastRow + 1, 1, implementedFeatures.length, 14).setValues(implementedFeatures);

  // Apply formatting to completed features
  const completedRange = feedbackSheet.getRange(lastRow + 1, 1, implementedFeatures.length, 14);
  completedRange.setBackground("#D1FAE5");  // Light green background
  completedRange.setFontColor("#059669");   // Dark green text

  SpreadsheetApp.getUi().alert(
    '✅ Features 79-94 Archived',
    'All 14 implemented security and enhancement features have been added to the Feedback & Development sheet.\n\n' +
    'They are marked as "Completed" with:\n' +
    '• Implementation date: 2025-11-27\n' +
    '• Progress: 100%\n' +
    '• Status: Completed\n' +
    '• Green highlighting for easy identification\n\n' +
    'These features are now documented and separated from active development items.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Removes all completed features from Feedback & Development sheet
 * (Use with caution - this is destructive)
 */
function clearCompletedFeatures() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Warning: Clear Completed Features',
    'This will PERMANENTLY DELETE all rows marked as "Completed" from the Feedback & Development sheet.\n\n' +
    'This action CANNOT be undone!\n\n' +
    'Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('✅ Operation cancelled. No features were deleted.');
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const feedbackSheet = ss.getSheetByName("Feedback & Development");

  if (!feedbackSheet) {
    ui.alert('❌ Error', 'Feedback & Development sheet not found.', ui.ButtonSet.OK);
    return;
  }

  const data = feedbackSheet.getDataRange().getValues();
  const rowsToDelete = [];

  // Find all rows with "Completed" status (column G, index 6)
  for (let i = 3; i < data.length; i++) {  // Start from row 4 (index 3), skip headers
    if (data[i][6] === "Completed") {
      rowsToDelete.push(i + 1);  // Store row number (1-indexed)
    }
  }

  if (rowsToDelete.length === 0) {
    ui.alert('ℹ️ No Completed Features', 'No completed features found to delete.', ui.ButtonSet.OK);
    return;
  }

  // Delete rows in reverse order to maintain row indices
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    feedbackSheet.deleteRow(rowsToDelete[i]);
  }

  ui.alert(
    '✅ Completed Features Cleared',
    `Deleted ${rowsToDelete.length} completed feature(s) from the Feedback & Development sheet.`,
    ui.ButtonSet.OK
  );
}

/**
 * Moves completed features to a separate "Completed Features" sheet
 */
function moveCompletedToArchive() {
  const ss = SpreadsheetApp.getActive();
  const feedbackSheet = ss.getSheetByName("Feedback & Development");

  if (!feedbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Feedback & Development sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Create or get "Completed Features" sheet
  let archiveSheet = ss.getSheetByName("📦 Completed Features");
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("📦 Completed Features");

    // Add headers
    const headers = feedbackSheet.getRange(3, 1, 1, 14).getValues();
    archiveSheet.getRange(1, 1, 1, 14).setValues(headers)
      .setFontWeight("bold")
      .setBackground("#E5E7EB")
      .setFontColor("#1F2937");

    archiveSheet.setFrozenRows(1);
    archiveSheet.setTabColor("#6B7280");
  }

  const data = feedbackSheet.getDataRange().getValues();
  const completedRows = [];
  const rowsToDelete = [];

  // Find all completed features
  for (let i = 3; i < data.length; i++) {
    if (data[i][6] === "Completed") {
      completedRows.push(data[i]);
      rowsToDelete.push(i + 1);
    }
  }

  if (completedRows.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Completed Features', 'No completed features found to archive.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Add completed features to archive sheet
  const archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, completedRows.length, 14).setValues(completedRows);

  // Delete from feedback sheet (in reverse order)
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    feedbackSheet.deleteRow(rowsToDelete[i]);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Completed Features Archived',
    `Moved ${completedRows.length} completed feature(s) to the "📦 Completed Features" sheet.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
