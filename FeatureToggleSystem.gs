/**
 * ============================================================================
 * FEATURE TOGGLE SYSTEM
 * ============================================================================
 *
 * Allows admins to enable/disable features from the Admin menu.
 * Features are controlled via the "Feature Controls" sheet.
 *
 * Features:
 * - Enable/disable menu items dynamically
 * - Control dashboard visibility
 * - Manage tool availability
 * - Role-based feature access
 *
 * ============================================================================
 */

/**
 * Creates the Feature Controls sheet
 */
function createFeatureControlsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet("⚙️ Feature Controls");
  sheet.clear();

  // Title
  sheet.getRange("A1:G1").merge()
    .setValue("⚙️ FEATURE CONTROLS - Admin Feature Management")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#7C3AED")
    .setFontColor("#FFFFFF");

  sheet.getRange("A2:G2").merge()
    .setValue("💡 Enable or disable features by changing the 'Enabled' column. Refresh the page to see changes in the menu.")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground("#F3F4F6")
    .setFontColor("#6B7280");

  // Headers
  const headers = [
    "Feature Name",
    "Description",
    "Category",
    "Enabled",
    "Menu Location",
    "Function Name",
    "Required Role"
  ];

  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground("#E5E7EB")
    .setFontColor("#1F2937");

  // Feature list with all features
  const features = [
    // VIEW MENU
    ["Unified Operations Monitor", "Terminal-style comprehensive dashboard", "View", "Yes", "View", "showUnifiedOperationsMonitor", "All"],
    ["Main Dashboard", "Overview dashboard with key metrics", "View", "Yes", "View", "goToDashboard", "All"],
    ["Interactive Dashboard", "Customizable dashboard with metric/chart dropdowns", "View", "Yes", "View", "openInteractiveDashboard", "All"],
    ["Refresh Interactive Dashboard", "Rebuild interactive dashboard charts", "View", "Yes", "View", "rebuildInteractiveDashboard", "All"],
    ["Hide Gridlines (Focus Mode)", "Hide all gridlines for cleaner view", "View", "Yes", "View", "hideAllGridlines", "All"],
    ["Show Gridlines", "Show all gridlines", "View", "Yes", "View", "showAllGridlines", "All"],
    ["Reorder Sheets Logically", "Organize sheets in logical order", "View", "Yes", "View", "reorderSheetsLogically", "All"],
    ["Setup ADHD Defaults", "Configure ADHD-friendly defaults", "View", "Yes", "View", "setupADHDDefaults", "All"],
    ["Toggle Advanced Grievance Columns", "Show/hide advanced grievance columns", "View", "Yes", "View", "toggleGrievanceColumns", "All"],
    ["Toggle Level 2 Member Columns", "Show/hide level 2 member columns", "View", "Yes", "View", "toggleLevel2Columns", "All"],
    ["Show All Member Columns", "Show all hidden member columns", "View", "Yes", "View", "showAllMemberColumns", "All"],

    // GRIEVANCES MENU
    ["Start New Grievance", "Launch new grievance workflow", "Grievances", "Yes", "Grievances", "showStartGrievanceDialog", "Steward"],
    ["Advanced Search", "Search grievances by multiple criteria", "Grievances", "Yes", "Grievances", "showSearchDialog", "All"],
    ["Advanced Filter", "Filter grievances by status/type/date", "Grievances", "Yes", "Grievances", "showFilterDialog", "All"],

    // REPORTS & EXPORT
    ["Export Wizard", "Guided data export (CSV/Excel/PDF)", "Reports", "Yes", "Reports & Export", "showExportWizard", "All"],
    ["Import Wizard", "Bulk data import from CSV/Excel", "Reports", "Yes", "Reports & Export", "showImportWizard", "Steward"],

    // SECURITY (ADMIN ONLY)
    ["Generate Audit Report", "Create audit trail report", "Security", "Yes", "Security", "showAuditReportDialog", "Steward"],
    ["Create Backup", "Manual backup to Google Drive", "Security", "Yes", "Security", "createAutomatedBackup", "Admin"],
    ["Setup Daily Backups", "Configure automated daily backups", "Security", "Yes", "Security", "setupDailyBackupTrigger", "Admin"],
    ["Configure RBAC", "Set up role-based access control", "Security", "Yes", "Security", "setupRBACConfiguration", "Admin"],
    ["Enforce Data Retention", "Apply data retention policy", "Security", "Yes", "Security", "enforceDataRetention", "Admin"],
    ["Initialize Security Features", "Set up security sheets and features", "Security", "Yes", "Security", "initializeSecurityFeatures", "Admin"],

    // ADMIN & TESTING
    ["Seed 5k Members", "Generate 5,000 test member records", "Admin", "Yes", "Admin & Testing", "SEED_20K_MEMBERS", "Admin"],
    ["Seed 1k Grievances", "Generate 1,000 test grievance records", "Admin", "Yes", "Admin & Testing", "SEED_5K_GRIEVANCES", "Admin"],
    ["Clear All Data", "Delete all members and grievances", "Admin", "Yes", "Admin & Testing", "clearAllData", "Admin"],
    ["Nuke All Seed Data", "Remove all seed data and exit demo mode", "Admin", "Yes", "Admin & Testing", "nukeSeedData", "Admin"],
    ["Diagnose Setup", "Run system diagnostics", "Admin", "Yes", "Admin & Testing", "DIAGNOSE_SETUP", "Admin"],

    // HELP
    ["Quick Actions Sidebar", "Interactive sidebar with common actions", "Help", "Yes", "Help", "showQuickActionsSidebar", "All"],
    ["Getting Started Guide", "New user onboarding guide", "Help", "Yes", "Help", "showGettingStartedGuide", "All"],
    ["Help", "System help and documentation", "Help", "Yes", "Help", "showHelp", "All"],
    ["Keyboard Shortcuts", "Available keyboard shortcuts", "Help", "Yes", "Help", "setupKeyboardShortcuts", "All"]
  ];

  sheet.getRange(4, 1, features.length, 7).setValues(features);

  // Add data validation for Enabled column
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(4, 4, features.length, 1).setDataValidation(enabledRule);

  // Add data validation for Category column
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['View', 'Grievances', 'Reports', 'Security', 'Admin', 'Help'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(4, 3, features.length, 1).setDataValidation(categoryRule);

  // Add data validation for Required Role column
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All', 'Steward', 'Admin'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(4, 7, features.length, 1).setDataValidation(roleRule);

  // Formatting
  sheet.setFrozenRows(3);
  sheet.setColumnWidth(1, 250);  // Feature Name
  sheet.setColumnWidth(2, 350);  // Description
  sheet.setColumnWidth(3, 100);  // Category
  sheet.setColumnWidth(4, 80);   // Enabled
  sheet.setColumnWidth(5, 150);  // Menu Location
  sheet.setColumnWidth(6, 200);  // Function Name
  sheet.setColumnWidth(7, 120);  // Required Role

  // Conditional formatting for Enabled column
  const enabledRange = sheet.getRange(4, 4, features.length, 1);

  // Green for "Yes"
  const yesRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Yes")
    .setBackground("#D1FAE5")
    .setFontColor("#059669")
    .setRanges([enabledRange])
    .build();

  // Red for "No"
  const noRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("No")
    .setBackground("#FEE2E2")
    .setFontColor("#DC2626")
    .setRanges([enabledRange])
    .build();

  sheet.setConditionalFormatRules([yesRule, noRule]);

  sheet.setTabColor("#7C3AED");

  SpreadsheetApp.getUi().alert(
    '✅ Feature Controls Created',
    'The Feature Controls sheet has been created.\n\n' +
    'You can now enable/disable features by changing the "Enabled" column.\n\n' +
    'Refresh the page (Ctrl+R or Cmd+R) to see menu changes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Checks if a feature is enabled
 * @param {string} featureName - Name of the feature to check
 * @returns {boolean} - True if feature is enabled
 */
function isFeatureEnabled(featureName) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("⚙️ Feature Controls");

    if (!sheet) {
      // If Feature Controls sheet doesn't exist, all features are enabled by default
      return true;
    }

    const data = sheet.getDataRange().getValues();

    // Find the feature row
    for (let i = 3; i < data.length; i++) {  // Start from row 4 (index 3)
      if (data[i][0] === featureName) {
        return data[i][3] === "Yes";  // Column D (index 3) is Enabled
      }
    }

    // If feature not found, enable by default
    return true;
  } catch (e) {
    Logger.log('Error checking feature status: ' + e.message);
    return true;  // Enable by default on error
  }
}

/**
 * Checks if a feature is enabled by function name
 * @param {string} functionName - Function name to check
 * @returns {boolean} - True if feature is enabled
 */
function isFeatureEnabledByFunction(functionName) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("⚙️ Feature Controls");

    if (!sheet) {
      return true;
    }

    const data = sheet.getDataRange().getValues();

    // Find the feature row by function name
    for (let i = 3; i < data.length; i++) {
      if (data[i][5] === functionName) {  // Column F (index 5) is Function Name
        return data[i][3] === "Yes";
      }
    }

    return true;
  } catch (e) {
    Logger.log('Error checking feature by function: ' + e.message);
    return true;
  }
}

/**
 * Shows a dialog to manage feature toggles
 */
function showFeatureManagementDialog() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (!sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚙️ Feature Controls Not Found',
      'The Feature Controls sheet does not exist.\n\n' +
      'Would you like to create it now?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      createFeatureControlsSheet();
    }
    return;
  }

  // Navigate to Feature Controls sheet
  sheet.activate();

  SpreadsheetApp.getUi().alert(
    '⚙️ Feature Management',
    'Edit the "Enabled" column to enable or disable features.\n\n' +
    'Changes will take effect after you refresh the page (Ctrl+R or Cmd+R).\n\n' +
    'Feature Categories:\n' +
    '• View - Dashboard and display features\n' +
    '• Grievances - Grievance management tools\n' +
    '• Reports - Data import/export\n' +
    '• Security - Audit and security features\n' +
    '• Admin - Testing and data seeding\n' +
    '• Help - User assistance tools',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Enables all features
 */
function enableAllFeatures() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Feature Controls sheet not found. Please create it first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // Set all Enabled values to "Yes"
  const range = sheet.getRange(4, 4, lastRow - 3, 1);
  const values = [];
  for (let i = 0; i < lastRow - 3; i++) {
    values.push(["Yes"]);
  }
  range.setValues(values);

  SpreadsheetApp.getUi().alert(
    '✅ All Features Enabled',
    'All features have been enabled.\n\nRefresh the page to see changes in the menu.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Disables all non-essential features
 */
function disableNonEssentialFeatures() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Feature Controls sheet not found. Please create it first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();

  // Essential features that should always be enabled
  const essentialFeatures = [
    "Main Dashboard",
    "Interactive Dashboard",
    "Start New Grievance",
    "Advanced Search",
    "Help"
  ];

  // Update enabled status
  for (let i = 3; i < data.length; i++) {
    const featureName = data[i][0];
    const isEssential = essentialFeatures.includes(featureName);
    sheet.getRange(i + 1, 4).setValue(isEssential ? "Yes" : "No");
  }

  SpreadsheetApp.getUi().alert(
    '✅ Non-Essential Features Disabled',
    'Only essential features are now enabled:\n\n' +
    essentialFeatures.join('\n') + '\n\n' +
    'Refresh the page to see changes in the menu.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Exports feature configuration
 */
function exportFeatureConfig() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Feature Controls sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 3; i < data.length; i++) {
    const featureName = data[i][0];
    const enabled = data[i][3] === "Yes";
    config[featureName] = enabled;
  }

  // Save to Script Properties
  PropertiesService.getScriptProperties().setProperty('FEATURE_CONFIG', JSON.stringify(config));

  SpreadsheetApp.getUi().alert(
    '✅ Configuration Exported',
    'Feature configuration has been saved to Script Properties.\n\n' +
    'You can now restore this configuration later.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Imports feature configuration
 */
function importFeatureConfig() {
  const configJson = PropertiesService.getScriptProperties().getProperty('FEATURE_CONFIG');

  if (!configJson) {
    SpreadsheetApp.getUi().alert(
      '❌ No Saved Configuration',
      'No saved feature configuration found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const config = JSON.parse(configJson);
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("⚙️ Feature Controls");

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Error', 'Feature Controls sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 3; i < data.length; i++) {
    const featureName = data[i][0];
    if (config.hasOwnProperty(featureName)) {
      sheet.getRange(i + 1, 4).setValue(config[featureName] ? "Yes" : "No");
    }
  }

  SpreadsheetApp.getUi().alert(
    '✅ Configuration Imported',
    'Feature configuration has been restored.\n\nRefresh the page to see changes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
