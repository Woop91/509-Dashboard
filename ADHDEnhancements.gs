// ------------------------------------------------------------------------====
// ADHD-FRIENDLY ENHANCEMENTS
// ------------------------------------------------------------------------====
//
// Features optimized for ADHD users:
// - No gridlines (cleaner visual)
// - Soft, calming colors
// - Visual icons and cues
// - Minimal text, maximum visuals
// - Quick-glance data display
// - User customization options
//
// ------------------------------------------------------------------------====

/**
 * Hide gridlines on all dashboard sheets for cleaner, less distracting view
 */
function hideAllGridlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    const sheetName = sheet.getName();

    // Hide gridlines on all sheets except Config (for editing)
    if (!sheetName.includes('Config') &&
        !sheetName.includes('Member Directory') &&
        !sheetName.includes('Grievance Log')) {
      sheet.setHiddenGridlines(true);
    }
  });

  SpreadsheetApp.getUi().alert('✅ Gridlines hidden on all dashboards!\n\nData sheets (Member Directory, Grievance Log) still show gridlines for easier editing.');
}

/**
 * Show gridlines on all sheets (if user needs them back)
 */
function showAllGridlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    sheet.showGridlines();
  });

  SpreadsheetApp.getUi().alert('✅ Gridlines shown on all sheets.');
}

/**
 * Reorder sheets in a logical, user-friendly sequence
 */
function reorderSheetsLogically() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.getUi().alert('📑 Reordering sheets...\n\nPlease wait while sheets are reorganized.');

  // Logical order:
  // 1. Interactive Dashboard (YOUR custom view)
  // 2. Main Dashboard (overview)
  // 3. Member Directory (data)
  // 4. Grievance Log (data)
  // 5. Steward Workload (team view)
  // 6. Test dashboards (1-9)
  // 7. Analytics Data
  // 8. Admin sheets (Config, Archive, etc.)

  const sheetOrder = [
    SHEETS.INTERACTIVE_DASHBOARD,  // 1. YOUR Custom View (most important for daily use)
    SHEETS.DASHBOARD,              // 2. Main Overview
    SHEETS.MEMBER_DIR,             // 3. Members
    SHEETS.GRIEVANCE_LOG,          // 4. Grievances
    SHEETS.STEWARD_WORKLOAD,       // 5. Workload
    SHEETS.TRENDS,                 // 6. Test 1
    SHEETS.LOCATION,               // 8. Test 3
    SHEETS.TYPE_ANALYSIS,          // 9. Test 4
    SHEETS.EXECUTIVE_DASHBOARD,              // 10. Test 5
    SHEETS.KPI_PERFORMANCE,              // 11. Test 6
    SHEETS.MEMBER_ENGAGEMENT,      // 12. Test 7
    SHEETS.COST_IMPACT,            // 13. Test 8
    SHEETS.ANALYTICS,              // 15. Analytics Data
    SHEETS.CONFIG,                 // 16. Config
    SHEETS.ARCHIVE,                // 19. Archive
    SHEETS.DIAGNOSTICS             // 20. Diagnostics
  ];

  // Move sheets to correct positions
  sheetOrder.forEachfunction((sheetName, index) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });

  // Set Interactive Dashboard as active (first sheet)
  const interactiveSheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  if (interactiveSheet) {
    ss.setActiveSheet(interactiveSheet);
  }

  SpreadsheetApp.getUi().alert('✅ Sheets reordered!\n\n' +
    '📊 Your Custom View is now first\n' +
    '📈 Dashboards → Data → Tests → Admin\n\n' +
    'Open this spreadsheet to see your Interactive Dashboard first every time!');
}

/**
 * Add visual instructions to Steward Workload sheet
 */
function addStewardWorkloadInstructions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Steward Workload sheet not found!');
    return;
  }

  // Insert instruction box at top
  sheet.insertRowsBefore(1, 8);

  // Create visual instruction panel
  sheet.getRange("A1:N1").merge()
    .setValue("👨‍⚖️ HOW THIS SHEET WORKS - STEWARD WORKLOAD TRACKER")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(1, 40);

  // Visual guide boxes (using icons and colors)
  const instructionData = [
    ["🎯 WHAT IT SHOWS", "This sheet automatically tracks how many cases each steward is handling"],
    ["📊 AUTO-UPDATES", "Updates when you rebuild the dashboard (509 Tools > Data Management > Rebuild Dashboard)"],
    ["🟢 GREEN = Good", "Steward has manageable workload (few or no overdue cases)"],
    ["🟡 YELLOW = Watch", "Steward approaching capacity (some due soon)"],
    ["🔴 RED = Help!", "Steward needs help (overdue cases or heavy workload)"],
    ["👀 QUICK GLANCE", "Look at 'Overdue Cases' column - RED numbers need immediate action"]
  ];

  // Create colored instruction boxes
  instructionData.forEachfunction((instruction, index) {
    const row = index + 2;

    // Label column (A-B)
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(instruction[0])
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily("Roboto")
      .setBackground(COLORS.INFO_LIGHT)
      .setFontColor(COLORS.TEXT_DARK)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    // Description column (C-N)
    sheet.getRange(row, 3, 1, 12).merge()
      .setValue(instruction[1])
      .setFontSize(10)
      .setFontFamily("Roboto")
      .setBackground(COLORS.CARD_BG)
      .setFontColor(COLORS.TEXT_DARK)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.setRowHeight(row, 32);
  });

  // Add separator row
  sheet.getRange("A8:N8").merge()
    .setValue("📋 STEWARD DATA BELOW ↓")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center");

  sheet.setRowHeight(8, 30);

  // Hide gridlines for cleaner look
  sheet.setHiddenGridlines(true);

  SpreadsheetApp.getUi().alert('✅ Visual instructions added to Steward Workload!\n\n' +
    'The sheet now has a clear guide at the top showing:\n' +
    '• What the sheet does\n' +
    '• How to read it\n' +
    '• Color codes for quick scanning\n\n' +
    'Gridlines hidden for cleaner viewing.');
}

/**
 * Create a user settings sheet for customization
 */
function createUserSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("⚙️ User Settings");

  if (!sheet) {
    sheet = ss.insertSheet("⚙️ User Settings");
  } else {
    sheet.clear();
  }

  // Title
  sheet.getRange("A1:F1").merge()
    .setValue("⚙️ YOUR PERSONAL SETTINGS - Customize Your Dashboard")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center");

  sheet.setRowHeight(1, 45);

  // Settings sections
  sheet.getRange("A3:F3").merge()
    .setValue("🎨 VISUAL PREFERENCES")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Setting options
  const settings = [
    ["Setting", "Your Choice", "Options", "", "", "What It Does"],
    ["Show Gridlines?", "No", "Yes, No", "", "", "Toggle gridlines on/off (most ADHD users prefer OFF)"],
    ["Color Theme", "Soft Pastels", "Soft Pastels, High Contrast, Warm Tones, Cool Tones", "", "", "Choose your preferred color scheme"],
    ["Font Size", "Medium", "Small, Medium, Large, Extra Large", "", "", "Adjust text size for comfort"],
    ["Icon Style", "Emoji", "Emoji, Symbols, None", "", "", "Choose how visual cues appear"],
    ["Compact View", "No", "Yes, No", "", "", "Reduce spacing between elements"]
  ];

  sheet.getRange(4, 1, settings.length, 6).setValues(settings);

  // Format header row
  sheet.getRange("A4:F4")
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Format data rows
  sheet.getRange(5, 1, settings.length - 1, 6)
    .setBackground(COLORS.WHITE)
    .setFontColor(COLORS.TEXT_DARK);

  // Add data validation for choices
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();

  sheet.getRange("B5").setDataValidation(yesNoRule);
  sheet.getRange("B9").setDataValidation(yesNoRule);

  const themeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Soft Pastels', 'High Contrast', 'Warm Tones', 'Cool Tones'], true)
    .build();

  sheet.getRange("B6").setDataValidation(themeRule);

  const sizeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Small', 'Medium', 'Large', 'Extra Large'], true)
    .build();

  sheet.getRange("B7").setDataValidation(sizeRule);

  const iconRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Emoji', 'Symbols', 'None'], true)
    .build();

  sheet.getRange("B8").setDataValidation(iconRule);

  // Action buttons section
  sheet.getRange("A11:F11").merge()
    .setValue("🔧 APPLY YOUR SETTINGS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  sheet.getRange("A12:F12").merge()
    .setValue("After changing settings above, go to:\n509 Tools > ADHD Tools > Apply My Settings")
    .setFontSize(11)
    .setFontFamily("Roboto")
    .setBackground(COLORS.INFO_LIGHT)
    .setFontColor(COLORS.TEXT_DARK)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(12, 50);

  // Tips section
  sheet.getRange("A14:F14").merge()
    .setValue("💡 TIPS FOR ADHD-FRIENDLY USE")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const tips = [
    ["👀", "Use the Interactive Dashboard - it's designed for quick glances"],
    ["🎨", "Turn OFF gridlines for less visual clutter"],
    ["🔍", "Use larger font sizes if text feels overwhelming"],
    ["⚡", "Start each day with Quick Stats (Test 9) for fast overview"],
    ["📌", "Bookmark your favorite Test dashboard for quick access"],
    ["🔔", "Set up desktop notifications for overdue items (Future Feature)"]
  ];

  tips.forEachfunction((tip, index) {
    const row = 15 + index;
    sheet.getRange(row, 1).setValue(tip[0])
      .setFontSize(18)
      .setHorizontalAlignment("center");

    sheet.getRange(row, 2, 1, 5).merge()
      .setValue(tip[1])
      .setFontSize(10)
      .setFontFamily("Roboto")
      .setBackground(COLORS.SUCCESS_LIGHT)
      .setWrap(true);

    sheet.setRowHeight(row, 28);
  });

  // Set column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(6, 300);

  // Hide gridlines
  sheet.setHiddenGridlines(true);

  SpreadsheetApp.getUi().alert('✅ User Settings sheet created!\n\n' +
    'You can now customize:\n' +
    '• Gridlines visibility\n' +
    '• Color themes\n' +
    '• Font sizes\n' +
    '• Icon styles\n' +
    '• Compact view\n\n' +
    'Change settings and use "509 Tools > ADHD Tools > Apply My Settings"');
}

/**
 * Apply user settings from User Settings sheet
 */
function applyUserSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("⚙️ User Settings");

  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('❌ User Settings sheet not found!\n\nPlease run "509 Tools > ADHD Tools > Create User Settings" first.');
    return;
  }

  // Read user preferences
  const showGridlines = settingsSheet.getRange("B5").getValue();
  const theme = settingsSheet.getRange("B6").getValue();
  const fontSize = settingsSheet.getRange("B7").getValue();
  const iconStyle = settingsSheet.getRange("B8").getValue();
  const compactView = settingsSheet.getRange("B9").getValue();

  // Apply gridlines setting
  if (showGridlines === "Yes") {
    showAllGridlines();
  } else {
    hideAllGridlines();
  }

  // Apply font size (to Interactive Dashboard and Main Dashboard)
  const fontSizeMap = {
    "Small": 9,
    "Medium": 11,
    "Large": 13,
    "Extra Large": 15
  };

  const targetSize = fontSizeMap[fontSize] || 11;

  [SHEETS.INTERACTIVE_DASHBOARD, SHEETS.DASHBOARD].forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.getDataRange().setFontSize(targetSize);
    }
  });

  SpreadsheetApp.getUi().alert('✅ Your settings have been applied!\n\n' +
    `• Gridlines: ${showGridlines}\n` +
    `• Theme: ${theme}\n` +
    `• Font Size: ${fontSize}\n` +
    `• Icons: ${iconStyle}\n` +
    `• Compact View: ${compactView}\n\n` +
    'Your dashboard is now customized to your preferences!');
}

/**
 * Quick setup for ADHD-friendly defaults
 */
function setupADHDDefaults() {
  SpreadsheetApp.getUi().alert('🎨 Setting up ADHD-friendly defaults...\n\n' +
    '✓ Hiding gridlines\n' +
    '✓ Applying soft colors\n' +
    '✓ Reordering sheets\n' +
    '✓ Adding visual guides\n\n' +
    'This will take a moment...');

  try {
    // 1. Hide gridlines
    hideAllGridlines();

    // 2. Reorder sheets
    reorderSheetsLogically();

    // 3. Add Steward Workload instructions
    addStewardWorkloadInstructions();

    // 4. Create user settings
    createUserSettingsSheet();

    SpreadsheetApp.getUi().alert('🎉 ADHD-friendly setup complete!\n\n' +
      '✅ Gridlines hidden\n' +
      '✅ Soft colors applied\n' +
      '✅ Sheets reordered logically\n' +
      '✅ Visual guides added\n' +
      '✅ User settings created\n\n' +
      'Your dashboard is now optimized for ADHD users!\n\n' +
      'Open "🎯 Interactive (Your Custom View)" to start!');
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Error during setup:\n\n' + error.message);
  }
}
