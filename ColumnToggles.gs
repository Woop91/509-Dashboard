/**
 * ============================================================================
 * COLUMN VISIBILITY TOGGLES
 * ============================================================================
 *
 * Functions to show/hide column groups in Member Directory
 * - Grievance columns toggle
 * - Level 2 columns toggle
 *
 * ============================================================================
 */

/**
 * Toggles visibility of grievance columns in Member Directory
 */
function toggleGrievanceColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  // Grievance columns are L through U (columns 12-21)
  // L: Total Grievances Filed
  // M: Active Grievances
  // N: Resolved Grievances
  // O: Grievances Won
  // P: Grievances Lost
  // Q: Last Grievance Date
  // R: Has Open Grievance?
  // S: # Open Grievances
  // T: Last Grievance Status
  // U: Next Deadline (Soonest)

  const firstGrievanceCol = 12; // Column L
  const lastGrievanceCol = 21;  // Column U
  const numCols = lastGrievanceCol - firstGrievanceCol + 1;

  // Check if columns are currently hidden
  const isHidden = sheet.isColumnHiddenByUser(firstGrievanceCol);

  if (isHidden) {
    // Show columns
    sheet.showColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now visible');
  } else {
    // Hide columns
    sheet.hideColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now hidden');
  }
}

/**
 * Toggles visibility of Level 2 (engagement tracking) columns in Member Directory
 */
function toggleLevel2Columns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  // Level 2 columns start after the basic columns
  // These will be added at the end of the existing columns
  // Starting at column AH (34) based on current structure
  const firstLevel2Col = 34;

  // Count: 14 Level 2 columns (AH through AU)
  const numCols = 14;

  // Check if the Level 2 columns exist
  const lastCol = sheet.getLastColumn();
  if (lastCol < firstLevel2Col) {
    SpreadsheetApp.getUi().alert('❌ Level 2 columns have not been added yet.\n\nPlease run "Create Dashboard" to add Level 2 columns.');
    return;
  }

  // Check if columns are currently hidden
  const isHidden = sheet.isColumnHiddenByUser(firstLevel2Col);

  if (isHidden) {
    // Show columns
    sheet.showColumns(firstLevel2Col, numCols);
    SpreadsheetApp.getUi().alert('✅ Level 2 columns are now visible');
  } else {
    // Hide columns
    sheet.hideColumns(firstLevel2Col, numCols);
    SpreadsheetApp.getUi().alert('✅ Level 2 columns are now hidden');
  }
}

/**
 * Shows all hidden columns in Member Directory
 */
function showAllMemberColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  const lastCol = sheet.getLastColumn();

  // Show all columns
  for (let i = 1; i <= lastCol; i++) {
    if (sheet.isColumnHiddenByUser(i)) {
      sheet.showColumns(i);
    }
  }

  SpreadsheetApp.getUi().alert('✅ All columns are now visible');
}
