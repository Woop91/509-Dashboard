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

  // Note: This feature expects specific grievance tracking columns (Total Grievances Filed,
  // Active Grievances, etc.) at columns 12-21 which are not present in the current Member
  // Directory structure. The current columns 12-21 contain engagement/contact data.

  SpreadsheetApp.getUi().alert(
    '⚠️ Feature Not Available',
    'This column toggle feature is designed for a different Member Directory structure.\n\n' +
    'The current Member Directory (31 columns) does not include the advanced grievance tracking columns this feature expects.\n\n' +
    'You can manually hide/show columns by right-clicking column headers.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return;

  /*
  // Legacy code - disabled because columns don't match current structure
  const firstGrievanceCol = 12; // Column L
  const lastGrievanceCol = 21;  // Column U
  const numCols = lastGrievanceCol - firstGrievanceCol + 1;

  const isHidden = sheet.isColumnHiddenByUser(firstGrievanceCol);

  if (isHidden) {
    sheet.showColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now visible');
  } else {
    sheet.hideColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now hidden');
  }
  */
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
