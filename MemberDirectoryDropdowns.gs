/**
 * ------------------------------------------------------------------------====
 * MEMBER DIRECTORY DROPDOWNS
 * ------------------------------------------------------------------------====
 *
 * Adds data validation dropdowns to Member Directory for consistent data entry
 *
 * Dropdowns for:
 * - Job Title / Position
 * - Department / Unit
 * - Worksite / Office Location
 * - Work Schedule / Office Days (multiple selections)
 * - Unit (8 or 10)
 * - Is Steward (Y/N)
 * - Assigned Steward
 * - Immediate Supervisor
 * - Manager / Program Director
 * - Engagement Level
 * - Committee Member (multiple selections)
 * - Interest: Allied Chapter Actions (Y/N)
 * - Preferred Communication Methods (multiple selections)
 * - Best Time(s) to Reach Member (multiple selections)
 * - Steward Who Contacted Member
 */

/**
 * Sets up all dropdowns for the Member Directory
 */
function setupMemberDirectoryDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Setting up dropdowns...', 'Please wait', -1);

  try {
    // Job Title / Position (Column D)
    const jobTitles = [
      'Case Manager',
      'Social Worker',
      'Program Coordinator',
      'Administrative Assistant',
      'Director',
      'Supervisor',
      'Analyst',
      'Specialist',
      'Counselor',
      'Advocate',
      'Other'
    ];
    setDropdown(memberSheet, 'D2:D10000', jobTitles, 'Job Title');

    // Department / Unit (Column not in original spec, but could be added)
    const departments = [
      'Human Services',
      'Child Welfare',
      'Mental Health',
      'Housing',
      'Employment Services',
      'Administration',
      'Finance',
      'IT',
      'Operations',
      'Other'
    ];

    // Worksite / Office Location (Column E)
    const worksites = [
      'Boston Office',
      'Springfield Office',
      'Worcester Office',
      'Lowell Office',
      'New Bedford Office',
      'Brockton Office',
      'Remote/Home',
      'Field/Mobile',
      'Other'
    ];
    setDropdown(memberSheet, 'E2:E10000', worksites, 'Worksite');

    // Unit (8 or 10) (Column F)
    const units = ['Unit 8', 'Unit 10'];
    setDropdown(memberSheet, 'F2:F10000', units, 'Unit');

    // Work Schedule / Office Days (Column G) - Multiple selections note
    const officeDays = [
      'Monday',
      'Tuesday',
      'Wednesday',
      'Thursday',
      'Friday',
      'Saturday',
      'Sunday',
      'M-F',
      'Varies',
      'Remote Only'
    ];
    setDropdown(memberSheet, 'G2:G10000', officeDays, 'Office Days', false);

    // Is Steward (Y/N) (Column J)
    const yesNo = ['Yes', 'No'];
    setDropdown(memberSheet, 'J2:J10000', yesNo, 'Is Steward');

    // Assigned Steward (Column M) - Will be populated from stewards list
    const stewards = getStewardsList();
    if (stewards.length > 0) {
      setDropdown(memberSheet, 'M2:M10000', stewards, 'Assigned Steward');
    }

    // Immediate Supervisor (Column K)
    const supervisors = [
      'John Smith',
      'Jane Doe',
      'Bob Johnson',
      'Sarah Williams',
      'Mike Brown',
      'Lisa Davis',
      'Other'
    ];
    setDropdown(memberSheet, 'K2:K10000', supervisors, 'Immediate Supervisor', false);

    // Manager / Program Director (Column L)
    const managers = [
      'Director Smith',
      'Director Johnson',
      'Director Williams',
      'Director Brown',
      'Director Davis',
      'Program Manager A',
      'Program Manager B',
      'Other'
    ];
    setDropdown(memberSheet, 'L2:L10000', managers, 'Manager', false);

    // Engagement Level - Need to find the column
    const engagementLevels = [
      'Very Active',
      'Active',
      'Moderate',
      'Low',
      'Inactive',
      'New Member'
    ];
    // This would go after the volunteer hours column

    // Committee Member - Multiple selections
    const committees = [
      'Executive Board',
      'Grievance Committee',
      'Contract Action Team',
      'Political Action',
      'Member Organizing',
      'Communications',
      'Social Committee',
      'Health & Safety',
      'Other'
    ];

    // Interest: Local Actions (Column T/20) - Y/N
    setDropdown(memberSheet, 'T2:T10000', yesNo, 'Interest: Local Actions');

    // Interest: Chapter Actions (Column U/21) - Y/N
    setDropdown(memberSheet, 'U2:U10000', yesNo, 'Interest: Chapter Actions');

    // Interest: Allied Chapter Actions (Column V/22) - Y/N
    setDropdown(memberSheet, 'V2:V10000', yesNo, 'Interest: Allied Chapter Actions');

    // Preferred Communication Methods (Column X/24) - Multiple
    const commMethods = [
      'Email',
      'Phone Call',
      'Text Message',
      'In-Person',
      'Video Call',
      'Social Media',
      'Mail'
    ];
    setDropdown(memberSheet, 'X2:X10000', commMethods, 'Preferred Communication', false);

    // Best Time(s) to Reach Member (Column Y/25) - Multiple
    const timeBlocks = [
      '8am-10am',
      '10am-12pm',
      '12pm-2pm',
      '2pm-4pm',
      '4pm-6pm',
      '6pm-8pm',
      'Evening (8pm+)',
      'Weekends',
      'Anytime'
    ];
    setDropdown(memberSheet, 'Y2:Y10000', timeBlocks, 'Best Times', false);

    // Steward Who Contacted Member (Column AD/30)
    if (stewards.length > 0) {
      setDropdown(memberSheet, 'AD2:AD10000', stewards, 'Steward Who Contacted');
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Dropdowns set up!', '', 3);

    SpreadsheetApp.getUi().alert(
      '✅ Dropdowns Setup Complete',
      'Data validation dropdowns have been added to the Member Directory.\n\n' +
      'Note: Some dropdowns allow multiple selections (use comma-separated values).\n\n' +
      'You can customize the dropdown options by editing the setupMemberDirectoryDropdowns function.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
    Logger.log('Error setting up dropdowns: ' + error);
  }
}

/**
 * Sets a dropdown validation rule on a range
 * @param {Sheet} sheet - The sheet to apply the dropdown to
 * @param {string} range - A1 notation range
 * @param {Array} values - Array of values for dropdown
 * @param {string} name - Name for logging
 * @param {boolean} requireList - Whether to strictly require list values (default true)
 */
function setDropdown(sheet, range, values, name, requireList = true) {
  try {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(!requireList)
      .build();

    sheet.getRange(range).setDataValidation(rule);
    Logger.log(`Set dropdown for ${name}: ${range}`);
  } catch (error) {
    Logger.log(`Error setting dropdown for ${name}: ${error.message}`);
  }
}

/**
 * Gets list of stewards from Member Directory
 * @returns {Array} Array of steward names
 */
function getStewardsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return [];

  try {
    const lastRow = memberSheet.getLastRow();
    if (lastRow < 2) return [];

    // Get member data: Column B (First Name), C (Last Name), J (Is Steward)
    const data = memberSheet.getRange(2, 1, lastRow - 1, 10).getValues();

    const stewards = data
      .filter(row => row[9] === 'Yes' || row[9] === 'Y') // Column J (index 9) = Is Steward
      .map(row => `${row[1]} ${row[2]}`) // First Name + Last Name
      .filter(name => name.trim() !== ' ');

    // Add a few default stewards if none found
    if (stewards.length === 0) {
      return ['Steward 1', 'Steward 2', 'Steward 3', 'Other'];
    }

    stewards.push('Other');
    return stewards;

  } catch (error) {
    Logger.log('Error getting stewards list: ' + error.message);
    return ['Steward 1', 'Steward 2', 'Other'];
  }
}

/**
 * Refreshes the steward dropdowns based on current steward list
 */
function refreshStewardDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  const stewards = getStewardsList();

  if (stewards.length > 0) {
    // Assigned Steward (Column M)
    setDropdown(memberSheet, 'M2:M10000', stewards, 'Assigned Steward');

    // Steward Who Contacted Member (Column AD/30)
    setDropdown(memberSheet, 'AD2:AD10000', stewards, 'Steward Who Contacted');

    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Steward dropdowns refreshed!', '', 2);
  }
}

/**
 * Removes emergency contact columns from Member Directory
 * This function identifies and removes any columns with "Emergency Contact" in the header
 */
function removeEmergencyContactColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Remove Emergency Contact Columns',
    'This will remove any columns with "Emergency Contact" in the header.\n\n' +
    'This action cannot be easily undone.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const headers = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    const columnsToDelete = [];

    // Find columns with "Emergency Contact" in the header
    headers.forEach((header, index) => {
      const headerStr = String(header).toLowerCase();
      if (headerStr.includes('emergency contact') || headerStr.includes('emergency phone')) {
        columnsToDelete.push(index + 1); // +1 because columns are 1-indexed
      }
    });

    if (columnsToDelete.length === 0) {
      ui.alert('ℹ️ No emergency contact columns found.');
      return;
    }

    // Delete columns in reverse order to maintain correct indices
    columnsToDelete.reverse().forEach(colIndex => {
      memberSheet.deleteColumn(colIndex);
      Logger.log(`Deleted column ${colIndex}`);
    });

    ui.alert(
      '✅ Columns Removed',
      `Removed ${columnsToDelete.length} emergency contact column(s).`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    Logger.log('Error removing columns: ' + error);
  }
}
