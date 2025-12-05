/**
 * ------------------------------------------------------------------------====
 * MEMBER DIRECTORY GOOGLE FORM LINK
 * ------------------------------------------------------------------------====
 *
 * Adds a button/checkbox feature to link to a Google Form
 * that auto-populates with member information
 *
 * Features:
 * - Button in Member Directory to open Google Form
 * - Auto-populates form with member data
 * - Can be wired to specific use cases later
 * - Added to Pending Features tab for future implementation
 */

// Configuration for the member info Google Form
MEMBER_FORM_CONFIG = {
  // Replace with your actual Google Form URL
  FORM_URL: "https://docs.google.com/forms/d/e/YOUR_MEMBER_FORM_ID/viewform",

  // Form field entry IDs (found by inspecting your form)
  FIELD_IDS: {
    MEMBER_ID: "entry.2000000001",
    FIRST_NAME: "entry.2000000002",
    LAST_NAME: "entry.2000000003",
    EMAIL: "entry.2000000004",
    PHONE: "entry.2000000005",
    JOB_TITLE: "entry.2000000006",
    WORK_LOCATION: "entry.2000000007",
    UNIT: "entry.2000000008",
    IS_STEWARD: "entry.2000000009",
    ASSIGNED_STEWARD: "entry.2000000010"
  }
};

/**
 * Shows a dialog with a link to Google Form for selected member
 */
function openMemberGoogleForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.MEMBER_DIR) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a member row in the Member Directory sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a member row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const memberData = getMemberDataFromRow(activeRow);

    if (!memberData.memberId) {
      ui.alert(
        '⚠️ No Member ID',
        'Selected row does not have a member ID.',
        ui.ButtonSet.OK
      );
      return;
    }

    const formUrl = generatePreFilledMemberForm(memberData);

    const html = createMemberFormDialog(memberData, formUrl);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(600)
      .setHeight(400);

    ui.showModalDialog(htmlOutput, '📋 Member Information Form');

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    Logger.log('Error opening member form: ' + error);
  }
}

/**
 * Gets member data from a specific row
 * Uses MEMBER_COLS constants for correct column mapping
 * @param {number} row - Row number
 * @returns {Object} Member data object
 */
function getMemberDataFromRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    throw new Error('Member Directory not found');
  }

  // Get member data from row - read enough columns to cover ASSIGNED_STEWARD (col 16)
  const data = memberSheet.getRange(row, 1, 1, MEMBER_COLS.ASSIGNED_STEWARD).getValues()[0];

  // Use MEMBER_COLS constants for correct column mapping (subtract 1 for 0-based array index)
  return {
    memberId: data[MEMBER_COLS.MEMBER_ID - 1],          // A: Member ID
    firstName: data[MEMBER_COLS.FIRST_NAME - 1],        // B: First Name
    lastName: data[MEMBER_COLS.LAST_NAME - 1],          // C: Last Name
    jobTitle: data[MEMBER_COLS.JOB_TITLE - 1],          // D: Job Title
    workLocation: data[MEMBER_COLS.WORK_LOCATION - 1],  // E: Work Location
    unit: data[MEMBER_COLS.UNIT - 1],                   // F: Unit
    officeDays: data[MEMBER_COLS.OFFICE_DAYS - 1],      // G: Office Days
    email: data[MEMBER_COLS.EMAIL - 1],                 // H: Email
    phone: data[MEMBER_COLS.PHONE - 1],                 // I: Phone
    isSteward: data[MEMBER_COLS.IS_STEWARD - 1],        // N: Is Steward
    supervisor: data[MEMBER_COLS.SUPERVISOR - 1],       // L: Supervisor
    manager: data[MEMBER_COLS.MANAGER - 1],             // M: Manager
    assignedSteward: data[MEMBER_COLS.ASSIGNED_STEWARD - 1]  // P: Assigned Steward
  };
}

/**
 * Generates a pre-filled Google Form URL for a member
 * @param {Object} memberData - Member data object
 * @returns {string} Pre-filled form URL
 */
function generatePreFilledMemberForm(memberData) {
  const baseUrl = MEMBER_FORM_CONFIG.FORM_URL;
  const params = new URLSearchParams();

  // Add member information to form URL
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.MEMBER_ID, memberData.memberId || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.FIRST_NAME, memberData.firstName || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.LAST_NAME, memberData.lastName || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.EMAIL, memberData.email || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.PHONE, memberData.phone || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.JOB_TITLE, memberData.jobTitle || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.WORK_LOCATION, memberData.workLocation || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.UNIT, memberData.unit || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.IS_STEWARD, memberData.isSteward || '');
  params.append(MEMBER_FORM_CONFIG.FIELD_IDS.ASSIGNED_STEWARD, memberData.assignedSteward || '');

  return baseUrl + '?' + params.toString();
}

/**
 * Creates HTML dialog for member form access
 * @param {Object} memberData - Member data
 * @param {string} formUrl - Pre-filled form URL
 * @returns {string} HTML content
 */
function createMemberFormDialog(memberData, formUrl) {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Roboto', Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f5f5f5;
    }
    .container {
      background: white;
      padding: 25px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .info-box strong {
      color: #1557b0;
    }
    .member-info {
      background: #fafafa;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
    }
    .member-info-item {
      margin: 8px 0;
      font-size: 14px;
    }
    .button {
      background: #1a73e8;
      color: white;
      padding: 12px 24px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 16px;
      font-weight: bold;
      text-decoration: none;
      display: inline-block;
      margin: 10px 5px;
    }
    .button:hover {
      background: #1557b0;
    }
    .button.secondary {
      background: #6b7280;
    }
    .button.secondary:hover {
      background: #4b5563;
    }
    .note {
      background: #fff3cd;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #ffc107;
      font-size: 13px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📋 Member Information Form</h2>

    <div class="info-box">
      <strong>ℹ️ This feature is in development</strong><br>
      This button will open a Google Form with auto-populated member information.
      The form can be used for various purposes such as surveys, updates, or data collection.
    </div>

    <div class="member-info">
      <h3>Member Information:</h3>
      <div class="member-info-item"><strong>Name:</strong> ${memberData.firstName} ${memberData.lastName}</div>
      <div class="member-info-item"><strong>Member ID:</strong> ${memberData.memberId}</div>
      <div class="member-info-item"><strong>Email:</strong> ${memberData.email || 'Not provided'}</div>
      <div class="member-info-item"><strong>Work Location:</strong> ${memberData.workLocation || 'Not provided'}</div>
      <div class="member-info-item"><strong>Unit:</strong> ${memberData.unit || 'Not provided'}</div>
    </div>

    <div class="note">
      <strong>📝 Note:</strong> Before using this feature, you'll need to set up a Google Form and configure the MEMBER_FORM_CONFIG in the script.
      See the MemberDirectoryGoogleFormLink.gs file for configuration instructions.
    </div>

    <a href="${formUrl}" target="_blank" class="button">Open Form with Member Data</a>
    <button class="button secondary" onclick="google.script.host.close()">Close</button>
  </div>
</body>
</html>
  `;
}

/**
 * Adds a "pending feature" entry for the Member Google Form Link
 * This adds an entry to the Feedback & Development sheet
 */
function addMemberFormLinkToPendingFeatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feedbackSheet = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!feedbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Feedback sheet not found!');
    return;
  }

  const today = new Date();
  const featureData = [
    'Future Feature',                          // Type
    today,                                     // Submitted/Started
    'System',                                  // Submitted By
    'Medium',                                  // Priority
    'Member Directory Google Form Link',       // Title
    'Add button/checkbox to Member Directory that opens a Google Form with auto-populated member information. ' +
    'Form can be used for surveys, data updates, or other member interactions. ' +
    'Requires Google Form setup and field ID configuration.',  // Description
    'Planned',                                 // Status
    '30%',                                     // Progress %
    'Moderate',                                // Complexity
    '',                                        // Target Completion
    '',                                        // Assigned To
    'Requires Google Form creation and configuration',  // Blockers
    'Feature placeholder created. Script functions ready. Needs form URL and field configuration.',  // Resolution/Notes
    today                                      // Last Updated
  ];

  feedbackSheet.appendRow(featureData);

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Added to Pending Features!', '', 3);
}

/**
 * Adds the Grievance Float Toggle feature to pending features
 */
function addGrievanceFloatToPendingFeatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feedbackSheet = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!feedbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Feedback sheet not found!');
    return;
  }

  const today = new Date();
  const featureData = [
    'Future Feature',                          // Type
    today,                                     // Submitted/Started
    'System',                                  // Submitted By
    'High',                                    // Priority
    'Grievance Float/Sort Toggle',             // Title
    'Toggle feature that sends closed/settled/inactive grievances to bottom. ' +
    'Prioritizes Step 3 > Step 2 > Step 1, then by soonest due date within each step. ' +
    'Helps users focus on active, high-priority grievances.',  // Description
    'Completed',                               // Status
    '100%',                                    // Progress %
    'Moderate',                                // Complexity
    today,                                     // Target Completion
    'System',                                  // Assigned To
    '',                                        // Blockers
    'Feature implemented in GrievanceFloatToggle.gs. Available via menu or control panel.',  // Resolution/Notes
    today                                      // Last Updated
  ];

  feedbackSheet.appendRow(featureData);

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Added to Pending Features!', '', 3);
}
