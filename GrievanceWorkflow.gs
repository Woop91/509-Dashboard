/**
 * ------------------------------------------------------------------------====
 * GRIEVANCE WORKFLOW - Start Grievance from Member Directory
 * ------------------------------------------------------------------------====
 *
 * Allows stewards to click a member in the directory and start a grievance
 * with pre-filled member and steward information via Google Form.
 *
 * Features:
 * - Pre-filled Google Form with member and steward details
 * - Automatic submission to Grievance Log
 * - PDF generation with fillable grievance form
 * - Email to multiple addresses or download option
 *
 * ------------------------------------------------------------------------====
 */

/**
 * Configuration for grievance workflow and Google Form integration
 * @type {Object}
 */
GRIEVANCE_FORM_CONFIG = {
  // Replace with your actual Google Form URL
  // To find: Create a Google Form, then copy the URL from the browser
  FORM_URL: "https://docs.google.com/forms/d/e/YOUR_FORM_ID/viewform",

  // Email address for grievance notifications
  // Configure this with your union's actual grievance email
  GRIEVANCE_EMAIL: "grievances@seiu509.org",

  // Form field entry IDs (found by inspecting your form)
  // To find these: Open your form, right-click a field, inspect element, find "entry.XXXXXXX"
  FIELD_IDS: {
    MEMBER_ID: "entry.1000000001",
    MEMBER_FIRST_NAME: "entry.1000000002",
    MEMBER_LAST_NAME: "entry.1000000003",
    MEMBER_EMAIL: "entry.1000000004",
    MEMBER_PHONE: "entry.1000000005",
    MEMBER_JOB_TITLE: "entry.1000000006",
    MEMBER_LOCATION: "entry.1000000007",
    MEMBER_UNIT: "entry.1000000008",
    STEWARD_NAME: "entry.1000000009",
    STEWARD_EMAIL: "entry.1000000010",
    STEWARD_PHONE: "entry.1000000011"
  }
};

/**
 * Shows a dialog to start a grievance from the Member Directory
 */
function showStartGrievanceDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory not found!');
    return;
  }

  // Get all members for selection
  const members = getMemberList();

  if (members.length === 0) {
    SpreadsheetApp.getUi().alert('❌ No members found in the directory.');
    return;
  }

  const html = createMemberSelectionDialog(members);

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(500),
    '🚀 Start New Grievance'
  );
}

/**
 * Gets list of all members from Member Directory
 * @returns {Array<Object>} Array of member objects with properties
 */
function getMemberList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!memberSheet) {
      logWarning('getMemberList', 'Member Directory sheet not found');
      return [];
    }

    const lastRow = memberSheet.getLastRow();
    if (lastRow < 2) return [];

    // Get all member data - using MEMBER_COLS constants
    const numCols = Math.max(
      MEMBER_COLS.MEMBER_ID,
      MEMBER_COLS.FIRST_NAME,
      MEMBER_COLS.LAST_NAME,
      MEMBER_COLS.JOB_TITLE,
      MEMBER_COLS.WORK_LOCATION,
      MEMBER_COLS.UNIT,
      MEMBER_COLS.OFFICE_DAYS,
      MEMBER_COLS.EMAIL,
      MEMBER_COLS.PHONE,
      MEMBER_COLS.IS_STEWARD,
      MEMBER_COLS.SUPERVISOR
    );

    const data = memberSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    return data.map(function(row, index) { return ({
      rowIndex: index + 2,
      memberId: safeArrayGet(row, MEMBER_COLS.MEMBER_ID - 1, ''),
      firstName: safeArrayGet(row, MEMBER_COLS.FIRST_NAME - 1, ''),
      lastName: safeArrayGet(row, MEMBER_COLS.LAST_NAME - 1, ''),
      jobTitle: safeArrayGet(row, MEMBER_COLS.JOB_TITLE - 1, ''),
      location: safeArrayGet(row, MEMBER_COLS.WORK_LOCATION - 1, ''),
      unit: safeArrayGet(row, MEMBER_COLS.UNIT - 1, ''),
      officeDays: safeArrayGet(row, MEMBER_COLS.OFFICE_DAYS - 1, ''),
      email: safeArrayGet(row, MEMBER_COLS.EMAIL - 1, ''),
      phone: safeArrayGet(row, MEMBER_COLS.PHONE - 1, ''),
      isSteward: safeArrayGet(row, MEMBER_COLS.IS_STEWARD - 1, ''),
      supervisor: safeArrayGet(row, MEMBER_COLS.SUPERVISOR - 1, '')
    })).filter(function(member) { return member.memberId; }); // Filter out empty rows
  } catch (error) {
    return handleError(error, 'getMemberList', true, true) || [];
  }
}

/**
 * Creates HTML dialog for member selection
 * @param {Array<Object>} members - Array of member objects
 * @returns {string} HTML content for dialog
 */
function createMemberSelectionDialog(members) {
  // Use HTML escaping to prevent XSS
  const memberOptions = members.map(function(m) {
    `<option value="${escapeHtmlAttribute(m.rowIndex)}">${escapeHtml(m.lastName)}, ${escapeHtml(m.firstName)} (${escapeHtml(m.memberId)}) - ${escapeHtml(m.location)}</option>`
  ).join('');

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Roboto', Arial, sans-serif;
      padding: 20px;
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
    .form-group {
      margin-bottom: 20px;
    }
    label {
      display: block;
      font-weight: bold;
      margin-bottom: 8px;
      color: #333;
    }
    select, input {
      width: 100%;
      padding: 10px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    select:focus, input:focus {
      outline: none;
      border-color: #1a73e8;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .button-group {
      display: flex;
      gap: 10px;
      margin-top: 25px;
    }
    button {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 4px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(26,115,232,0.4);
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-secondary:hover {
      background: #e8eaed;
    }
    .member-details {
      background: #fafafa;
      padding: 12px;
      border-radius: 4px;
      margin-top: 10px;
      display: none;
    }
    .detail-item {
      margin: 5px 0;
      font-size: 13px;
    }
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📋 Start New Grievance</h2>

    <div class="info-box">
      <strong>ℹ️ Instructions:</strong><br>
      1. Select the member filing the grievance<br>
      2. Review the pre-filled information<br>
      3. Click "Start Grievance" to open the form with pre-filled data
    </div>

    <div class="form-group">
      <label for="memberSelect">Select Member:</label>
      <select id="memberSelect" onchange="showMemberDetails()">
        <option value="">-- Select a Member --</option>
        ${memberOptions}
      </select>
    </div>

    <div id="memberDetails" class="member-details"></div>

    <div class="button-group">
      <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
      <button class="btn-primary" id="startBtn" onclick="startGrievance()" disabled>Start Grievance</button>
    </div>

    <div class="loading" id="loading">
      <p>⏳ Generating pre-filled form...</p>
    </div>
  </div>

  <script>
    var members = ${JSON.stringify(members)};

    function showMemberDetails() {
      const select = document.getElementById('memberSelect');
      const detailsDiv = document.getElementById('memberDetails');
      const startBtn = document.getElementById('startBtn');

      if (select.value) {
        const member = members.find(function(m) { return m.rowIndex == select.value; });
        if (member) {
          detailsDiv.innerHTML = \`
            <div class="detail-item"><strong>Name:</strong> \${member.firstName} \${member.lastName}</div>
            <div class="detail-item"><strong>Member ID:</strong> \${member.memberId}</div>
            <div class="detail-item"><strong>Email:</strong> \${member.email || 'Not provided'}</div>
            <div class="detail-item"><strong>Phone:</strong> \${member.phone || 'Not provided'}</div>
            <div class="detail-item"><strong>Job Title:</strong> \${member.jobTitle}</div>
            <div class="detail-item"><strong>Location:</strong> \${member.location}</div>
            <div class="detail-item"><strong>Unit:</strong> \${member.unit}</div>
          \`;
          detailsDiv.style.display = 'block';
          startBtn.disabled = false;
        }
      } else {
        detailsDiv.style.display = 'none';
        startBtn.disabled = true;
      }
    }

    function startGrievance() {
      const select = document.getElementById('memberSelect');
      if (!select.value) {
        alert('Please select a member first.');
        return;
      }

      const rowIndex = parseInt(select.value);
      document.getElementById('loading').style.display = 'block';

      google.script.run
        .withSuccessHandler(onFormUrlGenerated)
        .withFailureHandler(onError)
        .generatePreFilledGrievanceForm(rowIndex);
    }

    function onFormUrlGenerated(url) {
      document.getElementById('loading').style.display = 'none';
      if (url) {
        window.open(url, '_blank');
        alert('✅ Pre-filled form opened in new window.\\n\\nAfter submitting the form, the grievance will be automatically added to the Grievance Log.');
        google.script.host.close();
      } else {
        alert('❌ Could not generate form URL. Please check the configuration.');
      }
    }

    function onError(error) {
      document.getElementById('loading').style.display = 'none';
      alert('❌ Error: ' + error.message);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Generates a pre-filled Google Form URL for the selected member
 */
function generatePreFilledGrievanceForm(memberRowIndex) {
  // Validate configuration before proceeding
  if (GRIEVANCE_FORM_CONFIG.FORM_URL.includes('YOUR_FORM_ID')) {
    throw new Error(
      'Grievance form is not configured yet. Please update GRIEVANCE_FORM_CONFIG.FORM_URL in GrievanceWorkflow.gs with your actual Google Form URL. ' +
      'Instructions: 1) Create a Google Form, 2) Get the form URL, 3) Replace the placeholder in the code.'
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    throw new Error('Member Directory not found');
  }

  // Get member data
  const memberData = memberSheet.getRange(memberRowIndex, 1, 1, 11).getValues()[0];
  const member = {
    id: memberData[0],
    firstName: memberData[1],
    lastName: memberData[2],
    jobTitle: memberData[3],
    location: memberData[4],
    unit: memberData[5],
    officeDays: memberData[6],
    email: memberData[7],
    phone: memberData[8]
  };

  // Get steward contact info from Config
  const steward = getStewardContactInfo();

  // Build pre-filled form URL
  const baseUrl = GRIEVANCE_FORM_CONFIG.FORM_URL;
  const params = new URLSearchParams();

  // Add member information
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_ID, member.id || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_FIRST_NAME, member.firstName || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_LAST_NAME, member.lastName || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_EMAIL, member.email || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_PHONE, member.phone || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_JOB_TITLE, member.jobTitle || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_LOCATION, member.location || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.MEMBER_UNIT, member.unit || '');

  // Add steward information
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.STEWARD_NAME, steward.name || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.STEWARD_EMAIL, steward.email || '');
  params.append(GRIEVANCE_FORM_CONFIG.FIELD_IDS.STEWARD_PHONE, steward.phone || '');

  return baseUrl + '?' + params.toString();
}

/**
 * Gets steward contact information from Config sheet
 */
function getStewardContactInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    return { name: '', email: '', phone: '', location: '' };
  }

  // Steward contact info is stored in Config sheet
  // Define column position (update if Config structure changes)
  const CONFIG_STEWARD_INFO_COL = 21;  // Column U - Steward contact details

  try {
    const stewardData = configSheet.getRange(2, CONFIG_STEWARD_INFO_COL, 3, 1).getValues();
    return {
      name: stewardData[0][0] || '',
      email: stewardData[1][0] || '',
      phone: stewardData[2][0] || '',
      location: stewardData[0][0] || '' // Can be added if needed
    };
  } catch (e) {
    Logger.log('Error getting steward info: ' + e.message);
    return { name: '', email: '', phone: '', location: '' };
  }
}

/**
 * Gets grievance coordinator emails from Config sheet
 */
function getGrievanceCoordinators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    return { coordinator1: '', coordinator2: '', coordinator3: '' };
  }

  try {
    // Grievance Coordinators are in columns P (16), Q (17), R (18)
    const coordinatorData = configSheet.getRange(2, 16, 10, 3).getValues();

    // Get unique coordinator names
    const coordinator1List = coordinatorData.map(function(row) { return row[0]).filter(String; });
    const coordinator2List = coordinatorData.map(function(row) { return row[1]).filter(String; });
    const coordinator3List = coordinatorData.map(function(row) { return row[2]).filter(String; });

    return {
      coordinator1: coordinator1List.length > 0 ? coordinator1List[0] : '',
      coordinator2: coordinator2List.length > 0 ? coordinator2List[0] : '',
      coordinator3: coordinator3List.length > 0 ? coordinator3List[0] : ''
    };
  } catch (e) {
    Logger.log('Error getting grievance coordinators: ' + e.message);
    return { coordinator1: '', coordinator2: '', coordinator3: '' };
  }
}

/**
 * Updates Config sheet to include Steward Contact Info section
 */
function addStewardContactInfoToConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('❌ Config sheet not found!');
    return;
  }

  // Add Steward Contact Info section starting at column U (21)
  const startCol = 21;

  // Section header
  configSheet.getRange(1, startCol, 1, 4).merge()
    .setValue("STEWARD CONTACT INFORMATION")
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontFamily("Roboto")
    .setBackground("#7EC8E3") // PRIMARY_BLUE
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Field labels
  const labels = [
    ["Steward Name:"],
    ["Steward Email:"],
    ["Steward Phone:"],
    ["Instructions:"]
  ];

  configSheet.getRange(2, startCol, labels.length, 1)
    .setValues(labels)
    .setFontWeight("bold")
    .setBackground("#E8F0FE")
    .setHorizontalAlignment("right");

  // Instruction text
  configSheet.getRange(5, startCol, 1, 4).merge()
    .setValue("Enter the primary steward contact info above. This will be used when starting grievances from the Member Directory.")
    .setFontStyle("italic")
    .setFontSize(10)
    .setWrap(true)
    .setBackground("#FFF9E6");

  // Auto-resize columns
  for (let i = 0; i < 4; i++) {
    configSheet.setColumnWidth(startCol + i, 180);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Steward Contact Info Section Added',
    'A new section has been added to the Config sheet.\n\n' +
    'Please enter your steward contact information in column U (rows 2-4).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * ------------------------------------------------------------------------====
 * GOOGLE FORM SUBMISSION HANDLING
 * ------------------------------------------------------------------------====
 */

/**
 * Handles form submissions from the Google Form
 * This should be set up as an installable trigger on form submission
 *
 * To set up:
 * 1. In the Apps Script editor, go to Triggers (clock icon)
 * 2. Add trigger: onGrievanceFormSubmit, From spreadsheet, On form submit
 */
function onGrievanceFormSubmit(e) {
  try {
    // Extract form responses
    const formData = extractFormData(e);

    // Add to Grievance Log
    const grievanceId = addGrievanceToLog(formData);

    // Create Google Drive folder for the grievance
    const folder = createGrievanceFolder(grievanceId, formData);

    // Generate PDF
    const pdfBlob = generateGrievancePDF(grievanceId);

    // Save PDF to folder
    if (folder) {
      folder.createFile(pdfBlob);
    }

    // Show sharing options dialog with folder link
    showSharingOptionsDialog(grievanceId, pdfBlob, folder);

  } catch (error) {
    Logger.log('Error in onGrievanceFormSubmit: ' + error.message);
    SpreadsheetApp.getUi().alert('❌ Error processing form submission: ' + error.message);
  }
}

/**
 * Creates a Google Drive folder for a grievance
 */
function createGrievanceFolder(grievanceId, formData) {
  try {
    // Get or create main Grievances folder
    const mainFolderName = 'SEIU 509 - Grievances';
    var mainFolder;

    const folders = DriveApp.getFoldersByName(mainFolderName);
    if (folders.hasNext()) {
      mainFolder = folders.next();
    } else {
      mainFolder = DriveApp.createFolder(mainFolderName);
    }

    // Create grievance-specific folder
    const folderName = `${grievanceId} - ${formData.lastName}, ${formData.firstName}`;
    const grievanceFolder = mainFolder.createFolder(folderName);

    // Store folder URL in Grievance Log (if there's a column for it)
    // This would need a column added to the Grievance Log sheet

    Logger.log(`Created folder: ${folderName}`);
    return grievanceFolder;

  } catch (error) {
    Logger.log('Error creating grievance folder: ' + error.message);
    return null;
  }
}

/**
 * Shows dialog with sharing options for grievance form and folder
 */
function showSharingOptionsDialog(grievanceId, pdfBlob, folder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Get grievance data
  const grievanceData = grievanceLog.getDataRange().getValues();
  var grievanceRow = grievanceData.find(function(row) { return row[0] === grievanceId; });

  if (!grievanceRow) {
    SpreadsheetApp.getUi().alert('❌ Grievance not found');
    return;
  }

  const memberId = grievanceRow[1];
  const memberEmail = grievanceRow[23];
  const stewardEmail = getStewardContactInfo().email;
  const coordinators = getGrievanceCoordinators();
  const folderUrl = folder ? folder.getUrl() : '';

  // TODO: Configure actual grievance email address in settings
  // This email address needs to be set up by the organization
  const grievanceEmail = GRIEVANCE_FORM_CONFIG.GRIEVANCE_EMAIL || 'grievances@seiu509.org';

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
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
    .folder-link {
      background: #fef7e0;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #f9ab00;
    }
    .folder-link a {
      color: #1a73e8;
      font-weight: bold;
      text-decoration: none;
    }
    .folder-link a:hover {
      text-decoration: underline;
    }
    .checkbox-group {
      margin: 20px 0;
    }
    .checkbox-item {
      display: flex;
      align-items: center;
      padding: 10px;
      margin: 8px 0;
      border: 2px solid #ddd;
      border-radius: 4px;
      cursor: pointer;
    }
    .checkbox-item:hover {
      background: #f0f0f0;
      border-color: #1a73e8;
    }
    .checkbox-item input[type="checkbox"] {
      margin-right: 10px;
      width: 20px;
      height: 20px;
      cursor: pointer;
    }
    .checkbox-item label {
      cursor: pointer;
      flex: 1;
      font-size: 14px;
    }
    .button-group {
      display: flex;
      gap: 10px;
      margin-top: 25px;
    }
    button {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 4px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s;
    }
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    .btn-primary:hover {
      background: #1557b0;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(26,115,232,0.4);
    }
    .btn-secondary {
      background: #f1f3f4;
      color: #333;
    }
    .btn-secondary:hover {
      background: #e8eaed;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>✅ Grievance ${grievanceId} Created</h2>

    <div class="info-box">
      <strong>ℹ️ Grievance folder created successfully!</strong><br>
      Select the recipients you want to share the grievance form and/or folder with.
    </div>

    ${folderUrl ? `
    <div class="folder-link">
      📁 <strong>Folder Link:</strong><br>
      <a href="${folderUrl}" target="_blank">${folderUrl}</a>
    </div>
    ` : ''}

    <h3>Select Recipients:</h3>
    <div class="checkbox-group">
      ${memberEmail ? `
      <div class="checkbox-item">
        <input type="checkbox" id="member" value="${memberEmail}">
        <label for="member">Member (${memberEmail})</label>
      </div>
      ` : ''}

      ${stewardEmail ? `
      <div class="checkbox-item">
        <input type="checkbox" id="steward" value="${stewardEmail}">
        <label for="steward">Steward (${stewardEmail})</label>
      </div>
      ` : ''}

      ${coordinators.coordinator1 ? `
      <div class="checkbox-item">
        <input type="checkbox" id="coordinator1" value="${coordinators.coordinator1}">
        <label for="coordinator1">Grievance Coordinator 1 (${coordinators.coordinator1})</label>
      </div>
      ` : ''}

      ${coordinators.coordinator2 ? `
      <div class="checkbox-item">
        <input type="checkbox" id="coordinator2" value="${coordinators.coordinator2}">
        <label for="coordinator2">Grievance Coordinator 2 (${coordinators.coordinator2})</label>
      </div>
      ` : ''}

      ${coordinators.coordinator3 ? `
      <div class="checkbox-item">
        <input type="checkbox" id="coordinator3" value="${coordinators.coordinator3}">
        <label for="coordinator3">Grievance Coordinator 3 (${coordinators.coordinator3})</label>
      </div>
      ` : ''}

      ${grievanceEmail ? `
      <div class="checkbox-item">
        <input type="checkbox" id="grievanceEmail" value="${grievanceEmail}">
        <label for="grievanceEmail">Grievance Email (${grievanceEmail})</label>
      </div>
      ` : ''}
    </div>

    <div class="button-group">
      <button class="btn-secondary" onclick="google.script.host.close()">Close</button>
      <button class="btn-primary" onclick="shareWithSelected()">Share with Selected</button>
    </div>
  </div>

  <script>
    function shareWithSelected() {
      const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
      const recipients = Array.from(checkboxes).map(function(cb) { return cb.value; });

      if (recipients.length === 0) {
        alert('Please select at least one recipient.');
        return;
      }

      const folderUrl = '${folderUrl}';

      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Sharing invitations sent successfully!');
          google.script.host.close();
        })
        .withFailureHandler(function(error) {
          alert('❌ Error: ' + error.message);
        })
        .shareGrievanceWithRecipients('${grievanceId}', recipients, folderUrl);
    }
  </script>
</body>
</html>
  `).setWidth(600).setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Share Grievance');
}

/**
 * Shares grievance folder and sends notification to selected recipients
 */
function shareGrievanceWithRecipients(grievanceId, recipients, folderUrl) {
  try {
    const pdfBlob = generateGrievancePDF(grievanceId);

    // Get folder by URL
    var folder = null;
    if (folderUrl) {
      const folderId = folderUrl.match(/[-\w]{25,}/);
      if (folderId) {
        folder = DriveApp.getFolderById(folderId[0]);
      }
    }

    // Share folder with each recipient
    // Note: Requires valid email addresses. Configure coordinator emails in Config sheet.
    if (folder && recipients && recipients.length > 0) {
      recipients.forEach(function(email) {
        try {
          // Validate email before sharing
          const validEmail = sanitizeEmail(email);
          if (validEmail) {
            folder.addEditor(validEmail);
            logInfo('shareGrievanceWithRecipients', `Shared folder with: ${validEmail}`);
          } else {
            logWarning('shareGrievanceWithRecipients', `Invalid email address: ${email}`);
          }
        } catch (e) {
          handleError(e, `shareGrievanceWithRecipients - sharing with ${email}`, false);
        }
      });
    }

    // Send email notification
    const subject = `SEIU Local 509 - Grievance ${grievanceId}`;
    const body = `A new grievance has been filed and requires your attention.\n\n` +
                 `Grievance ID: ${grievanceId}\n` +
                 `Folder Link: ${folderUrl}\n\n` +
                 `Please review the attached grievance form and folder for details.\n\n` +
                 `This is an automated message from the SEIU Local 509 Dashboard.`;

    // Send email notification with grievance PDF
    // Note: Requires valid email addresses and Gmail permissions
    recipients.forEach(function(email) {
      try {
        const validEmail = sanitizeEmail(email);
        if (validEmail) {
          GmailApp.sendEmail(validEmail, subject, body, {
            attachments: [pdfBlob],
            name: 'SEIU Local 509 Dashboard'
          });
          logInfo('shareGrievanceWithRecipients', `Sent email to: ${validEmail}`);
        } else {
          logWarning('shareGrievanceWithRecipients', `Skipped invalid email: ${email}`);
        }
      } catch (e) {
        handleError(e, `shareGrievanceWithRecipients - emailing ${email}`, false);
      }
    });

  } catch (error) {
    Logger.log('Error sharing grievance: ' + error.message);
    throw error;
  }
}

/**
 * Extracts and structures data from form submission
 */
function extractFormData(e) {
  const responses = e.namedValues;

  // Map form responses to grievance data structure
  // Adjust field names based on your actual form questions
  return {
    memberId: responses['Member ID'] ? responses['Member ID'][0] : '',
    firstName: responses['First Name'] ? responses['First Name'][0] : '',
    lastName: responses['Last Name'] ? responses['Last Name'][0] : '',
    email: responses['Email'] ? responses['Email'][0] : '',
    phone: responses['Phone'] ? responses['Phone'][0] : '',
    jobTitle: responses['Job Title'] ? responses['Job Title'][0] : '',
    location: responses['Work Location'] ? responses['Work Location'][0] : '',
    unit: responses['Unit'] ? responses['Unit'][0] : '',
    stewardName: responses['Steward Name'] ? responses['Steward Name'][0] : '',
    incidentDate: responses['Incident Date'] ? new Date(responses['Incident Date'][0]) : new Date(),
    grievanceType: responses['Grievance Type'] ? responses['Grievance Type'][0] : 'Other',
    description: responses['Description'] ? responses['Description'][0] : '',
    desiredResolution: responses['Desired Resolution'] ? responses['Desired Resolution'][0] : ''
  };
}

/**
 * Adds grievance data to the Grievance Log
 */
function addGrievanceToLog(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  // Generate grievance ID
  const grievanceId = generateUniqueGrievanceId();

  // Prepare row data (32 columns A-AF)
  const today = new Date();
  const incidentDate = formData.incidentDate || today;

  const rowData = [
    grievanceId,                    // A: Grievance ID
    formData.memberId,              // B: Member ID
    formData.firstName,             // C: First Name
    formData.lastName,              // D: Last Name
    "Draft",                        // E: Status
    "Informal",                     // F: Current Step
    incidentDate,                   // G: Incident Date
    '',                             // H: Filing Deadline (calculated)
    '',                             // I: Date Filed
    '',                             // J: Step I Decision Due
    '',                             // K: Step I Decision Date
    '',                             // L: Step I Outcome
    '',                             // M: Step II Appeal Deadline
    '',                             // N: Step II Filed Date
    '',                             // O: Step II Decision Due
    '',                             // P: Step II Decision Date
    '',                             // Q: Step II Outcome
    '',                             // R: Step III Appeal Deadline
    '',                             // S: Step III Filed Date
    '',                             // T: Step III Decision Date
    '',                             // U: Mediation Date
    '',                             // V: Arbitration Date
    '',                             // W: Final Outcome
    formData.grievanceType,         // X: Grievance Type
    formData.description,           // Y: Description
    formData.stewardName,           // Z: Steward Name
    '',                             // AA: Days Open (calculated)
    '',                             // AB: Next Action Due (calculated)
    '',                             // AC: Days to Deadline (calculated)
    '',                             // AD: Overdue? (calculated)
    formData.desiredResolution || '', // AE: Notes
    today                           // AF: Last Updated
  ];

  // Add to sheet
  grievanceSheet.appendRow(rowData);

  // Recalculate the new grievance row
  const lastRow = grievanceSheet.getLastRow();
  recalcGrievanceRow(lastRow);

  // Update member directory
  const memberRow = findMemberRow(formData.memberId);
  if (memberRow > 0) {
    recalcMemberRow(memberRow);
  }

  // Automatically create Google Drive folder for the grievance
  try {
    const grievantName = `${formData.firstName}_${formData.lastName}`;
    const folder = createGrievanceFolder(grievanceId, grievantName);
    linkFolderToGrievance(grievanceId, folder.getId());
    Logger.log(`Auto-created Drive folder for grievance ${grievanceId}`);
  } catch (error) {
    Logger.log(`Error auto-creating Drive folder for ${grievanceId}: ${error.message}`);
    // Continue even if folder creation fails
  }

  return grievanceId;
}

/**
 * Recalculates formulas for a specific grievance row
 * Sets deadline formulas (Filing Deadline, Step deadlines, Days Open, etc.)
 */
function recalcGrievanceRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceLog) return;

  // Set all deadline formulas for this row
  // Filing Deadline (Column H): Incident Date + 21 days
  grievanceLog.getRange(row, 8).setFormula(`=IF(G${row}<>"",G${row}+21,"")`);

  // Step I Decision Due (Column J): Date Filed + 30 days
  grievanceLog.getRange(row, 10).setFormula(`=IF(I${row}<>"",I${row}+30,"")`);

  // Step II Appeal Due (Column L): Step I Decision + 10 days
  grievanceLog.getRange(row, 12).setFormula(`=IF(K${row}<>"",K${row}+10,"")`);

  // Step II Decision Due (Column N): Step II Appeal + 30 days
  grievanceLog.getRange(row, 14).setFormula(`=IF(M${row}<>"",M${row}+30,"")`);

  // Step III Appeal Due (Column P): Step II Decision + 30 days
  grievanceLog.getRange(row, 16).setFormula(`=IF(O${row}<>"",O${row}+30,"")`);

  // Days Open (Column S): Today - Date Filed (or Date Closed - Date Filed)
  grievanceLog.getRange(row, 19).setFormula(
    `=IF(I${row}<>"",IF(R${row}<>"",R${row}-I${row},TODAY()-I${row}),"")`
  );

  // Next Action Due (Column T): Based on current step
  grievanceLog.getRange(row, 20).setFormula(
    `=IF(E${row}="Open",IF(F${row}="Step I",J${row},IF(F${row}="Step II",N${row},IF(F${row}="Step III",P${row},H${row}))),"")`
  );

  // Days to Deadline (Column U): Next Action - Today
  grievanceLog.getRange(row, 21).setFormula(`=IF(T${row}<>"",T${row}-TODAY(),"")`);
}

/**
 * Recalculates formulas for a specific member row
 * Updates grievance snapshot columns (Has Open Grievance, Status, Next Deadline)
 */
function recalcMemberRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberDir) return;

  // Dynamic column references
  const memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  const grievanceMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  const nextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);

  // Has Open Grievance? (Column Z / 26)
  memberDir.getRange(row, MEMBER_COLS.HAS_OPEN_GRIEVANCE).setFormula(
    `=IF(COUNTIFS('Grievance Log'!${grievanceMemberIdCol}:${grievanceMemberIdCol},${memberIdCol}${row},'Grievance Log'!${statusCol}:${statusCol},"Open")>0,"Yes","No")`
  );

  // Grievance Status Snapshot (Column AA / 27)
  memberDir.getRange(row, MEMBER_COLS.GRIEVANCE_STATUS).setFormula(
    `=IFERROR(INDEX('Grievance Log'!${statusCol}:${statusCol},MATCH(${memberIdCol}${row},'Grievance Log'!${grievanceMemberIdCol}:${grievanceMemberIdCol},0)),"")`
  );

  // Next Grievance Deadline (Column AB / 28)
  memberDir.getRange(row, MEMBER_COLS.NEXT_DEADLINE).setFormula(
    `=IFERROR(INDEX('Grievance Log'!${nextActionCol}:${nextActionCol},MATCH(${memberIdCol}${row},'Grievance Log'!${grievanceMemberIdCol}:${grievanceMemberIdCol},0)),"")`
  );
}

/**
 * Generates a unique grievance ID
 */
function generateUniqueGrievanceId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return 'GRV-001-A';

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return 'GRV-001-A';

  // Get existing IDs
  const existingIds = grievanceSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  // Generate new ID
  var counter = 1;
  var letter = 'A';
  var newId;

  do {
    newId = `GRV-${String(counter).padStart(3, '0')}-${letter}`;
    if (!existingIds.includes(newId)) break;

    // Increment letter
    if (letter === 'Z') {
      letter = 'A';
      counter++;
    } else {
      letter = String.fromCharCode(letter.charCodeAt(0) + 1);
    }
  } while (existingIds.includes(newId));

  return newId;
}

/**
 * Finds member row by member ID
 */
function findMemberRow(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return -1;

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return -1;

  const memberIds = memberSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const index = memberIds.indexOf(memberId);

  return index >= 0 ? index + 2 : -1;
}

/**
 * ------------------------------------------------------------------------====
 * PDF GENERATION AND DISTRIBUTION
 * ------------------------------------------------------------------------====
 */

/**
 * Generates a PDF for a grievance
 */
function generateGrievancePDF(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log not found');
  }

  // Find grievance row
  const data = grievanceSheet.getDataRange().getValues();
  var grievanceRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      grievanceRow = i;
      break;
    }
  }

  if (grievanceRow === -1) {
    throw new Error('Grievance not found');
  }

  const grievanceData = data[grievanceRow];

  // Create temporary sheet for PDF
  const tempSheet = ss.insertSheet('Temp_PDF_' + new Date().getTime());

  try {
    // Format the grievance data
    formatGrievancePDFSheet(tempSheet, grievanceData);

    // Generate PDF
    const url = ss.getUrl();
    const sheetId = tempSheet.getSheetId();
    const pdfUrl = url.replace(/edit.*$/, '') +
      'export?exportFormat=pdf&format=pdf' +
      '&size=letter' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false&printtitle=false' +
      '&pagenumbers=false&gridlines=false' +
      '&fzr=false' +
      '&gid=' + sheetId;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    const pdfBlob = response.getBlob().setName('Grievance_' + grievanceId + '.pdf');

    return pdfBlob;

  } finally {
    // Clean up temp sheet
    ss.deleteSheet(tempSheet);
  }
}

/**
 * Formats temporary sheet for PDF generation
 */
function formatGrievancePDFSheet(sheet, data) {
  sheet.clear();

  // Header
  sheet.getRange('A1:D1').merge()
    .setValue('SEIU LOCAL 509 - GRIEVANCE FORM')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#7EC8E3')
    .setFontColor('white');

  var row = 3;

  // Member Information
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('MEMBER INFORMATION')
    .setFontWeight('bold')
    .setBackground('#E8F0FE');
  row++;

  addPDFField(sheet, row++, 'Grievance ID:', data[0]);
  addPDFField(sheet, row++, 'Member ID:', data[1]);
  addPDFField(sheet, row++, 'Name:', data[2] + ' ' + data[3]);
  row++;

  // Grievance Details
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('GRIEVANCE DETAILS')
    .setFontWeight('bold')
    .setBackground('#E8F0FE');
  row++;

  addPDFField(sheet, row++, 'Status:', data[4]);
  addPDFField(sheet, row++, 'Current Step:', data[5]);
  addPDFField(sheet, row++, 'Incident Date:', data[6] ? Utilities.formatDate(new Date(data[6]), Session.getScriptTimeZone(), 'MM/dd/yyyy') : '');
  addPDFField(sheet, row++, 'Grievance Type:', data[23]);
  addPDFField(sheet, row++, 'Steward:', data[25]);
  row++;

  // Description
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('DESCRIPTION')
    .setFontWeight('bold')
    .setBackground('#E8F0FE');
  row++;

  sheet.getRange(row, 1, 3, 4).merge()
    .setValue(data[24] || '')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, false, false);

  // Auto-resize
  sheet.autoResizeColumns(1, 4);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 200);
}

/**
 * Helper to add field to PDF
 */
function addPDFField(sheet, row, label, value) {
  sheet.getRange(row, 1).setValue(label).setFontWeight('bold');
  sheet.getRange(row, 2, 1, 3).merge().setValue(value || '');
}

/**
 * Shows dialog with PDF email/download options
 */
function showPDFOptionsDialog(grievanceId, pdfBlob) {
  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: Arial, sans-serif;
      padding: 20px;
    }
    h2 { color: #1a73e8; }
    .options { margin: 20px 0; }
    .option {
      padding: 15px;
      margin: 10px 0;
      border: 2px solid #ddd;
      border-radius: 5px;
      cursor: pointer;
    }
    .option:hover {
      background: #f0f0f0;
      border-color: #1a73e8;
    }
    input[type="email"] {
      width: 100%;
      padding: 8px;
      margin: 5px 0;
    }
    button {
      padding: 10px 20px;
      margin: 5px;
      border: none;
      border-radius: 4px;
      background: #1a73e8;
      color: white;
      cursor: pointer;
    }
    button:hover { background: #1557b0; }
  </style>
</head>
<body>
  <h2>✅ Grievance Created</h2>
  <p>Grievance ID: <strong>${grievanceId}</strong></p>

  <div class="options">
    <div class="option" onclick="showEmailForm()">
      <h3>📧 Email PDF</h3>
      <p>Send the grievance PDF to one or more recipients</p>
    </div>

    <div class="option" onclick="downloadPDF()">
      <h3>💾 Download PDF</h3>
      <p>Download the grievance PDF to your computer</p>
    </div>
  </div>

  <div id="emailForm" style="display:none;">
    <h3>Email Recipients</h3>
    <p>Enter email addresses (comma-separated for multiple):</p>
    <input type="email" id="emailAddresses" placeholder="email1@example.com, email2@example.com">
    <br>
    <button onclick="sendEmail()">Send</button>
    <button onclick="document.getElementById('emailForm').style.display='none'">Cancel</button>
  </div>

  <script>
    function showEmailForm() {
      document.getElementById('emailForm').style.display = 'block';
    }

    function sendEmail() {
      const emails = document.getElementById('emailAddresses').value;
      if (!emails) {
        alert('Please enter at least one email address');
        return;
      }

      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Email sent successfully!');
          google.script.host.close();
        })
        .withFailureHandler(function(e) { return alert('❌ Error: ' + e.message); })
        .emailGrievancePDF('${grievanceId}', emails);
    }

    function downloadPDF() {
      alert('PDF download will start automatically.\\n\\nPlease check your Downloads folder.');
      google.script.run
        .withSuccessHandler(function(url) {
          window.open(url, '_blank');
          google.script.host.close();
        })
        .withFailureHandler(function(e) { return alert('❌ Error: ' + e.message); })
        .getGrievancePDFUrl('${grievanceId}');
    }
  </script>
</body>
</html>
  `).setWidth(500).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance PDF Options');
}

/**
 * Emails the grievance PDF
 */
function emailGrievancePDF(grievanceId, emailAddresses) {
  const pdfBlob = generateGrievancePDF(grievanceId);
  const emails = emailAddresses.split(',').map(function(e) { return e.trim(); });

  const subject = 'SEIU Local 509 - Grievance ' + grievanceId;
  const body = 'Please find attached the grievance form for ' + grievanceId + '.\n\n' +
               'This grievance was automatically generated from the SEIU Local 509 Dashboard.\n\n' +
               'For questions, please contact your steward.';

  emails.forEach(function(email) {
    if (email) {
      GmailApp.sendEmail(email, subject, body, {
        attachments: [pdfBlob]
      });
    }
  });
}

/**
 * Gets URL for grievance PDF
 */
function getGrievancePDFUrl(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return null;

  const url = ss.getUrl();
  const sheetId = grievanceSheet.getSheetId();

  return url.replace(/edit.*$/, '') +
    'export?exportFormat=pdf&format=pdf' +
    '&size=letter' +
    '&portrait=true' +
    '&fitw=true' +
    '&gid=' + sheetId;
}
