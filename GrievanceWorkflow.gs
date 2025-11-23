/**
 * ============================================================================
 * GRIEVANCE WORKFLOW - Start Grievance from Member Directory
 * ============================================================================
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
 * ============================================================================
 */

// Configuration - Update this with your Google Form URL
const GRIEVANCE_FORM_CONFIG = {
  // Replace with your actual Google Form URL
  FORM_URL: "https://docs.google.com/forms/d/e/YOUR_FORM_ID/viewform",

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
 */
function getMemberList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return [];

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return [];

  // Get all member data (columns A-K: ID, First, Last, Job, Location, Unit, Office Days, Email, Phone, Is Steward, Status)
  const data = memberSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  return data.map((row, index) => ({
    rowIndex: index + 2,
    memberId: row[0],
    firstName: row[1],
    lastName: row[2],
    jobTitle: row[3],
    location: row[4],
    unit: row[5],
    officeDays: row[6],
    email: row[7],
    phone: row[8],
    isSteward: row[9],
    status: row[10]
  })).filter(member => member.memberId); // Filter out empty rows
}

/**
 * Creates HTML dialog for member selection
 */
function createMemberSelectionDialog(members) {
  const memberOptions = members.map(m =>
    `<option value="${m.rowIndex}">${m.lastName}, ${m.firstName} (${m.memberId}) - ${m.location}</option>`
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
    let members = ${JSON.stringify(members)};

    function showMemberDetails() {
      const select = document.getElementById('memberSelect');
      const detailsDiv = document.getElementById('memberDetails');
      const startBtn = document.getElementById('startBtn');

      if (select.value) {
        const member = members.find(m => m.rowIndex == select.value);
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

  // Steward info is in columns U-X (21-24), rows 2-4
  try {
    const stewardData = configSheet.getRange(2, 21, 3, 1).getValues();
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
 * Updates Config sheet to include Steward Contact Info section
 */
function addStewardContactInfoToConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEETS.CONFIG);

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
 * ============================================================================
 * GOOGLE FORM SUBMISSION HANDLING
 * ============================================================================
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

    // Generate PDF
    const pdfBlob = generateGrievancePDF(grievanceId);

    // Show email/download dialog
    showPDFOptionsDialog(grievanceId, pdfBlob);

  } catch (error) {
    Logger.log('Error in onGrievanceFormSubmit: ' + error.message);
    SpreadsheetApp.getUi().alert('❌ Error processing form submission: ' + error.message);
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

  return grievanceId;
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
  let counter = 1;
  let letter = 'A';
  let newId;

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
 * ============================================================================
 * PDF GENERATION AND DISTRIBUTION
 * ============================================================================
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
  let grievanceRow = -1;

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

  let row = 3;

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
        .withSuccessHandler(() => {
          alert('✅ Email sent successfully!');
          google.script.host.close();
        })
        .withFailureHandler(e => alert('❌ Error: ' + e.message))
        .emailGrievancePDF('${grievanceId}', emails);
    }

    function downloadPDF() {
      alert('PDF download will start automatically.\\n\\nPlease check your Downloads folder.');
      google.script.run
        .withSuccessHandler(url => {
          window.open(url, '_blank');
          google.script.host.close();
        })
        .withFailureHandler(e => alert('❌ Error: ' + e.message))
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
  const emails = emailAddresses.split(',').map(e => e.trim());

  const subject = 'SEIU Local 509 - Grievance ' + grievanceId;
  const body = 'Please find attached the grievance form for ' + grievanceId + '.\n\n' +
               'This grievance was automatically generated from the SEIU Local 509 Dashboard.\n\n' +
               'For questions, please contact your steward.';

  emails.forEach(email => {
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
