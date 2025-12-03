/**
 * ------------------------------------------------------------------------====
 * BATCH OPERATIONS
 * ------------------------------------------------------------------------====
 *
 * Bulk operations for efficient mass updates
 * Features:
 * - Bulk assign steward
 * - Bulk update status
 * - Bulk export to PDF
 * - Bulk email notifications
 * - Selection-based operations
 */

/**
 * Shows batch operations menu
 */
function showBatchOperationsMenu() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createHtmlOutput(`
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
      max-width: 600px;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .operation-card {
      background: #f8f9fa;
      padding: 15px;
      margin: 15px 0;
      border-radius: 4px;
      border-left: 4px solid #1a73e8;
      cursor: pointer;
      transition: all 0.2s;
    }
    .operation-card:hover {
      background: #e9ecef;
      transform: translateX(5px);
    }
    .operation-title {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 5px;
    }
    .operation-desc {
      font-size: 13px;
      color: #666;
    }
    .warning {
      background: #fff3cd;
      border: 1px solid #ffc107;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      color: #856404;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚡ Batch Operations</h2>

    <div class="warning">
      <strong>⚠️ Note:</strong> Select rows in the Grievance Log before using batch operations.
      Operations will apply to all selected grievances.
    </div>

    <div class="operation-card" onclick="runBatchAssignSteward()">
      <div class="operation-title">👤 Bulk Assign Steward</div>
      <div class="operation-desc">Assign a steward to multiple grievances at once</div>
    </div>

    <div class="operation-card" onclick="runBatchUpdateStatus()">
      <div class="operation-title">📊 Bulk Update Status</div>
      <div class="operation-desc">Change status for multiple grievances</div>
    </div>

    <div class="operation-card" onclick="runBatchExportPDF()">
      <div class="operation-title">📄 Bulk Export to PDF</div>
      <div class="operation-desc">Generate PDF reports for selected grievances</div>
    </div>

    <div class="operation-card" onclick="runBatchEmail()">
      <div class="operation-title">📧 Bulk Email Notifications</div>
      <div class="operation-desc">Send email updates to multiple stewards</div>
    </div>

    <div class="operation-card" onclick="runBatchAddNotes()">
      <div class="operation-title">📝 Bulk Add Notes</div>
      <div class="operation-desc">Add the same note to multiple grievances</div>
    </div>
  </div>

  <script>
    function runBatchAssignSteward() {
      google.script.run.withSuccessHandler(function(() {
        google.script.host.close();
      }).batchAssignSteward();
    }

    function runBatchUpdateStatus() {
      google.script.run.withSuccessHandler(function(() {
        google.script.host.close();
      }).batchUpdateStatus();
    }

    function runBatchExportPDF() {
      google.script.run.withSuccessHandler(function(() {
        google.script.host.close();
      }).batchExportPDF();
    }

    function runBatchEmail() {
      google.script.run.withSuccessHandler(function(() {
        google.script.host.close();
      }).batchEmailNotifications();
    }

    function runBatchAddNotes() {
      google.script.run.withSuccessHandler(function(() {
        google.script.host.close();
      }).batchAddNotes();
    }
  </script>
</body>
</html>
  `)
    .setWidth(650)
    .setHeight(550);

  ui.showModalDialog(html, '⚡ Batch Operations');
}

/**
 * Gets selected grievance rows from the active selection
 * @returns {Array} Array of selected row numbers
 */
function getSelectedGrievanceRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Check if we're on the Grievance Log sheet
  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Wrong Sheet',
      'Please select rows in the Grievance Log sheet first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return [];
  }

  const selection = activeSheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  // Don't include header row
  if (startRow === 1) {
    if (numRows === 1) {
      SpreadsheetApp.getUi().alert(
        '⚠️ No Data Selected',
        'Please select one or more grievance rows (not the header).',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return [];
    }
    return Array.fromfunction({ length: numRows - 1 }, (_, i) { return startRow + i + 1; });
  }

  return Array.fromfunction({ length: numRows }, (_, i) { return startRow + i; });
}

/**
 * Bulk assigns a steward to multiple grievances
 */
function batchAssignSteward() {
  const ui = SpreadsheetApp.getUi();
  const selectedRows = getSelectedGrievanceRows();

  if (selectedRows.length === 0) {
    return;
  }

  // Get list of stewards
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const stewardData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, 10).getValues();

  const stewards = stewardData
    .filter(function(row) { return row[9] === 'Yes'; }) // Column J: Is Steward?
    .map(function(row) { return `${row[1]} ${row[2]}`; }) // First + Last Name
    .filter(function(name) { return name.trim() !== ''; });

  if (stewards.length === 0) {
    ui.alert('❌ No stewards found in the Member Directory.');
    return;
  }

  // Show steward selection dialog
  const response = ui.prompt(
    '👤 Bulk Assign Steward',
    `You have selected ${selectedRows.length} grievance(s).\n\n` +
    `Available stewards:\n${stewards.slice(0, 10).join('\n')}${stewards.length > 10 ? '\n...' : ''}\n\n` +
    'Enter the steward name to assign:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const stewardName = response.getResponseText().trim();

  if (!stewards.includes(stewardName)) {
    ui.alert(
      '⚠️ Invalid Steward',
      `"${stewardName}" is not a valid steward.\n\nPlease enter a name exactly as it appears in the list.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Confirm operation
  const confirmResponse = ui.alert(
    '⚠️ Confirm Bulk Assignment',
    `This will assign "${stewardName}" to ${selectedRows.length} grievance(s).\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  // Perform bulk assignment
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var updated = 0;

  selectedRows.forEach(function(row) {
    grievanceSheet.getRange(row, GRIEVANCE_COLS.ASSIGNED_STEWARD).setValue(stewardName);
    updated++;
  });

  ui.alert(
    '✅ Bulk Assignment Complete',
    `Successfully assigned "${stewardName}" to ${updated} grievance(s).`,
    ui.ButtonSet.OK
  );
}

/**
 * Bulk updates status for multiple grievances
 */
function batchUpdateStatus() {
  const ui = SpreadsheetApp.getUi();
  const selectedRows = getSelectedGrievanceRows();

  if (selectedRows.length === 0) {
    return;
  }

  const statusOptions = ['Open', 'In Progress', 'Resolved', 'Closed', 'Withdrawn', 'Escalated'];

  // Show status selection dialog
  const response = ui.prompt(
    '📊 Bulk Update Status',
    `You have selected ${selectedRows.length} grievance(s).\n\n` +
    `Available statuses:\n${statusOptions.join('\n')}\n\n` +
    'Enter the new status:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const newStatus = response.getResponseText().trim();

  if (!statusOptions.includes(newStatus)) {
    ui.alert(
      '⚠️ Invalid Status',
      `"${newStatus}" is not a valid status.\n\nPlease enter one of:\n${statusOptions.join(', ')}`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Confirm operation
  const confirmResponse = ui.alert(
    '⚠️ Confirm Bulk Status Update',
    `This will change the status to "${newStatus}" for ${selectedRows.length} grievance(s).\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  // Perform bulk status update
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var updated = 0;

  selectedRows.forEach(function(row) {
    grievanceSheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newStatus);

    // Update closed date if status is Closed or Resolved
    if (newStatus === 'Closed' || newStatus === 'Resolved') {
      grievanceSheet.getRange(row, GRIEVANCE_COLS.DATE_CLOSED).setValue(new Date());
    }

    updated++;
  });

  ui.alert(
    '✅ Bulk Status Update Complete',
    `Successfully updated status to "${newStatus}" for ${updated} grievance(s).`,
    ui.ButtonSet.OK
  );
}

/**
 * Bulk exports selected grievances to PDF
 */
function batchExportPDF() {
  const ui = SpreadsheetApp.getUi();
  const selectedRows = getSelectedGrievanceRows();

  if (selectedRows.length === 0) {
    return;
  }

  if (selectedRows.length > 20) {
    ui.alert(
      '⚠️ Too Many Selections',
      `You selected ${selectedRows.length} grievances.\n\nFor performance reasons, bulk PDF export is limited to 20 grievances at a time.\n\nPlease select fewer rows.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const confirmResponse = ui.alert(
    '📄 Confirm Bulk PDF Export',
    `This will generate ${selectedRows.length} PDF report(s).\n\nPDFs will be saved to your Google Drive.\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('📄 Generating PDFs...', 'Please wait', -1);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const pdfs = [];

  selectedRows.forEach(function(row) {
    try {
      const grievanceId = grievanceSheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
      const pdf = exportGrievanceToPDF(grievanceId);
      if (pdf) {
        pdfs.push(pdf);
      }
    } catch (error) {
      Logger.log(`Error exporting row ${row}: ${error.message}`);
    }
  });

  ui.alert(
    '✅ Bulk PDF Export Complete',
    `Successfully generated ${pdfs.length} PDF report(s).\n\nCheck your Google Drive for the files.`,
    ui.ButtonSet.OK
  );
}

/**
 * Exports a single grievance to PDF (helper function)
 * @param {string} grievanceId - Grievance ID to export
 * @returns {File} PDF file object
 */
function exportGrievanceToPDF(grievanceId) {
  // Creates a PDF using Google Docs conversion

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Find the grievance
  const lastRow = grievanceSheet.getLastRow();
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      const row = data[i];

      // Create a Google Doc
      const doc = DocumentApp.create(`Grievance_${grievanceId}_Report`);
      const body = doc.getBody();

      // Add content
      body.appendParagraph('SEIU Local 509 - Grievance Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph(''); // Spacing

      body.appendParagraph(`Grievance ID: ${row[0]}`);
      body.appendParagraph(`Member: ${row[2]} ${row[3]}`);
      body.appendParagraph(`Status: ${row[4]}`);
      body.appendParagraph(`Issue Type: ${row[5]}`);
      body.appendParagraph(`Filed Date: ${row[6]}`);
      body.appendParagraph(`Assigned Steward: ${row[13]}`);

      doc.saveAndClose();

      // Convert to PDF
      const docFile = DriveApp.getFileById(doc.getId());
      const pdfBlob = docFile.getAs('application/pdf');
      const pdfFile = DriveApp.createFile(pdfBlob);

      // Delete the temporary doc
      DriveApp.getFileById(doc.getId()).setTrashed(true);

      return pdfFile;
    }
  }

  return null;
}

/**
 * Bulk sends email notifications for selected grievances
 */
function batchEmailNotifications() {
  const ui = SpreadsheetApp.getUi();
  const selectedRows = getSelectedGrievanceRows();

  if (selectedRows.length === 0) {
    return;
  }

  if (selectedRows.length > 50) {
    ui.alert(
      '⚠️ Too Many Selections',
      `You selected ${selectedRows.length} grievances.\n\nFor email quota reasons, bulk email is limited to 50 recipients at a time.\n\nPlease select fewer rows.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // Get email template
  const response = ui.prompt(
    '📧 Bulk Email Notifications',
    `You have selected ${selectedRows.length} grievance(s).\n\n` +
    'This will email the assigned steward for each grievance.\n\n' +
    'Enter email subject:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const subject = response.getResponseText().trim();

  if (!subject) {
    ui.alert('⚠️ Subject is required.');
    return;
  }

  const messageResponse = ui.prompt(
    '📧 Email Message',
    'Enter email message body:',
    ui.ButtonSet.OK_CANCEL
  );

  if (messageResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const message = messageResponse.getResponseText().trim();

  if (!message) {
    ui.alert('⚠️ Message is required.');
    return;
  }

  // Confirm operation
  const confirmResponse = ui.alert(
    '⚠️ Confirm Bulk Email',
    `This will send ${selectedRows.length} email(s).\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('📧 Sending emails...', 'Please wait', -1);

  // Send emails
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var sent = 0;

  selectedRows.forEach(function(row) {
    try {
      const grievanceId = grievanceSheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
      const steward = grievanceSheet.getRange(row, GRIEVANCE_COLS.ASSIGNED_STEWARD).getValue();
      const memberName = `${grievanceSheet.getRange(row, GRIEVANCE_COLS.FIRST_NAME).getValue()} ${grievanceSheet.getRange(row, GRIEVANCE_COLS.LAST_NAME).getValue()}`;

      if (steward && steward.toString().includes('@')) {
        const fullMessage = `${message}\n\n---\nGrievance ID: ${grievanceId}\nMember: ${memberName}\n\nSEIU Local 509 Dashboard`;

        MailApp.sendEmail({
          to: steward,
          subject: subject,
          body: fullMessage,
          name: 'SEIU Local 509 Dashboard'
        });

        sent++;
      }
    } catch (error) {
      Logger.log(`Error sending email for row ${row}: ${error.message}`);
    }
  });

  ui.alert(
    '✅ Bulk Email Complete',
    `Successfully sent ${sent} email(s).`,
    ui.ButtonSet.OK
  );
}

/**
 * Bulk adds the same note to multiple grievances
 */
function batchAddNotes() {
  const ui = SpreadsheetApp.getUi();
  const selectedRows = getSelectedGrievanceRows();

  if (selectedRows.length === 0) {
    return;
  }

  // Get note text
  const response = ui.prompt(
    '📝 Bulk Add Notes',
    `You have selected ${selectedRows.length} grievance(s).\n\n` +
    'Enter note to add to all selected grievances:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const noteText = response.getResponseText().trim();

  if (!noteText) {
    ui.alert('⚠️ Note text is required.');
    return;
  }

  // Confirm operation
  const confirmResponse = ui.alert(
    '⚠️ Confirm Bulk Note Addition',
    `This will add the note to ${selectedRows.length} grievance(s).\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  // Add notes
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const timestamp = new Date().toLocaleString();
  const user = Session.getActiveUser().getEmail() || 'User';
  const formattedNote = `[${timestamp}] ${user}: ${noteText}`;

  var updated = 0;

  selectedRows.forEach(function(row) {
    const existingNotes = grievanceSheet.getRange(row, GRIEVANCE_COLS.NOTES).getValue() || '';
    const newNotes = existingNotes ? `${existingNotes}\n${formattedNote}` : formattedNote;
    grievanceSheet.getRange(row, GRIEVANCE_COLS.NOTES).setValue(newNotes);
    updated++;
  });

  ui.alert(
    '✅ Bulk Note Addition Complete',
    `Successfully added note to ${updated} grievance(s).`,
    ui.ButtonSet.OK
  );
}
