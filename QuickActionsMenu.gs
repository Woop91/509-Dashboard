/**
 * ============================================================================
 * QUICK ACTIONS MENU
 * ============================================================================
 *
 * Right-click context menu for common actions on Member Directory and
 * Grievance Log rows.
 *
 * Features:
 * - Context-aware actions based on selected row
 * - Start Grievance from member row
 * - Send Email to member
 * - View History for grievance
 * - Quick status updates
 * - Copy member/grievance ID
 *
 * @module QuickActionsMenu
 * @version 1.0.0
 * @author SEIU Local 509 Tech Team
 */

/**
 * Shows quick actions menu for the currently selected row
 */
function showQuickActionsMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  const selection = sheet.getActiveRange();

  if (!selection) {
    SpreadsheetApp.getUi().alert('Please select a row first.');
    return;
  }

  const row = selection.getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a data row, not the header.');
    return;
  }

  if (sheetName === SHEETS.MEMBER_DIR) {
    showMemberQuickActions(row);
  } else if (sheetName === SHEETS.GRIEVANCE_LOG) {
    showGrievanceQuickActions(row);
  } else {
    SpreadsheetApp.getUi().alert(
      'Quick Actions',
      'Quick Actions are available for Member Directory and Grievance Log sheets.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Shows quick actions for a member row
 * @param {number} row - Row number
 */
function showMemberQuickActions(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Get member data
  const lastCol = memberSheet.getLastColumn();
  const rowData = memberSheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const memberId = rowData[MEMBER_COLS.MEMBER_ID - 1];
  const firstName = rowData[MEMBER_COLS.FIRST_NAME - 1];
  const lastName = rowData[MEMBER_COLS.LAST_NAME - 1];
  const email = rowData[MEMBER_COLS.EMAIL - 1];
  const hasOpenGrievance = rowData[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1];

  const memberName = `${firstName} ${lastName}`;

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; display: flex; align-items: center; gap: 10px; }
    .member-info { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
    .member-name { font-size: 18px; font-weight: bold; color: #333; }
    .member-id { color: #666; font-size: 14px; }
    .member-status { margin-top: 10px; font-size: 13px; }
    .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .status-open { background: #ffebee; color: #c62828; }
    .status-none { background: #e8f5e9; color: #2e7d32; }
    .actions { display: flex; flex-direction: column; gap: 10px; }
    .action-btn {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 15px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      text-align: left;
      transition: all 0.2s;
      background: #f8f9fa;
    }
    .action-btn:hover {
      background: #e8f4fd;
      transform: translateX(5px);
    }
    .action-icon { font-size: 24px; }
    .action-text { flex: 1; }
    .action-title { font-weight: bold; color: #333; }
    .action-desc { font-size: 12px; color: #666; margin-top: 2px; }
    .divider { height: 1px; background: #e0e0e0; margin: 10px 0; }
    .close-btn { width: 100%; margin-top: 15px; padding: 12px; background: #6c757d; color: white; border: none; border-radius: 8px; cursor: pointer; }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚡ Quick Actions</h2>

    <div class="member-info">
      <div class="member-name">${memberName}</div>
      <div class="member-id">${memberId} | ${email || 'No email'}</div>
      <div class="member-status">
        ${hasOpenGrievance === 'Yes'
          ? '<span class="status-badge status-open">🔴 Has Open Grievance</span>'
          : '<span class="status-badge status-none">🟢 No Open Grievances</span>'}
      </div>
    </div>

    <div class="actions">
      <button class="action-btn" onclick="startGrievance()">
        <span class="action-icon">📋</span>
        <span class="action-text">
          <div class="action-title">Start New Grievance</div>
          <div class="action-desc">Create a grievance for this member</div>
        </span>
      </button>

      ${email ? `
      <button class="action-btn" onclick="sendEmail()">
        <span class="action-icon">📧</span>
        <span class="action-text">
          <div class="action-title">Send Email</div>
          <div class="action-desc">Compose email to ${email}</div>
        </span>
      </button>
      ` : ''}

      <button class="action-btn" onclick="viewGrievances()">
        <span class="action-icon">📁</span>
        <span class="action-text">
          <div class="action-title">View Grievance History</div>
          <div class="action-desc">See all grievances for this member</div>
        </span>
      </button>

      <div class="divider"></div>

      <button class="action-btn" onclick="copyMemberId()">
        <span class="action-icon">📋</span>
        <span class="action-text">
          <div class="action-title">Copy Member ID</div>
          <div class="action-desc">${memberId}</div>
        </span>
      </button>

      <button class="action-btn" onclick="viewMemberDetails()">
        <span class="action-icon">👤</span>
        <span class="action-text">
          <div class="action-title">View Full Details</div>
          <div class="action-desc">See all member information</div>
        </span>
      </button>
    </div>

    <button class="close-btn" onclick="google.script.host.close()">Close</button>
  </div>

  <script>
    const memberId = '${memberId}';
    const memberRow = ${row};

    function startGrievance() {
      google.script.run.openGrievanceFormForMember(memberRow);
      google.script.host.close();
    }

    function sendEmail() {
      google.script.run.composeEmailForMember(memberId);
      google.script.host.close();
    }

    function viewGrievances() {
      google.script.run.showMemberGrievanceHistory(memberId);
      google.script.host.close();
    }

    function copyMemberId() {
      navigator.clipboard.writeText(memberId);
      alert('Member ID copied to clipboard!');
    }

    function viewMemberDetails() {
      google.script.run.showMemberDetailsDialog(memberId);
      google.script.host.close();
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '');
}

/**
 * Shows quick actions for a grievance row
 * @param {number} row - Row number
 */
function showGrievanceQuickActions(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Get grievance data
  const lastCol = grievanceSheet.getLastColumn();
  const rowData = grievanceSheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const grievanceId = rowData[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
  const memberId = rowData[GRIEVANCE_COLS.MEMBER_ID - 1];
  const firstName = rowData[GRIEVANCE_COLS.FIRST_NAME - 1];
  const lastName = rowData[GRIEVANCE_COLS.LAST_NAME - 1];
  const status = rowData[GRIEVANCE_COLS.STATUS - 1];
  const currentStep = rowData[GRIEVANCE_COLS.CURRENT_STEP - 1];
  const nextActionDue = rowData[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];

  const memberName = `${firstName} ${lastName}`;
  const isOpen = status === 'Open';
  const daysToDeadline = nextActionDue ? Math.floor((new Date(nextActionDue) - new Date()) / (1000 * 60 * 60 * 24)) : null;

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #DC2626; margin-top: 0; display: flex; align-items: center; gap: 10px; }
    .grievance-info { background: #fff5f5; padding: 15px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #DC2626; }
    .grievance-id { font-size: 18px; font-weight: bold; color: #333; }
    .grievance-member { color: #666; font-size: 14px; }
    .grievance-status { margin-top: 10px; font-size: 13px; display: flex; gap: 10px; flex-wrap: wrap; }
    .status-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .status-open { background: #ffebee; color: #c62828; }
    .status-pending { background: #fff3e0; color: #ef6c00; }
    .status-settled { background: #e8f5e9; color: #2e7d32; }
    .deadline-badge { background: ${daysToDeadline !== null && daysToDeadline < 7 ? '#ffebee' : '#e3f2fd'}; color: ${daysToDeadline !== null && daysToDeadline < 7 ? '#c62828' : '#1565c0'}; }
    .actions { display: flex; flex-direction: column; gap: 10px; }
    .action-btn {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 15px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      text-align: left;
      transition: all 0.2s;
      background: #f8f9fa;
    }
    .action-btn:hover {
      background: #fff5f5;
      transform: translateX(5px);
    }
    .action-btn.disabled { opacity: 0.5; cursor: not-allowed; }
    .action-icon { font-size: 24px; }
    .action-text { flex: 1; }
    .action-title { font-weight: bold; color: #333; }
    .action-desc { font-size: 12px; color: #666; margin-top: 2px; }
    .divider { height: 1px; background: #e0e0e0; margin: 10px 0; }
    .status-section { margin-top: 15px; padding: 15px; background: #f8f9fa; border-radius: 8px; }
    .status-section h4 { margin: 0 0 10px 0; color: #333; }
    select { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 14px; }
    .close-btn { width: 100%; margin-top: 15px; padding: 12px; background: #6c757d; color: white; border: none; border-radius: 8px; cursor: pointer; }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚡ Grievance Actions</h2>

    <div class="grievance-info">
      <div class="grievance-id">${grievanceId}</div>
      <div class="grievance-member">${memberName} (${memberId})</div>
      <div class="grievance-status">
        <span class="status-badge status-${status.toLowerCase().replace(' ', '-')}">${status}</span>
        <span class="status-badge">${currentStep}</span>
        ${daysToDeadline !== null ? `
          <span class="status-badge deadline-badge">
            ${daysToDeadline < 0 ? '⚠️ Overdue' : daysToDeadline === 0 ? '⚠️ Due Today' : `📅 ${daysToDeadline} days`}
          </span>
        ` : ''}
      </div>
    </div>

    <div class="actions">
      <button class="action-btn" onclick="sendEmail()">
        <span class="action-icon">📧</span>
        <span class="action-text">
          <div class="action-title">Send Email Update</div>
          <div class="action-desc">Email the member about this grievance</div>
        </span>
      </button>

      <button class="action-btn" onclick="viewFiles()">
        <span class="action-icon">📁</span>
        <span class="action-text">
          <div class="action-title">View Drive Folder</div>
          <div class="action-desc">Open associated Google Drive folder</div>
        </span>
      </button>

      <button class="action-btn" onclick="addToCalendar()">
        <span class="action-icon">📅</span>
        <span class="action-text">
          <div class="action-title">Sync to Calendar</div>
          <div class="action-desc">Add deadlines to Google Calendar</div>
        </span>
      </button>

      <button class="action-btn" onclick="viewComms()">
        <span class="action-icon">📞</span>
        <span class="action-text">
          <div class="action-title">View Communications</div>
          <div class="action-desc">See all logged communications</div>
        </span>
      </button>

      <div class="divider"></div>

      <button class="action-btn" onclick="copyGrievanceId()">
        <span class="action-icon">📋</span>
        <span class="action-text">
          <div class="action-title">Copy Grievance ID</div>
          <div class="action-desc">${grievanceId}</div>
        </span>
      </button>
    </div>

    ${isOpen ? `
    <div class="status-section">
      <h4>Quick Status Update</h4>
      <select id="statusSelect">
        <option value="">-- Select New Status --</option>
        <option value="Open">Open</option>
        <option value="Pending Info">Pending Info</option>
        <option value="Settled">Settled</option>
        <option value="Withdrawn">Withdrawn</option>
        <option value="Closed">Closed</option>
        <option value="Appealed">Appealed</option>
      </select>
      <button class="action-btn" style="margin-top: 10px;" onclick="updateStatus()">
        <span class="action-icon">✓</span>
        <span class="action-text">
          <div class="action-title">Update Status</div>
        </span>
      </button>
    </div>
    ` : ''}

    <button class="close-btn" onclick="google.script.host.close()">Close</button>
  </div>

  <script>
    const grievanceId = '${grievanceId}';
    const grievanceRow = ${row};

    function sendEmail() {
      google.script.run.composeGrievanceEmail(grievanceId);
      google.script.host.close();
    }

    function viewFiles() {
      google.script.run.openGrievanceDriveFolder(grievanceId);
      google.script.host.close();
    }

    function addToCalendar() {
      google.script.run.syncSingleGrievanceToCalendar(grievanceId);
      google.script.host.close();
    }

    function viewComms() {
      google.script.run.showGrievanceCommunications(grievanceId);
      google.script.host.close();
    }

    function copyGrievanceId() {
      navigator.clipboard.writeText(grievanceId);
      alert('Grievance ID copied to clipboard!');
    }

    function updateStatus() {
      const newStatus = document.getElementById('statusSelect').value;
      if (!newStatus) {
        alert('Please select a status');
        return;
      }
      google.script.run
        .withSuccessHandler(function() {
          alert('Status updated to: ' + newStatus);
          google.script.host.close();
        })
        .quickUpdateGrievanceStatus(grievanceRow, newStatus);
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '');
}

/**
 * Quickly updates grievance status from the Quick Actions menu
 * @param {number} row - Grievance row number
 * @param {string} newStatus - New status value
 */
function quickUpdateGrievanceStatus(row, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const statusCol = GRIEVANCE_COLS.STATUS;
  grievanceSheet.getRange(row, statusCol).setValue(newStatus);

  // If closed/settled, set the close date
  if (['Closed', 'Settled', 'Withdrawn'].includes(newStatus)) {
    const closeDateCol = GRIEVANCE_COLS.DATE_CLOSED;
    if (!grievanceSheet.getRange(row, closeDateCol).getValue()) {
      grievanceSheet.getRange(row, closeDateCol).setValue(new Date());
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Grievance status updated to: ${newStatus}`,
    'Status Updated',
    3
  );
}

/**
 * Composes email for a specific member (called from Quick Actions)
 * @param {string} memberId - Member ID
 */
function composeEmailForMember(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  const lastRow = memberSheet.getLastRow();
  const data = memberSheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.EMAIL).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      const email = data[i][MEMBER_COLS.EMAIL - 1];
      const firstName = data[i][MEMBER_COLS.FIRST_NAME - 1];
      const lastName = data[i][MEMBER_COLS.LAST_NAME - 1];

      if (email) {
        // Open email composition with member info
        const html = createMemberEmailHTML(memberId, `${firstName} ${lastName}`, email);
        const htmlOutput = HtmlService.createHtmlOutput(html)
          .setWidth(700)
          .setHeight(600);
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📧 Compose Email');
      } else {
        SpreadsheetApp.getUi().alert('This member does not have an email address on file.');
      }
      return;
    }
  }
}

/**
 * Creates email composition HTML for member
 */
function createMemberEmailHTML(memberId, memberName, email) {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; }
    .form-group { margin: 15px 0; }
    label { display: block; font-weight: bold; margin-bottom: 5px; color: #333; }
    input, textarea { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; }
    textarea { min-height: 200px; font-family: inherit; }
    input:focus, textarea:focus { outline: none; border-color: #1a73e8; }
    .buttons { display: flex; gap: 10px; margin-top: 20px; }
    button { padding: 12px 24px; font-size: 14px; border: none; border-radius: 4px; cursor: pointer; flex: 1; }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-secondary { background: #6c757d; color: white; }
    .info-box { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📧 Email to Member</h2>

    <div class="info-box">
      <strong>${memberName}</strong> (${memberId})<br>
      ${email}
    </div>

    <div class="form-group">
      <label>Subject:</label>
      <input type="text" id="subject" placeholder="Email subject">
    </div>

    <div class="form-group">
      <label>Message:</label>
      <textarea id="message" placeholder="Type your message here..."></textarea>
    </div>

    <div class="buttons">
      <button class="btn-primary" onclick="sendEmail()">📤 Send Email</button>
      <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
    </div>
  </div>

  <script>
    function sendEmail() {
      const subject = document.getElementById('subject').value.trim();
      const message = document.getElementById('message').value.trim();

      if (!subject || !message) {
        alert('Please fill in subject and message');
        return;
      }

      google.script.run
        .withSuccessHandler(function() {
          alert('Email sent successfully!');
          google.script.host.close();
        })
        .withFailureHandler(function(error) {
          alert('Error sending email: ' + error.message);
        })
        .sendQuickEmail('${email}', subject, message, '${memberId}');
    }
  </script>
</body>
</html>
  `;
}

/**
 * Sends a quick email from the Quick Actions menu
 * @param {string} to - Recipient email
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @param {string} memberId - Member ID for logging
 */
function sendQuickEmail(to, subject, body, memberId) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      name: EMAIL_CONFIG.FROM_NAME
    });

    // Log communication
    try {
      logCommunication(memberId, 'Email Sent', `To: ${to}\nSubject: ${subject}\n\n${body}`);
    } catch (e) {
      // Logging not critical
    }

    return { success: true };
  } catch (error) {
    throw new Error('Failed to send email: ' + error.message);
  }
}

/**
 * Shows member grievance history
 * @param {string} memberId - Member ID
 */
function showMemberGrievanceHistory(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found!');
    return;
  }

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No grievances found.');
    return;
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.RESOLUTION).getValues();

  const memberGrievances = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
      memberGrievances.push({
        grievanceId: data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        status: data[i][GRIEVANCE_COLS.STATUS - 1],
        step: data[i][GRIEVANCE_COLS.CURRENT_STEP - 1],
        issueCategory: data[i][GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
        dateFiled: data[i][GRIEVANCE_COLS.DATE_FILED - 1],
        dateClosed: data[i][GRIEVANCE_COLS.DATE_CLOSED - 1]
      });
    }
  }

  if (memberGrievances.length === 0) {
    SpreadsheetApp.getUi().alert('No grievances found for this member.');
    return;
  }

  // Build HTML
  const grievanceList = memberGrievances.map(g => `
    <div style="background: #f8f9fa; padding: 12px; margin: 8px 0; border-radius: 4px; border-left: 4px solid ${g.status === 'Open' ? '#f44336' : '#4caf50'};">
      <strong>${g.grievanceId}</strong><br>
      <span style="color: #666;">Status: ${g.status} | Step: ${g.step}</span><br>
      <span style="color: #888; font-size: 12px;">
        ${g.issueCategory} |
        Filed: ${g.dateFiled ? new Date(g.dateFiled).toLocaleDateString() : 'N/A'} |
        ${g.dateClosed ? 'Closed: ' + new Date(g.dateClosed).toLocaleDateString() : 'Ongoing'}
      </span>
    </div>
  `).join('');

  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  .summary { background: #e8f4fd; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
</style>
</head>
<body>
  <h2>📁 Grievance History</h2>
  <div class="summary">
    <strong>Member ID:</strong> ${memberId}<br>
    <strong>Total Grievances:</strong> ${memberGrievances.length}<br>
    <strong>Open:</strong> ${memberGrievances.filter(g => g.status === 'Open').length}<br>
    <strong>Closed:</strong> ${memberGrievances.filter(g => g.status !== 'Open').length}
  </div>
  ${grievanceList}
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Grievance History - ' + memberId);
}

/**
 * Shows member details dialog
 * @param {string} memberId - Member ID
 */
function showMemberDetailsDialog(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  const lastRow = memberSheet.getLastRow();
  const lastCol = memberSheet.getLastColumn();
  const headers = memberSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = memberSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let memberData = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      memberData = data[i];
      break;
    }
  }

  if (!memberData) {
    SpreadsheetApp.getUi().alert('Member not found.');
    return;
  }

  // Build details HTML
  let detailsHTML = '';
  for (let i = 0; i < headers.length; i++) {
    const value = memberData[i];
    if (value !== '' && value !== null && value !== undefined) {
      detailsHTML += `
        <tr>
          <td style="font-weight: bold; padding: 8px; background: #f8f9fa;">${headers[i]}</td>
          <td style="padding: 8px;">${value}</td>
        </tr>
      `;
    }
  }

  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  table { width: 100%; border-collapse: collapse; }
  tr:nth-child(even) { background: #f9f9f9; }
  td { border: 1px solid #ddd; }
</style>
</head>
<body>
  <h2>👤 Member Details - ${memberId}</h2>
  <table>${detailsHTML}</table>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Member Details');
}

/**
 * Opens grievance Drive folder
 * @param {string} grievanceId - Grievance ID
 */
function openGrievanceDriveFolder(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  const lastRow = grievanceSheet.getLastRow();
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId) {
      const folderUrl = data[i][GRIEVANCE_COLS.DRIVE_FOLDER_URL - 1];
      if (folderUrl) {
        const html = HtmlService.createHtmlOutput(
          `<script>window.open('${folderUrl}', '_blank'); google.script.host.close();</script>`
        );
        SpreadsheetApp.getUi().showModalDialog(html, 'Opening Drive Folder...');
      } else {
        SpreadsheetApp.getUi().alert(
          'No Drive folder found for this grievance. Would you like to create one?',
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );
      }
      return;
    }
  }

  SpreadsheetApp.getUi().alert('Grievance not found.');
}

/**
 * Syncs a single grievance to calendar
 * @param {string} grievanceId - Grievance ID
 */
function syncSingleGrievanceToCalendar(grievanceId) {
  try {
    // This would call the calendar integration function
    // For now, show a confirmation
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Syncing grievance deadlines to calendar...',
      'Calendar Sync',
      3
    );

    // Call the actual sync function if it exists
    if (typeof syncDeadlinesToCalendar === 'function') {
      syncDeadlinesToCalendar();
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error syncing to calendar: ' + error.message);
  }
}
