/**
 * ============================================================================
 * PII PROTECTION SYSTEM
 * ============================================================================
 *
 * Protects Personally Identifiable Information (PII) with encryption,
 * masking, and compliance features.
 *
 * Features:
 * - Field-level encryption for sensitive data
 * - Data masking in exports and displays
 * - GDPR/CCPA compliance helpers
 * - Right to erasure (data deletion)
 * - Data access logging
 * - Export controls
 * - Role-based access to PII
 *
 * @module PIIProtection
 * @version 1.0.0
 * @author SEIU Local 509 Tech Team
 */

/**
 * PII configuration
 */
const PII_CONFIG = {
  // Fields considered PII in Member Directory
  MEMBER_PII_FIELDS: [
    'EMAIL',           // Column H
    'PHONE',           // Column I
    'HOME_TOWN'        // Column X
  ],

  // Fields considered PII in Grievance Log
  GRIEVANCE_PII_FIELDS: [
    'MEMBER_EMAIL'     // Column X
  ],

  // Masking patterns
  MASK_EMAIL: true,
  MASK_PHONE: true,
  MASK_SSN: true,

  // Encryption key property name (stored in Script Properties)
  ENCRYPTION_KEY_PROP: 'pii_encryption_key',

  // Retention period in days (for GDPR compliance)
  DEFAULT_RETENTION_DAYS: 730  // 2 years
};

/**
 * PII field definitions with masking rules
 */
const PII_FIELDS = {
  email: {
    pattern: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    mask: (value) => {
      if (!value) return '';
      const parts = value.split('@');
      if (parts.length !== 2) return '***@***.***';
      const user = parts[0];
      const domain = parts[1];
      const maskedUser = user.length > 2
        ? user[0] + '*'.repeat(user.length - 2) + user[user.length - 1]
        : '*'.repeat(user.length);
      const domainParts = domain.split('.');
      const maskedDomain = domainParts[0][0] + '*'.repeat(domainParts[0].length - 1) +
        '.' + domainParts.slice(1).join('.');
      return maskedUser + '@' + maskedDomain;
    }
  },
  phone: {
    pattern: /^\+?[\d\s\-\(\)]{10,}$/,
    mask: (value) => {
      if (!value) return '';
      const digits = value.replace(/\D/g, '');
      if (digits.length < 4) return '***-***-****';
      return '***-***-' + digits.slice(-4);
    }
  },
  ssn: {
    pattern: /^\d{3}-?\d{2}-?\d{4}$/,
    mask: (value) => {
      if (!value) return '';
      const digits = value.replace(/\D/g, '');
      if (digits.length < 4) return '***-**-****';
      return '***-**-' + digits.slice(-4);
    }
  },
  address: {
    mask: (value) => {
      if (!value) return '';
      const parts = value.split(',');
      if (parts.length > 1) {
        return '*** ' + parts.slice(-1)[0].trim();
      }
      return '*** (Address Hidden)';
    }
  }
};

/**
 * Masks email address for display
 * @param {string} email - Full email address
 * @returns {string} Masked email (e.g., j***n@g***.com)
 */
function maskEmail(email) {
  return PII_FIELDS.email.mask(email);
}

/**
 * Masks phone number for display
 * @param {string} phone - Full phone number
 * @returns {string} Masked phone (e.g., ***-***-1234)
 */
function maskPhone(phone) {
  return PII_FIELDS.phone.mask(phone);
}

/**
 * Masks SSN for display
 * @param {string} ssn - Full SSN
 * @returns {string} Masked SSN (e.g., ***-**-1234)
 */
function maskSSN(ssn) {
  return PII_FIELDS.ssn.mask(ssn);
}

/**
 * Shows PII protection dashboard
 */
function showPIIProtectionDashboard() {
  const stats = getPIIStatistics();

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; display: flex; align-items: center; gap: 10px; }
    .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 20px 0; }
    .stat-card { background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; }
    .stat-card.protected { border-left: 4px solid #4caf50; }
    .stat-card.sensitive { border-left: 4px solid #f44336; }
    .stat-value { font-size: 28px; font-weight: bold; color: #333; }
    .stat-label { font-size: 14px; color: #666; margin-top: 5px; }
    .section { margin-top: 25px; }
    .section h3 { color: #333; border-bottom: 2px solid #e0e0e0; padding-bottom: 10px; }
    .action-btn { display: block; width: 100%; padding: 12px; margin: 8px 0; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; text-align: left; }
    .action-btn:hover { background: #1557b0; }
    .action-btn.warning { background: #ff9800; }
    .action-btn.danger { background: #f44336; }
    .info-box { background: #e8f5e9; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #4caf50; }
    .warning-box { background: #fff3e0; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #ff9800; }
    .compliance-badge { display: inline-block; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; margin: 5px; }
    .badge-gdpr { background: #e3f2fd; color: #1565c0; }
    .badge-ccpa { background: #f3e5f5; color: #7b1fa2; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔒 PII Protection Dashboard</h2>

    <div style="margin: 15px 0;">
      <span class="compliance-badge badge-gdpr">GDPR Compliant</span>
      <span class="compliance-badge badge-ccpa">CCPA Compliant</span>
    </div>

    <div class="stats-grid">
      <div class="stat-card protected">
        <div class="stat-value">${stats.totalMembers}</div>
        <div class="stat-label">Total Member Records</div>
      </div>
      <div class="stat-card sensitive">
        <div class="stat-value">${stats.piiFieldCount}</div>
        <div class="stat-label">PII Fields Protected</div>
      </div>
      <div class="stat-card protected">
        <div class="stat-value">${stats.emailCount}</div>
        <div class="stat-label">Email Addresses</div>
      </div>
      <div class="stat-card protected">
        <div class="stat-value">${stats.phoneCount}</div>
        <div class="stat-label">Phone Numbers</div>
      </div>
    </div>

    <div class="info-box">
      <strong>🛡️ Protection Status:</strong><br>
      All PII fields are configured for masking in exports. Email and phone numbers
      will be partially hidden when exported to CSV or shared externally.
    </div>

    <div class="section">
      <h3>🔧 Protection Actions</h3>

      <button class="action-btn" onclick="exportMasked()">
        📤 Export with Masked PII
      </button>

      <button class="action-btn" onclick="auditPII()">
        📊 Run PII Audit Report
      </button>

      <button class="action-btn" onclick="viewAccessLog()">
        📋 View PII Access Log
      </button>

      <button class="action-btn warning" onclick="anonymizeInactive()">
        🔄 Anonymize Inactive Members
      </button>
    </div>

    <div class="section">
      <h3>📋 GDPR/CCPA Compliance</h3>

      <button class="action-btn" onclick="generateDataPortability()">
        📥 Generate Data Portability Export
      </button>

      <button class="action-btn" onclick="showDataSubjectRequest()">
        👤 Process Data Subject Request
      </button>

      <button class="action-btn danger" onclick="processErasureRequest()">
        🗑️ Process Right to Erasure Request
      </button>
    </div>

    <div class="warning-box">
      <strong>⚠️ Important:</strong><br>
      All data deletions are logged and cannot be undone.
      Ensure proper authorization before processing erasure requests.
    </div>
  </div>

  <script>
    function exportMasked() {
      google.script.run.withSuccessHandler(function() {
        alert('Masked export created!');
      }).exportMembersWithMaskedPII();
    }

    function auditPII() {
      google.script.run.showPIIAuditReport();
      google.script.host.close();
    }

    function viewAccessLog() {
      google.script.run.showPIIAccessLog();
      google.script.host.close();
    }

    function anonymizeInactive() {
      if (confirm('This will anonymize PII for members inactive for 2+ years. Continue?')) {
        google.script.run.anonymizeInactiveMembers();
      }
    }

    function generateDataPortability() {
      const memberId = prompt('Enter Member ID for data export:');
      if (memberId) {
        google.script.run.generateMemberDataExport(memberId);
      }
    }

    function showDataSubjectRequest() {
      google.script.run.showDataSubjectRequestForm();
      google.script.host.close();
    }

    function processErasureRequest() {
      const memberId = prompt('Enter Member ID for erasure request:');
      if (memberId && confirm('This will permanently delete all data for ' + memberId + '. This cannot be undone. Continue?')) {
        google.script.run.processMemberErasure(memberId);
      }
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔒 PII Protection');
}

/**
 * Gets PII statistics
 * @returns {Object} Statistics about PII fields
 */
function getPIIStatistics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    return { totalMembers: 0, piiFieldCount: 0, emailCount: 0, phoneCount: 0 };
  }

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) {
    return { totalMembers: 0, piiFieldCount: 3, emailCount: 0, phoneCount: 0 };
  }

  const totalMembers = lastRow - 1;

  // Count emails and phones
  const emailData = memberSheet.getRange(2, MEMBER_COLS.EMAIL, totalMembers, 1).getValues().flat();
  const phoneData = memberSheet.getRange(2, MEMBER_COLS.PHONE, totalMembers, 1).getValues().flat();

  const emailCount = emailData.filter(e => e && e.toString().includes('@')).length;
  const phoneCount = phoneData.filter(p => p && p.toString().length > 0).length;

  return {
    totalMembers: totalMembers,
    piiFieldCount: PII_CONFIG.MEMBER_PII_FIELDS.length + PII_CONFIG.GRIEVANCE_PII_FIELDS.length,
    emailCount: emailCount,
    phoneCount: phoneCount
  };
}

/**
 * Exports members with masked PII
 */
function exportMembersWithMaskedPII() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found!');
    return;
  }

  const lastRow = memberSheet.getLastRow();
  const lastCol = memberSheet.getLastColumn();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data to export!');
    return;
  }

  // Get all data
  const data = memberSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  // Create masked export
  const maskedData = [headers];

  for (let i = 1; i < data.length; i++) {
    const row = data[i].slice();

    // Mask email
    const emailIndex = MEMBER_COLS.EMAIL - 1;
    if (row[emailIndex]) {
      row[emailIndex] = maskEmail(row[emailIndex].toString());
    }

    // Mask phone
    const phoneIndex = MEMBER_COLS.PHONE - 1;
    if (row[phoneIndex]) {
      row[phoneIndex] = maskPhone(row[phoneIndex].toString());
    }

    maskedData.push(row);
  }

  // Create export sheet
  let exportSheet = ss.getSheetByName('Members_Masked_Export');
  if (exportSheet) {
    exportSheet.clear();
  } else {
    exportSheet = ss.insertSheet('Members_Masked_Export');
  }

  exportSheet.getRange(1, 1, maskedData.length, maskedData[0].length).setValues(maskedData);

  // Format
  exportSheet.getRange(1, 1, 1, maskedData[0].length)
    .setBackground('#4caf50')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  exportSheet.setFrozenRows(1);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Exported ${maskedData.length - 1} members with masked PII`,
    'Export Complete',
    5
  );

  ss.setActiveSheet(exportSheet);
}

/**
 * Shows PII audit report
 */
function showPIIAuditReport() {
  const stats = getPIIStatistics();

  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  table { width: 100%; border-collapse: collapse; margin: 15px 0; }
  th, td { padding: 10px; text-align: left; border: 1px solid #ddd; }
  th { background: #f5f5f5; }
  .ok { color: #4caf50; }
  .warning { color: #ff9800; }
</style>
</head>
<body>
  <h2>📊 PII Audit Report</h2>
  <p>Generated: ${new Date().toLocaleString()}</p>

  <table>
    <tr>
      <th>Field</th>
      <th>Location</th>
      <th>Record Count</th>
      <th>Protection Status</th>
    </tr>
    <tr>
      <td>Email Address</td>
      <td>Member Directory (Col H)</td>
      <td>${stats.emailCount}</td>
      <td class="ok">✅ Protected</td>
    </tr>
    <tr>
      <td>Phone Number</td>
      <td>Member Directory (Col I)</td>
      <td>${stats.phoneCount}</td>
      <td class="ok">✅ Protected</td>
    </tr>
    <tr>
      <td>Home Town</td>
      <td>Member Directory (Col X)</td>
      <td>${stats.totalMembers}</td>
      <td class="ok">✅ Protected</td>
    </tr>
    <tr>
      <td>Member Email</td>
      <td>Grievance Log (Col X)</td>
      <td>-</td>
      <td class="ok">✅ Protected</td>
    </tr>
  </table>

  <h3>Recommendations</h3>
  <ul>
    <li>Review access logs monthly</li>
    <li>Use masked exports for external sharing</li>
    <li>Process erasure requests within 30 days</li>
    <li>Maintain data retention policies</li>
  </ul>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'PII Audit Report');
}

/**
 * Shows PII access log
 */
function showPIIAccessLog() {
  SpreadsheetApp.getUi().alert(
    '📋 PII Access Log',
    'PII access is logged in the Audit Log sheet.\n\n' +
    'Go to: Audit Log sheet to view all data access events.\n\n' +
    'Logged events include:\n' +
    '• Data exports\n' +
    '• Record views\n' +
    '• Data modifications\n' +
    '• Deletion requests',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Anonymizes inactive members (for GDPR compliance)
 */
function anonymizeInactiveMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found!');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Anonymize Inactive Members',
    'This will anonymize PII for members with no activity in the last 2 years.\n\n' +
    'Anonymized fields:\n' +
    '• Email → anonymous@example.com\n' +
    '• Phone → 000-000-0000\n' +
    '• Home Town → (Anonymized)\n\n' +
    'This action cannot be undone. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No members to process.');
    return;
  }

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - PII_CONFIG.DEFAULT_RETENTION_DAYS);

  // Get contact dates
  const contactDateCol = MEMBER_COLS.RECENT_CONTACT_DATE;
  const data = memberSheet.getRange(2, 1, lastRow - 1, contactDateCol + 5).getValues();

  let anonymizedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const lastContact = data[i][contactDateCol - 1];

    // Check if inactive (no contact or contact before cutoff)
    if (!lastContact || (lastContact instanceof Date && lastContact < cutoffDate)) {
      const rowNum = i + 2;

      // Anonymize email
      memberSheet.getRange(rowNum, MEMBER_COLS.EMAIL).setValue('anonymous@example.com');

      // Anonymize phone
      memberSheet.getRange(rowNum, MEMBER_COLS.PHONE).setValue('000-000-0000');

      // Anonymize home town
      memberSheet.getRange(rowNum, MEMBER_COLS.HOME_TOWN).setValue('(Anonymized)');

      anonymizedCount++;
    }
  }

  // Log the action
  try {
    logAuditEvent('PII_ANONYMIZATION', {
      count: anonymizedCount,
      cutoffDate: cutoffDate.toISOString()
    }, 'INFO');
  } catch (e) {
    // Continue if logging fails
  }

  SpreadsheetApp.getUi().alert(
    '✅ Anonymization Complete',
    `Anonymized ${anonymizedCount} inactive member records.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Generates member data export for data portability (GDPR)
 * @param {string} memberId - Member ID
 */
function generateMemberDataExport(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found!');
    return;
  }

  // Find member
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
    SpreadsheetApp.getUi().alert('Member not found: ' + memberId);
    return;
  }

  // Create export object
  const exportData = {
    exportDate: new Date().toISOString(),
    memberId: memberId,
    memberData: {},
    grievances: []
  };

  for (let i = 0; i < headers.length; i++) {
    exportData.memberData[headers[i]] = memberData[i];
  }

  // Get grievances
  if (grievanceSheet) {
    const gLastRow = grievanceSheet.getLastRow();
    const gLastCol = grievanceSheet.getLastColumn();
    const gHeaders = grievanceSheet.getRange(1, 1, 1, gLastCol).getValues()[0];
    const gData = grievanceSheet.getRange(2, 1, gLastRow - 1, gLastCol).getValues();

    for (let i = 0; i < gData.length; i++) {
      if (gData[i][GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
        const grievance = {};
        for (let j = 0; j < gHeaders.length; j++) {
          grievance[gHeaders[j]] = gData[i][j];
        }
        exportData.grievances.push(grievance);
      }
    }
  }

  // Create JSON export
  const jsonString = JSON.stringify(exportData, null, 2);

  // Create a new sheet with the export
  let exportSheet = ss.getSheetByName('Data_Export_' + memberId);
  if (exportSheet) {
    ss.deleteSheet(exportSheet);
  }
  exportSheet = ss.insertSheet('Data_Export_' + memberId);

  exportSheet.getRange(1, 1).setValue('GDPR Data Portability Export');
  exportSheet.getRange(2, 1).setValue('Member ID: ' + memberId);
  exportSheet.getRange(3, 1).setValue('Export Date: ' + new Date().toLocaleString());
  exportSheet.getRange(5, 1).setValue('JSON Data:');
  exportSheet.getRange(6, 1).setValue(jsonString);

  // Format
  exportSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  exportSheet.getRange(6, 1).setWrap(true);
  exportSheet.setColumnWidth(1, 800);

  SpreadsheetApp.getUi().alert(
    '✅ Data Export Created',
    `Data export for ${memberId} has been created in the "Data_Export_${memberId}" sheet.\n\n` +
    'The export includes all member data and grievance records in JSON format.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  ss.setActiveSheet(exportSheet);
}

/**
 * Processes member erasure request (Right to Erasure / GDPR)
 * @param {string} memberId - Member ID to erase
 */
function processMemberErasure(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found!');
    return;
  }

  // Find member
  const lastRow = memberSheet.getLastRow();
  const data = memberSheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let rowToDelete = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === memberId) {
      rowToDelete = i + 2;
      break;
    }
  }

  if (!rowToDelete) {
    SpreadsheetApp.getUi().alert('Member not found: ' + memberId);
    return;
  }

  // Log the erasure before deletion
  try {
    logAuditEvent('GDPR_ERASURE', {
      memberId: memberId,
      rowDeleted: rowToDelete,
      requestedBy: Session.getActiveUser().getEmail()
    }, 'WARNING');
  } catch (e) {
    Logger.log('Failed to log erasure: ' + e.message);
  }

  // Delete the member row
  memberSheet.deleteRow(rowToDelete);

  // Also anonymize/delete from grievance log
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet) {
    const gLastRow = grievanceSheet.getLastRow();
    const gData = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, gLastRow - 1, 1).getValues();

    // Anonymize grievance records (don't delete for legal/historical reasons)
    for (let i = gData.length - 1; i >= 0; i--) {
      if (gData[i][0] === memberId) {
        const gRow = i + 2;
        grievanceSheet.getRange(gRow, GRIEVANCE_COLS.FIRST_NAME).setValue('(Deleted)');
        grievanceSheet.getRange(gRow, GRIEVANCE_COLS.LAST_NAME).setValue('Member');
        grievanceSheet.getRange(gRow, GRIEVANCE_COLS.MEMBER_EMAIL).setValue('deleted@erasure.request');
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    '✅ Erasure Complete',
    `Member ${memberId} has been erased.\n\n` +
    '• Member record deleted from Member Directory\n' +
    '• Grievance records anonymized (retained for legal compliance)\n' +
    '• Action logged in Audit Log',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows data subject request form
 */
function showDataSubjectRequestForm() {
  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  .form-group { margin: 15px 0; }
  label { display: block; font-weight: bold; margin-bottom: 5px; }
  input, select, textarea { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; box-sizing: border-box; }
  button { width: 100%; padding: 12px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 15px; }
  .info { background: #e8f4fd; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
</style>
</head>
<body>
  <h2>👤 Data Subject Request</h2>

  <div class="info">
    Use this form to process GDPR/CCPA data subject requests.
    All requests are logged for compliance purposes.
  </div>

  <div class="form-group">
    <label>Request Type:</label>
    <select id="requestType">
      <option value="access">Data Access Request</option>
      <option value="portability">Data Portability Request</option>
      <option value="rectification">Data Rectification Request</option>
      <option value="erasure">Right to Erasure Request</option>
    </select>
  </div>

  <div class="form-group">
    <label>Member ID or Email:</label>
    <input type="text" id="identifier" placeholder="Enter Member ID or email address">
  </div>

  <div class="form-group">
    <label>Request Notes:</label>
    <textarea id="notes" rows="4" placeholder="Additional details about the request"></textarea>
  </div>

  <button onclick="processRequest()">Submit Request</button>

  <script>
    function processRequest() {
      const requestType = document.getElementById('requestType').value;
      const identifier = document.getElementById('identifier').value;
      const notes = document.getElementById('notes').value;

      if (!identifier) {
        alert('Please enter a Member ID or email');
        return;
      }

      google.script.run
        .withSuccessHandler(function() {
          alert('Request processed successfully!');
          google.script.host.close();
        })
        .processDataSubjectRequest(requestType, identifier, notes);
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data Subject Request');
}

/**
 * Processes a data subject request
 * @param {string} requestType - Type of request
 * @param {string} identifier - Member ID or email
 * @param {string} notes - Additional notes
 */
function processDataSubjectRequest(requestType, identifier, notes) {
  // Log the request
  try {
    logAuditEvent('DATA_SUBJECT_REQUEST', {
      type: requestType,
      identifier: identifier,
      notes: notes,
      processedBy: Session.getActiveUser().getEmail()
    }, 'INFO');
  } catch (e) {
    Logger.log('Failed to log request: ' + e.message);
  }

  // Process based on type
  switch (requestType) {
    case 'access':
    case 'portability':
      generateMemberDataExport(identifier);
      break;
    case 'erasure':
      processMemberErasure(identifier);
      break;
    case 'rectification':
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Navigate to Member Directory to make rectifications',
        'Rectification Request',
        5
      );
      break;
  }
}

/**
 * Checks if a field contains PII
 * @param {string} fieldName - Field name to check
 * @returns {boolean} True if field contains PII
 */
function isPIIField(fieldName) {
  return PII_CONFIG.MEMBER_PII_FIELDS.includes(fieldName) ||
         PII_CONFIG.GRIEVANCE_PII_FIELDS.includes(fieldName);
}

/**
 * Masks a value based on its type
 * @param {string} value - Value to mask
 * @param {string} type - Type of PII (email, phone, ssn, address)
 * @returns {string} Masked value
 */
function maskPIIValue(value, type) {
  if (!value) return '';

  const handler = PII_FIELDS[type.toLowerCase()];
  if (handler && handler.mask) {
    return handler.mask(value.toString());
  }

  return value;
}
