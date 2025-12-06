/**
 * ============================================================================
 * ENHANCED VALIDATION SYSTEM
 * ============================================================================
 *
 * Comprehensive email and phone validation with real-time feedback.
 *
 * Features:
 * - Email format validation
 * - Phone number format validation (US/International)
 * - Duplicate detection with warnings
 * - Real-time validation on edit
 * - Validation report generation
 * - Bulk validation tool
 *
 * @module EnhancedValidation
 * @version 1.0.0
 * @author SEIU Local 509 Tech Team
 */

/**
 * Validation patterns
 */
const VALIDATION_PATTERNS = {
  // Standard email pattern
  EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,

  // US phone patterns
  PHONE_US: /^[\+]?1?[-.\s]?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4}$/,

  // International phone (basic)
  PHONE_INTL: /^[\+]?[0-9]{1,4}[-.\s]?\(?[0-9]{1,4}\)?[-.\s]?[0-9]{1,4}[-.\s]?[0-9]{1,9}$/,

  // Member ID format
  MEMBER_ID: /^M\d{6}$/,

  // Grievance ID format
  GRIEVANCE_ID: /^G-\d{6}(-[A-Z])?$/
};

/**
 * Validation error messages
 */
const VALIDATION_MESSAGES = {
  EMAIL_INVALID: 'Invalid email format. Please use format: name@domain.com',
  EMAIL_EMPTY: 'Email address is required',
  PHONE_INVALID: 'Invalid phone format. Please use format: (555) 555-1234 or 555-555-1234',
  PHONE_EMPTY: 'Phone number is required',
  MEMBER_ID_INVALID: 'Invalid Member ID format. Please use format: M123456',
  MEMBER_ID_DUPLICATE: 'This Member ID already exists',
  GRIEVANCE_ID_INVALID: 'Invalid Grievance ID format. Please use format: G-123456',
  GRIEVANCE_ID_DUPLICATE: 'This Grievance ID already exists'
};

/**
 * Validates an email address
 * @param {string} email - Email to validate
 * @returns {Object} {valid: boolean, message: string}
 */
function validateEmailAddress(email) {
  if (!email || email.toString().trim() === '') {
    return { valid: false, message: VALIDATION_MESSAGES.EMAIL_EMPTY };
  }

  const cleanEmail = email.toString().trim().toLowerCase();

  if (!VALIDATION_PATTERNS.EMAIL.test(cleanEmail)) {
    return { valid: false, message: VALIDATION_MESSAGES.EMAIL_INVALID };
  }

  // Additional checks
  const parts = cleanEmail.split('@');
  if (parts.length !== 2) {
    return { valid: false, message: VALIDATION_MESSAGES.EMAIL_INVALID };
  }

  const domain = parts[1];
  const domainParts = domain.split('.');

  // Check for valid TLD (at least 2 chars)
  if (domainParts[domainParts.length - 1].length < 2) {
    return { valid: false, message: 'Invalid domain extension' };
  }

  // Check for common typos
  const commonDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com'];
  const commonTypos = {
    'gmial.com': 'gmail.com',
    'gmai.com': 'gmail.com',
    'gamil.com': 'gmail.com',
    'yaho.com': 'yahoo.com',
    'yahooo.com': 'yahoo.com',
    'hotmal.com': 'hotmail.com',
    'outlok.com': 'outlook.com'
  };

  if (commonTypos[domain]) {
    return {
      valid: true,
      message: `Did you mean ${parts[0]}@${commonTypos[domain]}?`,
      suggestion: `${parts[0]}@${commonTypos[domain]}`
    };
  }

  return { valid: true, message: 'Valid email address' };
}

/**
 * Validates a phone number
 * @param {string} phone - Phone number to validate
 * @returns {Object} {valid: boolean, message: string, formatted: string}
 */
function validatePhoneNumber(phone) {
  if (!phone || phone.toString().trim() === '') {
    return { valid: false, message: VALIDATION_MESSAGES.PHONE_EMPTY };
  }

  const cleanPhone = phone.toString().trim();

  // Remove all non-digit characters for validation
  const digits = cleanPhone.replace(/\D/g, '');

  // Check length
  if (digits.length < 10) {
    return { valid: false, message: 'Phone number must have at least 10 digits' };
  }

  if (digits.length > 15) {
    return { valid: false, message: 'Phone number is too long' };
  }

  // Validate US format
  if (digits.length === 10 || (digits.length === 11 && digits[0] === '1')) {
    const formatted = formatUSPhone(digits);
    return { valid: true, message: 'Valid phone number', formatted: formatted };
  }

  // Accept international
  if (digits.length > 10) {
    return { valid: true, message: 'Valid international phone number', formatted: cleanPhone };
  }

  return { valid: false, message: VALIDATION_MESSAGES.PHONE_INVALID };
}

/**
 * Formats a US phone number
 * @param {string} digits - Phone digits only
 * @returns {string} Formatted phone number
 */
function formatUSPhone(digits) {
  // Remove country code if present
  if (digits.length === 11 && digits[0] === '1') {
    digits = digits.substring(1);
  }

  if (digits.length === 10) {
    return `(${digits.substring(0, 3)}) ${digits.substring(3, 6)}-${digits.substring(6)}`;
  }

  return digits;
}

/**
 * Validates data on edit trigger
 * @param {Object} e - Edit event
 */
function onEditValidation(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    return;
  }

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (row < 2) return; // Skip header

  const value = e.value;

  // Member Directory validations
  if (sheetName === SHEETS.MEMBER_DIR) {
    // Email validation
    if (col === MEMBER_COLS.EMAIL && value) {
      const result = validateEmailAddress(value);
      if (!result.valid) {
        e.range.setNote('⚠️ ' + result.message);
        e.range.setBackground('#fff3e0');
      } else if (result.suggestion) {
        e.range.setNote('💡 ' + result.message);
        e.range.setBackground('#e3f2fd');
      } else {
        e.range.clearNote();
        e.range.setBackground(null);
      }
    }

    // Phone validation
    if (col === MEMBER_COLS.PHONE && value) {
      const result = validatePhoneNumber(value);
      if (!result.valid) {
        e.range.setNote('⚠️ ' + result.message);
        e.range.setBackground('#fff3e0');
      } else {
        e.range.clearNote();
        e.range.setBackground(null);
        if (result.formatted && result.formatted !== value) {
          e.range.setValue(result.formatted);
        }
      }
    }

    // Member ID duplicate check
    if (col === MEMBER_COLS.MEMBER_ID && value) {
      const isDuplicate = checkDuplicateMemberID(value);
      if (isDuplicate) {
        e.range.setNote('⚠️ ' + VALIDATION_MESSAGES.MEMBER_ID_DUPLICATE);
        e.range.setBackground('#ffebee');
      } else if (!VALIDATION_PATTERNS.MEMBER_ID.test(value)) {
        e.range.setNote('⚠️ ' + VALIDATION_MESSAGES.MEMBER_ID_INVALID);
        e.range.setBackground('#fff3e0');
      } else {
        e.range.clearNote();
        e.range.setBackground(null);
      }
    }
  }

  // Grievance Log validations
  if (sheetName === SHEETS.GRIEVANCE_LOG) {
    // Grievance ID duplicate check
    if (col === GRIEVANCE_COLS.GRIEVANCE_ID && value) {
      const isDuplicate = checkDuplicateGrievanceID(value);
      if (isDuplicate) {
        e.range.setNote('⚠️ ' + VALIDATION_MESSAGES.GRIEVANCE_ID_DUPLICATE);
        e.range.setBackground('#ffebee');
      } else if (!VALIDATION_PATTERNS.GRIEVANCE_ID.test(value)) {
        e.range.setNote('⚠️ ' + VALIDATION_MESSAGES.GRIEVANCE_ID_INVALID);
        e.range.setBackground('#fff3e0');
      } else {
        e.range.clearNote();
        e.range.setBackground(null);
      }
    }
  }
}

/**
 * Runs bulk validation on Member Directory
 */
function runBulkValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) {
    SpreadsheetApp.getUi().alert('Member Directory not found!');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Running bulk validation...',
    'Please wait',
    -1
  );

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No data to validate', 'Info', 3);
    return;
  }

  const emailData = memberSheet.getRange(2, MEMBER_COLS.EMAIL, lastRow - 1, 1).getValues();
  const phoneData = memberSheet.getRange(2, MEMBER_COLS.PHONE, lastRow - 1, 1).getValues();
  const memberIdData = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1).getValues();

  const issues = [];
  const seenMemberIds = {};

  for (let i = 0; i < lastRow - 1; i++) {
    const row = i + 2;
    const email = emailData[i][0];
    const phone = phoneData[i][0];
    const memberId = memberIdData[i][0];

    // Validate email
    if (email) {
      const emailResult = validateEmailAddress(email);
      if (!emailResult.valid) {
        issues.push({ row, field: 'Email', value: email, message: emailResult.message });
      }
    }

    // Validate phone
    if (phone) {
      const phoneResult = validatePhoneNumber(phone);
      if (!phoneResult.valid) {
        issues.push({ row, field: 'Phone', value: phone, message: phoneResult.message });
      }
    }

    // Check for duplicate Member IDs
    if (memberId) {
      if (seenMemberIds[memberId]) {
        issues.push({
          row,
          field: 'Member ID',
          value: memberId,
          message: `Duplicate of row ${seenMemberIds[memberId]}`
        });
      } else {
        seenMemberIds[memberId] = row;
      }
    }
  }

  // Show results
  showValidationReport(issues, lastRow - 1);
}

/**
 * Shows validation report
 * @param {Array} issues - Array of validation issues
 * @param {number} totalRecords - Total records validated
 */
function showValidationReport(issues, totalRecords) {
  const issueCount = issues.length;
  const passRate = totalRecords > 0 ? (((totalRecords - issueCount) / totalRecords) * 100).toFixed(1) : 100;

  let issuesHtml = '';
  if (issues.length > 0) {
    issuesHtml = issues.slice(0, 50).map(issue => `
      <tr>
        <td>${issue.row}</td>
        <td>${issue.field}</td>
        <td>${issue.value}</td>
        <td>${issue.message}</td>
      </tr>
    `).join('');

    if (issues.length > 50) {
      issuesHtml += `<tr><td colspan="4">...and ${issues.length - 50} more issues</td></tr>`;
    }
  }

  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  .summary { display: flex; gap: 20px; margin: 20px 0; }
  .stat-card { flex: 1; padding: 20px; border-radius: 8px; text-align: center; }
  .stat-card.good { background: #e8f5e9; border-left: 4px solid #4caf50; }
  .stat-card.warning { background: #fff3e0; border-left: 4px solid #ff9800; }
  .stat-card.bad { background: #ffebee; border-left: 4px solid #f44336; }
  .stat-value { font-size: 32px; font-weight: bold; color: #333; }
  .stat-label { color: #666; margin-top: 5px; }
  table { width: 100%; border-collapse: collapse; margin-top: 20px; }
  th, td { padding: 10px; text-align: left; border: 1px solid #ddd; }
  th { background: #f5f5f5; }
  .no-issues { text-align: center; padding: 40px; color: #4caf50; font-size: 18px; }
</style>
</head>
<body>
  <h2>📊 Validation Report</h2>

  <div class="summary">
    <div class="stat-card ${issueCount === 0 ? 'good' : issueCount < 10 ? 'warning' : 'bad'}">
      <div class="stat-value">${passRate}%</div>
      <div class="stat-label">Pass Rate</div>
    </div>
    <div class="stat-card ${issueCount === 0 ? 'good' : 'warning'}">
      <div class="stat-value">${totalRecords}</div>
      <div class="stat-label">Records Validated</div>
    </div>
    <div class="stat-card ${issueCount === 0 ? 'good' : 'bad'}">
      <div class="stat-value">${issueCount}</div>
      <div class="stat-label">Issues Found</div>
    </div>
  </div>

  ${issueCount > 0 ? `
    <table>
      <tr>
        <th>Row</th>
        <th>Field</th>
        <th>Value</th>
        <th>Issue</th>
      </tr>
      ${issuesHtml}
    </table>
  ` : '<div class="no-issues">✅ No validation issues found!</div>'}
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Validation Report');
}

/**
 * Shows validation settings dialog
 */
function showValidationSettings() {
  const html = `
<!DOCTYPE html>
<html>
<head><base target="_top">
<style>
  body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
  h2 { color: #1a73e8; margin-top: 0; }
  .setting { margin: 15px 0; padding: 15px; background: #f8f9fa; border-radius: 8px; }
  .setting-header { display: flex; justify-content: space-between; align-items: center; }
  .setting-title { font-weight: bold; }
  .setting-desc { color: #666; font-size: 13px; margin-top: 5px; }
  .toggle { position: relative; width: 50px; height: 26px; }
  .toggle input { opacity: 0; width: 0; height: 0; }
  .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background: #ccc; border-radius: 26px; transition: .4s; }
  .slider:before { position: absolute; content: ""; height: 20px; width: 20px; left: 3px; bottom: 3px; background: white; border-radius: 50%; transition: .4s; }
  input:checked + .slider { background: #4caf50; }
  input:checked + .slider:before { transform: translateX(24px); }
  button { width: 100%; padding: 12px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 20px; }
</style>
</head>
<body>
  <h2>⚙️ Validation Settings</h2>

  <div class="setting">
    <div class="setting-header">
      <div>
        <div class="setting-title">Real-time Email Validation</div>
        <div class="setting-desc">Validate email format as you type</div>
      </div>
      <label class="toggle">
        <input type="checkbox" id="emailValidation" checked>
        <span class="slider"></span>
      </label>
    </div>
  </div>

  <div class="setting">
    <div class="setting-header">
      <div>
        <div class="setting-title">Auto-format Phone Numbers</div>
        <div class="setting-desc">Automatically format phone numbers to (XXX) XXX-XXXX</div>
      </div>
      <label class="toggle">
        <input type="checkbox" id="phoneFormat" checked>
        <span class="slider"></span>
      </label>
    </div>
  </div>

  <div class="setting">
    <div class="setting-header">
      <div>
        <div class="setting-title">Duplicate ID Detection</div>
        <div class="setting-desc">Warn when entering duplicate Member or Grievance IDs</div>
      </div>
      <label class="toggle">
        <input type="checkbox" id="duplicateCheck" checked>
        <span class="slider"></span>
      </label>
    </div>
  </div>

  <div class="setting">
    <div class="setting-header">
      <div>
        <div class="setting-title">Visual Indicators</div>
        <div class="setting-desc">Highlight cells with validation issues</div>
      </div>
      <label class="toggle">
        <input type="checkbox" id="visualIndicators" checked>
        <span class="slider"></span>
      </label>
    </div>
  </div>

  <button onclick="runValidation()">🔍 Run Bulk Validation Now</button>

  <script>
    function runValidation() {
      google.script.run.runBulkValidation();
      google.script.host.close();
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Validation Settings');
}

/**
 * Clears all validation notes and highlighting
 */
function clearValidationIndicators() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return;

  // Clear email column notes and background
  const emailRange = memberSheet.getRange(2, MEMBER_COLS.EMAIL, lastRow - 1, 1);
  emailRange.clearNote();
  emailRange.setBackground(null);

  // Clear phone column notes and background
  const phoneRange = memberSheet.getRange(2, MEMBER_COLS.PHONE, lastRow - 1, 1);
  phoneRange.clearNote();
  phoneRange.setBackground(null);

  // Clear member ID column notes and background
  const memberIdRange = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastRow - 1, 1);
  memberIdRange.clearNote();
  memberIdRange.setBackground(null);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Validation indicators cleared',
    'Done',
    3
  );
}

/**
 * Installs validation trigger
 */
function installValidationTrigger() {
  // Remove existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditValidation') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger
  ScriptApp.newTrigger('onEditValidation')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Validation Trigger Installed',
    'Real-time validation is now active.\n\n' +
    'Email and phone numbers will be validated as you type.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
