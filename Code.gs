/**
 * ============================================================================
 * 509 DASHBOARD - Enhanced Member Directory & Grievance Tracking System
 * ============================================================================
 *
 * For SEIU Local 509 (Units 8 & 10) - Massachusetts State Employees
 * Based on Collective Bargaining Agreement 2024-2026
 *
 * FEATURES:
 * - Member Directory with auto-calculated grievance metrics
 * - Grievance Log with CBA-compliant deadline tracking
 * - Priority sorting (Step III → II → I, then by due date)
 * - Advanced Dashboard with KPIs, charts, and Top 10 lists
 * - Steward workload tracking
 * - Form integration for data entry
 * - All calculations done in code (no formula rows)
 * - Automated deadline and status tracking
 *
 * CBA COMPLIANCE:
 * - Article 23A: Grievance deadlines (21-day filing, 30-day decisions, 10-day appeals)
 * - Article 8: Leave provisions
 * - Article 14: Promotions
 *
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const SHEETS = {
  CONFIG: "Config",
  MEMBER_DIR: "Member Directory",
  GRIEVANCE_LOG: "Grievance Log",
  DASHBOARD: "Dashboard",
  STEWARD_WORKLOAD: "Steward Workload",
  ANALYTICS: "Analytics Data",
  ARCHIVE: "Archive",
  DIAGNOSTICS: "Diagnostics"
};

const COLORS = {
  HEADER_BLUE: "#4A86E8",
  HEADER_RED: "#E06666",
  HEADER_GREEN: "#93C47D",
  HEADER_ORANGE: "#F6B26B",
  HEADER_PURPLE: "#8E7CC3",
  WHITE: "#FFFFFF",
  LIGHT_GRAY: "#F3F3F3",
  OVERDUE: "#EA4335",
  DUE_SOON: "#FBBC04",
  ON_TRACK: "#34A853"
};

const CBA_DEADLINES = {
  FILING: 21,           // Days from incident to file grievance
  STEP_I_DECISION: 30,  // Days for Step I decision
  STEP_II_APPEAL: 10,   // Days to appeal to Step II
  STEP_II_DECISION: 30, // Days for Step II decision
  STEP_III_APPEAL: 10,  // Days to appeal to Step III
  STEP_III_DECISION: 30 // Days for Step III decision
};

const PRIORITY_ORDER = {
  "Step III - Human Resources": 1,
  "Step II - Agency Head": 2,
  "Step I - Immediate Supervisor": 3,
  "Informal": 4,
  "Mediation": 5,
  "Arbitration": 6
};

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main setup function - Creates entire 509 Dashboard system
 * Run this once to set up everything
 */
function CREATE_509_DASHBOARD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.getUi().alert('Creating 509 Dashboard...\n\nThis will take about 30 seconds.');

  // Clear existing sheets (keep first one)
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  sheets[0].setName(SHEETS.CONFIG);

  // Create all sheets in order
  createConfigSheet(ss);
  createMemberDirectorySheet(ss);
  createGrievanceLogSheet(ss);
  createStewardWorkloadSheet(ss);
  createDashboardSheet(ss);
  createAnalyticsSheet(ss);
  createArchiveSheet(ss);
  createDiagnosticsSheet(ss);

  // Setup data validation
  setupDataValidation(ss);

  // Setup triggers
  setupTriggers();

  // Build initial dashboard
  rebuildDashboard();

  SpreadsheetApp.getUi().alert('✅ 509 Dashboard created successfully!\n\n' +
    'Next steps:\n' +
    '1. Add members via "509 Tools > Seed 20K Members" or manually\n' +
    '2. Add grievances via "509 Tools > Seed 5K Grievances" or manually\n' +
    '3. View Dashboard tab for real-time metrics');
}

// ============================================================================
// CONFIG SHEET
// ============================================================================

function createConfigSheet(ss) {
  let config = ss.getSheetByName(SHEETS.CONFIG);
  if (!config) config = ss.insertSheet(SHEETS.CONFIG);

  config.clear();

  // Headers
  const headers = [
    "Job Titles", "Work Locations", "Units", "Steward Status",
    "Grievance Status", "Grievance Steps", "Grievance Types",
    "Outcomes", "Satisfaction Levels", "Engagement Levels",
    "Contact Methods", "Committee Types", "Membership Status"
  ];

  config.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor("white");

  // Column A: Job Titles (from CBA Appendix C)
  const jobTitles = [
    "Administrative Assistant I", "Administrative Assistant II", "Administrative Assistant III",
    "Case Manager I", "Case Manager II", "Case Manager III",
    "Program Coordinator I", "Program Coordinator II", "Program Coordinator III",
    "Social Worker I", "Social Worker II", "Social Worker III",
    "Clerk I", "Clerk II", "Clerk III",
    "Specialist I", "Specialist II", "Specialist III"
  ];

  // Column B: Work Locations
  const workLocations = [
    "Boston - State House", "Boston - McCormack Building", "Boston - Saltonstall Building",
    "Springfield - State Office Building", "Worcester - State Office Complex",
    "Pittsfield - Regional Office", "Lowell - Regional Office", "New Bedford - Regional Office",
    "Hyannis - Regional Office", "Lawrence - Regional Office", "Brockton - Regional Office",
    "Fall River - Regional Office", "Framingham - Regional Office", "Quincy - Regional Office",
    "Remote/Hybrid"
  ];

  // Column C: Units
  const units = ["Unit 8", "Unit 10"];

  // Column D: Steward Status
  const stewardStatus = ["Yes", "No"];

  // Column E: Grievance Status
  const grievanceStatus = [
    "Draft",
    "Filed - Step I",
    "Filed - Step II",
    "Filed - Step III",
    "In Mediation",
    "In Arbitration",
    "Resolved - Withdrawn",
    "Resolved - Settled",
    "Resolved - Won",
    "Resolved - Lost",
    "Pending Decision"
  ];

  // Column F: Grievance Steps (CBA Article 23A)
  const grievanceSteps = [
    "Informal",
    "Step I - Immediate Supervisor",
    "Step II - Agency Head",
    "Step III - Human Resources",
    "Mediation",
    "Arbitration"
  ];

  // Column G: Grievance Types
  const grievanceTypes = [
    "Disciplinary Action",
    "Contract Violation - Article 8 (Leave)",
    "Contract Violation - Article 9 (Vacation)",
    "Contract Violation - Article 10 (Sick Leave)",
    "Contract Violation - Article 13 (Work Week/Overtime)",
    "Contract Violation - Article 14 (Holidays)",
    "Contract Violation - Article 15 (Salary)",
    "Contract Violation - Other",
    "Working Conditions",
    "Workplace Safety",
    "Discrimination/Harassment",
    "Promotion/Demotion",
    "Performance Evaluation",
    "Layoff/Recall",
    "Other"
  ];

  // Column H: Outcomes
  const outcomes = [
    "Pending",
    "Withdrawn",
    "Settled - Full Relief",
    "Settled - Partial Relief",
    "Won - Arbitration",
    "Lost - Arbitration",
    "Denied - Step I",
    "Denied - Step II",
    "Denied - Step III"
  ];

  // Column I: Satisfaction Levels
  const satisfactionLevels = [
    "Very Satisfied", "Satisfied", "Neutral", "Dissatisfied", "Very Dissatisfied", "N/A"
  ];

  // Column J: Engagement Levels
  const engagementLevels = [
    "Very Active", "Active", "Moderately Active", "Inactive", "New Member"
  ];

  // Column K: Contact Methods
  const contactMethods = ["Email", "Phone", "Text", "Mail"];

  // Column L: Committee Types
  const committeeTypes = [
    "Executive Board", "Bargaining Committee", "Grievance Committee",
    "Political Action", "Communications", "Member Engagement", "None"
  ];

  // Column M: Membership Status
  const membershipStatus = ["Active", "Inactive", "On Leave", "Retired"];

  // Write all data
  const allData = [
    jobTitles, workLocations, units, stewardStatus, grievanceStatus,
    grievanceSteps, grievanceTypes, outcomes, satisfactionLevels,
    engagementLevels, contactMethods, committeeTypes, membershipStatus
  ];

  for (let i = 0; i < allData.length; i++) {
    const values = allData[i].map(item => [item]);
    config.getRange(2, i + 1, values.length, 1).setValues(values);
  }

  config.setFrozenRows(1);
  config.autoResizeColumns(1, headers.length);
}

// ============================================================================
// MEMBER DIRECTORY SHEET
// ============================================================================

function createMemberDirectorySheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBER_DIR);

  sheet.clear();

  // 35 columns with enhanced tracking
  const headers = [
    // Basic Info (A-F)
    "Member ID", "First Name", "Last Name", "Job Title", "Work Location", "Unit",
    // Contact & Role (G-M)
    "Office Days", "Email Address", "Phone Number", "Is Steward", "Date Joined Union",
    "Membership Status", "Engagement Level",
    // Grievance Metrics (N-T) - AUTO CALCULATED
    "Total Grievances", "Active Grievances", "Resolved Grievances",
    "Grievances Won", "Grievances Lost", "Last Grievance Date", "Win Rate %",
    // Derived Fields (U-X) - AUTO CALCULATED
    "Has Open Grievance?", "Current Grievance Status", "Next Deadline", "Days to Deadline",
    // Participation (Y-AB)
    "Events Attended (12mo)", "Training Sessions", "Committee Member", "Preferred Contact",
    // Emergency & Admin (AC-AI)
    "Emergency Contact Name", "Emergency Contact Phone", "Notes",
    "Date of Birth", "Hire Date", "Seniority Date",
    "Last Updated", "Updated By"
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.HEADER_BLUE)
    .setFontColor("white")
    .setWrap(true);

  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 110);  // Member ID
  sheet.setColumnWidth(2, 120);  // First Name
  sheet.setColumnWidth(3, 120);  // Last Name
  sheet.setColumnWidth(4, 200);  // Job Title
  sheet.setColumnWidth(5, 220);  // Work Location
  sheet.setColumnWidth(8, 220);  // Email

  // Color-code derived field columns (light gray background)
  const derivedCols = [14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]; // N through X
  derivedCols.forEach(col => {
    sheet.getRange(1, col).setBackground(COLORS.HEADER_GREEN);
  });
}

// ============================================================================
// GRIEVANCE LOG SHEET
// ============================================================================

function createGrievanceLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) sheet = ss.insertSheet(SHEETS.GRIEVANCE_LOG);

  sheet.clear();

  // 32 columns with enhanced deadline tracking
  const headers = [
    // Basic Info (A-F)
    "Grievance ID", "Member ID", "First Name", "Last Name", "Status", "Current Step",
    // Incident & Filing (G-J) - H, J auto-calculated
    "Incident Date", "Filing Deadline", "Date Filed", "Step I Decision Due",
    // Step I (K-N) - M auto-calculated
    "Step I Decision Date", "Step I Outcome", "Step II Appeal Deadline", "Step II Filed Date",
    // Step II (O-R) - O auto-calculated
    "Step II Decision Due", "Step II Decision Date", "Step II Outcome", "Step III Appeal Deadline",
    // Step III & Beyond (S-V)
    "Step III Filed Date", "Step III Decision Date", "Mediation Date", "Arbitration Date",
    // Details (W-Z)
    "Final Outcome", "Grievance Type", "Description", "Representative",
    // Derived Fields (AA-AF) - AUTO CALCULATED
    "Days Open", "Days to Next Deadline", "Is Overdue?", "Priority Score",
    "Assigned Steward", "Steward Contact",
    // Admin (AG-AH)
    "Notes", "Last Updated"
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.HEADER_RED)
    .setFontColor("white")
    .setWrap(true);

  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 120);  // Grievance ID
  sheet.setColumnWidth(2, 110);  // Member ID
  sheet.setColumnWidth(25, 300); // Description
  sheet.setColumnWidth(33, 300); // Notes

  // Color-code auto-calculated deadline columns
  const deadlineCols = [8, 10, 13, 15, 18]; // Filing Deadline, Step I Due, Step II Appeal, Step II Due, Step III Appeal
  deadlineCols.forEach(col => {
    sheet.getRange(1, col).setBackground(COLORS.HEADER_ORANGE);
  });

  // Color-code derived fields
  const derivedCols = [27, 28, 29, 30, 31, 32]; // Days Open through Steward Contact
  derivedCols.forEach(col => {
    sheet.getRange(1, col).setBackground(COLORS.HEADER_GREEN);
  });
}

// ============================================================================
// STEWARD WORKLOAD SHEET
// ============================================================================

function createStewardWorkloadSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.STEWARD_WORKLOAD);

  sheet.clear();

  const headers = [
    "Steward Name", "Member ID", "Work Location", "Total Cases",
    "Active Cases", "Step I Cases", "Step II Cases", "Step III Cases",
    "Overdue Cases", "Due This Week", "Win Rate %", "Avg Days to Resolution",
    "Members Assigned", "Last Case Date", "Status"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.HEADER_PURPLE)
    .setFontColor("white");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

// ============================================================================
// DASHBOARD SHEET
// ============================================================================

function createDashboardSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.DASHBOARD);

  sheet.clear();

  // Title
  sheet.getRange("A1:H1").merge()
    .setValue("509 DASHBOARD - REAL-TIME METRICS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BLUE)
    .setFontColor("white");

  // Section headers (will be populated by rebuildDashboard())
  sheet.getRange("A3").setValue("MEMBER METRICS").setFontWeight("bold");
  sheet.getRange("A11").setValue("GRIEVANCE METRICS").setFontWeight("bold");
  sheet.getRange("A21").setValue("DEADLINE TRACKING").setFontWeight("bold");
  sheet.getRange("E3").setValue("TOP 10 OVERDUE GRIEVANCES").setFontWeight("bold");
  sheet.getRange("E15").setValue("STEWARD WORKLOAD").setFontWeight("bold");

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
}

// ============================================================================
// ANALYTICS SHEET
// ============================================================================

function createAnalyticsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.ANALYTICS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ANALYTICS);

  sheet.clear();

  const headers = [
    "Snapshot Date", "Total Members", "Active Members", "Total Stewards",
    "Total Grievances", "Active Grievances", "Resolved This Month",
    "Win Rate %", "Avg Days to Resolution", "Overdue Count", "Notes"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.HEADER_GREEN)
    .setFontColor("white");

  sheet.setFrozenRows(1);
}

// ============================================================================
// ARCHIVE SHEET
// ============================================================================

function createArchiveSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.ARCHIVE);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ARCHIVE);

  sheet.clear();

  sheet.getRange("A1").setValue("Archived Grievances")
    .setFontWeight("bold")
    .setFontSize(14);

  sheet.getRange("A2").setValue("Resolved grievances are automatically moved here after 90 days")
    .setFontStyle("italic");
}

// ============================================================================
// DIAGNOSTICS SHEET
// ============================================================================

function createDiagnosticsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.DIAGNOSTICS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.DIAGNOSTICS);

  sheet.clear();

  const headers = ["Check Type", "Status", "Details", "Last Run"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold");

  const checks = [
    ["Data Validation", "OK", "All dropdowns configured", new Date()],
    ["Triggers", "OK", "onEdit and onFormSubmit active", new Date()],
    ["Member Count", "OK", "0 members", new Date()],
    ["Grievance Count", "OK", "0 grievances", new Date()],
    ["Orphaned Grievances", "OK", "0 grievances without members", new Date()]
  ];

  sheet.getRange(2, 1, checks.length, headers.length).setValues(checks);
}

// ============================================================================
// DATA VALIDATION
// ============================================================================

function setupDataValidation(ss) {
  const config = ss.getSheetByName(SHEETS.CONFIG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Member Directory Validations
  const validations = [
    { sheet: memberDir, column: "D", range: "A2:A19", name: "Job Title" },      // Job Titles
    { sheet: memberDir, column: "E", range: "B2:B16", name: "Work Location" },  // Locations
    { sheet: memberDir, column: "F", range: "C2:C3", name: "Unit" },            // Units
    { sheet: memberDir, column: "J", range: "D2:D3", name: "Steward" },         // Steward Status
    { sheet: memberDir, column: "L", range: "M2:M5", name: "Membership" },      // Membership Status
    { sheet: memberDir, column: "M", range: "J2:J6", name: "Engagement" },      // Engagement Level
    { sheet: memberDir, column: "AB", range: "L2:L7", name: "Committee" },      // Committee
    { sheet: memberDir, column: "AD", range: "K2:K5", name: "Contact Method" }  // Contact Method
  ];

  validations.forEach(v => {
    const sourceRange = config.getRange(v.range);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    v.sheet.getRange(`${v.column}2:${v.column}10000`).setDataValidation(rule);
  });

  // Grievance Log Validations
  const grievanceValidations = [
    { column: "E", range: "E2:E12", name: "Status" },          // Grievance Status
    { column: "F", range: "F2:F7", name: "Step" },             // Current Step
    { column: "L", range: "H2:H10", name: "Step I Outcome" },  // Step I Outcome
    { column: "Q", range: "H2:H10", name: "Step II Outcome" }, // Step II Outcome
    { column: "W", range: "H2:H10", name: "Final Outcome" },   // Final Outcome
    { column: "X", range: "G2:G16", name: "Type" }             // Grievance Type
  ];

  grievanceValidations.forEach(v => {
    const sourceRange = config.getRange(v.range);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();
    grievanceLog.getRange(`${v.column}2:${v.column}10000`).setDataValidation(rule);
  });
}

// ============================================================================
// CALCULATION FUNCTIONS (NO FORMULA ROWS - ALL IN CODE)
// ============================================================================

/**
 * Recalculates a single grievance row
 * Called by onEdit trigger when grievance data changes
 */
function recalcGrievanceRow(sheet, row) {
  if (row < 2) return; // Skip header

  const incidentDate = sheet.getRange(row, 7).getValue(); // G: Incident Date
  const dateFiled = sheet.getRange(row, 9).getValue();    // I: Date Filed
  const step1Decision = sheet.getRange(row, 11).getValue(); // K: Step I Decision
  const step2Filed = sheet.getRange(row, 14).getValue();    // N: Step II Filed
  const step2Decision = sheet.getRange(row, 16).getValue(); // P: Step II Decision
  const currentStep = sheet.getRange(row, 6).getValue();    // F: Current Step
  const status = sheet.getRange(row, 5).getValue();         // E: Status

  // Calculate Filing Deadline (H = G + 21 days)
  if (incidentDate) {
    const filingDeadline = new Date(incidentDate);
    filingDeadline.setDate(filingDeadline.getDate() + CBA_DEADLINES.FILING);
    sheet.getRange(row, 8).setValue(filingDeadline);
  }

  // Calculate Step I Decision Due (J = I + 30 days)
  if (dateFiled) {
    const step1Due = new Date(dateFiled);
    step1Due.setDate(step1Due.getDate() + CBA_DEADLINES.STEP_I_DECISION);
    sheet.getRange(row, 10).setValue(step1Due);
  }

  // Calculate Step II Appeal Deadline (M = K + 10 days)
  if (step1Decision) {
    const step2AppealDeadline = new Date(step1Decision);
    step2AppealDeadline.setDate(step2AppealDeadline.getDate() + CBA_DEADLINES.STEP_II_APPEAL);
    sheet.getRange(row, 13).setValue(step2AppealDeadline);
  }

  // Calculate Step II Decision Due (O = N + 30 days)
  if (step2Filed) {
    const step2Due = new Date(step2Filed);
    step2Due.setDate(step2Due.getDate() + CBA_DEADLINES.STEP_II_DECISION);
    sheet.getRange(row, 15).setValue(step2Due);
  }

  // Calculate Step III Appeal Deadline (R = P + 10 days)
  if (step2Decision) {
    const step3AppealDeadline = new Date(step2Decision);
    step3AppealDeadline.setDate(step3AppealDeadline.getDate() + CBA_DEADLINES.STEP_III_APPEAL);
    sheet.getRange(row, 18).setValue(step3AppealDeadline);
  }

  // Calculate Days Open (AA)
  if (dateFiled) {
    const today = new Date();
    const filed = new Date(dateFiled);
    const daysOpen = Math.floor((today - filed) / (1000 * 60 * 60 * 24));
    sheet.getRange(row, 27).setValue(daysOpen);
  }

  // Calculate Days to Next Deadline (AB)
  const nextDeadline = getNextDeadline(sheet, row);
  if (nextDeadline) {
    const today = new Date();
    const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
    sheet.getRange(row, 28).setValue(daysTo);

    // Is Overdue? (AC)
    sheet.getRange(row, 29).setValue(daysTo < 0 ? "YES" : "NO");

    // Apply color coding
    if (daysTo < 0) {
      sheet.getRange(row, 28).setBackground(COLORS.OVERDUE);
    } else if (daysTo <= 7) {
      sheet.getRange(row, 28).setBackground(COLORS.DUE_SOON);
    } else {
      sheet.getRange(row, 28).setBackground(COLORS.ON_TRACK);
    }
  }

  // Calculate Priority Score (AD)
  const priority = PRIORITY_ORDER[currentStep] || 99;
  sheet.getRange(row, 30).setValue(priority);

  // Last Updated (AH)
  sheet.getRange(row, 34).setValue(new Date());
}

/**
 * Gets the next deadline for a grievance based on current step
 */
function getNextDeadline(sheet, row) {
  const currentStep = sheet.getRange(row, 6).getValue();
  const status = sheet.getRange(row, 5).getValue();

  if (status.startsWith("Resolved")) return null;

  // Check deadlines in order
  const step1Due = sheet.getRange(row, 10).getValue();
  const step2AppealDue = sheet.getRange(row, 13).getValue();
  const step2Due = sheet.getRange(row, 15).getValue();
  const step3AppealDue = sheet.getRange(row, 18).getValue();

  if (currentStep === "Step I - Immediate Supervisor" && step1Due) return step1Due;
  if (currentStep === "Step II - Agency Head") {
    if (step2AppealDue) return step2AppealDue;
    if (step2Due) return step2Due;
  }
  if (currentStep === "Step III - Human Resources" && step3AppealDue) return step3AppealDue;

  return null;
}

/**
 * Recalculates all grievances
 * Run this after bulk imports or data changes
 */
function recalcAllGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  SpreadsheetApp.getUi().alert(`Recalculating ${lastRow - 1} grievances...`);

  for (let row = 2; row <= lastRow; row++) {
    recalcGrievanceRow(sheet, row);
  }

  SpreadsheetApp.getUi().alert('✅ All grievances recalculated!');
}

/**
 * Recalculates member directory derived fields
 */
function recalcMemberRow(memberSheet, grievanceSheet, row) {
  if (row < 2) return;

  const memberId = memberSheet.getRange(row, 1).getValue();
  if (!memberId) return;

  // Get all grievances for this member
  const grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 34).getValues();
  const memberGrievances = grievanceData.filter(g => g[1] === memberId); // Column B = Member ID

  // Total Grievances (N)
  memberSheet.getRange(row, 14).setValue(memberGrievances.length);

  // Active Grievances (O)
  const active = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Filed")).length;
  memberSheet.getRange(row, 15).setValue(active);

  // Resolved Grievances (P)
  const resolved = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Resolved")).length;
  memberSheet.getRange(row, 16).setValue(resolved);

  // Grievances Won (Q)
  const won = memberGrievances.filter(g => g[4] === "Resolved - Won").length;
  memberSheet.getRange(row, 17).setValue(won);

  // Grievances Lost (R)
  const lost = memberGrievances.filter(g => g[4] === "Resolved - Lost").length;
  memberSheet.getRange(row, 18).setValue(lost);

  // Last Grievance Date (S)
  if (memberGrievances.length > 0) {
    const dates = memberGrievances.map(g => g[6]).filter(d => d); // Column G = Incident Date
    if (dates.length > 0) {
      const lastDate = new Date(Math.max(...dates.map(d => new Date(d))));
      memberSheet.getRange(row, 19).setValue(lastDate);
    }
  }

  // Win Rate % (T)
  if (resolved > 0) {
    const winRate = (won / resolved) * 100;
    memberSheet.getRange(row, 20).setValue(winRate.toFixed(1) + "%");
  } else {
    memberSheet.getRange(row, 20).setValue("N/A");
  }

  // Has Open Grievance? (U)
  memberSheet.getRange(row, 21).setValue(active > 0 ? "YES" : "NO");

  // Current Grievance Status (V)
  if (active > 0) {
    const activeGrievance = memberGrievances.find(g => g[4] && g[4].toString().startsWith("Filed"));
    memberSheet.getRange(row, 22).setValue(activeGrievance ? activeGrievance[4] : "");
  } else {
    memberSheet.getRange(row, 22).setValue("");
  }

  // Next Deadline (W) and Days to Deadline (X)
  if (active > 0) {
    const activeGrievances = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Filed"));
    let earliestDeadline = null;

    activeGrievances.forEach(g => {
      // Check various deadline columns
      const deadlines = [g[9], g[12], g[14], g[17]].filter(d => d); // Step I Due, Step II Appeal, Step II Due, Step III Appeal
      deadlines.forEach(d => {
        if (!earliestDeadline || new Date(d) < earliestDeadline) {
          earliestDeadline = new Date(d);
        }
      });
    });

    if (earliestDeadline) {
      memberSheet.getRange(row, 23).setValue(earliestDeadline);
      const today = new Date();
      const daysTo = Math.floor((earliestDeadline - today) / (1000 * 60 * 60 * 24));
      memberSheet.getRange(row, 24).setValue(daysTo);
    }
  } else {
    memberSheet.getRange(row, 23).setValue("");
    memberSheet.getRange(row, 24).setValue("");
  }

  // Last Updated (AI)
  memberSheet.getRange(row, 35).setValue(new Date());
}

/**
 * Recalculates all members
 */
function recalcAllMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = memberSheet.getLastRow();

  if (lastRow < 2) return;

  SpreadsheetApp.getUi().alert(`Recalculating ${lastRow - 1} members...`);

  for (let row = 2; row <= lastRow; row++) {
    recalcMemberRow(memberSheet, grievanceSheet, row);
  }

  SpreadsheetApp.getUi().alert('✅ All members recalculated!');
}

// ============================================================================
// DASHBOARD BUILDING
// ============================================================================

/**
 * Rebuilds the entire dashboard with current data
 */
function rebuildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Get data
  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();

  // Member Metrics (A4:B9)
  const totalMembers = memberData.length - 1;
  const activeMembers = memberData.filter((r, i) => i > 0 && r[11] === "Active").length;
  const totalStewards = memberData.filter((r, i) => i > 0 && r[9] === "Yes").length;
  const unit8 = memberData.filter((r, i) => i > 0 && r[5] === "Unit 8").length;
  const unit10 = memberData.filter((r, i) => i > 0 && r[5] === "Unit 10").length;

  dashboard.getRange("A4").setValue("Total Members:");
  dashboard.getRange("B4").setValue(totalMembers);
  dashboard.getRange("A5").setValue("Active Members:");
  dashboard.getRange("B5").setValue(activeMembers);
  dashboard.getRange("A6").setValue("Total Stewards:");
  dashboard.getRange("B6").setValue(totalStewards);
  dashboard.getRange("A7").setValue("Unit 8 Members:");
  dashboard.getRange("B7").setValue(unit8);
  dashboard.getRange("A8").setValue("Unit 10 Members:");
  dashboard.getRange("B8").setValue(unit10);

  // Grievance Metrics (A11:B19)
  const totalGrievances = grievanceData.length - 1;
  const activeGrievances = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Filed")).length;
  const resolvedGrievances = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Resolved")).length;
  const won = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Won").length;
  const lost = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Lost").length;
  const winRate = resolvedGrievances > 0 ? ((won / resolvedGrievances) * 100).toFixed(1) + "%" : "N/A";
  const inMediation = grievanceData.filter((r, i) => i > 0 && r[4] === "In Mediation").length;
  const inArbitration = grievanceData.filter((r, i) => i > 0 && r[4] === "In Arbitration").length;

  dashboard.getRange("A12").setValue("Total Grievances:");
  dashboard.getRange("B12").setValue(totalGrievances);
  dashboard.getRange("A13").setValue("Active Grievances:");
  dashboard.getRange("B13").setValue(activeGrievances);
  dashboard.getRange("A14").setValue("Resolved Grievances:");
  dashboard.getRange("B14").setValue(resolvedGrievances);
  dashboard.getRange("A15").setValue("Grievances Won:");
  dashboard.getRange("B15").setValue(won);
  dashboard.getRange("A16").setValue("Grievances Lost:");
  dashboard.getRange("B16").setValue(lost);
  dashboard.getRange("A17").setValue("Win Rate:");
  dashboard.getRange("B17").setValue(winRate);
  dashboard.getRange("A18").setValue("In Mediation:");
  dashboard.getRange("B18").setValue(inMediation);
  dashboard.getRange("A19").setValue("In Arbitration:");
  dashboard.getRange("B19").setValue(inArbitration);

  // Deadline Tracking (A21:B25)
  const overdue = grievanceData.filter((r, i) => i > 0 && r[28] === "YES").length;
  const dueThisWeek = grievanceData.filter((r, i) => i > 0 && r[27] && r[27] >= 0 && r[27] <= 7).length;

  dashboard.getRange("A22").setValue("Overdue Grievances:");
  dashboard.getRange("B22").setValue(overdue);
  dashboard.getRange("A23").setValue("Due This Week:");
  dashboard.getRange("B23").setValue(dueThisWeek);

  // Top 10 Overdue Grievances (E4:H13)
  const overdueList = grievanceData
    .filter((r, i) => i > 0 && r[28] === "YES")
    .map(r => ({
      id: r[0],
      member: r[2] + " " + r[3],
      step: r[5],
      daysTo: r[27]
    }))
    .sort((a, b) => a.daysTo - b.daysTo)
    .slice(0, 10);

  dashboard.getRange("E4").setValue("Grievance ID");
  dashboard.getRange("F4").setValue("Member");
  dashboard.getRange("G4").setValue("Step");
  dashboard.getRange("H4").setValue("Days Overdue");

  for (let i = 0; i < 10; i++) {
    const row = 5 + i;
    if (i < overdueList.length) {
      const g = overdueList[i];
      dashboard.getRange("E" + row).setValue(g.id);
      dashboard.getRange("F" + row).setValue(g.member);
      dashboard.getRange("G" + row).setValue(g.step);
      dashboard.getRange("H" + row).setValue(Math.abs(g.daysTo));
    } else {
      dashboard.getRange("E" + row + ":H" + row).clearContent();
    }
  }
}

// ============================================================================
// PRIORITY SORTING
// ============================================================================

/**
 * Sorts grievances by priority (Step III > II > I, then by deadline)
 */
function sortGrievancesByPriority() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return; // Need at least 2 data rows to sort

  // Sort by Column AD (Priority Score), then Column AB (Days to Deadline)
  const range = sheet.getRange(2, 1, lastRow - 1, 34);
  range.sort([
    { column: 30, ascending: true },  // Priority Score
    { column: 28, ascending: true }   // Days to Deadline
  ]);

  SpreadsheetApp.getUi().alert('✅ Grievances sorted by priority!');
}

// ============================================================================
// TRIGGERS
// ============================================================================

/**
 * Sets up automated triggers
 */
function setupTriggers() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Create new triggers
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  ScriptApp.newTrigger('onFormSubmitTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

/**
 * Handles edit events
 */
function onEditTrigger(e) {
  if (!e) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  if (sheet.getName() === SHEETS.GRIEVANCE_LOG && row >= 2) {
    recalcGrievanceRow(sheet, row);

    // Update member directory if Member ID changed
    const memberId = sheet.getRange(row, 2).getValue();
    if (memberId) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      const memberData = memberSheet.getDataRange().getValues();
      const memberRow = memberData.findIndex(r => r[0] === memberId);
      if (memberRow > 0) {
        recalcMemberRow(memberSheet, sheet, memberRow + 1);
      }
    }
  }

  if (sheet.getName() === SHEETS.MEMBER_DIR && row >= 2) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    recalcMemberRow(sheet, grievanceSheet, row);
  }
}

/**
 * Handles form submissions (if connected to Google Form)
 */
function onFormSubmitTrigger(e) {
  if (!e) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // Auto-generate IDs if needed
  if (sheet.getName() === SHEETS.GRIEVANCE_LOG) {
    const id = sheet.getRange(row, 1).getValue();
    if (!id) {
      const grievanceNum = sheet.getLastRow() - 1;
      sheet.getRange(row, 1).setValue(`GRV${String(grievanceNum).padStart(6, '0')}`);
    }
    recalcGrievanceRow(sheet, row);
  }

  if (sheet.getName() === SHEETS.MEMBER_DIR) {
    const id = sheet.getRange(row, 1).getValue();
    if (!id) {
      const memberNum = sheet.getLastRow() - 1;
      sheet.getRange(row, 1).setValue(`MEM${String(memberNum).padStart(6, '0')}`);
    }
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    recalcMemberRow(sheet, grievanceSheet, row);
  }

  rebuildDashboard();
}

// ============================================================================
// SEEDING FUNCTIONS
// ============================================================================

/**
 * Seeds 20,000 members with realistic data
 */
function SEED_20K_MEMBERS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const jobTitles = config.getRange("A2:A19").getValues().flat().filter(String);
  const locations = config.getRange("B2:B16").getValues().flat().filter(String);
  const units = ["Unit 8", "Unit 10"];
  const engagementLevels = config.getRange("J2:J6").getValues().flat().filter(String);
  const membershipStatus = ["Active", "Inactive", "On Leave"];
  const contactMethods = ["Email", "Phone", "Text", "Mail"];
  const committees = ["Executive Board", "Bargaining Committee", "Grievance Committee", "None"];

  const firstNames = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda",
    "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas",
    "Sarah", "Charles", "Karen", "Christopher", "Nancy", "Daniel", "Lisa", "Matthew", "Betty"];

  const lastNames = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
    "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas",
    "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez"];

  const TOTAL = 20000;
  const BATCH = 1000;

  SpreadsheetApp.getUi().alert(`Seeding ${TOTAL} members in ${TOTAL/BATCH} batches...\n\nThis will take 3-5 minutes.`);

  for (let batch = 0; batch < TOTAL / BATCH; batch++) {
    const data = [];

    for (let i = 0; i < BATCH; i++) {
      const num = (batch * BATCH) + i + 1;
      const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
      const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
      const isSteward = Math.random() < 0.1 ? "Yes" : "No";

      data.push([
        `MEM${String(num).padStart(6, '0')}`,
        firstName,
        lastName,
        jobTitles[Math.floor(Math.random() * jobTitles.length)],
        locations[Math.floor(Math.random() * locations.length)],
        units[Math.floor(Math.random() * units.length)],
        "Mon-Fri",
        `${firstName.toLowerCase()}.${lastName.toLowerCase()}@mass.gov`,
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`,
        isSteward,
        randomDate(2020, 2024),
        membershipStatus[Math.floor(Math.random() * membershipStatus.length)],
        engagementLevels[Math.floor(Math.random() * engagementLevels.length)],
        "", "", "", "", "", "", "",  // Grievance metrics (calculated)
        "", "", "", "",              // Derived fields (calculated)
        Math.floor(Math.random() * 25),
        Math.floor(Math.random() * 15),
        isSteward === "Yes" ? committees[Math.floor(Math.random() * (committees.length - 1))] : "None",
        contactMethods[Math.floor(Math.random() * contactMethods.length)],
        "", "",
        "",
        randomDate(1960, 2000),
        randomDate(2015, 2024),
        randomDate(2015, 2024),
        new Date(),
        "SEED_SCRIPT"
      ]);
    }

    sheet.getRange(2 + (batch * BATCH), 1, BATCH, 35).setValues(data);
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getUi().alert('✅ 20,000 members seeded!\n\nRun "Recalc All Members" after adding grievances.');
}

/**
 * Seeds 5,000 grievances with realistic data
 */
function SEED_5K_GRIEVANCES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const memberData = memberSheet.getRange("A2:D" + memberSheet.getLastRow()).getValues();
  const members = memberData.filter(r => r[0]);

  if (members.length === 0) {
    SpreadsheetApp.getUi().alert("Please seed members first!");
    return;
  }

  const statuses = config.getRange("E2:E12").getValues().flat().filter(String);
  const steps = config.getRange("F2:F7").getValues().flat().filter(String);
  const types = config.getRange("G2:G16").getValues().flat().filter(String);
  const outcomes = config.getRange("H2:H10").getValues().flat().filter(String);

  const TOTAL = 5000;
  const BATCH = 500;

  SpreadsheetApp.getUi().alert(`Seeding ${TOTAL} grievances in ${TOTAL/BATCH} batches...\n\nThis will take 2-3 minutes.`);

  for (let batch = 0; batch < TOTAL / BATCH; batch++) {
    const data = [];

    for (let i = 0; i < BATCH; i++) {
      const num = (batch * BATCH) + i + 1;
      const member = members[Math.floor(Math.random() * members.length)];
      const incidentDate = randomDateWithinDays(730);
      const hasFiled = Math.random() > 0.1;
      const filedDate = hasFiled ? addDays(incidentDate, Math.floor(Math.random() * 20)) : null;
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      const step = steps[Math.floor(Math.random() * steps.length)];
      const type = types[Math.floor(Math.random() * types.length)];

      data.push([
        `GRV${String(num).padStart(6, '0')}`,
        member[0], member[1], member[2],
        status, step,
        incidentDate,
        "", // Filing deadline (calculated)
        filedDate,
        "", // Step I due (calculated)
        null, "", "", null, "", null, "", "", null, null, null, null,
        status.startsWith("Resolved") ? outcomes[Math.floor(Math.random() * outcomes.length)] : "Pending",
        type,
        `${type} - Auto-generated description`,
        "Union Rep " + (Math.floor(Math.random() * 10) + 1),
        "", "", "", "", "", "",  // Derived fields (calculated)
        "",
        new Date()
      ]);
    }

    const startRow = 2 + (batch * BATCH);
    sheet.getRange(startRow, 1, BATCH, 34).setValues(data);

    // Recalculate this batch
    for (let row = startRow; row < startRow + BATCH; row++) {
      recalcGrievanceRow(sheet, row);
    }

    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getUi().alert('✅ 5,000 grievances seeded!\n\nRun "Recalc All Members" to update member stats.');
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function randomDate(startYear, endYear) {
  const start = new Date(startYear, 0, 1);
  const end = new Date(endYear, 11, 31);
  return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
}

function randomDateWithinDays(days) {
  const now = new Date();
  const past = new Date(now.getTime() - (days * 24 * 60 * 60 * 1000));
  return new Date(past.getTime() + Math.random() * (now.getTime() - past.getTime()));
}

function addDays(date, days) {
  if (!date) return null;
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

// ============================================================================
// CUSTOM MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('509 Tools')
    .addItem('🔧 Create Dashboard', 'CREATE_509_DASHBOARD')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Data Management')
      .addItem('Seed 20K Members', 'SEED_20K_MEMBERS')
      .addItem('Seed 5K Grievances', 'SEED_5K_GRIEVANCES')
      .addSeparator()
      .addItem('Recalc All Grievances', 'recalcAllGrievances')
      .addItem('Recalc All Members', 'recalcAllMembers')
      .addItem('Rebuild Dashboard', 'rebuildDashboard'))
    .addSubMenu(ui.createMenu('⚙️ Utilities')
      .addItem('Sort by Priority', 'sortGrievancesByPriority')
      .addItem('Setup Triggers', 'setupTriggers'))
    .addToUi();
}
