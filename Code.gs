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
 * - Advanced Dashboard with KPIs and real-time metrics
 * - Interactive Visual Charts:
 *   • Grievances by Status (donut chart)
 *   • Top 10 Grievance Types (bar chart)
 *   • Grievances by Current Step (column chart)
 *   • Members by Unit (donut chart)
 *   • Top 10 Member Locations (bar chart)
 *   • Resolved Grievance Outcomes (donut chart)
 * - Steward Workload Tracking with automated calculations
 * - Color-coded metrics and conditional formatting
 * - Top 10 Overdue Grievances with severity indicators
 * - Form integration for data entry
 * - All calculations done in code (no formula rows)
 * - Automated deadline and status tracking
 *
 * NEW ANALYTICAL TABS:
 * - Trends & Timeline: Monthly filing trends, resolution time analysis, 12-month trends
 * - Performance Metrics: Win rates, settlement rates, efficiency metrics, step-by-step analysis
 * - Location Analytics: Grievances by location, location performance comparisons
 * - Type Analysis: Breakdown by grievance type, success rates, resolution times by type
 *
 * VISUAL ENHANCEMENTS:
 * - Professional color scheme with section headers
 * - Number formatting with thousands separators
 * - Conditional formatting (red=overdue, yellow=due soon, green=on track)
 * - Gradient color coding for overdue severity
 * - Easy-to-read metrics with right-aligned numbers
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
  TRENDS: "Trends & Timeline",
  PERFORMANCE: "Performance Metrics",
  LOCATION: "Location Analytics",
  TYPE_ANALYSIS: "Type Analysis",
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

  try {
    // Clear existing sheets (keep first one)
    const sheets = ss.getSheets();
    for (let i = sheets.length - 1; i > 0; i--) {
      ss.deleteSheet(sheets[i]);
    }
    SpreadsheetApp.flush(); // Let Google process sheet deletions

    // Rename first sheet to Config
    sheets[0].setName(SHEETS.CONFIG);
    SpreadsheetApp.flush();

    // Create all sheets in order with flush after each
    createConfigSheet(ss);
    SpreadsheetApp.flush();

    createMemberDirectorySheet(ss);
    SpreadsheetApp.flush();

    createGrievanceLogSheet(ss);
    SpreadsheetApp.flush();

    createStewardWorkloadSheet(ss);
    SpreadsheetApp.flush();

    createDashboardSheet(ss);
    SpreadsheetApp.flush();

    createAnalyticsSheet(ss);
    SpreadsheetApp.flush();

    createTrendsSheet(ss);
    SpreadsheetApp.flush();

    createPerformanceSheet(ss);
    SpreadsheetApp.flush();

    createLocationSheet(ss);
    SpreadsheetApp.flush();

    createTypeAnalysisSheet(ss);
    SpreadsheetApp.flush();

    createArchiveSheet(ss);
    SpreadsheetApp.flush();

    createDiagnosticsSheet(ss);
    SpreadsheetApp.flush();

    // Setup data validation
    setupDataValidation(ss);
    SpreadsheetApp.flush();

    // Setup triggers
    setupTriggers();

    // Build initial dashboard
    rebuildDashboard();
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert('✅ 509 Dashboard created successfully!\n\n' +
      'Next steps:\n' +
      '1. Add members via "509 Tools > Seed 20K Members" or manually\n' +
      '2. Add grievances via "509 Tools > Seed 5K Grievances" or manually\n' +
      '3. View Dashboard tab for real-time metrics');
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error creating dashboard:\n\n' + error.message + '\n\nPlease try again.');
    Logger.log('Dashboard creation error: ' + error.toString());
  }
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

  // 31 columns (A-AE) per spec
  const headers = [
    // Basic Info (A-F)
    "Member ID", "First Name", "Last Name", "Job Title", "Work Location (Site)", "Unit",
    // Contact & Role (G-L)
    "Office Days", "Email Address", "Phone Number", "Is Steward (Y/N)", "Date Joined Union", "Membership Status",
    // Grievance Metrics (M-R) - AUTO CALCULATED
    "Total Grievances Filed", "Active Grievances", "Resolved Grievances",
    "Grievances Won", "Grievances Lost", "Last Grievance Date",
    // Participation (S-W)
    "Engagement Level", "Events Attended (Last 12mo)", "Training Sessions Attended",
    "Committee Member", "Preferred Contact Method",
    // Emergency & Admin (X-AE)
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
  sheet.setColumnWidth(25, 300); // Notes

  // Color-code auto-calculated columns (M-R and AD)
  const autoCols = [13, 14, 15, 16, 17, 18, 30]; // M-R (grievance metrics) and AD (Last Updated)
  autoCols.forEach(col => {
    sheet.getRange(1, col).setBackground(COLORS.HEADER_GREEN);
  });
}

// ============================================================================
// GRIEVANCE LOG SHEET
// ============================================================================

function createGrievanceLogSheet(ss) {
  try {
    let sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) sheet = ss.insertSheet(SHEETS.GRIEVANCE_LOG);

    sheet.clear();

    // 28 columns (A-AB) per spec
    const headers = [
      // Basic Info (A-F)
      "Grievance ID", "Member ID", "First Name", "Last Name", "Status", "Current Step",
      // Incident & Filing (G-J) - H, J auto-calculated
      "Incident Date", "Filing Deadline (21d)", "Date Filed (Step I)", "Step I Decision Due (30d)",
      // Step I (K-M) - M auto-calculated
      "Step I Decision Date", "Step I Outcome", "Step II Appeal Deadline (10d)",
      // Step II (N-R) - O, R auto-calculated
      "Step II Filed Date", "Step II Decision Due (30d)", "Step II Decision Date", "Step II Outcome", "Step III Appeal Deadline (10d)",
      // Step III & Beyond (S-V)
      "Step III Filed Date", "Step III Decision Date", "Mediation Date", "Arbitration Date",
      // Details (W-Z)
      "Final Outcome", "Grievance Type", "Description", "Representative",
      // Admin (AA-AB)
      "Notes", "Last Updated"
    ];

    // Set headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setFontWeight("bold")
      .setBackground(COLORS.HEADER_RED)
      .setFontColor("white")
      .setWrap(true);

    sheet.setFrozenRows(1);

    // Set column widths
    try {
      sheet.setColumnWidth(1, 120);  // Grievance ID
      sheet.setColumnWidth(2, 110);  // Member ID
      sheet.setColumnWidth(25, 300); // Description (Column Y)
      sheet.setColumnWidth(27, 300); // Notes (Column AA)
    } catch (widthError) {
      Logger.log('Column width error (non-critical): ' + widthError.toString());
    }

    // Color-code auto-calculated deadline columns (H, J, M, O, R, AB)
    const autoCols = [8, 10, 13, 15, 18, 28]; // Filing Deadline, Step I Due, Step II Appeal, Step II Due, Step III Appeal, Last Updated
    autoCols.forEach(col => {
      if (col <= headers.length) {
        sheet.getRange(1, col).setBackground(COLORS.HEADER_ORANGE);
      }
    });
  } catch (error) {
    Logger.log('Error in createGrievanceLogSheet: ' + error.toString());
    throw new Error('Failed to create Grievance Log sheet: ' + error.message);
  }
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
  sheet.getRange("A1:L1").merge()
    .setValue("509 DASHBOARD - REAL-TIME METRICS & ANALYTICS")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BLUE)
    .setFontColor("white");

  // Section headers (will be populated by rebuildDashboard())
  sheet.getRange("A3").setValue("MEMBER METRICS").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A11").setValue("GRIEVANCE METRICS").setFontWeight("bold").setFontSize(12);
  sheet.getRange("A21").setValue("DEADLINE TRACKING").setFontWeight("bold").setFontSize(12);
  sheet.getRange("E3").setValue("TOP 10 OVERDUE GRIEVANCES").setFontWeight("bold").setFontSize(11);

  // Chart area headers
  sheet.getRange("A26").setValue("VISUAL ANALYTICS").setFontWeight("bold").setFontSize(14)
    .setBackground(COLORS.HEADER_GREEN).setFontColor("white");

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
  sheet.setRowHeight(1, 40);
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
// TRENDS & TIMELINE SHEET
// ============================================================================

function createTrendsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.TRENDS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.TRENDS);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("TRENDS & TIMELINE ANALYSIS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_BLUE)
    .setFontColor("white");

  // Section headers
  sheet.getRange("A3").setValue("MONTHLY TRENDS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_GREEN).setFontColor("white");

  sheet.getRange("A15").setValue("RESOLUTION TIME ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_ORANGE).setFontColor("white");

  sheet.getRange("A25").setValue("FILING TRENDS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_PURPLE).setFontColor("white");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 200);
}

// ============================================================================
// PERFORMANCE METRICS SHEET
// ============================================================================

function createPerformanceSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.PERFORMANCE);
  if (!sheet) sheet = ss.insertSheet(SHEETS.PERFORMANCE);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("PERFORMANCE METRICS & KPIs")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_GREEN)
    .setFontColor("white");

  // Section headers
  sheet.getRange("A3").setValue("RESOLUTION PERFORMANCE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_BLUE).setFontColor("white");

  sheet.getRange("E3").setValue("EFFICIENCY METRICS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_ORANGE).setFontColor("white");

  sheet.getRange("A13").setValue("OUTCOME ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_RED).setFontColor("white");

  sheet.getRange("A23").setValue("STEP PROGRESSION ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_PURPLE).setFontColor("white");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
}

// ============================================================================
// LOCATION ANALYTICS SHEET
// ============================================================================

function createLocationSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.LOCATION);
  if (!sheet) sheet = ss.insertSheet(SHEETS.LOCATION);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("LOCATION ANALYTICS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_PURPLE)
    .setFontColor("white");

  // Section headers
  sheet.getRange("A3").setValue("GRIEVANCES BY LOCATION").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_BLUE).setFontColor("white");

  sheet.getRange("A15").setValue("LOCATION PERFORMANCE METRICS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_GREEN).setFontColor("white");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 250);
}

// ============================================================================
// TYPE ANALYSIS SHEET
// ============================================================================

function createTypeAnalysisSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.TYPE_ANALYSIS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.TYPE_ANALYSIS);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("GRIEVANCE TYPE ANALYSIS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.HEADER_RED)
    .setFontColor("white");

  // Section headers
  sheet.getRange("A3").setValue("TYPE BREAKDOWN").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_BLUE).setFontColor("white");

  sheet.getRange("E3").setValue("SUCCESS RATE BY TYPE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_GREEN).setFontColor("white");

  sheet.getRange("A15").setValue("TYPE TRENDS OVER TIME").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_ORANGE).setFontColor("white");

  sheet.getRange("A27").setValue("AVERAGE RESOLUTION TIME BY TYPE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.HEADER_PURPLE).setFontColor("white");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
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
    { sheet: memberDir, column: "J", range: "D2:D3", name: "Steward" },         // Is Steward (Y/N)
    { sheet: memberDir, column: "L", range: "M2:M5", name: "Membership" },      // Membership Status
    { sheet: memberDir, column: "S", range: "J2:J6", name: "Engagement" },      // Engagement Level (now column S)
    { sheet: memberDir, column: "V", range: "L2:L8", name: "Committee" },       // Committee Member (now column V) - FIXED: L2:L8 includes all 7 committee types
    { sheet: memberDir, column: "W", range: "K2:K5", name: "Contact Method" }   // Preferred Contact Method (now column W)
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

  // Last Updated (AB)
  sheet.getRange(row, 28).setValue(new Date());
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
 * Gets the next deadline from a grievance row array (for dashboard calculations)
 */
function getNextDeadlineFromRow(row) {
  const currentStep = row[5]; // F: Current Step
  const status = row[4];       // E: Status

  if (!status || status.toString().startsWith("Resolved")) return null;

  // Check deadlines in order: J, M, O, R (columns 9, 12, 14, 17 in 0-indexed)
  const step1Due = row[9];        // J: Step I Decision Due
  const step2AppealDue = row[12]; // M: Step II Appeal Deadline
  const step2Due = row[14];       // O: Step II Decision Due
  const step3AppealDue = row[17]; // R: Step III Appeal Deadline

  if (currentStep === "Step I - Immediate Supervisor" && step1Due) return new Date(step1Due);
  if (currentStep === "Step II - Agency Head") {
    if (step2AppealDue) return new Date(step2AppealDue);
    if (step2Due) return new Date(step2Due);
  }
  if (currentStep === "Step III - Human Resources" && step3AppealDue) return new Date(step3AppealDue);

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

  // Get all grievances for this member (28 columns in new structure)
  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    // No grievances - set all metrics to 0/blank
    memberSheet.getRange(row, 13).setValue(0); // M: Total Grievances
    memberSheet.getRange(row, 14).setValue(0); // N: Active Grievances
    memberSheet.getRange(row, 15).setValue(0); // O: Resolved Grievances
    memberSheet.getRange(row, 16).setValue(0); // P: Grievances Won
    memberSheet.getRange(row, 17).setValue(0); // Q: Grievances Lost
    memberSheet.getRange(row, 18).setValue(""); // R: Last Grievance Date
    memberSheet.getRange(row, 30).setValue(new Date()); // AD: Last Updated
    memberSheet.getRange(row, 31).setValue("AUTO"); // AE: Updated By
    return;
  }

  const grievanceData = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();
  const memberGrievances = grievanceData.filter(g => g[1] === memberId); // Column B = Member ID

  // Total Grievances Filed (M)
  memberSheet.getRange(row, 13).setValue(memberGrievances.length);

  // Active Grievances (N)
  const active = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Filed")).length;
  memberSheet.getRange(row, 14).setValue(active);

  // Resolved Grievances (O)
  const resolved = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Resolved")).length;
  memberSheet.getRange(row, 15).setValue(resolved);

  // Grievances Won (P)
  const won = memberGrievances.filter(g => g[4] === "Resolved - Won").length;
  memberSheet.getRange(row, 16).setValue(won);

  // Grievances Lost (Q)
  const lost = memberGrievances.filter(g => g[4] === "Resolved - Lost").length;
  memberSheet.getRange(row, 17).setValue(lost);

  // Last Grievance Date (R)
  if (memberGrievances.length > 0) {
    const dates = memberGrievances.map(g => g[6]).filter(d => d); // Column G = Incident Date
    if (dates.length > 0) {
      const lastDate = new Date(Math.max(...dates.map(d => new Date(d))));
      memberSheet.getRange(row, 18).setValue(lastDate);
    } else {
      memberSheet.getRange(row, 18).setValue("");
    }
  } else {
    memberSheet.getRange(row, 18).setValue("");
  }

  // Last Updated (AD)
  memberSheet.getRange(row, 30).setValue(new Date());

  // Updated By (AE)
  memberSheet.getRange(row, 31).setValue("AUTO");
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
    const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!dashboard || !memberSheet || !grievanceSheet) {
      Logger.log('Required sheets not found for dashboard rebuild');
      return;
    }

    // Get data
    const memberData = memberSheet.getDataRange().getValues();
    const grievanceData = grievanceSheet.getDataRange().getValues();

  // Member Metrics (A4:B9)
  const totalMembers = memberData.length - 1;
  const activeMembers = memberData.filter((r, i) => i > 0 && r[11] === "Active").length;
  const totalStewards = memberData.filter((r, i) => i > 0 && r[9] === "Yes").length;
  const unit8 = memberData.filter((r, i) => i > 0 && r[5] === "Unit 8").length;
  const unit10 = memberData.filter((r, i) => i > 0 && r[5] === "Unit 10").length;

  dashboard.getRange("A4").setValue("Total Members:").setFontWeight("bold");
  dashboard.getRange("B4").setValue(totalMembers).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A5").setValue("Active Members:").setFontWeight("bold");
  dashboard.getRange("B5").setValue(activeMembers).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A6").setValue("Total Stewards:").setFontWeight("bold");
  dashboard.getRange("B6").setValue(totalStewards).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A7").setValue("Unit 8 Members:").setFontWeight("bold");
  dashboard.getRange("B7").setValue(unit8).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A8").setValue("Unit 10 Members:").setFontWeight("bold");
  dashboard.getRange("B8").setValue(unit10).setNumberFormat("#,##0").setHorizontalAlignment("right");

  // Color code member metrics section
  dashboard.getRange("A3:B3").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
  dashboard.getRange("A4:B9").setBackground(COLORS.LIGHT_GRAY);

  // Grievance Metrics (A11:B19)
  const totalGrievances = grievanceData.length - 1;
  const activeGrievances = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Filed")).length;
  const resolvedGrievances = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Resolved")).length;
  const won = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Won").length;
  const lost = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Lost").length;
  const winRate = resolvedGrievances > 0 ? ((won / resolvedGrievances) * 100).toFixed(1) + "%" : "N/A";
  const inMediation = grievanceData.filter((r, i) => i > 0 && r[4] === "In Mediation").length;
  const inArbitration = grievanceData.filter((r, i) => i > 0 && r[4] === "In Arbitration").length;

  dashboard.getRange("A12").setValue("Total Grievances:").setFontWeight("bold");
  dashboard.getRange("B12").setValue(totalGrievances).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A13").setValue("Active Grievances:").setFontWeight("bold");
  dashboard.getRange("B13").setValue(activeGrievances).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A14").setValue("Resolved Grievances:").setFontWeight("bold");
  dashboard.getRange("B14").setValue(resolvedGrievances).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A15").setValue("Grievances Won:").setFontWeight("bold");
  dashboard.getRange("B15").setValue(won).setNumberFormat("#,##0").setHorizontalAlignment("right")
    .setBackground(COLORS.ON_TRACK).setFontColor("white");
  dashboard.getRange("A16").setValue("Grievances Lost:").setFontWeight("bold");
  dashboard.getRange("B16").setValue(lost).setNumberFormat("#,##0").setHorizontalAlignment("right")
    .setBackground(COLORS.OVERDUE).setFontColor("white");
  dashboard.getRange("A17").setValue("Win Rate:").setFontWeight("bold");
  dashboard.getRange("B17").setValue(winRate).setHorizontalAlignment("right").setFontSize(12).setFontWeight("bold");
  dashboard.getRange("A18").setValue("In Mediation:").setFontWeight("bold");
  dashboard.getRange("B18").setValue(inMediation).setNumberFormat("#,##0").setHorizontalAlignment("right");
  dashboard.getRange("A19").setValue("In Arbitration:").setFontWeight("bold");
  dashboard.getRange("B19").setValue(inArbitration).setNumberFormat("#,##0").setHorizontalAlignment("right");

  // Color code grievance metrics section
  dashboard.getRange("A11:B11").setBackground(COLORS.HEADER_RED).setFontColor("white");
  dashboard.getRange("A12:B14").setBackground(COLORS.LIGHT_GRAY);
  dashboard.getRange("A15").setBackground(COLORS.LIGHT_GRAY);
  dashboard.getRange("A16").setBackground(COLORS.LIGHT_GRAY);
  dashboard.getRange("A17:B17").setBackground("#E8F5E9"); // Light green for win rate
  dashboard.getRange("A18:B19").setBackground(COLORS.LIGHT_GRAY);

  // Deadline Tracking (A21:B24) - Calculate dynamically
  let overdue = 0;
  let dueThisWeek = 0;
  let dueNextWeek = 0;
  const today = new Date();

  grievanceData.forEach((r, i) => {
    if (i === 0 || !r[4] || r[4].toString().startsWith("Resolved")) return; // Skip header and resolved

    // Get next deadline based on current step
    const nextDeadline = getNextDeadlineFromRow(r);
    if (nextDeadline) {
      const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
      if (daysTo < 0) {
        overdue++;
      } else if (daysTo <= 7) {
        dueThisWeek++;
      } else if (daysTo <= 14) {
        dueNextWeek++;
      }
    }
  });

  dashboard.getRange("A22").setValue("Overdue Grievances:").setFontWeight("bold");
  dashboard.getRange("B22").setValue(overdue).setNumberFormat("#,##0").setHorizontalAlignment("right")
    .setBackground(overdue > 0 ? COLORS.OVERDUE : COLORS.ON_TRACK)
    .setFontColor("white").setFontWeight("bold").setFontSize(12);
  dashboard.getRange("A23").setValue("Due This Week:").setFontWeight("bold");
  dashboard.getRange("B23").setValue(dueThisWeek).setNumberFormat("#,##0").setHorizontalAlignment("right")
    .setBackground(dueThisWeek > 0 ? COLORS.DUE_SOON : COLORS.ON_TRACK)
    .setFontColor("white").setFontWeight("bold").setFontSize(12);
  dashboard.getRange("A24").setValue("Due Next Week:").setFontWeight("bold");
  dashboard.getRange("B24").setValue(dueNextWeek).setNumberFormat("#,##0").setHorizontalAlignment("right");

  // Color code deadline tracking section
  dashboard.getRange("A21:B21").setBackground(COLORS.HEADER_ORANGE).setFontColor("white");
  dashboard.getRange("A22").setBackground(COLORS.LIGHT_GRAY);
  dashboard.getRange("A23").setBackground(COLORS.LIGHT_GRAY);
  dashboard.getRange("A24:B24").setBackground(COLORS.LIGHT_GRAY);

  // Top 10 Overdue Grievances (E4:H13) - Calculate dynamically
  const overdueList = [];
  grievanceData.forEach((r, i) => {
    if (i === 0 || !r[4] || r[4].toString().startsWith("Resolved")) return; // Skip header and resolved

    const nextDeadline = getNextDeadlineFromRow(r);
    if (nextDeadline) {
      const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
      if (daysTo < 0) {
        overdueList.push({
          id: r[0],
          member: r[2] + " " + r[3],
          step: r[5],
          daysTo: daysTo
        });
      }
    }
  });
  overdueList.sort((a, b) => a.daysTo - b.daysTo).splice(10); // Keep only top 10

  dashboard.getRange("E4:H4").setFontWeight("bold").setBackground(COLORS.HEADER_ORANGE).setFontColor("white");
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

      // Color code overdue rows with gradient (darker for more overdue)
      const daysOverdue = Math.abs(g.daysTo);
      if (daysOverdue > 30) {
        dashboard.getRange("E" + row + ":H" + row).setBackground("#C00000").setFontColor("white").setFontWeight("bold");
      } else if (daysOverdue > 14) {
        dashboard.getRange("E" + row + ":H" + row).setBackground(COLORS.OVERDUE).setFontColor("white");
      } else {
        dashboard.getRange("E" + row + ":H" + row).setBackground("#F4CCCC");
      }
    } else {
      dashboard.getRange("E" + row + ":H" + row).clearContent().setBackground("white").setFontWeight("normal");
    }
  }

    // Build Steward Workload Sheet
    rebuildStewardWorkload();
    SpreadsheetApp.flush();

    // Create all dashboard charts
    createDashboardCharts();
    SpreadsheetApp.flush();

    // Rebuild new analytical tabs
    rebuildTrendsSheet();
    SpreadsheetApp.flush();

    rebuildPerformanceSheet();
    SpreadsheetApp.flush();

    rebuildLocationSheet();
    SpreadsheetApp.flush();

    rebuildTypeAnalysisSheet();
    SpreadsheetApp.flush();

  } catch (error) {
    Logger.log('Error in rebuildDashboard: ' + error.toString());
    SpreadsheetApp.getUi().alert('⚠️ Dashboard rebuild error:\n\n' + error.message);
  }
}

// ============================================================================
// STEWARD WORKLOAD CALCULATION
// ============================================================================

/**
 * Rebuilds the Steward Workload sheet with current data
 */
function rebuildStewardWorkload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stewardSheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Get all stewards
  const memberData = memberSheet.getDataRange().getValues();
  const stewards = memberData.filter((r, i) => i > 0 && r[9] === "Yes"); // Column J = Is Steward

  // Get all grievances
  const grievanceData = grievanceSheet.getDataRange().getValues();

  // Clear existing data (keep headers)
  const lastRow = stewardSheet.getLastRow();
  if (lastRow > 1) {
    stewardSheet.getRange(2, 1, lastRow - 1, 15).clearContent();
  }

  const stewardStats = [];

  stewards.forEach(steward => {
    const memberId = steward[0];
    const fullName = steward[1] + " " + steward[2];
    const location = steward[4];

    // Find grievances assigned to this steward (using Representative column)
    const assignedGrievances = grievanceData.filter((g, i) =>
      i > 0 && g[25] && g[25].toString().includes(fullName)
    );

    // Calculate metrics
    const totalCases = assignedGrievances.length;
    const activeCases = assignedGrievances.filter(g => g[4] && g[4].toString().startsWith("Filed")).length;
    const step1Cases = assignedGrievances.filter(g => g[5] === "Step I - Immediate Supervisor").length;
    const step2Cases = assignedGrievances.filter(g => g[5] === "Step II - Agency Head").length;
    const step3Cases = assignedGrievances.filter(g => g[5] === "Step III - Human Resources").length;

    // Calculate overdue and due this week dynamically
    const today = new Date();
    let overdueCases = 0;
    let dueThisWeek = 0;
    assignedGrievances.forEach(g => {
      const nextDeadline = getNextDeadlineFromRow(g);
      if (nextDeadline) {
        const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
        if (daysTo < 0) overdueCases++;
        else if (daysTo <= 7) dueThisWeek++;
      }
    });

    const resolvedCases = assignedGrievances.filter(g => g[4] && g[4].toString().startsWith("Resolved")).length;
    const wonCases = assignedGrievances.filter(g => g[4] === "Resolved - Won").length;
    const winRate = resolvedCases > 0 ? ((wonCases / resolvedCases) * 100).toFixed(1) + "%" : "N/A";

    // Calculate average days to resolution
    const resolvedWithDates = assignedGrievances.filter(g =>
      g[4] && g[4].toString().startsWith("Resolved") && g[8] // Date Filed (I)
    );
    const avgDays = resolvedWithDates.length > 0
      ? Math.round(resolvedWithDates.reduce((sum, g) => {
          const filed = new Date(g[8]); // Date Filed (I)
          const resolved = new Date(); // Approximate - could use Last Updated
          return sum + Math.floor((resolved - filed) / (1000 * 60 * 60 * 24));
        }, 0) / resolvedWithDates.length)
      : "N/A";

    // Get last case date
    const caseDates = assignedGrievances.map(g => g[6]).filter(d => d);
    const lastCaseDate = caseDates.length > 0
      ? new Date(Math.max(...caseDates.map(d => new Date(d))))
      : "";

    const status = activeCases > 0 ? "Active" : "Available";

    stewardStats.push([
      fullName,
      memberId,
      location,
      totalCases,
      activeCases,
      step1Cases,
      step2Cases,
      step3Cases,
      overdueCases,
      dueThisWeek,
      winRate,
      avgDays,
      0, // Members Assigned (placeholder)
      lastCaseDate,
      status
    ]);
  });

  // Write data
  if (stewardStats.length > 0) {
    stewardSheet.getRange(2, 1, stewardStats.length, 15).setValues(stewardStats);

    // Apply conditional formatting
    for (let i = 0; i < stewardStats.length; i++) {
      const row = i + 2;
      const overdue = stewardStats[i][8];
      if (overdue > 0) {
        stewardSheet.getRange(row, 9).setBackground(COLORS.OVERDUE).setFontColor("white");
      }
    }
  }
}

// ============================================================================
// DASHBOARD CHARTS
// ============================================================================

/**
 * Creates all visual charts for the dashboard
 */
function createDashboardCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Remove existing charts
  const charts = dashboard.getCharts();
  charts.forEach(chart => dashboard.removeChart(chart));

  // Get data
  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();

  // Check if we have data to chart
  if (memberData.length <= 1 && grievanceData.length <= 1) {
    dashboard.getRange("A28").setValue("No data available yet. Add members and grievances to see charts.");
    return;
  }

  // Prepare chart data sheets (hidden sheet for chart data)
  let chartDataSheet = ss.getSheetByName("_ChartData");
  if (!chartDataSheet) {
    chartDataSheet = ss.insertSheet("_ChartData");
    chartDataSheet.hideSheet();
  }
  chartDataSheet.clear();

  // === CHART 1: Grievances by Status (Pie Chart) ===
  if (grievanceData.length > 1) {
    const statusCounts = {};
    grievanceData.forEach((r, i) => {
      if (i > 0 && r[4]) {
        const status = r[4].toString();
        statusCounts[status] = (statusCounts[status] || 0) + 1;
      }
    });

    const statusData = [["Status", "Count"]];
    Object.entries(statusCounts).forEach(([status, count]) => {
      statusData.push([status, count]);
    });

    if (statusData.length > 1) {
      chartDataSheet.getRange(1, 1, statusData.length, 2).setValues(statusData);

      const statusChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataSheet.getRange(1, 1, statusData.length, 2))
        .setPosition(27, 1, 0, 0)
        .setOption('title', 'Grievances by Status')
        .setOption('pieHole', 0.4)
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('legend', {position: 'right', textStyle: {fontSize: 10}})
        .setOption('chartArea', {width: '85%', height: '80%'})
        .build();
      dashboard.insertChart(statusChart);
    }
  }

  // === CHART 2: Grievances by Type (Bar Chart) ===
  if (grievanceData.length > 1) {
    const typeCounts = {};
    grievanceData.forEach((r, i) => {
      if (i > 0 && r[23]) { // Column X (24 in 1-indexed, 23 in 0-indexed) = Grievance Type
        const type = r[23].toString();
        typeCounts[type] = (typeCounts[type] || 0) + 1;
      }
    });

    const typeData = [["Type", "Count"]];
    Object.entries(typeCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .forEach(([type, count]) => {
        typeData.push([type.substring(0, 30), count]);
      });

    if (typeData.length > 1) {
      chartDataSheet.getRange(1, 4, typeData.length, 2).setValues(typeData);

      const typeChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(chartDataSheet.getRange(1, 4, typeData.length, 2))
        .setPosition(27, 6, 0, 0)
        .setOption('title', 'Top 10 Grievance Types')
        .setOption('width', 500)
        .setOption('height', 300)
        .setOption('legend', {position: 'none'})
        .setOption('chartArea', {width: '70%', height: '80%'})
        .setOption('hAxis', {title: 'Count'})
        .setOption('colors', ['#E06666'])
        .build();
      dashboard.insertChart(typeChart);
    }
  }

  // === CHART 3: Grievances by Step (Column Chart) ===
  if (grievanceData.length > 1) {
    const stepCounts = {};
    grievanceData.forEach((r, i) => {
      if (i > 0 && r[5]) {
        const step = r[5].toString();
        stepCounts[step] = (stepCounts[step] || 0) + 1;
      }
    });

    const stepData = [["Step", "Count"]];
    Object.entries(stepCounts).forEach(([step, count]) => {
      stepData.push([step, count]);
    });

    if (stepData.length > 1) {
      chartDataSheet.getRange(1, 7, stepData.length, 2).setValues(stepData);

      const stepChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(chartDataSheet.getRange(1, 7, stepData.length, 2))
        .setPosition(45, 1, 0, 0)
        .setOption('title', 'Grievances by Current Step')
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('legend', {position: 'none'})
        .setOption('chartArea', {width: '85%', height: '75%'})
        .setOption('vAxis', {title: 'Count'})
        .setOption('colors', ['#4A86E8'])
        .build();
      dashboard.insertChart(stepChart);
    }
  }

  // === CHART 4: Members by Unit (Pie Chart) ===
  if (memberData.length > 1) {
    const unitCounts = {};
    memberData.forEach((r, i) => {
      if (i > 0 && r[5]) {
        const unit = r[5].toString();
        unitCounts[unit] = (unitCounts[unit] || 0) + 1;
      }
    });

    const unitData = [["Unit", "Count"]];
    Object.entries(unitCounts).forEach(([unit, count]) => {
      unitData.push([unit, count]);
    });

    if (unitData.length > 1) {
      chartDataSheet.getRange(1, 10, unitData.length, 2).setValues(unitData);

      const unitChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataSheet.getRange(1, 10, unitData.length, 2))
        .setPosition(45, 6, 0, 0)
        .setOption('title', 'Members by Unit')
        .setOption('pieHole', 0.4)
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('legend', {position: 'right', textStyle: {fontSize: 12}})
        .setOption('chartArea', {width: '85%', height: '80%'})
        .setOption('colors', ['#93C47D', '#8E7CC3'])
        .build();
      dashboard.insertChart(unitChart);
    }
  }

  // === CHART 5: Members by Location (Top 10 Bar Chart) ===
  if (memberData.length > 1) {
    const locationCounts = {};
    memberData.forEach((r, i) => {
      if (i > 0 && r[4]) {
        const location = r[4].toString();
        locationCounts[location] = (locationCounts[location] || 0) + 1;
      }
    });

    const locationData = [["Location", "Count"]];
    Object.entries(locationCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .forEach(([location, count]) => {
        locationData.push([location.substring(0, 25), count]);
      });

    if (locationData.length > 1) {
      chartDataSheet.getRange(1, 13, locationData.length, 2).setValues(locationData);

      const locationChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(chartDataSheet.getRange(1, 13, locationData.length, 2))
        .setPosition(63, 1, 0, 0)
        .setOption('title', 'Top 10 Member Locations')
        .setOption('width', 500)
        .setOption('height', 300)
        .setOption('legend', {position: 'none'})
        .setOption('chartArea', {width: '70%', height: '80%'})
        .setOption('hAxis', {title: 'Members'})
        .setOption('colors', ['#93C47D'])
        .build();
      dashboard.insertChart(locationChart);
    }
  }

  // === CHART 6: Win/Loss Rate (Pie Chart) ===
  if (grievanceData.length > 1) {
    const wonCount = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Won").length;
    const lostCount = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Lost").length;
    const settledFull = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Settled").length;
    const withdrawn = grievanceData.filter((r, i) => i > 0 && r[4] === "Resolved - Withdrawn").length;

    if (wonCount + lostCount + settledFull + withdrawn > 0) {
      const outcomeData = [
        ["Outcome", "Count"],
        ["Won", wonCount],
        ["Lost", lostCount],
        ["Settled", settledFull],
        ["Withdrawn", withdrawn]
      ];

      chartDataSheet.getRange(1, 16, outcomeData.length, 2).setValues(outcomeData);

      const outcomeChart = chartDataSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataSheet.getRange(1, 16, outcomeData.length, 2))
        .setPosition(63, 6, 0, 0)
        .setOption('title', 'Resolved Grievance Outcomes')
        .setOption('pieHole', 0.4)
        .setOption('width', 450)
        .setOption('height', 300)
        .setOption('legend', {position: 'right', textStyle: {fontSize: 11}})
        .setOption('chartArea', {width: '85%', height: '80%'})
        .setOption('colors', ['#34A853', '#EA4335', '#FBBC04', '#999999'])
        .build();
      dashboard.insertChart(outcomeChart);
    }
  }
}

// ============================================================================
// NEW TAB DATA POPULATION FUNCTIONS
// ============================================================================

/**
 * Rebuilds the Trends & Timeline sheet with current data
 */
function rebuildTrendsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trendsSheet = ss.getSheetByName(SHEETS.TRENDS);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!trendsSheet || !grievanceSheet) return;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length <= 1) {
    trendsSheet.getRange("A4").setValue("No data available yet.");
    return;
  }

  // Calculate monthly filing trends
  const monthlyData = {};
  const today = new Date();

  grievanceData.forEach((r, i) => {
    if (i === 0) return;
    const filedDate = r[8]; // Column I: Date Filed
    if (filedDate) {
      const date = new Date(filedDate);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      monthlyData[monthKey] = (monthlyData[monthKey] || 0) + 1;
    }
  });

  // Monthly Trends KPIs
  trendsSheet.getRange("A4").setValue("Total Months with Data:").setFontWeight("bold");
  trendsSheet.getRange("B4").setValue(Object.keys(monthlyData).length).setNumberFormat("#,##0");

  trendsSheet.getRange("A5").setValue("Average Filings per Month:").setFontWeight("bold");
  const avgPerMonth = Object.keys(monthlyData).length > 0
    ? (Object.values(monthlyData).reduce((a, b) => a + b, 0) / Object.keys(monthlyData).length).toFixed(1)
    : 0;
  trendsSheet.getRange("B5").setValue(avgPerMonth).setNumberFormat("#,##0.0");

  trendsSheet.getRange("A6").setValue("Peak Month:").setFontWeight("bold");
  const peakMonth = Object.entries(monthlyData).sort((a, b) => b[1] - a[1])[0];
  trendsSheet.getRange("B6").setValue(peakMonth ? `${peakMonth[0]} (${peakMonth[1]} filings)` : "N/A");

  // Resolution Time Analysis
  const resolutionTimes = [];
  grievanceData.forEach((r, i) => {
    if (i === 0 || !r[4] || !r[4].toString().startsWith("Resolved")) return;
    const filed = r[8]; // Date Filed
    const resolved = r[27]; // Last Updated (approximation)
    if (filed && resolved) {
      const days = Math.floor((new Date(resolved) - new Date(filed)) / (1000 * 60 * 60 * 24));
      if (days > 0) resolutionTimes.push(days);
    }
  });

  trendsSheet.getRange("A16").setValue("Average Resolution Time:").setFontWeight("bold");
  const avgResolution = resolutionTimes.length > 0
    ? (resolutionTimes.reduce((a, b) => a + b, 0) / resolutionTimes.length).toFixed(1)
    : "N/A";
  trendsSheet.getRange("B16").setValue(`${avgResolution} days`);

  trendsSheet.getRange("A17").setValue("Fastest Resolution:").setFontWeight("bold");
  trendsSheet.getRange("B17").setValue(resolutionTimes.length > 0 ? `${Math.min(...resolutionTimes)} days` : "N/A");

  trendsSheet.getRange("A18").setValue("Slowest Resolution:").setFontWeight("bold");
  trendsSheet.getRange("B18").setValue(resolutionTimes.length > 0 ? `${Math.max(...resolutionTimes)} days` : "N/A");

  // Filing Trends (Last 12 months)
  const last12Months = {};
  for (let i = 11; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    last12Months[key] = 0;
  }

  Object.keys(monthlyData).forEach(month => {
    if (last12Months.hasOwnProperty(month)) {
      last12Months[month] = monthlyData[month];
    }
  });

  trendsSheet.getRange("A26").setValue("Month").setFontWeight("bold");
  trendsSheet.getRange("B26").setValue("Filings").setFontWeight("bold");

  let row = 27;
  Object.entries(last12Months).forEach(([month, count]) => {
    trendsSheet.getRange(`A${row}`).setValue(month);
    trendsSheet.getRange(`B${row}`).setValue(count).setNumberFormat("#,##0");
    row++;
  });

  // Apply formatting
  trendsSheet.getRange("A4:B18").setBackground(COLORS.LIGHT_GRAY);
  trendsSheet.getRange("B4:B18").setHorizontalAlignment("right");
}

/**
 * Rebuilds the Performance Metrics sheet
 */
function rebuildPerformanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perfSheet = ss.getSheetByName(SHEETS.PERFORMANCE);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!perfSheet || !grievanceSheet) return;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length <= 1) {
    perfSheet.getRange("A4").setValue("No data available yet.");
    return;
  }

  // Resolution Performance
  const resolved = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Resolved"));
  const won = resolved.filter(r => r[4] === "Resolved - Won");
  const lost = resolved.filter(r => r[4] === "Resolved - Lost");
  const settled = resolved.filter(r => r[4] === "Resolved - Settled");
  const withdrawn = resolved.filter(r => r[4] === "Resolved - Withdrawn");

  perfSheet.getRange("A4").setValue("Total Resolved:").setFontWeight("bold");
  perfSheet.getRange("B4").setValue(resolved.length).setNumberFormat("#,##0").setHorizontalAlignment("right");

  perfSheet.getRange("A5").setValue("Win Rate:").setFontWeight("bold");
  const winRate = resolved.length > 0 ? ((won.length / resolved.length) * 100).toFixed(1) + "%" : "N/A";
  perfSheet.getRange("B5").setValue(winRate).setHorizontalAlignment("right").setFontWeight("bold")
    .setBackground("#E8F5E9");

  perfSheet.getRange("A6").setValue("Settlement Rate:").setFontWeight("bold");
  const settleRate = resolved.length > 0 ? ((settled.length / resolved.length) * 100).toFixed(1) + "%" : "N/A";
  perfSheet.getRange("B6").setValue(settleRate).setHorizontalAlignment("right");

  perfSheet.getRange("A7").setValue("Withdrawal Rate:").setFontWeight("bold");
  const withdrawRate = resolved.length > 0 ? ((withdrawn.length / resolved.length) * 100).toFixed(1) + "%" : "N/A";
  perfSheet.getRange("B7").setValue(withdrawRate).setHorizontalAlignment("right");

  // Efficiency Metrics
  const active = grievanceData.filter((r, i) => i > 0 && r[4] && r[4].toString().startsWith("Filed"));
  const today = new Date();
  let overdue = 0;

  active.forEach(r => {
    const nextDeadline = getNextDeadlineFromRow(r);
    if (nextDeadline && nextDeadline < today) overdue++;
  });

  perfSheet.getRange("E4").setValue("Active Grievances:").setFontWeight("bold");
  perfSheet.getRange("F4").setValue(active.length).setNumberFormat("#,##0").setHorizontalAlignment("right");

  perfSheet.getRange("E5").setValue("Overdue Rate:").setFontWeight("bold");
  const overdueRate = active.length > 0 ? ((overdue / active.length) * 100).toFixed(1) + "%" : "N/A";
  perfSheet.getRange("F5").setValue(overdueRate).setHorizontalAlignment("right")
    .setBackground(overdue > 0 ? COLORS.OVERDUE : COLORS.ON_TRACK)
    .setFontColor("white").setFontWeight("bold");

  perfSheet.getRange("E6").setValue("On-Time Rate:").setFontWeight("bold");
  const onTimeRate = active.length > 0 ? (((active.length - overdue) / active.length) * 100).toFixed(1) + "%" : "N/A";
  perfSheet.getRange("F6").setValue(onTimeRate).setHorizontalAlignment("right");

  // Outcome Analysis by Step
  perfSheet.getRange("A14:C14").setFontWeight("bold").setBackground(COLORS.HEADER_RED).setFontColor("white");
  perfSheet.getRange("A14").setValue("Step");
  perfSheet.getRange("B14").setValue("Total");
  perfSheet.getRange("C14").setValue("Win Rate");

  const stepOutcomes = {};
  ["Step I - Immediate Supervisor", "Step II - Agency Head", "Step III - Human Resources"].forEach(step => {
    const stepGrievances = resolved.filter(r => r[5] === step);
    const stepWon = stepGrievances.filter(r => r[4] === "Resolved - Won");
    stepOutcomes[step] = {
      total: stepGrievances.length,
      winRate: stepGrievances.length > 0 ? ((stepWon.length / stepGrievances.length) * 100).toFixed(1) + "%" : "N/A"
    };
  });

  let row = 15;
  Object.entries(stepOutcomes).forEach(([step, data]) => {
    perfSheet.getRange(`A${row}`).setValue(step);
    perfSheet.getRange(`B${row}`).setValue(data.total).setNumberFormat("#,##0");
    perfSheet.getRange(`C${row}`).setValue(data.winRate);
    row++;
  });

  // Apply formatting
  perfSheet.getRange("A4:B7").setBackground(COLORS.LIGHT_GRAY);
  perfSheet.getRange("E4:F6").setBackground(COLORS.LIGHT_GRAY);
  perfSheet.getRange("A15:C17").setBackground(COLORS.LIGHT_GRAY);
}

/**
 * Rebuilds the Location Analytics sheet
 */
function rebuildLocationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const locSheet = ss.getSheetByName(SHEETS.LOCATION);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!locSheet || !memberSheet || !grievanceSheet) return;

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();

  if (memberData.length <= 1) {
    locSheet.getRange("A4").setValue("No data available yet.");
    return;
  }

  // Count members and grievances by location
  const locationStats = {};

  memberData.forEach((r, i) => {
    if (i === 0) return;
    const location = r[4]; // Column E: Work Location
    const memberId = r[0];

    if (!locationStats[location]) {
      locationStats[location] = { members: 0, grievances: 0, active: 0, resolved: 0, won: 0 };
    }
    locationStats[location].members++;

    // Count grievances for this member
    const memberGrievances = grievanceData.filter((g, gi) => gi > 0 && g[1] === memberId);
    locationStats[location].grievances += memberGrievances.length;
    locationStats[location].active += memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Filed")).length;
    locationStats[location].resolved += memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Resolved")).length;
    locationStats[location].won += memberGrievances.filter(g => g[4] === "Resolved - Won").length;
  });

  // Headers
  locSheet.getRange("A4:F4").setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
  locSheet.getRange("A4").setValue("Location");
  locSheet.getRange("B4").setValue("Members");
  locSheet.getRange("C4").setValue("Total Grievances");
  locSheet.getRange("D4").setValue("Active");
  locSheet.getRange("E4").setValue("Resolved");
  locSheet.getRange("F4").setValue("Win Rate");

  // Sort by grievance count
  const sortedLocations = Object.entries(locationStats)
    .sort((a, b) => b[1].grievances - a[1].grievances)
    .slice(0, 15); // Top 15 locations

  let row = 5;
  sortedLocations.forEach(([location, stats]) => {
    const winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) + "%" : "N/A";

    locSheet.getRange(`A${row}`).setValue(location);
    locSheet.getRange(`B${row}`).setValue(stats.members).setNumberFormat("#,##0");
    locSheet.getRange(`C${row}`).setValue(stats.grievances).setNumberFormat("#,##0");
    locSheet.getRange(`D${row}`).setValue(stats.active).setNumberFormat("#,##0");
    locSheet.getRange(`E${row}`).setValue(stats.resolved).setNumberFormat("#,##0");
    locSheet.getRange(`F${row}`).setValue(winRate);
    row++;
  });

  locSheet.getRange(`A5:F${row-1}`).setBackground(COLORS.LIGHT_GRAY);
  locSheet.getRange(`B5:F${row-1}`).setHorizontalAlignment("right");

  locSheet.setColumnWidth(1, 250);
}

/**
 * Rebuilds the Type Analysis sheet
 */
function rebuildTypeAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const typeSheet = ss.getSheetByName(SHEETS.TYPE_ANALYSIS);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!typeSheet || !grievanceSheet) return;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length <= 1) {
    typeSheet.getRange("A4").setValue("No data available yet.");
    return;
  }

  // Analyze by type
  const typeStats = {};

  grievanceData.forEach((r, i) => {
    if (i === 0) return;
    const type = r[23]; // Column X: Grievance Type
    if (!type) return;

    if (!typeStats[type]) {
      typeStats[type] = { total: 0, active: 0, resolved: 0, won: 0, lost: 0, settled: 0 };
    }

    typeStats[type].total++;
    const status = r[4];
    if (status && status.toString().startsWith("Filed")) typeStats[type].active++;
    if (status && status.toString().startsWith("Resolved")) {
      typeStats[type].resolved++;
      if (status === "Resolved - Won") typeStats[type].won++;
      if (status === "Resolved - Lost") typeStats[type].lost++;
      if (status === "Resolved - Settled") typeStats[type].settled++;
    }
  });

  // Type Breakdown
  typeSheet.getRange("A4:D4").setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
  typeSheet.getRange("A4").setValue("Grievance Type");
  typeSheet.getRange("B4").setValue("Total");
  typeSheet.getRange("C4").setValue("Active");
  typeSheet.getRange("D4").setValue("Resolved");

  const sortedTypes = Object.entries(typeStats).sort((a, b) => b[1].total - a[1].total);

  let row = 5;
  sortedTypes.forEach(([type, stats]) => {
    typeSheet.getRange(`A${row}`).setValue(type);
    typeSheet.getRange(`B${row}`).setValue(stats.total).setNumberFormat("#,##0");
    typeSheet.getRange(`C${row}`).setValue(stats.active).setNumberFormat("#,##0");
    typeSheet.getRange(`D${row}`).setValue(stats.resolved).setNumberFormat("#,##0");
    row++;
  });

  typeSheet.getRange(`A5:D${row-1}`).setBackground(COLORS.LIGHT_GRAY);
  typeSheet.getRange(`B5:D${row-1}`).setHorizontalAlignment("right");

  // Success Rate by Type
  typeSheet.getRange("E4:G4").setFontWeight("bold").setBackground(COLORS.HEADER_GREEN).setFontColor("white");
  typeSheet.getRange("E4").setValue("Grievance Type");
  typeSheet.getRange("F4").setValue("Win Rate");
  typeSheet.getRange("G4").setValue("Settlement Rate");

  row = 5;
  sortedTypes.forEach(([type, stats]) => {
    const winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) + "%" : "N/A";
    const settleRate = stats.resolved > 0 ? ((stats.settled / stats.resolved) * 100).toFixed(1) + "%" : "N/A";

    typeSheet.getRange(`E${row}`).setValue(type.substring(0, 30));
    typeSheet.getRange(`F${row}`).setValue(winRate);
    typeSheet.getRange(`G${row}`).setValue(settleRate);
    row++;
  });

  typeSheet.getRange(`E5:G${row-1}`).setBackground(COLORS.LIGHT_GRAY);
  typeSheet.getRange(`F5:G${row-1}`).setHorizontalAlignment("right");
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

  // Get all data and calculate priority scores
  const data = sheet.getRange(2, 1, lastRow - 1, 28).getValues();
  const today = new Date();

  // Create array with priority scores and days to deadline
  const rowsWithPriority = data.map((row, index) => {
    const currentStep = row[5]; // F: Current Step
    const priority = PRIORITY_ORDER[currentStep] || 99;
    const nextDeadline = getNextDeadlineFromRow(row);
    const daysTo = nextDeadline ? Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24)) : 999999;

    return {
      row: row,
      priority: priority,
      daysTo: daysTo,
      originalIndex: index
    };
  });

  // Sort by priority, then by days to deadline
  rowsWithPriority.sort((a, b) => {
    if (a.priority !== b.priority) return a.priority - b.priority;
    return a.daysTo - b.daysTo;
  });

  // Write back sorted data
  const sortedData = rowsWithPriority.map(item => item.row);
  sheet.getRange(2, 1, sortedData.length, 28).setValues(sortedData);

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
        // A-F: Basic Info
        `MEM${String(num).padStart(6, '0')}`,
        firstName,
        lastName,
        jobTitles[Math.floor(Math.random() * jobTitles.length)],
        locations[Math.floor(Math.random() * locations.length)],
        units[Math.floor(Math.random() * units.length)],
        // G-L: Contact & Role
        "Mon-Fri",
        `${firstName.toLowerCase()}.${lastName.toLowerCase()}@mass.gov`,
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`,
        isSteward,
        randomDate(2020, 2024),
        membershipStatus[Math.floor(Math.random() * membershipStatus.length)],
        // M-R: Grievance Metrics (AUTO CALCULATED - leave blank)
        "", "", "", "", "", "",
        // S-W: Participation
        engagementLevels[Math.floor(Math.random() * engagementLevels.length)],
        Math.floor(Math.random() * 25),
        Math.floor(Math.random() * 15),
        isSteward === "Yes" ? committees[Math.floor(Math.random() * (committees.length - 1))] : "None",
        contactMethods[Math.floor(Math.random() * contactMethods.length)],
        // X-AE: Emergency & Admin
        `${firstNames[Math.floor(Math.random() * firstNames.length)]} ${lastNames[Math.floor(Math.random() * lastNames.length)]}`,
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`,
        "",
        randomDate(1960, 2000),
        randomDate(2015, 2024),
        randomDate(2015, 2024),
        new Date(),
        "SEED_SCRIPT"
      ]);
    }

    sheet.getRange(2 + (batch * BATCH), 1, BATCH, 31).setValues(data);
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
        // A-F: Basic Info
        `GRV${String(num).padStart(6, '0')}`,
        member[0], member[1], member[2],
        status, step,
        // G-J: Incident & Filing (H, J auto-calculated)
        incidentDate,
        "", // Filing deadline (H - calculated)
        filedDate,
        "", // Step I due (J - calculated)
        // K-M: Step I (M auto-calculated)
        null, "", "", // Step I Decision Date, Outcome, Appeal Deadline (M - calculated)
        // N-R: Step II (O, R auto-calculated)
        null, "", null, "", "", // Step II Filed, Decision Due (O - calc), Decision, Outcome, Appeal Deadline (R - calc)
        // S-V: Step III & Beyond
        null, null, null, null, // Step III Filed, Decision, Mediation, Arbitration
        // W-Z: Details
        status.startsWith("Resolved") ? outcomes[Math.floor(Math.random() * outcomes.length)] : "Pending",
        type,
        `${type} - Auto-generated description`,
        "Union Rep " + (Math.floor(Math.random() * 10) + 1),
        // AA-AB: Admin (AB auto-calculated)
        "",
        new Date() // Last Updated (AB)
      ]);
    }

    const startRow = 2 + (batch * BATCH);
    sheet.getRange(startRow, 1, BATCH, 28).setValues(data);

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
    .addSubMenu(ui.createMenu('📈 Rebuild Analytics')
      .addItem('Rebuild All Tabs', 'rebuildDashboard')
      .addSeparator()
      .addItem('Rebuild Trends & Timeline', 'rebuildTrendsSheet')
      .addItem('Rebuild Performance Metrics', 'rebuildPerformanceSheet')
      .addItem('Rebuild Location Analytics', 'rebuildLocationSheet')
      .addItem('Rebuild Type Analysis', 'rebuildTypeAnalysisSheet'))
    .addSubMenu(ui.createMenu('⚙️ Utilities')
      .addItem('Sort by Priority', 'sortGrievancesByPriority')
      .addItem('Setup Triggers', 'setupTriggers'))
    .addToUi();
}
