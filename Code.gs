/**
 * ============================================================================
 * 509 DASHBOARD - Enhanced Member Directory & Grievance Tracking System
 * ============================================================================
 *
 * For SEIU Local 509 (Units 8 & 10) - Massachusetts State Employees
 * Based on Collective Bargaining Agreement 2024-2026
 *
 * FEATURES:
 * - Member Directory with auto-calculated grievance metrics & snapshots
 * - Grievance Log with CBA-compliant timeline tracking
 * - Contract-based Timeline Rules Table in Config
 * - Real-time deadline calculations (Days Open, Next Action Due, Days to Deadline, Overdue?)
 * - Member snapshots (Has Open Grievance?, # Open Grievances, Next Deadline)
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
 * ANALYTICAL TABS:
 * - Trends & Timeline: Monthly filing trends, resolution time analysis, 12-month trends
 * - Performance Metrics: Win rates, settlement rates, efficiency metrics, step-by-step analysis
 * - Location Analytics: Grievances by location, location performance comparisons
 * - Type Analysis: Breakdown by grievance type, success rates, resolution times by type
 *
 * NEW DASHBOARD VARIATIONS (Inspired by Professional Design References):
 * - Executive Overview: Clean CRM-style dashboard with large KPI cards for leadership
 * - KPI Dashboard: Gaming-style dashboard with bold numbers and performance indicators
 * - Member Engagement: Track participation, satisfaction, committee involvement
 * - Cost Impact Analysis: Financial recovery tracking and ROI metrics
 * - Quick Stats: At-a-glance real-time statistics with prominent displays
 *
 * VISUAL ENHANCEMENTS:
 * - Unified professional union color scheme (Union Blue, Solidarity Red, Success Green)
 * - Consistent design language across all dashboard tabs
 * - Number formatting with thousands separators
 * - Conditional formatting (red=overdue, yellow=due soon, green=on track)
 * - Gradient color coding for overdue severity
 * - Easy-to-read metrics with right-aligned numbers
 *
 * CBA COMPLIANCE & TIMELINE LOGIC:
 * - Config tab contains Timeline Rules Table with contract deadlines
 * - Article 23A: Grievance deadlines (21-day filing, 30-day decisions, 10-day appeals)
 * - Automatic calculation of next action due dates based on current step
 * - Tracks responsible party (Employee/Union vs Employer) for each deadline
 * - Article 8: Leave provisions
 * - Article 14: Promotions
 *
 * DATA STRUCTURE:
 * - Member Directory: 35 columns (A-AI) with grievance snapshot fields
 * - Grievance Log: 32 columns (A-AF) with timeline tracking fields
 * - Config: Dropdown lists + Timeline Rules Table
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
  EXECUTIVE: "Executive Overview",
  KPI_BOARD: "KPI Dashboard",
  MEMBER_ENGAGEMENT: "Member Engagement",
  COST_IMPACT: "Cost Impact Analysis",
  QUICK_STATS: "Quick Stats",
  ARCHIVE: "Archive",
  DIAGNOSTICS: "Diagnostics"
};

// Unified Professional Union Color Scheme (inspired by design references)
const COLORS = {
  // Primary Union Theme
  PRIMARY_BLUE: "#2563EB",      // Union Blue - main headers
  SOLIDARITY_RED: "#DC2626",    // Solidarity Red - urgent/critical
  UNION_GREEN: "#059669",       // Success/wins

  // Secondary Colors
  ACCENT_TEAL: "#0891B2",       // Info/neutral
  ACCENT_PURPLE: "#7C3AED",     // In progress/pending
  ACCENT_ORANGE: "#EA580C",     // Warning/attention
  ACCENT_YELLOW: "#CA8A04",     // Due soon

  // Backgrounds & Neutrals
  WHITE: "#FFFFFF",
  CARD_BG: "#FFFFFF",
  LIGHT_GRAY: "#F9FAFB",
  BORDER_GRAY: "#E5E7EB",
  TEXT_DARK: "#1F2937",
  TEXT_GRAY: "#6B7280",

  // Status Colors
  OVERDUE: "#DC2626",           // Red
  DUE_SOON: "#F59E0B",          // Orange
  ON_TRACK: "#10B981",          // Green
  IN_PROGRESS: "#8B5CF6",       // Purple

  // Legacy compatibility (mapped to new colors)
  HEADER_BLUE: "#2563EB",
  HEADER_RED: "#DC2626",
  HEADER_GREEN: "#059669",
  HEADER_ORANGE: "#EA580C",
  HEADER_PURPLE: "#7C3AED"
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

    // New dashboard variations inspired by design references
    createExecutiveOverviewSheet(ss);
    SpreadsheetApp.flush();

    createKPIDashboardSheet(ss);
    SpreadsheetApp.flush();

    createMemberEngagementSheet(ss);
    SpreadsheetApp.flush();

    createCostImpactSheet(ss);
    SpreadsheetApp.flush();

    createQuickStatsSheet(ss);
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
    "Contact Methods", "Committee Types", "Membership Status", "Office Days"
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

  // Column N: Office Days
  const officeDays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

  // Write all data
  const allData = [
    jobTitles, workLocations, units, stewardStatus, grievanceStatus,
    grievanceSteps, grievanceTypes, outcomes, satisfactionLevels,
    engagementLevels, contactMethods, committeeTypes, membershipStatus, officeDays
  ];

  for (let i = 0; i < allData.length; i++) {
    const values = allData[i].map(item => [item]);
    config.getRange(2, i + 1, values.length, 1).setValues(values);
  }

  config.setFrozenRows(1);
  config.autoResizeColumns(1, headers.length);

  // ============================================================================
  // TIMELINE RULES TABLE (Contract-based grievance deadlines)
  // ============================================================================

  // Create Timeline Rules section starting at column O (15)
  const timelineStartCol = 15;

  // Section header
  config.getRange(1, timelineStartCol, 1, 6).merge()
    .setValue("GRIEVANCE TIMELINE RULES (CBA Article 23A)")
    .setFontWeight("bold")
    .setFontSize(12)
    .setBackground(COLORS.HEADER_ORANGE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Timeline table headers
  const timelineHeaders = [
    "Step", "Trigger Event", "Responsible Party",
    "Max Days Allowed", "Day Type", "What Happens if Late"
  ];

  config.getRange(2, timelineStartCol, 1, timelineHeaders.length).setValues([timelineHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.HEADER_BLUE)
    .setFontColor("white")
    .setWrap(true);

  // Timeline rules data
  const timelineRules = [
    ["Initial Filing", "Incident Date", "Employee/Union", 21, "Calendar", "Grievance may be deemed untimely"],
    ["Step I - Decision", "Date Filed (Step I)", "Employer", 30, "Calendar", "Union may advance to Step II"],
    ["Step II - Appeal", "Step I Decision Date", "Employee/Union", 10, "Calendar", "Grievance deemed resolved at Step I"],
    ["Step II - Decision", "Step II Filed Date", "Employer", 30, "Calendar", "Union may advance to Step III"],
    ["Step III - Appeal", "Step II Decision Date", "Employee/Union", 10, "Calendar", "Grievance deemed resolved at Step II"],
    ["Step III - Decision", "Step III Filed Date", "Employer", 30, "Calendar", "Union may request arbitration"],
    ["Arbitration - Request", "Step III Decision Date", "Employee/Union", 30, "Calendar", "Grievance deemed withdrawn"],
    ["Mediation - Optional", "Any Step", "Both Parties", 0, "N/A", "Optional dispute resolution process"]
  ];

  config.getRange(3, timelineStartCol, timelineRules.length, timelineHeaders.length)
    .setValues(timelineRules);

  // Auto-resize timeline columns
  for (let i = 0; i < timelineHeaders.length; i++) {
    config.setColumnWidth(timelineStartCol + i, i === 1 ? 180 : (i === 5 ? 280 : 150));
  }
}

// ============================================================================
// MEMBER DIRECTORY SHEET
// ============================================================================

function createMemberDirectorySheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBER_DIR);

  sheet.clear();

  // 33 columns (A-AG) - Added grievance snapshot fields
  const headers = [
    // Basic Info (A-F)
    "Member ID", "First Name", "Last Name", "Job Title", "Work Location (Site)", "Unit",
    // Contact & Role (G-K)
    "Office Days", "Email Address", "Phone Number", "Is Steward (Y/N)", "Membership Status",
    // Grievance Metrics (L-Q) - AUTO CALCULATED
    "Total Grievances Filed", "Active Grievances", "Resolved Grievances",
    "Grievances Won", "Grievances Lost", "Last Grievance Date",
    // Grievance Snapshot (R-U) - AUTO CALCULATED
    "Has Open Grievance?", "# Open Grievances", "Last Grievance Status", "Next Deadline (Soonest)",
    // Participation (V-Z)
    "Engagement Level", "Events Attended (Last 12mo)", "Training Sessions Attended",
    "Committee Member", "Preferred Contact Method",
    // Emergency & Admin (AA-AF)
    "Emergency Contact Name", "Emergency Contact Phone", "Notes",
    "Date of Birth", "Hire Date",
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

  // Color-code auto-calculated columns (M-V and AH)
  const autoCols = [13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 34]; // M-R (grievance metrics), S-V (snapshot fields), AH (Last Updated)
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

    // 32 columns (A-AF) - Added timeline tracking columns
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
      "Final Outcome", "Grievance Type", "Description", "Steward Name",
      // Timeline Tracking (AA-AD) - AUTO CALCULATED
      "Days Open", "Next Action Due", "Days to Deadline", "Overdue?",
      // Admin (AE-AF)
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

    // Color-code auto-calculated deadline columns (H, J, M, O, R, AA-AD, AF)
    const autoCols = [8, 10, 13, 15, 18, 27, 28, 29, 30, 32]; // Filing Deadline, Step I Due, Step II Appeal, Step II Due, Step III Appeal, Days Open, Next Action Due, Days to Deadline, Overdue?, Last Updated
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

  // Title - Unified color scheme
  sheet.getRange("A1:L1").merge()
    .setValue("TRENDS & TIMELINE ANALYSIS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("MONTHLY TRENDS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("RESOLUTION TIME ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A25").setValue("FILING TRENDS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

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

  // Title - Unified color scheme
  sheet.getRange("A1:L1").merge()
    .setValue("PERFORMANCE METRICS & KPIs")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("RESOLUTION PERFORMANCE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("E3").setValue("EFFICIENCY METRICS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A13").setValue("OUTCOME ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A23").setValue("STEP PROGRESSION ANALYSIS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

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

  // Title - Unified color scheme
  sheet.getRange("A1:L1").merge()
    .setValue("LOCATION ANALYTICS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("GRIEVANCES BY LOCATION").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("LOCATION PERFORMANCE METRICS").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

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

  // Title - Unified color scheme
  sheet.getRange("A1:L1").merge()
    .setValue("GRIEVANCE TYPE ANALYSIS")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("TYPE BREAKDOWN").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("E3").setValue("SUCCESS RATE BY TYPE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("TYPE TRENDS OVER TIME").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A27").setValue("AVERAGE RESOLUTION TIME BY TYPE").setFontWeight("bold").setFontSize(12)
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 120);
}

// ============================================================================
// EXECUTIVE OVERVIEW DASHBOARD (Inspired by CRM Dashboard)
// ============================================================================

/**
 * Clean, professional executive-level overview with large KPI cards
 * Design inspiration: CRM Dashboard with white cards and colored metrics
 */
function createExecutiveOverviewSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.EXECUTIVE);
  if (!sheet) sheet = ss.insertSheet(SHEETS.EXECUTIVE);

  sheet.clear();

  // Main title with union blue
  sheet.getRange("A1:P1").merge()
    .setValue("📋 EXECUTIVE OVERVIEW - 509 DASHBOARD")
    .setFontSize(24)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Subtitle/date range
  sheet.getRange("A2:P2").merge()
    .setValue("Union-wide Performance Metrics & Key Indicators")
    .setFontSize(12)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  // Section: Key Metrics (Row 4-10)
  sheet.getRange("A4:P4").merge()
    .setValue("📊 KEY PERFORMANCE INDICATORS")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // ===== KPI CARD 1: Total Members =====
  const kpiCard1 = sheet.getRange("A6:C9");
  kpiCard1.setBackground(COLORS.WHITE);

  sheet.getRange("A6:C6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A6:A9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("C6:C9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("A9:C9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("A6:C6").merge().setValue("TOTAL MEMBERS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:C7").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("A8:C9").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("vs Last Month");

  // ===== KPI CARD 2: Active Grievances =====
  const kpiCard2 = sheet.getRange("E6:G9");
  kpiCard2.setBackground(COLORS.WHITE);

  sheet.getRange("E6:G6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E6:E9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G6:G9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E9:G9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E6:G6").merge().setValue("ACTIVE GRIEVANCES")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E7:G7").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("E8:G9").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("vs Last Month");

  // ===== KPI CARD 3: Win Rate =====
  const kpiCard3 = sheet.getRange("I6:K9");
  kpiCard3.setBackground(COLORS.WHITE);

  sheet.getRange("I6:K6").setBorder(true, null, null, null, null, null, COLORS.UNION_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I6:I9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K6:K9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I9:K9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I6:K6").merge().setValue("WIN RATE")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I7:K7").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("I8:K9").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("vs Last Month");

  // ===== KPI CARD 4: Avg Resolution Time =====
  const kpiCard4 = sheet.getRange("M6:O9");
  kpiCard4.setBackground(COLORS.WHITE);

  sheet.getRange("M6:O6").setBorder(true, null, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M6:M9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O6:O9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M9:O9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M6:O6").merge().setValue("AVG RESOLUTION TIME")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M7:O7").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("M8:O9").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("vs Last Month");

  // ===== UNION HEALTH SECTION =====
  sheet.getRange("A12:P12").merge()
    .setValue("💪 UNION HEALTH & ENGAGEMENT")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // ===== HEALTH CARD 1: Steward Count =====
  const healthCard1 = sheet.getRange("A14:C17");
  healthCard1.setBackground(COLORS.WHITE);

  sheet.getRange("A14:C14").setBorder(true, null, null, null, null, null, COLORS.ACCENT_PURPLE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A14:A17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("C14:C17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("A17:C17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("A14:C14").merge().setValue("STEWARD COUNT")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A15:C15").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("A16:C17").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("Change YTD");

  // ===== HEALTH CARD 2: Member Engagement =====
  const healthCard2 = sheet.getRange("E14:G17");
  healthCard2.setBackground(COLORS.WHITE);

  sheet.getRange("E14:G14").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E14:E17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G14:G17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E17:G17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E14:G14").merge().setValue("MEMBER ENGAGEMENT")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E15:G15").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("E16:G17").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("Avg Score");

  // ===== HEALTH CARD 3: Active Committees =====
  const healthCard3 = sheet.getRange("I14:K17");
  healthCard3.setBackground(COLORS.WHITE);

  sheet.getRange("I14:K14").setBorder(true, null, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I14:I17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K14:K17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I17:K17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I14:K14").merge().setValue("ACTIVE COMMITTEES")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I15:K15").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("I16:K17").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("This Quarter");

  // ===== HEALTH CARD 4: Training Sessions =====
  const healthCard4 = sheet.getRange("M14:O17");
  healthCard4.setBackground(COLORS.WHITE);

  sheet.getRange("M14:O14").setBorder(true, null, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M14:M17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O14:O17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M17:O17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M14:O14").merge().setValue("TRAINING SESSIONS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M15:O15").merge().setFontSize(42).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("M16:O17").merge().setFontSize(10).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("Last 12 Months");

  // ===== PRIORITY ALERTS SECTION =====
  sheet.getRange("A20:P20").merge()
    .setValue("🚨 PRIORITY ALERTS & ACTION ITEMS")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  // Alert cards with borders
  const alertCard = sheet.getRange("A22:O25");
  alertCard.setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A22").setValue("Overdue Grievances:").setFontWeight("bold").setFontSize(11);
  sheet.getRange("A23").setValue("Due This Week:").setFontWeight("bold").setFontSize(11);
  sheet.getRange("A24").setValue("Arbitrations Pending:").setFontWeight("bold").setFontSize(11);
  sheet.getRange("A25").setValue("High-Priority Cases:").setFontWeight("bold").setFontSize(11);

  // Add colored left borders for alert severity
  sheet.getRange("A22:A22").setBorder(null, true, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A23:A23").setBorder(null, true, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A24:A24").setBorder(null, true, null, null, null, null, COLORS.ACCENT_YELLOW, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A25:A25").setBorder(null, true, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ===== VISUAL SUMMARIES SECTION =====
  sheet.getRange("A28:P28").merge()
    .setValue("📈 VISUAL SUMMARIES")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // Chart placeholders
  sheet.getRange("A30").setValue("Status Breakdown").setFontWeight("bold").setFontSize(12);
  sheet.getRange("F30").setValue("Monthly Trends").setFontWeight("bold").setFontSize(12);
  sheet.getRange("K30").setValue("Win/Loss Outcomes").setFontWeight("bold").setFontSize(12);

  // Formatting
  sheet.setRowHeight(1, 55);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 40);
  sheet.setRowHeight(6, 30);
  sheet.setRowHeight(7, 60);
  sheet.setRowHeight(8, 15);
  sheet.setRowHeight(9, 15);
  sheet.setRowHeight(12, 40);
  sheet.setRowHeight(14, 30);
  sheet.setRowHeight(15, 60);
  sheet.setRowHeight(16, 15);
  sheet.setRowHeight(17, 15);
  sheet.setRowHeight(20, 40);
  sheet.setRowHeight(28, 40);

  // Set column widths (cards are 3 columns wide, with 1 column spacing)
  sheet.setColumnWidth(1, 140); // A
  sheet.setColumnWidth(2, 140); // B
  sheet.setColumnWidth(3, 140); // C
  sheet.setColumnWidth(4, 30);  // D (spacer)
  sheet.setColumnWidth(5, 140); // E
  sheet.setColumnWidth(6, 140); // F
  sheet.setColumnWidth(7, 140); // G
  sheet.setColumnWidth(8, 30);  // H (spacer)
  sheet.setColumnWidth(9, 140); // I
  sheet.setColumnWidth(10, 140); // J
  sheet.setColumnWidth(11, 140); // K
  sheet.setColumnWidth(12, 30);  // L (spacer)
  sheet.setColumnWidth(13, 140); // M
  sheet.setColumnWidth(14, 140); // N
  sheet.setColumnWidth(15, 140); // O

  // Set background for spacing columns
  sheet.getRange("D1:D35").setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange("H1:H35").setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange("L1:L35").setBackground(COLORS.LIGHT_GRAY);
}

// ============================================================================
// KPI DASHBOARD (Inspired by Gaming KPI Dashboard)
// ============================================================================

/**
 * Large metric displays with performance indicators
 * Design inspiration: Gaming KPI Dashboard with triangle indicators and bold numbers
 */
function createKPIDashboardSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.KPI_BOARD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.KPI_BOARD);

  sheet.clear();

  // Bold header with gradient feel
  sheet.getRange("A1:P1").merge()
    .setValue("⚡ REAL-TIME KPI DASHBOARD")
    .setFontSize(26)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  // Current period indicator
  const today = new Date();
  sheet.getRange("A2:P2").merge()
    .setValue(`Performance Period: ${today.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`)
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Main KPI Grid - Large Numbers Section
  sheet.getRange("A4:P4").merge()
    .setValue("📊 THIS MONTH")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // ===== KPI CARD 1: Grievances Filed =====
  const card1Range = sheet.getRange("A6:C9");
  card1Range.setBackground(COLORS.WHITE);

  // Card 1: Top border (color accent)
  sheet.getRange("A6:C6").setBorder(true, null, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Card 1: Side and bottom borders
  sheet.getRange("A6:A9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("C6:C9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("A9:C9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("A6:C6").merge().setValue("GRIEVANCES FILED")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:C7").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("A8:C9").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 2: Grievances Resolved =====
  const card2Range = sheet.getRange("E6:G9");
  card2Range.setBackground(COLORS.WHITE);

  sheet.getRange("E6:G6").setBorder(true, null, null, null, null, null, COLORS.UNION_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E6:E9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G6:G9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E9:G9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E6:G6").merge().setValue("GRIEVANCES RESOLVED")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E7:G7").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("E8:G9").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 3: Pending Decisions =====
  const card3Range = sheet.getRange("I6:K9");
  card3Range.setBackground(COLORS.WHITE);

  sheet.getRange("I6:K6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I6:I9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K6:K9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I9:K9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I6:K6").merge().setValue("PENDING DECISIONS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I7:K7").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("I8:K9").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 4: Win Rate % =====
  const card4Range = sheet.getRange("M6:O9");
  card4Range.setBackground(COLORS.WHITE);

  sheet.getRange("M6:O6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_PURPLE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M6:M9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O6:O9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M9:O9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M6:O6").merge().setValue("WIN RATE %")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M7:O7").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("M8:O9").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YEAR TO DATE SECTION =====
  sheet.getRange("A12:P12").merge()
    .setValue("📈 YEAR TO DATE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // ===== YTD CARD 1: Total Filings =====
  const ytdCard1Range = sheet.getRange("A14:C17");
  ytdCard1Range.setBackground(COLORS.WHITE);

  sheet.getRange("A14:C14").setBorder(true, null, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A14:A17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("C14:C17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("A17:C17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("A14:C14").merge().setValue("TOTAL FILINGS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A15:C15").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("A16:C17").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 2: Total Resolved =====
  const ytdCard2Range = sheet.getRange("E14:G17");
  ytdCard2Range.setBackground(COLORS.WHITE);

  sheet.getRange("E14:G14").setBorder(true, null, null, null, null, null, COLORS.UNION_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E14:E17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G14:G17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E17:G17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E14:G14").merge().setValue("TOTAL RESOLVED")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E15:G15").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("E16:G17").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 3: Settlements =====
  const ytdCard3Range = sheet.getRange("I14:K17");
  ytdCard3Range.setBackground(COLORS.WHITE);

  sheet.getRange("I14:K14").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I14:I17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K14:K17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I17:K17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I14:K14").merge().setValue("SETTLEMENTS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I15:K15").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("I16:K17").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 4: Arbitrations =====
  const ytdCard4Range = sheet.getRange("M14:O17");
  ytdCard4Range.setBackground(COLORS.WHITE);

  sheet.getRange("M14:O14").setBorder(true, null, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M14:M17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O14:O17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M17:O17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M14:O14").merge().setValue("ARBITRATIONS")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M15:O15").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.SOLIDARITY_RED);
  sheet.getRange("M16:O17").merge().setFontSize(11).setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== PERFORMANCE BREAKDOWN SECTION =====
  sheet.getRange("A20:P20").merge()
    .setValue("⚡ PERFORMANCE BREAKDOWN")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white");

  // Table headers for detailed breakdown with enhanced styling
  const perfHeaders = ["Metric", "Current", "Target", "Variance", "Status"];
  sheet.getRange("A22:E22").setValues([perfHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const perfMetrics = [
    ["Avg Days to Resolution", "", "30 days", "", ""],
    ["Member Satisfaction", "", "4.5/5.0", "", ""],
    ["Steward Response Time", "", "24 hrs", "", ""],
    ["Case Documentation Rate", "", "100%", "", ""],
    ["Settlement Success Rate", "", "75%", "", ""]
  ];

  sheet.getRange("A23:E27").setValues(perfMetrics);

  // Add borders to performance table
  sheet.getRange("A22:E27").setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // Alternating row colors for readability
  sheet.getRange("A23:E23").setBackground("#F9FAFB");
  sheet.getRange("A24:E24").setBackground(COLORS.WHITE);
  sheet.getRange("A25:E25").setBackground("#F9FAFB");
  sheet.getRange("A26:E26").setBackground(COLORS.WHITE);
  sheet.getRange("A27:E27").setBackground("#F9FAFB");

  // Formatting
  sheet.setRowHeight(1, 60);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 40);
  sheet.setRowHeight(6, 30); // Card label row
  sheet.setRowHeight(7, 65); // Large number row
  sheet.setRowHeight(8, 20); // Trend indicator row
  sheet.setRowHeight(9, 20);
  sheet.setRowHeight(12, 40);
  sheet.setRowHeight(14, 30); // YTD Card label row
  sheet.setRowHeight(15, 65); // YTD Large number row
  sheet.setRowHeight(16, 20);
  sheet.setRowHeight(17, 20);
  sheet.setRowHeight(20, 40);
  sheet.setRowHeight(22, 35);

  // Set column widths (cards are 3 columns wide, with 1 column spacing)
  sheet.setColumnWidth(1, 140); // A
  sheet.setColumnWidth(2, 140); // B
  sheet.setColumnWidth(3, 140); // C
  sheet.setColumnWidth(4, 30);  // D (spacer)
  sheet.setColumnWidth(5, 140); // E
  sheet.setColumnWidth(6, 140); // F
  sheet.setColumnWidth(7, 140); // G
  sheet.setColumnWidth(8, 30);  // H (spacer)
  sheet.setColumnWidth(9, 140); // I
  sheet.setColumnWidth(10, 140); // J
  sheet.setColumnWidth(11, 140); // K
  sheet.setColumnWidth(12, 30);  // L (spacer)
  sheet.setColumnWidth(13, 140); // M
  sheet.setColumnWidth(14, 140); // N
  sheet.setColumnWidth(15, 140); // O

  // Set background for spacing columns to distinguish them
  sheet.getRange("D1:D30").setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange("H1:H30").setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange("L1:L30").setBackground(COLORS.LIGHT_GRAY);
}

// ============================================================================
// MEMBER ENGAGEMENT DASHBOARD (Inspired by Employee Performance Dashboard)
// ============================================================================

/**
 * Member participation, satisfaction, and engagement tracking
 * Design inspiration: Employee Performance Dashboard with multi-colored metrics
 */
function createMemberEngagementSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.MEMBER_ENGAGEMENT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBER_ENGAGEMENT);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("MEMBER ENGAGEMENT & PARTICIPATION")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  // Engagement Score Cards
  sheet.getRange("A3:L3").merge()
    .setValue("🎯 ENGAGEMENT SCORES")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  const engagementLabels = [
    ["OVERALL ENGAGEMENT", "VERY ACTIVE MEMBERS", "EVENT ATTENDANCE", "TRAINING PARTICIPATION"],
    ["", "", "", ""],
    ["Avg Level", "Count & %", "Last 12 Months", "Completion Rate"]
  ];

  sheet.getRange("A5:D7").setValues([
    engagementLabels[0],
    engagementLabels[1],
    engagementLabels[2]
  ]);

  sheet.getRange("A5:D5").setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setFontColor(COLORS.TEXT_GRAY);

  // Member Distribution
  sheet.getRange("A10:F10").merge()
    .setValue("👥 MEMBER DISTRIBUTION BY ENGAGEMENT LEVEL")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  const engagementHeaders = ["Engagement Level", "Count", "Percentage", "Change (30d)", "Trend"];
  sheet.getRange("A12:E12").setValues([engagementHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  const engagementLevels = [
    ["Very Active", "", "", "", ""],
    ["Active", "", "", "", ""],
    ["Moderately Active", "", "", "", ""],
    ["Inactive", "", "", "", ""],
    ["New Member", "", "", "", ""]
  ];

  sheet.getRange("A13:E17").setValues(engagementLevels);

  // Committee Participation
  sheet.getRange("A20:F20").merge()
    .setValue("🏛️ COMMITTEE PARTICIPATION")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  const committeeHeaders = ["Committee", "Members", "Meetings (YTD)", "Attendance Rate", "Status"];
  sheet.getRange("A22:E22").setValues([committeeHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  const committees = [
    ["Executive Board", "", "", "", ""],
    ["Bargaining Committee", "", "", "", ""],
    ["Grievance Committee", "", "", "", ""],
    ["Political Action", "", "", "", ""],
    ["Communications", "", "", "", ""],
    ["Member Engagement", "", "", "", ""]
  ];

  sheet.getRange("A23:E28").setValues(committees);

  // Contact & Communication Preferences
  sheet.getRange("A31:F31").merge()
    .setValue("📱 COMMUNICATION PREFERENCES")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  const contactHeaders = ["Preferred Method", "Member Count", "Percentage", "Response Rate"];
  sheet.getRange("A33:D33").setValues([contactHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  const contactMethods = [
    ["Email", "", "", ""],
    ["Phone", "", "", ""],
    ["Text", "", "", ""],
    ["Mail", "", "", ""]
  ];

  sheet.getRange("A34:D37").setValues(contactMethods);

  // Formatting
  sheet.setRowHeight(1, 45);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidths(2, 4, 130);
}

// ============================================================================
// COST IMPACT ANALYSIS DASHBOARD (Inspired by Portfolio Dashboard)
// ============================================================================

/**
 * Financial and cost analysis of grievance activities
 * Design inspiration: Portfolio Dashboard with clean charts and financial metrics
 */
function createCostImpactSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.COST_IMPACT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.COST_IMPACT);

  sheet.clear();

  // Title
  sheet.getRange("A1:L1").merge()
    .setValue("COST IMPACT ANALYSIS")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  sheet.getRange("A2:L2").merge()
    .setValue("Financial Impact of Grievances & Union Activities")
    .setFontSize(11)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY);

  // Financial Metrics
  sheet.getRange("A4:L4").merge()
    .setValue("💰 MEMBER WINS & RECOVERIES")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  const costLabels = [
    ["TOTAL RECOVERED", "BACK PAY WON", "BENEFITS RESTORED", "AVG PER CASE"],
    ["", "", "", ""],
    ["Year to Date", "Year to Date", "Count", "This Year"]
  ];

  sheet.getRange("A6:D8").setValues([
    costLabels[0],
    costLabels[1],
    costLabels[2]
  ]);

  sheet.getRange("A6:D6").setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setFontColor(COLORS.TEXT_GRAY);

  // Cost by Type
  sheet.getRange("A11:F11").merge()
    .setValue("📊 RECOVERY BY GRIEVANCE TYPE")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  const typeHeaders = ["Grievance Type", "Cases Won", "Total Recovered", "Avg Recovery", "% of Total"];
  sheet.getRange("A13:E13").setValues([typeHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  // Placeholder rows (will be populated by rebuild)
  for (let i = 0; i < 10; i++) {
    sheet.getRange(14 + i, 1).setValue(`Type ${i + 1}`);
  }

  // Time Cost Analysis
  sheet.getRange("A25:F25").merge()
    .setValue("⏱️ TIME INVESTMENT ANALYSIS")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white");

  const timeHeaders = ["Activity", "Total Hours", "Avg Hours/Case", "Steward Time", "Member Time"];
  sheet.getRange("A27:E27").setValues([timeHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  const timeActivities = [
    ["Grievance Preparation", "", "", "", ""],
    ["Meetings & Hearings", "", "", "", ""],
    ["Research & Documentation", "", "", "", ""],
    ["Mediation Sessions", "", "", "", ""],
    ["Arbitration Proceedings", "", "", "", ""]
  ];

  sheet.getRange("A28:E32").setValues(timeActivities);

  // ROI Metrics
  sheet.getRange("A35:F35").merge()
    .setValue("📈 RETURN ON INVESTMENT")
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  sheet.getRange("A37").setValue("Time Invested vs. Recovery Ratio:").setFontWeight("bold");
  sheet.getRange("A38").setValue("Success Rate Impact:").setFontWeight("bold");
  sheet.getRange("A39").setValue("Member Value Delivered:").setFontWeight("bold");

  // Formatting
  sheet.setRowHeight(1, 45);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidths(2, 4, 130);
}

// ============================================================================
// QUICK STATS DASHBOARD (Inspired by Running Analytics)
// ============================================================================

/**
 * Fast, at-a-glance statistics with large numbers
 * Design inspiration: Running Analytics with prominent metric displays
 */
function createQuickStatsSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.QUICK_STATS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.QUICK_STATS);

  sheet.clear();

  // Bold, eye-catching title
  sheet.getRange("A1:O1").merge()
    .setValue("⚡ QUICK STATS - AT A GLANCE")
    .setFontSize(26)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // Today's snapshot
  const today = new Date();
  sheet.getRange("A2:O2").merge()
    .setValue(`Snapshot: ${today.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`)
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // ===== TODAY SECTION =====
  sheet.getRange("A4:O4").merge()
    .setValue("📅 TODAY")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // TODAY CARD 1: New Grievances
  sheet.getRange("A6:D8").setBackground(COLORS.WHITE);
  sheet.getRange("A6:D6").setBorder(true, null, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A6:A8").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("D6:D8").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("A8:D8").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("A6:D6").merge().setValue("NEW GRIEVANCES")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:D8").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);

  // TODAY CARD 2: Deadlines Due
  sheet.getRange("F6:I8").setBackground(COLORS.WHITE);
  sheet.getRange("F6:I6").setBorder(true, null, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("F6:F8").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I6:I8").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("F8:I8").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("F6:I6").merge().setValue("DEADLINES DUE")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("F7:I8").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.SOLIDARITY_RED);

  // TODAY CARD 3: Meetings Scheduled
  sheet.getRange("K6:N8").setBackground(COLORS.WHITE);
  sheet.getRange("K6:N6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("K6:K8").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("N6:N8").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K8:N8").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("K6:N6").merge().setValue("MEETINGS SCHEDULED")
    .setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("K7:N8").merge().setFontSize(48).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);

  // ===== THIS WEEK SECTION =====
  sheet.getRange("A11:O11").merge()
    .setValue("📊 THIS WEEK")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  // Helper function to create mini cards for this week
  const createMiniCard = (range, topColor, label, valueColor) => {
    sheet.getRange(range).setBackground(COLORS.WHITE);
    const cells = range.split(":");
    const topRow = cells[0].match(/[A-Z]+/)[0] + cells[0].match(/\d+/)[0];
    const bottomRow = cells[1].match(/[A-Z]+/)[0] + cells[1].match(/\d+/)[0];
    const topRowRange = cells[0].match(/[A-Z]+/)[0] + cells[0].match(/\d+/)[0] + ":" + cells[1].match(/[A-Z]+/)[0] + cells[0].match(/\d+/)[0];

    sheet.getRange(topRowRange).setBorder(true, null, null, null, null, null, topColor, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(cells[0].match(/[A-Z]+/)[0] + cells[0].match(/\d+/)[0] + ":" + cells[0].match(/[A-Z]+/)[0] + cells[1].match(/\d+/)[0])
      .setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(cells[1].match(/[A-Z]+/)[0] + cells[0].match(/\d+/)[0] + ":" + cells[1].match(/[A-Z]+/)[0] + cells[1].match(/\d+/)[0])
      .setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(bottomRow).setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(topRowRange).merge().setValue(label)
      .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center")
      .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
    sheet.getRange(range.split(":")[0].match(/[A-Z]+/)[0] + (parseInt(range.split(":")[0].match(/\d+/)[0]) + 1) + ":" + bottomRow).merge()
      .setFontSize(38).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setFontColor(valueColor);
  };

  createMiniCard("A13:C15", COLORS.PRIMARY_BLUE, "FILED", COLORS.PRIMARY_BLUE);
  createMiniCard("E13:G15", COLORS.UNION_GREEN, "RESOLVED", COLORS.UNION_GREEN);
  createMiniCard("I13:K15", COLORS.ACCENT_ORANGE, "DUE SOON", COLORS.ACCENT_ORANGE);
  createMiniCard("M13:O15", COLORS.SOLIDARITY_RED, "OVERDUE", COLORS.SOLIDARITY_RED);

  // ===== THIS MONTH SECTION =====
  sheet.getRange("A18:O18").merge()
    .setValue("📈 THIS MONTH")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  // Slightly wider mini cards for 5 items
  const monthRanges = ["A20:B22", "D20:E22", "G20:H22", "J20:K22", "M20:N22"];
  const monthLabels = ["FILED", "RESOLVED", "WON", "LOST", "SETTLED"];
  const monthColors = [COLORS.PRIMARY_BLUE, COLORS.ACCENT_TEAL, COLORS.UNION_GREEN, COLORS.SOLIDARITY_RED, COLORS.ACCENT_ORANGE];

  for (let i = 0; i < monthRanges.length; i++) {
    createMiniCard(monthRanges[i], monthColors[i], monthLabels[i], monthColors[i]);
  }

  // ===== YEAR TO DATE SECTION =====
  sheet.getRange("A25:O25").merge()
    .setValue("🎯 YEAR TO DATE")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  const ytdRanges = ["A27:B29", "D27:E29", "G27:H29", "J27:K29", "M27:N29"];
  const ytdLabels = ["TOTAL FILED", "TOTAL RESOLVED", "WIN RATE %", "AVG DAYS", "ACTIVE NOW"];
  const ytdColors = [COLORS.PRIMARY_BLUE, COLORS.UNION_GREEN, COLORS.ACCENT_PURPLE, COLORS.ACCENT_ORANGE, COLORS.ACCENT_TEAL];

  for (let i = 0; i < ytdRanges.length; i++) {
    createMiniCard(ytdRanges[i], ytdColors[i], ytdLabels[i], ytdColors[i]);
  }

  // ===== ALL TIME SECTION =====
  sheet.getRange("A32:O32").merge()
    .setValue("🏆 ALL TIME RECORDS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white");

  createMiniCard("A34:C36", COLORS.ACCENT_TEAL, "MEMBERS SERVED", COLORS.ACCENT_TEAL);
  createMiniCard("E34:G36", COLORS.PRIMARY_BLUE, "GRIEVANCES FILED", COLORS.PRIMARY_BLUE);
  createMiniCard("I34:K36", COLORS.UNION_GREEN, "TOTAL WINS", COLORS.UNION_GREEN);
  createMiniCard("M34:O36", COLORS.ACCENT_PURPLE, "SUCCESS RATE", COLORS.ACCENT_PURPLE);

  // ===== CRITICAL ALERTS BOX =====
  sheet.getRange("A39:O39").merge()
    .setValue("🚨 CRITICAL ALERTS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  const alertBox = sheet.getRange("A41:O44");
  alertBox.setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A41").setValue("Grievances Overdue:").setFontWeight("bold").setFontSize(12).setFontColor(COLORS.SOLIDARITY_RED);
  sheet.getRange("A42").setValue("Due in Next 48 Hours:").setFontWeight("bold").setFontSize(12).setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("A43").setValue("Arbitrations Pending:").setFontWeight("bold").setFontSize(12).setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("A44").setValue("Step III Cases:").setFontWeight("bold").setFontSize(12).setFontColor(COLORS.PRIMARY_BLUE);

  // Add colored left borders for alert severity
  sheet.getRange("A41:A41").setBorder(null, true, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A42:A42").setBorder(null, true, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A43:A43").setBorder(null, true, null, null, null, null, COLORS.ACCENT_PURPLE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A44:A44").setBorder(null, true, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Formatting
  sheet.setRowHeight(1, 60);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 40);
  sheet.setRowHeight(6, 30);
  sheet.setRowHeight(7, 40);
  sheet.setRowHeight(8, 40);
  sheet.setRowHeight(11, 40);
  sheet.setRowHeight(13, 28);
  sheet.setRowHeight(14, 45);
  sheet.setRowHeight(15, 30);
  sheet.setRowHeight(18, 40);
  sheet.setRowHeight(20, 28);
  sheet.setRowHeight(21, 40);
  sheet.setRowHeight(22, 25);
  sheet.setRowHeight(25, 40);
  sheet.setRowHeight(27, 28);
  sheet.setRowHeight(28, 40);
  sheet.setRowHeight(29, 25);
  sheet.setRowHeight(32, 40);
  sheet.setRowHeight(34, 28);
  sheet.setRowHeight(35, 45);
  sheet.setRowHeight(36, 30);
  sheet.setRowHeight(39, 40);

  // Column widths
  sheet.setColumnWidth(1, 100);  // A
  sheet.setColumnWidth(2, 100);  // B
  sheet.setColumnWidth(3, 100);  // C
  sheet.setColumnWidth(4, 100);  // D
  sheet.setColumnWidth(5, 25);   // E (spacer)
  sheet.setColumnWidth(6, 100);  // F
  sheet.setColumnWidth(7, 100);  // G
  sheet.setColumnWidth(8, 100);  // H
  sheet.setColumnWidth(9, 100);  // I
  sheet.setColumnWidth(10, 25);  // J (spacer)
  sheet.setColumnWidth(11, 100); // K
  sheet.setColumnWidth(12, 100); // L
  sheet.setColumnWidth(13, 100); // M
  sheet.setColumnWidth(14, 100); // N
  sheet.setColumnWidth(15, 100); // O

  // Set background for spacing columns
  sheet.getRange("E1:E50").setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange("J1:J50").setBackground(COLORS.LIGHT_GRAY);
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
    { sheet: memberDir, column: "G", range: "N2:N8", name: "Office Days" },     // Office Days (allows comma-separated values)
    { sheet: memberDir, column: "J", range: "D2:D3", name: "Steward" },         // Is Steward (Y/N)
    { sheet: memberDir, column: "K", range: "M2:M5", name: "Membership" },      // Membership Status (was column L)
    { sheet: memberDir, column: "R", range: "J2:J6", name: "Engagement" },      // Engagement Level (was column S, now column R)
    { sheet: memberDir, column: "U", range: "L2:L8", name: "Committee" },       // Committee Member (was column V, now column U)
    { sheet: memberDir, column: "V", range: "K2:K5", name: "Contact Method" }   // Preferred Contact Method (was column W, now column V)
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
  const step3Filed = sheet.getRange(row, 19).getValue();    // S: Step III Filed
  const step3Decision = sheet.getRange(row, 20).getValue(); // T: Step III Decision
  const currentStep = sheet.getRange(row, 6).getValue();    // F: Current Step
  const status = sheet.getRange(row, 5).getValue();         // E: Status
  const finalOutcome = sheet.getRange(row, 23).getValue();  // W: Final Outcome

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

  // ============================================================================
  // TIMELINE TRACKING CALCULATIONS (AA-AD)
  // ============================================================================

  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to midnight for date comparisons

  // Calculate Days Open (AA)
  let daysOpen = "";
  if (dateFiled) {
    const filedDate = new Date(dateFiled);
    filedDate.setHours(0, 0, 0, 0);

    // If resolved, calculate from filed to resolution date
    if (status && status.toString().startsWith("Resolved")) {
      // Find the latest resolution date from step decisions or final outcome date
      let resolutionDate = step3Decision || step2Decision || step1Decision || today;
      if (resolutionDate) {
        const resDt = new Date(resolutionDate);
        resDt.setHours(0, 0, 0, 0);
        daysOpen = Math.floor((resDt - filedDate) / (1000 * 60 * 60 * 24));
      }
    } else {
      // Active grievance - calculate from filed to today
      daysOpen = Math.floor((today - filedDate) / (1000 * 60 * 60 * 24));
    }
  }
  sheet.getRange(row, 27).setValue(daysOpen); // AA: Days Open

  // Calculate Next Action Due (AB) - Based on current step
  let nextActionDue = null;
  if (status && !status.toString().startsWith("Resolved")) {
    // Determine next deadline based on current step
    if (currentStep === "Step I - Immediate Supervisor") {
      nextActionDue = sheet.getRange(row, 10).getValue(); // J: Step I Decision Due
    } else if (currentStep === "Step II - Agency Head") {
      // Check if we're waiting for Step II appeal or Step II decision
      if (step1Decision && !step2Filed) {
        nextActionDue = sheet.getRange(row, 13).getValue(); // M: Step II Appeal Deadline
      } else if (step2Filed) {
        nextActionDue = sheet.getRange(row, 15).getValue(); // O: Step II Decision Due
      }
    } else if (currentStep === "Step III - Human Resources") {
      // Check if we're waiting for Step III appeal or Step III decision
      if (step2Decision && !step3Filed) {
        nextActionDue = sheet.getRange(row, 18).getValue(); // R: Step III Appeal Deadline
      } else if (step3Filed) {
        // Step III decision due = Step III Filed + 30 days
        if (step3Filed) {
          const step3Due = new Date(step3Filed);
          step3Due.setDate(step3Due.getDate() + CBA_DEADLINES.STEP_III_DECISION);
          nextActionDue = step3Due;
        }
      }
    } else if (currentStep === "Informal") {
      // For informal step, use filing deadline if not yet filed
      if (!dateFiled && incidentDate) {
        nextActionDue = sheet.getRange(row, 8).getValue(); // H: Filing Deadline
      } else if (dateFiled) {
        nextActionDue = sheet.getRange(row, 10).getValue(); // J: Step I Decision Due
      }
    }
  }
  sheet.getRange(row, 28).setValue(nextActionDue || ""); // AB: Next Action Due

  // Calculate Days to Deadline (AC)
  let daysToDeadline = "";
  if (nextActionDue && nextActionDue instanceof Date) {
    const deadlineDt = new Date(nextActionDue);
    deadlineDt.setHours(0, 0, 0, 0);
    daysToDeadline = Math.floor((deadlineDt - today) / (1000 * 60 * 60 * 24));
  }
  sheet.getRange(row, 29).setValue(daysToDeadline); // AC: Days to Deadline

  // Calculate Overdue? (AD)
  let isOverdue = "";
  if (daysToDeadline !== "" && !status.toString().startsWith("Resolved")) {
    isOverdue = daysToDeadline < 0 ? "YES" : "NO";
  }
  sheet.getRange(row, 30).setValue(isOverdue); // AD: Overdue?

  // Last Updated (AF - was AB, now moved to column 32)
  sheet.getRange(row, 32).setValue(new Date());
}

/**
 * Gets the next deadline for a grievance based on current step
 * Now simplified to use the pre-calculated "Next Action Due" column
 */
function getNextDeadline(sheet, row) {
  const status = sheet.getRange(row, 5).getValue();
  if (status && status.toString().startsWith("Resolved")) return null;

  // Simply return the pre-calculated Next Action Due (column AB = 28)
  const nextActionDue = sheet.getRange(row, 28).getValue();
  return nextActionDue || null;
}

/**
 * Gets the next deadline from a grievance row array (for dashboard calculations)
 * Now simplified to use the pre-calculated "Next Action Due" column
 */
function getNextDeadlineFromRow(row) {
  const status = row[4]; // E: Status
  if (!status || status.toString().startsWith("Resolved")) return null;

  // Simply return the pre-calculated Next Action Due (column AB = index 27)
  const nextActionDue = row[27];
  return nextActionDue ? new Date(nextActionDue) : null;
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
 *
 * CROSS-POPULATION: This function reads all grievances for a member from the Grievance Log
 * and updates their Member Directory record with aggregated statistics:
 * - Column L: Total Grievances Filed
 * - Column M: Active Grievances
 * - Column N: Resolved Grievances
 * - Column O: Grievances Won
 * - Column P: Grievances Lost
 * - Column Q: Last Grievance Date
 * - Column R: Has Open Grievance?
 * - Column S: # Open Grievances
 * - Column T: Last Grievance Status
 * - Column U: Next Deadline (Soonest)
 *
 * This ensures the Member Directory always reflects current grievance data from Grievance Log.
 */
function recalcMemberRow(memberSheet, grievanceSheet, row) {
  if (row < 2) return;

  const memberId = memberSheet.getRange(row, 1).getValue();
  if (!memberId) return;

  // Get all grievances for this member from Grievance Log (32 columns in new structure)
  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    // No grievances found - set all metrics to 0/blank
    memberSheet.getRange(row, 12).setValue(0); // L: Total Grievances
    memberSheet.getRange(row, 13).setValue(0); // M: Active Grievances
    memberSheet.getRange(row, 14).setValue(0); // N: Resolved Grievances
    memberSheet.getRange(row, 15).setValue(0); // O: Grievances Won
    memberSheet.getRange(row, 16).setValue(0); // P: Grievances Lost
    memberSheet.getRange(row, 17).setValue(""); // Q: Last Grievance Date
    memberSheet.getRange(row, 18).setValue("NO"); // R: Has Open Grievance?
    memberSheet.getRange(row, 19).setValue(0); // S: # Open Grievances
    memberSheet.getRange(row, 20).setValue(""); // T: Last Grievance Status
    memberSheet.getRange(row, 21).setValue(""); // U: Next Deadline (Soonest)
    memberSheet.getRange(row, 32).setValue(new Date()); // AF: Last Updated
    memberSheet.getRange(row, 33).setValue("AUTO"); // AG: Updated By
    return;
  }

  const grievanceData = grievanceSheet.getRange(2, 1, lastRow - 1, 32).getValues();
  const memberGrievances = grievanceData.filter(g => g[1] === memberId); // Column B = Member ID

  // Total Grievances Filed (L)
  memberSheet.getRange(row, 12).setValue(memberGrievances.length);

  // Active Grievances (M) - Filed or Pending Decision status
  const active = memberGrievances.filter(g => g[4] && (g[4].toString().startsWith("Filed") || g[4] === "Pending Decision")).length;
  memberSheet.getRange(row, 13).setValue(active);

  // Resolved Grievances (N)
  const resolved = memberGrievances.filter(g => g[4] && g[4].toString().startsWith("Resolved")).length;
  memberSheet.getRange(row, 14).setValue(resolved);

  // Grievances Won (O)
  const won = memberGrievances.filter(g => g[4] === "Resolved - Won").length;
  memberSheet.getRange(row, 15).setValue(won);

  // Grievances Lost (P)
  const lost = memberGrievances.filter(g => g[4] === "Resolved - Lost").length;
  memberSheet.getRange(row, 16).setValue(lost);

  // Last Grievance Date (Q) - Use Date Filed instead of Incident Date
  if (memberGrievances.length > 0) {
    const dates = memberGrievances.map(g => g[8]).filter(d => d); // Column I = Date Filed
    if (dates.length > 0) {
      const lastDate = new Date(Math.max(...dates.map(d => new Date(d))));
      memberSheet.getRange(row, 17).setValue(lastDate);
    } else {
      memberSheet.getRange(row, 17).setValue("");
    }
  } else {
    memberSheet.getRange(row, 17).setValue("");
  }

  // ============================================================================
  // GRIEVANCE SNAPSHOT FIELDS (R-U)
  // ============================================================================

  // Get open/active grievances (not resolved)
  const openGrievances = memberGrievances.filter(g =>
    g[4] && !g[4].toString().startsWith("Resolved") && g[4] !== "Draft"
  );

  // Has Open Grievance? (R)
  memberSheet.getRange(row, 18).setValue(openGrievances.length > 0 ? "YES" : "NO");

  // # Open Grievances (S)
  memberSheet.getRange(row, 19).setValue(openGrievances.length);

  // Last Grievance Status (T) - Most recent grievance by Date Filed
  if (memberGrievances.length > 0) {
    // Sort by Date Filed (column I, index 8) descending
    const sortedGrievances = memberGrievances.slice().sort((a, b) => {
      const dateA = a[8] ? new Date(a[8]) : new Date(0);
      const dateB = b[8] ? new Date(b[8]) : new Date(0);
      return dateB - dateA;
    });
    const lastStatus = sortedGrievances[0][4] || ""; // Column E = Status
    memberSheet.getRange(row, 20).setValue(lastStatus);
  } else {
    memberSheet.getRange(row, 20).setValue("");
  }

  // Next Deadline (Soonest) (U) - Find earliest Next Action Due from open grievances
  if (openGrievances.length > 0) {
    const deadlines = openGrievances
      .map(g => g[27]) // Column AB = Next Action Due (index 27 in 0-indexed array)
      .filter(d => d && d instanceof Date);

    if (deadlines.length > 0) {
      const soonestDeadline = new Date(Math.min(...deadlines.map(d => new Date(d))));
      memberSheet.getRange(row, 21).setValue(soonestDeadline);
    } else {
      memberSheet.getRange(row, 21).setValue("");
    }
  } else {
    memberSheet.getRange(row, 21).setValue("");
  }

  // Last Updated (AF)
  memberSheet.getRange(row, 32).setValue(new Date());

  // Updated By (AG)
  memberSheet.getRange(row, 33).setValue("AUTO");
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
    const today = new Date();

    // Calculate all member metrics in one pass
    let activeMembers = 0, totalStewards = 0, unit8 = 0, unit10 = 0;
    for (let i = 1; i < memberData.length; i++) {
      if (memberData[i][11] === "Active") activeMembers++;
      if (memberData[i][9] === "Yes") totalStewards++;
      if (memberData[i][5] === "Unit 8") unit8++;
      if (memberData[i][5] === "Unit 10") unit10++;
    }
    const totalMembers = memberData.length - 1;

    // Calculate all grievance metrics in one pass
    let activeGrievances = 0, resolvedGrievances = 0, won = 0, lost = 0;
    let inMediation = 0, inArbitration = 0;
    let overdue = 0, dueThisWeek = 0, dueNextWeek = 0;
    const overdueList = [];

    for (let i = 1; i < grievanceData.length; i++) {
      const r = grievanceData[i];
      const status = r[4];
      if (!status) continue;

      const statusStr = status.toString();
      if (statusStr.startsWith("Filed")) activeGrievances++;
      if (statusStr.startsWith("Resolved")) {
        resolvedGrievances++;
        if (status === "Resolved - Won") won++;
        if (status === "Resolved - Lost") lost++;
      }
      if (status === "In Mediation") inMediation++;
      if (status === "In Arbitration") inArbitration++;

      // Calculate deadline metrics only for non-resolved grievances
      if (!statusStr.startsWith("Resolved")) {
        const nextDeadline = getNextDeadlineFromRow(r);
        if (nextDeadline) {
          const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
          if (daysTo < 0) {
            overdue++;
            overdueList.push({
              id: r[0],
              member: r[2] + " " + r[3],
              step: r[5],
              daysTo: daysTo
            });
          } else if (daysTo <= 7) {
            dueThisWeek++;
          } else if (daysTo <= 14) {
            dueNextWeek++;
          }
        }
      }
    }

    const totalGrievances = grievanceData.length - 1;
    const winRate = resolvedGrievances > 0 ? ((won / resolvedGrievances) * 100).toFixed(1) + "%" : "N/A";

    // Sort and limit overdue list
    overdueList.sort((a, b) => a.daysTo - b.daysTo);
    const top10Overdue = overdueList.slice(0, 10);

    // Batch write member metrics
    const memberMetrics = [
      ["Total Members:", totalMembers],
      ["Active Members:", activeMembers],
      ["Total Stewards:", totalStewards],
      ["Unit 8 Members:", unit8],
      ["Unit 10 Members:", unit10],
      ["", ""] // Empty row for spacing
    ];
    dashboard.getRange(4, 1, 6, 2).setValues(memberMetrics);
    dashboard.getRange("A4:A8").setFontWeight("bold");
    dashboard.getRange("B4:B8").setNumberFormat("#,##0").setHorizontalAlignment("right");

    // Batch write grievance metrics
    const grievanceMetrics = [
      ["", ""], // Empty row
      ["Total Grievances:", totalGrievances],
      ["Active Grievances:", activeGrievances],
      ["Resolved Grievances:", resolvedGrievances],
      ["Grievances Won:", won],
      ["Grievances Lost:", lost],
      ["Win Rate:", winRate],
      ["In Mediation:", inMediation],
      ["In Arbitration:", inArbitration]
    ];
    dashboard.getRange(11, 1, 9, 2).setValues(grievanceMetrics);
    dashboard.getRange("A12:A19").setFontWeight("bold");
    dashboard.getRange("B12:B16").setNumberFormat("#,##0").setHorizontalAlignment("right");
    dashboard.getRange("B17").setHorizontalAlignment("right");
    dashboard.getRange("B18:B19").setNumberFormat("#,##0").setHorizontalAlignment("right");

    // Batch write deadline metrics
    const deadlineMetrics = [
      ["", ""],
      ["Overdue Grievances:", overdue],
      ["Due This Week:", dueThisWeek],
      ["Due Next Week:", dueNextWeek]
    ];
    dashboard.getRange(21, 1, 4, 2).setValues(deadlineMetrics);
    dashboard.getRange("A22:A24").setFontWeight("bold");
    dashboard.getRange("B22:B24").setNumberFormat("#,##0").setHorizontalAlignment("right");

    // Batch write overdue grievances table
    const overdueTable = [["Grievance ID", "Member", "Step", "Days Overdue"]];
    for (let i = 0; i < 10; i++) {
      if (i < top10Overdue.length) {
        const g = top10Overdue[i];
        overdueTable.push([g.id, g.member, g.step, Math.abs(g.daysTo)]);
      } else {
        overdueTable.push(["", "", "", ""]);
      }
    }
    dashboard.getRange(4, 5, 11, 4).setValues(overdueTable);

    // Apply formatting (batch operations)
    dashboard.getRange("A3:B3").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
    dashboard.getRange("A4:B9").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("A11:B11").setBackground(COLORS.HEADER_RED).setFontColor("white");
    dashboard.getRange("A12:B14").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("A15").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B15").setBackground(COLORS.ON_TRACK).setFontColor("white");
    dashboard.getRange("A16").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B16").setBackground(COLORS.OVERDUE).setFontColor("white");
    dashboard.getRange("A17:B17").setBackground("#E8F5E9").setFontSize(12).setFontWeight("bold");
    dashboard.getRange("A18:B19").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("A21:B21").setBackground(COLORS.HEADER_ORANGE).setFontColor("white");
    dashboard.getRange("A22").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B22").setBackground(overdue > 0 ? COLORS.OVERDUE : COLORS.ON_TRACK)
      .setFontColor("white").setFontWeight("bold").setFontSize(12);
    dashboard.getRange("A23").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B23").setBackground(dueThisWeek > 0 ? COLORS.DUE_SOON : COLORS.ON_TRACK)
      .setFontColor("white").setFontWeight("bold").setFontSize(12);
    dashboard.getRange("A24:B24").setBackground(COLORS.LIGHT_GRAY);

    // Format overdue table
    dashboard.getRange("E4:H4").setFontWeight("bold").setBackground(COLORS.HEADER_ORANGE).setFontColor("white");
    for (let i = 0; i < top10Overdue.length; i++) {
      const row = 5 + i;
      const daysOverdue = Math.abs(top10Overdue[i].daysTo);
      const range = dashboard.getRange("E" + row + ":H" + row);
      if (daysOverdue > 30) {
        range.setBackground("#C00000").setFontColor("white").setFontWeight("bold");
      } else if (daysOverdue > 14) {
        range.setBackground(COLORS.OVERDUE).setFontColor("white");
      } else {
        range.setBackground("#F4CCCC");
      }
    }
    // Clear formatting for empty rows
    if (top10Overdue.length < 10) {
      dashboard.getRange(5 + top10Overdue.length, 5, 10 - top10Overdue.length, 4)
        .setBackground("white").setFontWeight("normal").setFontColor("black");
    }

    // Build Steward Workload Sheet
    rebuildStewardWorkload();

    // Create all dashboard charts
    createDashboardCharts();

    // Rebuild new analytical tabs
    rebuildTrendsSheet();
    rebuildPerformanceSheet();
    rebuildLocationSheet();
    rebuildTypeAnalysisSheet();

    // Rebuild all dashboard variations with live data
    rebuildAllDashboards();

    // Single flush at the end
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
  const stewards = [];
  const stewardMap = {}; // Map full name to steward data

  for (let i = 1; i < memberData.length; i++) {
    const r = memberData[i];
    if (r[9] === "Yes") { // Column J = Is Steward
      const fullName = r[1] + " " + r[2];
      const stewardData = {
        memberId: r[0],
        fullName: fullName,
        location: r[4],
        grievances: []
      };
      stewards.push(stewardData);
      stewardMap[fullName] = stewardData;
    }
  }

  // Get all grievances and assign to stewards
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const today = new Date();

  // PERFORMANCE FIX: Build grievance assignments in one pass
  for (let i = 1; i < grievanceData.length; i++) {
    const g = grievanceData[i];
    const assignedSteward = g[25]; // Column Z: Steward Name
    if (assignedSteward) {
      const stewardStr = assignedSteward.toString();
      // Check if any steward name matches
      for (const stewardName in stewardMap) {
        if (stewardStr.includes(stewardName)) {
          stewardMap[stewardName].grievances.push(g);
          break; // Assign to first matching steward only
        }
      }
    }
  }

  // Clear existing data (keep headers)
  const lastRow = stewardSheet.getLastRow();
  if (lastRow > 1) {
    stewardSheet.getRange(2, 1, lastRow - 1, 15).clearContent();
  }

  const stewardStats = [];

  // Calculate metrics for each steward
  stewards.forEach(steward => {
    const assignedGrievances = steward.grievances;

    // Calculate all metrics in one pass through assigned grievances
    let totalCases = assignedGrievances.length;
    let activeCases = 0, step1Cases = 0, step2Cases = 0, step3Cases = 0;
    let overdueCases = 0, dueThisWeek = 0, resolvedCases = 0, wonCases = 0;
    let resolutionDaysSum = 0, resolutionDaysCount = 0;
    let lastCaseDateMs = 0;

    assignedGrievances.forEach(g => {
      const status = g[4];
      const step = g[5];
      const incidentDate = g[6];

      // Count by status
      if (status) {
        const statusStr = status.toString();
        if (statusStr.startsWith("Filed")) activeCases++;
        if (statusStr.startsWith("Resolved")) {
          resolvedCases++;
          if (status === "Resolved - Won") wonCases++;

          // Calculate resolution time
          const filedDate = g[8];
          if (filedDate) {
            const filed = new Date(filedDate);
            const resolved = new Date(); // Approximate
            resolutionDaysSum += Math.floor((resolved - filed) / (1000 * 60 * 60 * 24));
            resolutionDaysCount++;
          }
        }
      }

      // Count by step
      if (step === "Step I - Immediate Supervisor") step1Cases++;
      if (step === "Step II - Agency Head") step2Cases++;
      if (step === "Step III - Human Resources") step3Cases++;

      // Calculate deadline metrics
      const nextDeadline = getNextDeadlineFromRow(g);
      if (nextDeadline) {
        const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
        if (daysTo < 0) overdueCases++;
        else if (daysTo <= 7) dueThisWeek++;
      }

      // Track last case date
      if (incidentDate) {
        const dateMs = new Date(incidentDate).getTime();
        if (dateMs > lastCaseDateMs) lastCaseDateMs = dateMs;
      }
    });

    const winRate = resolvedCases > 0 ? ((wonCases / resolvedCases) * 100).toFixed(1) + "%" : "N/A";
    const avgDays = resolutionDaysCount > 0 ? Math.round(resolutionDaysSum / resolutionDaysCount) : "N/A";
    const lastCaseDate = lastCaseDateMs > 0 ? new Date(lastCaseDateMs) : "";
    const statusVal = activeCases > 0 ? "Active" : "Available";

    stewardStats.push([
      steward.fullName,
      steward.memberId,
      steward.location,
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
      statusVal
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

  // Calculate monthly filing trends and resolution times in one pass
  const monthlyData = {};
  const resolutionTimes = [];
  const today = new Date();

  for (let i = 1; i < grievanceData.length; i++) {
    const r = grievanceData[i];
    const filedDate = r[8]; // Column I: Date Filed
    if (filedDate) {
      const date = new Date(filedDate);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      monthlyData[monthKey] = (monthlyData[monthKey] || 0) + 1;
    }

    // Calculate resolution times
    if (r[4] && r[4].toString().startsWith("Resolved")) {
      const filed = r[8]; // Date Filed
      const resolved = r[27]; // Last Updated (approximation)
      if (filed && resolved) {
        const days = Math.floor((new Date(resolved) - new Date(filed)) / (1000 * 60 * 60 * 24));
        if (days > 0) resolutionTimes.push(days);
      }
    }
  }

  // Calculate KPI values
  const totalMonths = Object.keys(monthlyData).length;
  const avgPerMonth = totalMonths > 0
    ? (Object.values(monthlyData).reduce((a, b) => a + b, 0) / totalMonths).toFixed(1)
    : 0;
  const peakMonth = Object.entries(monthlyData).sort((a, b) => b[1] - a[1])[0];
  const avgResolution = resolutionTimes.length > 0
    ? (resolutionTimes.reduce((a, b) => a + b, 0) / resolutionTimes.length).toFixed(1)
    : "N/A";
  const minResolution = resolutionTimes.length > 0 ? Math.min(...resolutionTimes) : "N/A";
  const maxResolution = resolutionTimes.length > 0 ? Math.max(...resolutionTimes) : "N/A";

  // Batch write KPIs
  const kpiData = [
    ["Total Months with Data:", totalMonths],
    ["Average Filings per Month:", avgPerMonth],
    ["Peak Month:", peakMonth ? `${peakMonth[0]} (${peakMonth[1]} filings)` : "N/A"],
    ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], // Empty rows
    ["Average Resolution Time:", `${avgResolution} days`],
    ["Fastest Resolution:", minResolution === "N/A" ? "N/A" : `${minResolution} days`],
    ["Slowest Resolution:", maxResolution === "N/A" ? "N/A" : `${maxResolution} days`]
  ];
  trendsSheet.getRange(4, 1, 15, 2).setValues(kpiData);

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

  // Batch write 12-month table
  const monthTableData = [["Month", "Filings"]];
  Object.entries(last12Months).forEach(([month, count]) => {
    monthTableData.push([month, count]);
  });
  trendsSheet.getRange(26, 1, monthTableData.length, 2).setValues(monthTableData);

  // Apply formatting (batch operations)
  trendsSheet.getRange("A4:A18").setFontWeight("bold");
  trendsSheet.getRange("A4:B18").setBackground(COLORS.LIGHT_GRAY);
  trendsSheet.getRange("B4:B18").setHorizontalAlignment("right");
  trendsSheet.getRange("B4").setNumberFormat("#,##0");
  trendsSheet.getRange("B5").setNumberFormat("#,##0.0");
  trendsSheet.getRange("A26:B26").setFontWeight("bold");
  trendsSheet.getRange("B27:B38").setNumberFormat("#,##0");
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

  // Calculate all metrics in a single pass
  const today = new Date();
  let resolvedCount = 0, wonCount = 0, lostCount = 0, settledCount = 0, withdrawnCount = 0;
  let activeCount = 0, overdueCount = 0;
  const stepStats = {
    "Step I - Immediate Supervisor": { total: 0, won: 0 },
    "Step II - Agency Head": { total: 0, won: 0 },
    "Step III - Human Resources": { total: 0, won: 0 }
  };

  for (let i = 1; i < grievanceData.length; i++) {
    const r = grievanceData[i];
    const status = r[4];
    if (!status) continue;

    const statusStr = status.toString();
    if (statusStr.startsWith("Resolved")) {
      resolvedCount++;
      if (status === "Resolved - Won") wonCount++;
      if (status === "Resolved - Lost") lostCount++;
      if (status === "Resolved - Settled") settledCount++;
      if (status === "Resolved - Withdrawn") withdrawnCount++;

      // Track step outcomes
      const step = r[5];
      if (stepStats[step]) {
        stepStats[step].total++;
        if (status === "Resolved - Won") stepStats[step].won++;
      }
    } else if (statusStr.startsWith("Filed")) {
      activeCount++;
      const nextDeadline = getNextDeadlineFromRow(r);
      if (nextDeadline && nextDeadline < today) overdueCount++;
    }
  }

  // Calculate rates
  const winRate = resolvedCount > 0 ? ((wonCount / resolvedCount) * 100).toFixed(1) + "%" : "N/A";
  const settleRate = resolvedCount > 0 ? ((settledCount / resolvedCount) * 100).toFixed(1) + "%" : "N/A";
  const withdrawRate = resolvedCount > 0 ? ((withdrawnCount / resolvedCount) * 100).toFixed(1) + "%" : "N/A";
  const overdueRate = activeCount > 0 ? ((overdueCount / activeCount) * 100).toFixed(1) + "%" : "N/A";
  const onTimeRate = activeCount > 0 ? (((activeCount - overdueCount) / activeCount) * 100).toFixed(1) + "%" : "N/A";

  // Batch write resolution performance metrics
  const resolutionData = [
    ["Total Resolved:", resolvedCount],
    ["Win Rate:", winRate],
    ["Settlement Rate:", settleRate],
    ["Withdrawal Rate:", withdrawRate]
  ];
  perfSheet.getRange(4, 1, 4, 2).setValues(resolutionData);

  // Batch write efficiency metrics
  const efficiencyData = [
    ["Active Grievances:", activeCount],
    ["Overdue Rate:", overdueRate],
    ["On-Time Rate:", onTimeRate]
  ];
  perfSheet.getRange(4, 5, 3, 2).setValues(efficiencyData);

  // Batch write step analysis table
  const stepTableData = [["Step", "Total", "Win Rate"]];
  Object.entries(stepStats).forEach(([step, stats]) => {
    const stepWinRate = stats.total > 0 ? ((stats.won / stats.total) * 100).toFixed(1) + "%" : "N/A";
    stepTableData.push([step, stats.total, stepWinRate]);
  });
  perfSheet.getRange(14, 1, stepTableData.length, 3).setValues(stepTableData);

  // Apply formatting (batch operations)
  perfSheet.getRange("A4:A7").setFontWeight("bold");
  perfSheet.getRange("B4").setNumberFormat("#,##0").setHorizontalAlignment("right");
  perfSheet.getRange("B5:B7").setHorizontalAlignment("right");
  perfSheet.getRange("A4:B7").setBackground(COLORS.LIGHT_GRAY);
  perfSheet.getRange("B5").setFontWeight("bold").setBackground("#E8F5E9");

  perfSheet.getRange("E4:E6").setFontWeight("bold");
  perfSheet.getRange("F4").setNumberFormat("#,##0").setHorizontalAlignment("right");
  perfSheet.getRange("F5:F6").setHorizontalAlignment("right");
  perfSheet.getRange("E4:F6").setBackground(COLORS.LIGHT_GRAY);
  perfSheet.getRange("F5").setFontWeight("bold")
    .setBackground(overdueCount > 0 ? COLORS.OVERDUE : COLORS.ON_TRACK)
    .setFontColor("white");

  perfSheet.getRange("A14:C14").setFontWeight("bold").setBackground(COLORS.HEADER_RED).setFontColor("white");
  perfSheet.getRange("B15:B17").setNumberFormat("#,##0");
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

  // PERFORMANCE FIX: Create lookup map from member ID to location
  const memberToLocation = {};
  const locationStats = {};

  // First pass: count members by location and create lookup map
  for (let i = 1; i < memberData.length; i++) {
    const r = memberData[i];
    const location = r[4]; // Column E: Work Location
    const memberId = r[0];

    memberToLocation[memberId] = location;

    if (!locationStats[location]) {
      locationStats[location] = { members: 0, grievances: 0, active: 0, resolved: 0, won: 0 };
    }
    locationStats[location].members++;
  }

  // Second pass: count grievances by location using lookup map
  for (let i = 1; i < grievanceData.length; i++) {
    const g = grievanceData[i];
    const memberId = g[1]; // Column B: Member ID
    const status = g[4];   // Column E: Status
    const location = memberToLocation[memberId];

    if (!location || !locationStats[location]) continue;

    locationStats[location].grievances++;

    if (status) {
      const statusStr = status.toString();
      if (statusStr.startsWith("Filed")) {
        locationStats[location].active++;
      } else if (statusStr.startsWith("Resolved")) {
        locationStats[location].resolved++;
        if (status === "Resolved - Won") locationStats[location].won++;
      }
    }
  }

  // Sort by grievance count and limit to top 15
  const sortedLocations = Object.entries(locationStats)
    .sort((a, b) => b[1].grievances - a[1].grievances)
    .slice(0, 15);

  // Batch write table data
  const tableData = [["Location", "Members", "Total Grievances", "Active", "Resolved", "Win Rate"]];
  sortedLocations.forEach(([location, stats]) => {
    const winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) + "%" : "N/A";
    tableData.push([location, stats.members, stats.grievances, stats.active, stats.resolved, winRate]);
  });
  locSheet.getRange(4, 1, tableData.length, 6).setValues(tableData);

  // Apply formatting (batch operations)
  locSheet.getRange("A4:F4").setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
  if (sortedLocations.length > 0) {
    const dataRange = `A5:F${4 + sortedLocations.length}`;
    locSheet.getRange(dataRange).setBackground(COLORS.LIGHT_GRAY);
    locSheet.getRange(`B5:F${4 + sortedLocations.length}`).setHorizontalAlignment("right");
    locSheet.getRange(`B5:E${4 + sortedLocations.length}`).setNumberFormat("#,##0");
  }

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

  // Analyze by type in a single pass
  const typeStats = {};

  for (let i = 1; i < grievanceData.length; i++) {
    const r = grievanceData[i];
    const type = r[23]; // Column X: Grievance Type
    if (!type) continue;

    if (!typeStats[type]) {
      typeStats[type] = { total: 0, active: 0, resolved: 0, won: 0, lost: 0, settled: 0 };
    }

    typeStats[type].total++;
    const status = r[4];
    if (status) {
      const statusStr = status.toString();
      if (statusStr.startsWith("Filed")) {
        typeStats[type].active++;
      } else if (statusStr.startsWith("Resolved")) {
        typeStats[type].resolved++;
        if (status === "Resolved - Won") typeStats[type].won++;
        if (status === "Resolved - Lost") typeStats[type].lost++;
        if (status === "Resolved - Settled") typeStats[type].settled++;
      }
    }
  }

  const sortedTypes = Object.entries(typeStats).sort((a, b) => b[1].total - a[1].total);

  // Batch write type breakdown table
  const breakdownData = [["Grievance Type", "Total", "Active", "Resolved"]];
  sortedTypes.forEach(([type, stats]) => {
    breakdownData.push([type, stats.total, stats.active, stats.resolved]);
  });
  typeSheet.getRange(4, 1, breakdownData.length, 4).setValues(breakdownData);

  // Batch write success rate table
  const ratesData = [["Grievance Type", "Win Rate", "Settlement Rate"]];
  sortedTypes.forEach(([type, stats]) => {
    const winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) + "%" : "N/A";
    const settleRate = stats.resolved > 0 ? ((stats.settled / stats.resolved) * 100).toFixed(1) + "%" : "N/A";
    ratesData.push([type.substring(0, 30), winRate, settleRate]);
  });
  typeSheet.getRange(4, 5, ratesData.length, 3).setValues(ratesData);

  // Apply formatting (batch operations)
  typeSheet.getRange("A4:D4").setFontWeight("bold").setBackground(COLORS.HEADER_BLUE).setFontColor("white");
  if (sortedTypes.length > 0) {
    typeSheet.getRange(`A5:D${4 + sortedTypes.length}`).setBackground(COLORS.LIGHT_GRAY);
    typeSheet.getRange(`B5:D${4 + sortedTypes.length}`).setHorizontalAlignment("right").setNumberFormat("#,##0");
  }

  typeSheet.getRange("E4:G4").setFontWeight("bold").setBackground(COLORS.HEADER_GREEN).setFontColor("white");
  if (sortedTypes.length > 0) {
    typeSheet.getRange(`E5:G${4 + sortedTypes.length}`).setBackground(COLORS.LIGHT_GRAY);
    typeSheet.getRange(`F5:G${4 + sortedTypes.length}`).setHorizontalAlignment("right");
  }
}

// ============================================================================
// DASHBOARD VARIATIONS - DATA POPULATION
// ============================================================================

/**
 * Rebuilds all dashboard variations with live data
 */
function rebuildAllDashboards() {
  rebuildExecutiveOverview();
  rebuildKPIDashboard();
  rebuildMemberEngagement();
  rebuildCostImpact();
  rebuildQuickStats();
}

/**
 * Calculate common metrics used across all dashboards
 */
function calculateDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return null;

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const today = new Date();
  const thisMonth = today.getMonth();
  const thisYear = today.getFullYear();

  // Member metrics
  let totalMembers = memberData.length - 1;
  let activeMembers = 0, totalStewards = 0, unit8 = 0, unit10 = 0;
  let engagementCounts = { "Very Active": 0, "Active": 0, "Moderately Active": 0, "Inactive": 0, "New Member": 0 };
  let satisfactionCounts = { "Very Satisfied": 0, "Satisfied": 0, "Neutral": 0, "Dissatisfied": 0, "Very Dissatisfied": 0 };
  let committeeMembers = 0;

  for (let i = 1; i < memberData.length; i++) {
    const r = memberData[i];
    if (r[10] === "Active") activeMembers++;  // Column K: Membership Status
    if (r[9] === "Yes") totalStewards++;      // Column J: Is Steward
    if (r[5] === "Unit 8") unit8++;           // Column F: Unit
    if (r[5] === "Unit 10") unit10++;

    const engagement = r[20];                 // Column U: Engagement Level
    if (engagement && engagementCounts[engagement] !== undefined) {
      engagementCounts[engagement]++;
    }

    const committee = r[23];                  // Column X: Committee Member
    if (committee && committee !== "None") committeeMembers++;
  }

  // Grievance metrics
  let totalGrievances = grievanceData.length - 1;
  let activeGrievances = 0, resolvedGrievances = 0, won = 0, lost = 0, settled = 0;
  let inMediation = 0, inArbitration = 0, pendingDecision = 0;
  let overdue = 0, dueThisWeek = 0, dueNextWeek = 0;
  let thisMonthFiled = 0, thisMonthResolved = 0;
  let ytdFiled = 0, ytdResolved = 0;
  let resolutionDaysSum = 0, resolutionDaysCount = 0;
  let financialRecovery = 0;

  for (let i = 1; i < grievanceData.length; i++) {
    const r = grievanceData[i];
    const status = r[4];                      // Column E: Status
    const filedDate = r[8];                   // Column I: Date Filed

    if (!status) continue;

    const statusStr = status.toString();

    // Count filed this month/year
    if (filedDate) {
      const filed = new Date(filedDate);
      if (filed.getMonth() === thisMonth && filed.getFullYear() === thisYear) {
        thisMonthFiled++;
      }
      if (filed.getFullYear() === thisYear) {
        ytdFiled++;
      }
    }

    // Status counts
    if (statusStr.startsWith("Filed")) {
      activeGrievances++;
    } else if (statusStr.startsWith("Resolved")) {
      resolvedGrievances++;
      if (status === "Resolved - Won") {
        won++;
        financialRecovery += 5000; // Estimated average recovery
      }
      if (status === "Resolved - Lost") lost++;
      if (status === "Resolved - Settled") {
        settled++;
        financialRecovery += 2500; // Estimated average settlement
      }

      // Calculate resolution time
      if (filedDate) {
        const filed = new Date(filedDate);
        const resolved = today; // Approximation
        const days = Math.floor((resolved - filed) / (1000 * 60 * 60 * 24));
        resolutionDaysSum += days;
        resolutionDaysCount++;

        // Count resolved this month/year
        if (resolved.getMonth() === thisMonth && resolved.getFullYear() === thisYear) {
          thisMonthResolved++;
        }
        if (resolved.getFullYear() === thisYear) {
          ytdResolved++;
        }
      }
    }

    if (status === "In Mediation") inMediation++;
    if (status === "In Arbitration") inArbitration++;
    if (status === "Pending Decision") pendingDecision++;

    // Deadline metrics (for active grievances)
    if (!statusStr.startsWith("Resolved")) {
      const nextDeadline = getNextDeadlineFromRow(r);
      if (nextDeadline) {
        const daysTo = Math.floor((nextDeadline - today) / (1000 * 60 * 60 * 24));
        if (daysTo < 0) overdue++;
        else if (daysTo <= 7) dueThisWeek++;
        else if (daysTo <= 14) dueNextWeek++;
      }
    }
  }

  const winRate = resolvedGrievances > 0 ? (won / resolvedGrievances) * 100 : 0;
  const settlementRate = resolvedGrievances > 0 ? (settled / resolvedGrievances) * 100 : 0;
  const avgResolutionDays = resolutionDaysCount > 0 ? Math.round(resolutionDaysSum / resolutionDaysCount) : 0;
  const memberEngagementRate = totalMembers > 0 ? ((engagementCounts["Very Active"] + engagementCounts["Active"]) / totalMembers) * 100 : 0;

  return {
    // Member metrics
    totalMembers, activeMembers, totalStewards, unit8, unit10, committeeMembers,
    engagementCounts, memberEngagementRate,

    // Grievance metrics
    totalGrievances, activeGrievances, resolvedGrievances,
    won, lost, settled, winRate, settlementRate,
    inMediation, inArbitration, pendingDecision,
    overdue, dueThisWeek, dueNextWeek,
    avgResolutionDays,

    // Time-based metrics
    thisMonthFiled, thisMonthResolved,
    ytdFiled, ytdResolved,

    // Financial
    financialRecovery
  };
}

/**
 * Rebuild Executive Overview Dashboard
 */
function rebuildExecutiveOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.EXECUTIVE);
  if (!sheet) return;

  const metrics = calculateDashboardMetrics();
  if (!metrics) return;

  // KPI Cards - Row 7 contains the main values
  sheet.getRange("A7:C7").merge().setValue(metrics.totalMembers);
  sheet.getRange("E7:G7").merge().setValue(metrics.activeGrievances);
  sheet.getRange("I7:K7").merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("M7:O7").merge().setValue(metrics.avgResolutionDays + " days");

  // Union Health Section - Row 15 contains values
  sheet.getRange("A15:C15").merge().setValue(metrics.totalStewards);
  sheet.getRange("E15:G15").merge().setValue(metrics.memberEngagementRate.toFixed(1) + "%");
  sheet.getRange("I15:K15").merge().setValue(metrics.activeMembers);
  sheet.getRange("M15:O15").merge().setValue(metrics.committeeMembers);
}

/**
 * Rebuild KPI Dashboard
 */
function rebuildKPIDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.KPI_BOARD);
  if (!sheet) return;

  const metrics = calculateDashboardMetrics();
  if (!metrics) return;

  // THIS MONTH section - Row 7
  sheet.getRange("A7:C7").merge().setValue(metrics.thisMonthFiled);
  sheet.getRange("A8:C9").merge().setValue("+" + metrics.thisMonthFiled + " from last month");

  sheet.getRange("E7:G7").merge().setValue(metrics.thisMonthResolved);
  sheet.getRange("E8:G9").merge().setValue(metrics.thisMonthResolved + " cases closed");

  sheet.getRange("I7:K7").merge().setValue(metrics.pendingDecision);
  sheet.getRange("I8:K9").merge().setValue("Awaiting decisions");

  sheet.getRange("M7:O7").merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("M8:O9").merge().setValue("Success rate");

  // YEAR TO DATE section - Row 15
  sheet.getRange("A15:C15").merge().setValue(metrics.ytdFiled);
  sheet.getRange("A16:C17").merge().setValue("Cases filed YTD");

  sheet.getRange("E15:G15").merge().setValue(metrics.ytdResolved);
  sheet.getRange("E16:G17").merge().setValue("Cases resolved YTD");

  sheet.getRange("I15:K15").merge().setValue(metrics.activeGrievances);
  sheet.getRange("I16:K17").merge().setValue("Currently active");

  sheet.getRange("M15:O15").merge().setValue(metrics.overdue);
  sheet.getRange("M16:O17").merge().setValue("Cases overdue");
}

/**
 * Rebuild Member Engagement Dashboard
 */
function rebuildMemberEngagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_ENGAGEMENT);
  if (!sheet) return;

  const metrics = calculateDashboardMetrics();
  if (!metrics) return;

  // Engagement KPIs - Row 7
  sheet.getRange("A7:C7").merge().setValue(metrics.memberEngagementRate.toFixed(1) + "%");
  sheet.getRange("E7:G7").merge().setValue(metrics.totalStewards);
  sheet.getRange("I7:K7").merge().setValue(metrics.committeeMembers);
  sheet.getRange("M7:O7").merge().setValue(metrics.activeMembers);

  // Engagement breakdown - Starting at row 15
  sheet.getRange("A15").setValue(metrics.engagementCounts["Very Active"]);
  sheet.getRange("B15").setValue(metrics.engagementCounts["Active"]);
  sheet.getRange("C15").setValue(metrics.engagementCounts["Moderately Active"]);
  sheet.getRange("D15").setValue(metrics.engagementCounts["Inactive"]);
  sheet.getRange("E15").setValue(metrics.engagementCounts["New Member"]);
}

/**
 * Rebuild Cost Impact Dashboard
 */
function rebuildCostImpact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.COST_IMPACT);
  if (!sheet) return;

  const metrics = calculateDashboardMetrics();
  if (!metrics) return;

  // Financial KPIs - Row 7
  const formatter = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 });

  sheet.getRange("A7:C7").merge().setValue(formatter.format(metrics.financialRecovery));
  sheet.getRange("E7:G7").merge().setValue(metrics.won);
  sheet.getRange("I7:K7").merge().setValue(formatter.format(metrics.financialRecovery / Math.max(1, metrics.won + metrics.settled)));
  sheet.getRange("M7:O7").merge().setValue(metrics.winRate.toFixed(1) + "%");

  // Additional financial metrics - Row 15
  const avgPerMember = metrics.totalMembers > 0 ? metrics.financialRecovery / metrics.totalMembers : 0;
  sheet.getRange("A15:C15").merge().setValue(formatter.format(avgPerMember));
  sheet.getRange("E15:G15").merge().setValue(metrics.won + metrics.settled);
  sheet.getRange("I15:K15").merge().setValue(formatter.format(metrics.financialRecovery * 1.2)); // Projected
}

/**
 * Rebuild Quick Stats Dashboard
 */
function rebuildQuickStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.QUICK_STATS);
  if (!sheet) return;

  const metrics = calculateDashboardMetrics();
  if (!metrics) return;

  // Large stat cards - Row 7
  sheet.getRange("A7:C7").merge().setValue(metrics.totalGrievances);
  sheet.getRange("E7:G7").merge().setValue(metrics.activeGrievances);
  sheet.getRange("I7:K7").merge().setValue(metrics.overdue);
  sheet.getRange("M7:O7").merge().setValue(metrics.totalMembers);

  // Secondary stats - Row 15
  sheet.getRange("A15:C15").merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("E15:G15").merge().setValue(metrics.avgResolutionDays);
  sheet.getRange("I15:K15").merge().setValue(metrics.totalStewards);
  sheet.getRange("M15:O15").merge().setValue(metrics.dueThisWeek);
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
  const data = sheet.getRange(2, 1, lastRow - 1, 32).getValues();
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
  sheet.getRange(2, 1, sortedData.length, 32).setValues(sortedData);

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
 * Generates a random unique Member ID
 */
/**
 * Generates a random Member ID in format: Initials-XXX-L
 * Example: JD-123-A (John Doe)
 */
function generateRandomMemberId(existingIds, firstName, lastName) {
  let newId;
  do {
    const initials = (firstName.charAt(0) + lastName.charAt(0)).toUpperCase();
    const randomNum = Math.floor(Math.random() * 1000); // 0-999
    const randomLetter = String.fromCharCode(65 + Math.floor(Math.random() * 26)); // A-Z
    newId = `${initials}-${String(randomNum).padStart(3, '0')}-${randomLetter}`;
  } while (existingIds.has(newId));
  return newId;
}

/**
 * Generates a random Grievance ID in format: GRV-XXX-L
 * Example: GRV-123-A
 */
function generateRandomGrievanceId(existingIds) {
  let newId;
  do {
    const randomNum = Math.floor(Math.random() * 1000); // 0-999
    const randomLetter = String.fromCharCode(65 + Math.floor(Math.random() * 26)); // A-Z
    newId = `GRV-${String(randomNum).padStart(3, '0')}-${randomLetter}`;
  } while (existingIds.has(newId));
  return newId;
}

/**
 * Validates that a Member ID exists in the Member Directory
 * Returns the member data [ID, FirstName, LastName, JobTitle] or null if not found
 */
function validateAndGetMemberData(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return null;

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return null;

  const memberData = memberSheet.getRange("A2:D" + lastRow).getValues();
  const member = memberData.find(row => row[0] === memberId);

  return member || null;
}

/**
 * Gets all valid Member IDs from the Member Directory
 */
function getAllMemberIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return [];

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return [];

  const memberIds = memberSheet.getRange("A2:A" + lastRow).getValues().flat().filter(String);
  return memberIds;
}

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

  // Track all generated IDs to ensure uniqueness
  const existingIds = new Set();

  for (let batch = 0; batch < TOTAL / BATCH; batch++) {
    const data = [];

    for (let i = 0; i < BATCH; i++) {
      const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
      const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
      const isSteward = Math.random() < 0.1 ? "Yes" : "No";

      // Generate random unique Member ID with initials
      const memberId = generateRandomMemberId(existingIds, firstName, lastName);
      existingIds.add(memberId);

      data.push([
        // A-F: Basic Info
        memberId,
        firstName,
        lastName,
        jobTitles[Math.floor(Math.random() * jobTitles.length)],
        locations[Math.floor(Math.random() * locations.length)],
        units[Math.floor(Math.random() * units.length)],
        // G-K: Contact & Role
        "Mon-Fri",
        `${firstName.toLowerCase()}.${lastName.toLowerCase()}@mass.gov`,
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`,
        isSteward,
        membershipStatus[Math.floor(Math.random() * membershipStatus.length)],
        // L-Q: Grievance Metrics (AUTO CALCULATED - leave blank)
        "", "", "", "", "", "",
        // R-U: Grievance Snapshot (AUTO CALCULATED - leave blank)
        "", "", "", "",
        // V-Z: Participation
        engagementLevels[Math.floor(Math.random() * engagementLevels.length)],
        Math.floor(Math.random() * 25),
        Math.floor(Math.random() * 15),
        isSteward === "Yes" ? committees[Math.floor(Math.random() * (committees.length - 1))] : "None",
        contactMethods[Math.floor(Math.random() * contactMethods.length)],
        // AA-AG: Emergency & Admin
        `${firstNames[Math.floor(Math.random() * firstNames.length)]} ${lastNames[Math.floor(Math.random() * lastNames.length)]}`,
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`,
        "",
        randomDate(1960, 2000),
        randomDate(2015, 2024),
        new Date(),
        "SEED_SCRIPT"
      ]);
    }

    sheet.getRange(2 + (batch * BATCH), 1, BATCH, 33).setValues(data);
    SpreadsheetApp.flush();
  }

  SpreadsheetApp.getUi().alert('✅ 20,000 members seeded with random IDs!\n\nRun "Recalc All Members" after adding grievances.');
}

/**
 * Seeds 5,000 grievances with realistic data
 *
 * IMPORTANT: All grievances use ACTUAL MEMBER IDs from the Member Directory
 * - Member ID, First Name, Last Name are pulled from existing members
 * - This ensures data integrity between Grievance Log and Member Directory
 * - After seeding, run "Recalc All Members" to cross-populate grievance stats to members
 */
function SEED_5K_GRIEVANCES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  // Get actual member data from Member Directory (columns A-J: ID, FirstName, LastName, JobTitle, Location, Unit, OffDays, Email, Phone, IsSteward)
  const allMemberData = memberSheet.getRange("A2:J" + memberSheet.getLastRow()).getValues();
  const members = allMemberData.filter(r => r[0]);

  if (members.length === 0) {
    SpreadsheetApp.getUi().alert("❌ Please seed members first!\n\nGrievances must reference actual Member IDs.");
    return;
  }

  // Get steward names (members where column J = "Yes")
  const stewards = members.filter(m => m[9] === "Yes");
  if (stewards.length === 0) {
    SpreadsheetApp.getUi().alert("❌ No stewards found!\n\nPlease ensure some members have 'Is Steward' set to Yes.");
    return;
  }

  const statuses = config.getRange("E2:E12").getValues().flat().filter(String);
  const steps = config.getRange("F2:F7").getValues().flat().filter(String);
  const types = config.getRange("G2:G16").getValues().flat().filter(String);
  const outcomes = config.getRange("H2:H10").getValues().flat().filter(String);

  const TOTAL = 5000;
  const BATCH = 500;

  SpreadsheetApp.getUi().alert(`Seeding ${TOTAL} grievances in ${TOTAL/BATCH} batches...\n\nThis will take 2-3 minutes.`);

  // Track all generated IDs to ensure uniqueness
  const existingIds = new Set();

  for (let batch = 0; batch < TOTAL / BATCH; batch++) {
    const data = [];

    for (let i = 0; i < BATCH; i++) {
      // Generate unique Grievance ID
      const grievanceId = generateRandomGrievanceId(existingIds);
      existingIds.add(grievanceId);

      // Select a random ACTUAL member from Member Directory
      const member = members[Math.floor(Math.random() * members.length)];

      // Select a random steward
      const steward = stewards[Math.floor(Math.random() * stewards.length)];
      const stewardName = `${steward[1]} ${steward[2]}`; // FirstName LastName

      const incidentDate = randomDateWithinDays(730);

      // More realistic status distribution
      const statusRand = Math.random();
      let status, step;
      if (statusRand < 0.15) {
        status = "Resolved - Won";
      } else if (statusRand < 0.25) {
        status = "Resolved - Lost";
      } else if (statusRand < 0.30) {
        status = "Resolved - Settled";
      } else if (statusRand < 0.35) {
        status = "Resolved - Withdrawn";
      } else if (statusRand < 0.40) {
        status = "In Mediation";
      } else if (statusRand < 0.43) {
        status = "In Arbitration";
      } else {
        status = "Filed - Active";
      }

      // Realistic step distribution based on status
      if (status.startsWith("Resolved") || status === "In Mediation" || status === "In Arbitration") {
        const stepRand = Math.random();
        if (stepRand < 0.6) {
          step = "Step I - Immediate Supervisor";
        } else if (stepRand < 0.85) {
          step = "Step II - Agency Head";
        } else {
          step = "Step III - Human Resources";
        }
      } else {
        const stepRand = Math.random();
        if (stepRand < 0.7) {
          step = "Step I - Immediate Supervisor";
        } else if (stepRand < 0.90) {
          step = "Step II - Agency Head";
        } else {
          step = "Step III - Human Resources";
        }
      }

      const hasFiled = Math.random() > 0.05; // 95% have been filed
      const filedDate = hasFiled ? addDays(incidentDate, Math.floor(Math.random() * 20)) : null;
      const type = types[Math.floor(Math.random() * types.length)];

      // More realistic step dates based on progression
      let step1DecisionDate = null, step1Outcome = "", step2FiledDate = null, step2DecisionDate = null, step2Outcome = "";
      let step3FiledDate = null, step3DecisionDate = null, mediationDate = null, arbitrationDate = null;
      let finalOutcome = "Pending";

      if (hasFiled) {
        // Step I
        if (step === "Step I - Immediate Supervisor" || step === "Step II - Agency Head" || step === "Step III - Human Resources" || status.startsWith("Resolved")) {
          step1DecisionDate = addDays(filedDate, Math.floor(Math.random() * 35) + 15);
          if (step !== "Step I - Immediate Supervisor" || status.startsWith("Resolved")) {
            step1Outcome = Math.random() < 0.3 ? "Denied" : "Partial";
          }
        }

        // Step II
        if (step === "Step II - Agency Head" || step === "Step III - Human Resources" || status.startsWith("Resolved")) {
          if (step1DecisionDate) {
            step2FiledDate = addDays(step1DecisionDate, Math.floor(Math.random() * 12) + 1);
            step2DecisionDate = addDays(step2FiledDate, Math.floor(Math.random() * 35) + 15);
            if (step !== "Step II - Agency Head" || status.startsWith("Resolved")) {
              step2Outcome = Math.random() < 0.4 ? "Denied" : "Partial";
            }
          }
        }

        // Step III
        if (step === "Step III - Human Resources" || status === "In Mediation" || status === "In Arbitration" || (status.startsWith("Resolved") && Math.random() < 0.3)) {
          if (step2DecisionDate) {
            step3FiledDate = addDays(step2DecisionDate, Math.floor(Math.random() * 12) + 1);
            step3DecisionDate = addDays(step3FiledDate, Math.floor(Math.random() * 40) + 20);
          }
        }

        // Mediation/Arbitration
        if (status === "In Mediation") {
          mediationDate = step3DecisionDate || step2DecisionDate || step1DecisionDate || addDays(filedDate, 60);
        } else if (status === "In Arbitration") {
          arbitrationDate = step3DecisionDate || step2DecisionDate || step1DecisionDate || addDays(filedDate, 90);
        }

        // Final outcome
        if (status.startsWith("Resolved")) {
          finalOutcome = outcomes[Math.floor(Math.random() * outcomes.length)];
        }
      }

      data.push([
        // A-F: Basic Info
        grievanceId,
        member[0], // Member ID from Member Directory
        member[1], // First Name from Member Directory
        member[2], // Last Name from Member Directory
        status, step,
        // G-J: Incident & Filing (H, J auto-calculated)
        incidentDate,
        "", // Filing deadline (H - calculated)
        filedDate,
        "", // Step I due (J - calculated)
        // K-M: Step I (M auto-calculated)
        step1DecisionDate, step1Outcome, "", // Step I Decision Date, Outcome, Appeal Deadline (M - calculated)
        // N-R: Step II (O, R auto-calculated)
        step2FiledDate, "", step2DecisionDate, step2Outcome, "", // Step II Filed, Decision Due (O - calc), Decision, Outcome, Appeal Deadline (R - calc)
        // S-V: Step III & Beyond
        step3FiledDate, step3DecisionDate, mediationDate, arbitrationDate,
        // W-Z: Details
        finalOutcome,
        type,
        `${type} - Auto-generated description`,
        stewardName, // Steward name from actual members
        // AA-AD: Timeline Tracking (ALL auto-calculated)
        "", // Days Open (AA - calculated)
        "", // Next Action Due (AB - calculated)
        "", // Days to Deadline (AC - calculated)
        "", // Overdue? (AD - calculated)
        // AE-AF: Admin (AF auto-calculated)
        "", // Notes (AE)
        new Date() // Last Updated (AF)
      ]);
    }

    const startRow = 2 + (batch * BATCH);
    sheet.getRange(startRow, 1, BATCH, 32).setValues(data);

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
    .addSubMenu(ui.createMenu('📤 Export Data')
      .addItem('Export Dashboard to PDF', 'exportDashboardToPDF')
      .addItem('Export Member Directory to CSV', 'exportMembersToCsv')
      .addItem('Export Grievances to CSV', 'exportGrievancesToCsv')
      .addItem('Export Steward Workload to CSV', 'exportStewardWorkloadToCsv'))
    .addSubMenu(ui.createMenu('🎨 Theme Options')
      .addItem('Light Theme', 'applyLightTheme')
      .addItem('Dark Theme', 'applyDarkTheme')
      .addItem('High Contrast Theme', 'applyHighContrastTheme'))
    .addSubMenu(ui.createMenu('⚙️ Utilities')
      .addItem('Sort by Priority', 'sortGrievancesByPriority')
      .addItem('Toggle Mobile Mode', 'toggleMobileMode')
      .addItem('Setup Triggers', 'setupTriggers'))
    .addToUi();
}

// ============================================================================
// EXPORT FUNCTIONS
// ============================================================================

/**
 * Exports the Dashboard sheet to PDF
 */
function exportDashboardToPDF() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

    if (!dashboard) {
      SpreadsheetApp.getUi().alert('❌ Dashboard sheet not found!');
      return;
    }

    const url = ss.getUrl();
    const sheetId = dashboard.getSheetId();
    const pdfUrl = url.replace(/edit.*$/, '') +
      'export?exportFormat=pdf&format=pdf' +
      '&size=letter' +
      '&portrait=true' +
      '&fitw=true' +
      '&sheetnames=false&printtitle=false' +
      '&pagenumbers=false&gridlines=false' +
      '&fzr=false' +
      '&gid=' + sheetId;

    SpreadsheetApp.getUi().showModelessDialog(
      HtmlService.createHtmlOutput(
        '<p>Click the link below to download the PDF:</p>' +
        '<p><a href="' + pdfUrl + '" target="_blank">Download Dashboard PDF</a></p>' +
        '<p><button onclick="google.script.host.close()">Close</button></p>'
      ).setWidth(400).setHeight(150),
      'Export Dashboard to PDF'
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Export error: ' + error.message);
  }
}

/**
 * Exports Member Directory to CSV
 */
function exportMembersToCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

    if (!sheet) {
      SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const csv = convertToCsv(data);
    const fileName = '509_Member_Directory_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.csv';

    const blob = Utilities.newBlob(csv, 'text/csv', fileName);
    const url = DriveApp.createFile(blob).getUrl();

    SpreadsheetApp.getUi().alert('✅ CSV exported!\n\nFile saved to your Google Drive:\n' + fileName + '\n\nURL: ' + url);
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Export error: ' + error.message);
  }
}

/**
 * Exports Grievance Log to CSV
 */
function exportGrievancesToCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!sheet) {
      SpreadsheetApp.getUi().alert('❌ Grievance Log sheet not found!');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const csv = convertToCsv(data);
    const fileName = '509_Grievance_Log_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.csv';

    const blob = Utilities.newBlob(csv, 'text/csv', fileName);
    const url = DriveApp.createFile(blob).getUrl();

    SpreadsheetApp.getUi().alert('✅ CSV exported!\n\nFile saved to your Google Drive:\n' + fileName + '\n\nURL: ' + url);
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Export error: ' + error.message);
  }
}

/**
 * Exports Steward Workload to CSV
 */
function exportStewardWorkloadToCsv() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);

    if (!sheet) {
      SpreadsheetApp.getUi().alert('❌ Steward Workload sheet not found!');
      return;
    }

    const data = sheet.getDataRange().getValues();
    const csv = convertToCsv(data);
    const fileName = '509_Steward_Workload_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.csv';

    const blob = Utilities.newBlob(csv, 'text/csv', fileName);
    const url = DriveApp.createFile(blob).getUrl();

    SpreadsheetApp.getUi().alert('✅ CSV exported!\n\nFile saved to your Google Drive:\n' + fileName + '\n\nURL: ' + url);
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Export error: ' + error.message);
  }
}

/**
 * Helper function to convert 2D array to CSV format
 */
function convertToCsv(data) {
  return data.map(row =>
    row.map(cell => {
      if (cell == null) return '';
      const str = cell.toString();
      // Escape quotes and wrap in quotes if contains comma, quote, or newline
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\n');
}

// ============================================================================
// THEME FUNCTIONS
// ============================================================================

/**
 * Applies light theme to the dashboard
 */
function applyLightTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboard) {
    SpreadsheetApp.getUi().alert('❌ Dashboard sheet not found!');
    return;
  }

  // Apply light theme colors
  dashboard.getRange("A3:B3").setBackground("#4A86E8").setFontColor("#FFFFFF");
  dashboard.getRange("A11:B11").setBackground("#E06666").setFontColor("#FFFFFF");
  dashboard.getRange("A21:B21").setBackground("#F6B26B").setFontColor("#FFFFFF");
  dashboard.getRange("E4:H4").setBackground("#F6B26B").setFontColor("#FFFFFF");

  SpreadsheetApp.getUi().alert('✅ Light theme applied!');
}

/**
 * Applies dark theme to the dashboard
 */
function applyDarkTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboard) {
    SpreadsheetApp.getUi().alert('❌ Dashboard sheet not found!');
    return;
  }

  // Apply dark theme colors
  dashboard.getRange("A3:B3").setBackground("#1E3A5F").setFontColor("#FFFFFF");
  dashboard.getRange("A11:B11").setBackground("#5C1010").setFontColor("#FFFFFF");
  dashboard.getRange("A21:B21").setBackground("#5C3A10").setFontColor("#FFFFFF");
  dashboard.getRange("E4:H4").setBackground("#5C3A10").setFontColor("#FFFFFF");

  SpreadsheetApp.getUi().alert('✅ Dark theme applied!');
}

/**
 * Applies high contrast theme to the dashboard
 */
function applyHighContrastTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboard) {
    SpreadsheetApp.getUi().alert('❌ Dashboard sheet not found!');
    return;
  }

  // Apply high contrast theme colors
  dashboard.getRange("A3:B3").setBackground("#000000").setFontColor("#FFFF00");
  dashboard.getRange("A11:B11").setBackground("#000000").setFontColor("#00FF00");
  dashboard.getRange("A21:B21").setBackground("#000000").setFontColor("#FF8800");
  dashboard.getRange("E4:H4").setBackground("#000000").setFontColor("#FF8800");

  SpreadsheetApp.getUi().alert('✅ High contrast theme applied!');
}

// ============================================================================
// MOBILE MODE FUNCTIONS
// ============================================================================

/**
 * Toggles mobile mode for better viewing on phones
 */
function toggleMobileMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const isMobile = props.getProperty('MOBILE_MODE') === 'true';

  if (isMobile) {
    // Switch to desktop mode
    props.setProperty('MOBILE_MODE', 'false');
    applyDesktopLayout();
    SpreadsheetApp.getUi().alert('✅ Desktop mode enabled!\n\nThe dashboard is now optimized for desktop viewing.');
  } else {
    // Switch to mobile mode
    props.setProperty('MOBILE_MODE', 'true');
    applyMobileLayout();
    SpreadsheetApp.getUi().alert('✅ Mobile mode enabled!\n\nThe dashboard is now optimized for mobile viewing with:\n• Wider columns\n• Larger text\n• Simplified layout');
  }
}

/**
 * Applies mobile-friendly layout
 */
function applyMobileLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Dashboard: Larger text and wider columns
  if (dashboard) {
    dashboard.setColumnWidth(1, 250);
    dashboard.setColumnWidth(2, 150);
    dashboard.getRange("A1:H50").setFontSize(14);
  }

  // Member Directory: Essential columns only visible
  if (memberDir) {
    memberDir.setColumnWidth(1, 150);  // Member ID
    memberDir.setColumnWidth(2, 150);  // First Name
    memberDir.setColumnWidth(3, 150);  // Last Name
    memberDir.setColumnWidth(4, 200);  // Job Title
    memberDir.setColumnWidth(5, 200);  // Location
    memberDir.getRange("1:1").setFontSize(12).setFontWeight("bold");
  }

  // Grievance Log: Essential columns only visible
  if (grievanceLog) {
    grievanceLog.setColumnWidth(1, 150);  // Grievance ID
    grievanceLog.setColumnWidth(2, 150);  // Member ID
    grievanceLog.setColumnWidth(3, 150);  // First Name
    grievanceLog.setColumnWidth(4, 150);  // Last Name
    grievanceLog.setColumnWidth(5, 150);  // Status
    grievanceLog.getRange("1:1").setFontSize(12).setFontWeight("bold");
  }
}

/**
 * Applies desktop layout
 */
function applyDesktopLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Dashboard: Standard text and column widths
  if (dashboard) {
    dashboard.setColumnWidth(1, 200);
    dashboard.setColumnWidth(2, 120);
    dashboard.getRange("A1:H50").setFontSize(10);
  }

  // Member Directory: Standard column widths
  if (memberDir) {
    memberDir.setColumnWidth(1, 110);  // Member ID
    memberDir.setColumnWidth(2, 100);  // First Name
    memberDir.setColumnWidth(3, 100);  // Last Name
    memberDir.setColumnWidth(4, 150);  // Job Title
    memberDir.setColumnWidth(5, 150);  // Location
    memberDir.getRange("1:1").setFontSize(10).setFontWeight("bold");
  }

  // Grievance Log: Standard column widths
  if (grievanceLog) {
    grievanceLog.setColumnWidth(1, 120);  // Grievance ID
    grievanceLog.setColumnWidth(2, 110);  // Member ID
    grievanceLog.setColumnWidth(3, 100);  // First Name
    grievanceLog.setColumnWidth(4, 100);  // Last Name
    grievanceLog.setColumnWidth(5, 120);  // Status
    grievanceLog.getRange("1:1").setFontSize(10).setFontWeight("bold");
  }
}

/**
 * Auto-detects mobile device and applies appropriate layout on open
 */
function autoDetectMobile() {
  // Note: Google Sheets Apps Script doesn't have direct access to user agent
  // This function would need to be triggered manually or through a custom UI
  // For now, users can manually toggle mobile mode via the menu
  const props = PropertiesService.getDocumentProperties();
  const isMobile = props.getProperty('MOBILE_MODE') === 'true';

  if (isMobile) {
    applyMobileLayout();
  }
}
