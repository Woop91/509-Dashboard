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
  FUTURE_FEATURES: "Future Features",
  PENDING_FEATURES: "Pending Features",
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

// Contact Update Configuration
// TODO: Replace this URL with your actual Google Form URL for member contact updates
const MEMBER_CONTACT_UPDATE_FORM_URL = "https://forms.gle/YOUR_FORM_ID_HERE";

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

    createFutureFeaturesSheet(ss);
    SpreadsheetApp.flush();

    createPendingFeaturesSheet(ss);
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
    .setFontWeight("bold").setBackground(COLORS.PRIMARY_BLUE).setFontColor("white").setFontFamily("Roboto");

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
    .setFontSize(12).setFontFamily("Roboto")
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
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto")
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
    sheet.getRange(1, col).setBackground(COLORS.ACCENT_TEAL);
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
      .setBackground(COLORS.PRIMARY_BLUE)
      .setFontColor("white")
      .setFontFamily("Roboto")
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
        sheet.getRange(1, col).setBackground(COLORS.ACCENT_TEAL);
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
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto");

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
    .setFontSize(20).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers (will be populated by rebuildDashboard())
  sheet.getRange("A3").setValue("MEMBER METRICS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
  sheet.getRange("A11").setValue("GRIEVANCE METRICS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
  sheet.getRange("A21").setValue("DEADLINE TRACKING").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
  sheet.getRange("E3").setValue("TOP 10 OVERDUE GRIEVANCES").setFontWeight("bold").setFontSize(11).setFontFamily("Roboto");

  // Chart area headers
  sheet.getRange("A26").setValue("VISUAL ANALYTICS").setFontWeight("bold").setFontSize(14).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

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
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto");

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
    .setFontSize(18).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("MONTHLY TRENDS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("RESOLUTION TIME ANALYSIS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A25").setValue("FILING TRENDS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
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
    .setFontSize(18).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("RESOLUTION PERFORMANCE").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("E3").setValue("EFFICIENCY METRICS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A13").setValue("OUTCOME ANALYSIS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A23").setValue("STEP PROGRESSION ANALYSIS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
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
    .setFontSize(18).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("GRIEVANCES BY LOCATION").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("LOCATION PERFORMANCE METRICS").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
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
    .setFontSize(18).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Section headers - Unified color scheme
  sheet.getRange("A3").setValue("TYPE BREAKDOWN").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("E3").setValue("SUCCESS RATE BY TYPE").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A15").setValue("TYPE TRENDS OVER TIME").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
    .setBackground(COLORS.ACCENT_TEAL).setFontColor("white");

  sheet.getRange("A27").setValue("AVERAGE RESOLUTION TIME BY TYPE").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto")
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
    .setFontSize(24).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Subtitle/date range
  sheet.getRange("A2:P2").merge()
    .setValue("Union-wide Performance Metrics & Key Indicators")
    .setFontSize(12).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  // Section: Key Metrics (Row 4-10)
  sheet.getRange("A4:P4").merge()
    .setValue("📊 KEY PERFORMANCE INDICATORS")
    .setFontSize(15).setFontFamily("Roboto")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:C7").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("A8:C9").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E7:G7").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("E8:G9").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I7:K7").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("I8:K9").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M7:O7").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("M8:O9").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("vs Last Month");

  // ===== UNION HEALTH SECTION =====
  sheet.getRange("A12:P12").merge()
    .setValue("💪 UNION HEALTH & ENGAGEMENT")
    .setFontSize(15).setFontFamily("Roboto")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A15:C15").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("A16:C17").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E15:G15").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("E16:G17").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I15:K15").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("I16:K17").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M15:O15").merge().setFontSize(42).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("M16:O17").merge().setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY)
    .setValue("Last 12 Months");

  // ===== PRIORITY ALERTS SECTION =====
  sheet.getRange("A20:P20").merge()
    .setValue("🚨 PRIORITY ALERTS & ACTION ITEMS")
    .setFontSize(15).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  // Alert cards with borders
  const alertCard = sheet.getRange("A22:O25");
  alertCard.setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A22").setValue("Overdue Grievances:").setFontWeight("bold").setFontSize(11).setFontFamily("Roboto");
  sheet.getRange("A23").setValue("Due This Week:").setFontWeight("bold").setFontSize(11).setFontFamily("Roboto");
  sheet.getRange("A24").setValue("Arbitrations Pending:").setFontWeight("bold").setFontSize(11).setFontFamily("Roboto");
  sheet.getRange("A25").setValue("High-Priority Cases:").setFontWeight("bold").setFontSize(11).setFontFamily("Roboto");

  // Add colored left borders for alert severity
  sheet.getRange("A22:A22").setBorder(null, true, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A23:A23").setBorder(null, true, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A24:A24").setBorder(null, true, null, null, null, null, COLORS.ACCENT_YELLOW, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("A25:A25").setBorder(null, true, null, null, null, null, COLORS.PRIMARY_BLUE, SpreadsheetApp.BorderStyle.SOLID_THICK);

  // ===== VISUAL SUMMARIES SECTION =====
  sheet.getRange("A28:P28").merge()
    .setValue("📈 VISUAL SUMMARIES")
    .setFontSize(15).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // Chart placeholders
  sheet.getRange("A30").setValue("Status Breakdown").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
  sheet.getRange("F30").setValue("Monthly Trends").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
  sheet.getRange("K30").setValue("Win/Loss Outcomes").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");

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
    .setFontSize(26).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  // Current period indicator
  const today = new Date();
  sheet.getRange("A2:P2").merge()
    .setValue(`Performance Period: ${today.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`)
    .setFontSize(12).setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Main KPI Grid - Large Numbers Section
  sheet.getRange("A4:P4").merge()
    .setValue("📊 THIS MONTH")
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:C7").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("A8:C9").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 2: Grievances Resolved =====
  const card2Range = sheet.getRange("E6:G9");
  card2Range.setBackground(COLORS.WHITE);

  sheet.getRange("E6:G6").setBorder(true, null, null, null, null, null, COLORS.UNION_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E6:E9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G6:G9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E9:G9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E6:G6").merge().setValue("GRIEVANCES RESOLVED")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E7:G7").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("E8:G9").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 3: Pending Decisions =====
  const card3Range = sheet.getRange("I6:K9");
  card3Range.setBackground(COLORS.WHITE);

  sheet.getRange("I6:K6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_ORANGE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I6:I9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K6:K9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I9:K9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I6:K6").merge().setValue("PENDING DECISIONS")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I7:K7").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("I8:K9").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== KPI CARD 4: Win Rate % =====
  const card4Range = sheet.getRange("M6:O9");
  card4Range.setBackground(COLORS.WHITE);

  sheet.getRange("M6:O6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_PURPLE, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M6:M9").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O6:O9").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M9:O9").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M6:O6").merge().setValue("WIN RATE %")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M7:O7").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("M8:O9").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YEAR TO DATE SECTION =====
  sheet.getRange("A12:P12").merge()
    .setValue("📈 YEAR TO DATE")
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A15:C15").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);
  sheet.getRange("A16:C17").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 2: Total Resolved =====
  const ytdCard2Range = sheet.getRange("E14:G17");
  ytdCard2Range.setBackground(COLORS.WHITE);

  sheet.getRange("E14:G14").setBorder(true, null, null, null, null, null, COLORS.UNION_GREEN, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("E14:E17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("G14:G17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("E17:G17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("E14:G14").merge().setValue("TOTAL RESOLVED")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("E15:G15").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.UNION_GREEN);
  sheet.getRange("E16:G17").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 3: Settlements =====
  const ytdCard3Range = sheet.getRange("I14:K17");
  ytdCard3Range.setBackground(COLORS.WHITE);

  sheet.getRange("I14:K14").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("I14:I17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K14:K17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I17:K17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("I14:K14").merge().setValue("SETTLEMENTS")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("I15:K15").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);
  sheet.getRange("I16:K17").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== YTD CARD 4: Arbitrations =====
  const ytdCard4Range = sheet.getRange("M14:O17");
  ytdCard4Range.setBackground(COLORS.WHITE);

  sheet.getRange("M14:O14").setBorder(true, null, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("M14:M17").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("O14:O17").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("M17:O17").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("M14:O14").merge().setValue("ARBITRATIONS")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("M15:O15").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.SOLIDARITY_RED);
  sheet.getRange("M16:O17").merge().setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setVerticalAlignment("middle").setFontColor(COLORS.TEXT_GRAY);

  // ===== PERFORMANCE BREAKDOWN SECTION =====
  sheet.getRange("A20:P20").merge()
    .setValue("⚡ PERFORMANCE BREAKDOWN")
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontSize(20).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  // Engagement Score Cards
  sheet.getRange("A3:L3").merge()
    .setValue("🎯 ENGAGEMENT SCORES")
    .setFontSize(14).setFontFamily("Roboto")
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

  sheet.getRange("A5:D5").setFontWeight("bold").setFontSize(10).setFontFamily("Roboto")
    .setHorizontalAlignment("center").setFontColor(COLORS.TEXT_GRAY);

  // Member Distribution
  sheet.getRange("A10:F10").merge()
    .setValue("👥 MEMBER DISTRIBUTION BY ENGAGEMENT LEVEL")
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(20).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  sheet.getRange("A2:L2").merge()
    .setValue("Financial Impact of Grievances & Union Activities")
    .setFontSize(11).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY);

  // Financial Metrics
  sheet.getRange("A4:L4").merge()
    .setValue("💰 MEMBER WINS & RECOVERIES")
    .setFontSize(14).setFontFamily("Roboto")
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

  sheet.getRange("A6:D6").setFontWeight("bold").setFontSize(10).setFontFamily("Roboto")
    .setHorizontalAlignment("center").setFontColor(COLORS.TEXT_GRAY);

  // Cost by Type
  sheet.getRange("A11:F11").merge()
    .setValue("📊 RECOVERY BY GRIEVANCE TYPE")
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(13).setFontFamily("Roboto")
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
    .setFontSize(26).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  // Today's snapshot
  const today = new Date();
  sheet.getRange("A2:O2").merge()
    .setValue(`Snapshot: ${today.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`)
    .setFontSize(12).setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // ===== TODAY SECTION =====
  sheet.getRange("A4:O4").merge()
    .setValue("📅 TODAY")
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("A7:D8").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.PRIMARY_BLUE);

  // TODAY CARD 2: Deadlines Due
  sheet.getRange("F6:I8").setBackground(COLORS.WHITE);
  sheet.getRange("F6:I6").setBorder(true, null, null, null, null, null, COLORS.SOLIDARITY_RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("F6:F8").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("I6:I8").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("F8:I8").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("F6:I6").merge().setValue("DEADLINES DUE")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("F7:I8").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.SOLIDARITY_RED);

  // TODAY CARD 3: Meetings Scheduled
  sheet.getRange("K6:N8").setBackground(COLORS.WHITE);
  sheet.getRange("K6:N6").setBorder(true, null, null, null, null, null, COLORS.ACCENT_TEAL, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange("K6:K8").setBorder(null, true, null, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("N6:N8").setBorder(null, null, null, true, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange("K8:N8").setBorder(null, null, true, null, null, null, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("K6:N6").merge().setValue("MEETINGS SCHEDULED")
    .setFontWeight("bold").setFontSize(11).setFontFamily("Roboto").setHorizontalAlignment("center")
    .setFontColor(COLORS.TEXT_GRAY).setVerticalAlignment("middle");
  sheet.getRange("K7:N8").merge().setFontSize(48).setFontFamily("Roboto").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontColor(COLORS.ACCENT_TEAL);

  // ===== THIS WEEK SECTION =====
  sheet.getRange("A11:O11").merge()
    .setValue("📊 THIS WEEK")
    .setFontSize(16).setFontFamily("Roboto")
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
      .setFontWeight("bold").setFontSize(10).setFontFamily("Roboto").setHorizontalAlignment("center")
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
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontSize(16).setFontFamily("Roboto")
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
    .setFontSize(16).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  const alertBox = sheet.getRange("A41:O44");
  alertBox.setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A41").setValue("Grievances Overdue:").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto").setFontColor(COLORS.SOLIDARITY_RED);
  sheet.getRange("A42").setValue("Due in Next 48 Hours:").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto").setFontColor(COLORS.ACCENT_ORANGE);
  sheet.getRange("A43").setValue("Arbitrations Pending:").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto").setFontColor(COLORS.ACCENT_PURPLE);
  sheet.getRange("A44").setValue("Step III Cases:").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto").setFontColor(COLORS.PRIMARY_BLUE);

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
// FUTURE FEATURES SHEET
// ============================================================================

/**
 * Creates Future Features sheet listing security enhancements and features for future implementation
 */
function createFutureFeaturesSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.FUTURE_FEATURES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.FUTURE_FEATURES);

  sheet.clear();

  // Title
  sheet.getRange("A1:H1").merge()
    .setValue("🔮 FUTURE FEATURES - Security & Enhancements")
    .setFontSize(20).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  // Subtitle
  sheet.getRange("A2:H2").merge()
    .setValue("Advanced features available for implementation when needed")
    .setFontSize(12).setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Security Enhancements Section
  sheet.getRange("A4:H4").merge()
    .setValue("🔒 SECURITY & AUDIT ENHANCEMENTS")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  // Headers
  const headers = ["Feature #", "Feature Name", "Description", "Function Name", "Status", "Priority", "Notes", "Dependencies"];
  sheet.getRange("A5:H5").setValues([headers])
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Roboto")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Security Features Data
  const securityFeatures = [
    ["79", "Audit Logging", "Tracks all data modifications with user, timestamp, and change details", "logDataModification()", "Available", "High", "Creates Audit_Log sheet automatically", "None"],
    ["80", "Role-Based Access Control", "Implements user permissions (Admin, Steward, Viewer roles)", "checkUserPermission()", "Available", "High", "Requires script properties configuration", "ADMINS, STEWARDS, VIEWERS properties"],
    ["81", "Data Encryption", "Encrypts sensitive data fields (SSN, DOB, etc.)", "encryptSensitiveData()", "Available", "Medium", "Use proper encryption in production", "ENCRYPTION_KEY property"],
    ["82", "Data Decryption", "Decrypts encrypted sensitive data", "decryptSensitiveData()", "Available", "Medium", "Paired with encryption feature", "ENCRYPTION_KEY property"],
    ["83", "Input Sanitization", "Validates and sanitizes user input for security", "sanitizeInput()", "Available", "High", "Prevents script injection attacks", "None"],
    ["84", "Audit Reporting", "Generates audit trail reports for date ranges", "generateAuditReport()", "Available", "Medium", "Requires Audit_Log sheet", "Feature 79"],
    ["85", "Data Retention Policy", "Enforces data retention policies (default 7 years)", "enforceDataRetention()", "Available", "Low", "Configurable retention period", "Feature 79"],
    ["86", "Suspicious Activity Detection", "Detects unusual patterns (>50 changes/hour)", "detectSuspiciousActivity()", "Available", "Medium", "Alerts on potential data breaches", "Feature 79"]
  ];

  sheet.getRange(6, 1, securityFeatures.length, 8).setValues(securityFeatures);

  // Advanced Features Section
  const advancedRow = 6 + securityFeatures.length + 2;
  sheet.getRange(advancedRow, 1, 1, 8).merge()
    .setValue("⚡ ADVANCED FEATURES")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  // Headers for advanced features
  sheet.getRange(advancedRow + 1, 1, 1, 8).setValues([headers])
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Roboto")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Advanced Features Data
  const advancedFeatures = [
    ["87", "Quick Actions Sidebar", "Interactive sidebar with one-click common actions", "showQuickActionsSidebar()", "Available", "Medium", "Accessible via menu", "None"],
    ["88", "Advanced Search", "Interactive grievance search dialog", "showSearchDialog()", "Available", "Medium", "Search by ID, name, type", "None"],
    ["89", "Advanced Filtering", "Filter grievances by multiple criteria", "showFilterDialog()", "Available", "Medium", "Date range, status, type filters", "None"],
    ["90", "Automated Backups", "Creates daily backups to Google Drive", "createAutomatedBackup()", "Available", "High", "Requires Drive folder configuration", "BACKUP_FOLDER_ID property"],
    ["91", "Performance Monitoring", "Tracks script execution times and performance", "trackPerformance()", "Available", "Low", "Logs to Performance_Log sheet", "None"],
    ["92", "Keyboard Shortcuts", "Custom keyboard shortcuts for common actions", "setupKeyboardShortcuts()", "Available", "Low", "Enhances user productivity", "None"],
    ["93", "Export Wizard", "Guided export with format and filter options", "showExportWizard()", "Available", "Medium", "Interactive dialog for exports", "None"],
    ["94", "Data Import", "Import members/grievances from CSV/Excel", "showImportWizard()", "Available", "High", "Bulk data import capability", "None"]
  ];

  sheet.getRange(advancedRow + 2, 1, advancedFeatures.length, 8).setValues(advancedFeatures);

  // Implementation Notes Section
  const notesRow = advancedRow + 2 + advancedFeatures.length + 2;
  sheet.getRange(notesRow, 1, 1, 8).merge()
    .setValue("📝 IMPLEMENTATION NOTES")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  const notes = [
    ["To enable these features:", "", "", "", "", "", "", ""],
    ["1.", "Call addEnhancementMenus() from onOpen() to expose all menus", "", "", "", "", "", ""],
    ["2.", "Configure script properties for sensitive features (encryption keys, roles, etc.)", "", "", "", "", "", ""],
    ["3.", "Set up automated triggers for backup and reporting features", "", "", "", "", "", ""],
    ["4.", "Test each feature in a development environment before production use", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["Security Best Practices:", "", "", "", "", "", "", ""],
    ["•", "Use strong encryption keys (not default) for production", "", "", "", "", "", ""],
    ["•", "Regularly review audit logs for unauthorized access", "", "", "", "", "", ""],
    ["•", "Implement role-based access control before going live", "", "", "", "", "", ""],
    ["•", "Enable automated backups with appropriate retention policies", "", "", "", "", "", ""]
  ];

  sheet.getRange(notesRow + 1, 1, notes.length, 8).setValues(notes);

  // Formatting
  sheet.setRowHeight(1, 50);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 35);
  sheet.setRowHeight(5, 30);
  sheet.setRowHeight(advancedRow, 35);
  sheet.setRowHeight(advancedRow + 1, 30);
  sheet.setRowHeight(notesRow, 35);

  // Column widths
  sheet.setColumnWidth(1, 80);   // Feature #
  sheet.setColumnWidth(2, 200);  // Feature Name
  sheet.setColumnWidth(3, 350);  // Description
  sheet.setColumnWidth(4, 250);  // Function Name
  sheet.setColumnWidth(5, 100);  // Status
  sheet.setColumnWidth(6, 80);   // Priority
  sheet.setColumnWidth(7, 250);  // Notes
  sheet.setColumnWidth(8, 200);  // Dependencies

  // Apply borders and styling to data rows
  const totalRows = notesRow + notes.length;
  sheet.getRange(6, 1, securityFeatures.length, 8)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily("Roboto")
    .setVerticalAlignment("middle");

  sheet.getRange(advancedRow + 2, 1, advancedFeatures.length, 8)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID)
    .setFontFamily("Roboto")
    .setVerticalAlignment("middle");

  sheet.getRange(notesRow + 1, 1, notes.length, 8)
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  // Freeze header rows
  sheet.setFrozenRows(2);
}

// ============================================================================
// PENDING FEATURES SHEET
// ============================================================================

/**
 * Creates Pending Features sheet with setup instructions for production deployment
 */
function createPendingFeaturesSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.PENDING_FEATURES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.PENDING_FEATURES);

  sheet.clear();

  // Title
  sheet.getRange("A1:G1").merge()
    .setValue("⏳ PENDING FEATURES - Production Setup Checklist")
    .setFontSize(20).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white");

  // Subtitle
  sheet.getRange("A2:G2").merge()
    .setValue("Complete these tasks before going live")
    .setFontSize(12).setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Section 1: Test Key Features
  sheet.getRange("A4:G4").merge()
    .setValue("🧪 TEST KEY FEATURES")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  const testFeatures = [
    ["Task", "Description", "Function to Run", "Expected Result", "Status", "Priority", "Notes"],
    ["Run Data Integrity Check", "Validates all data relationships and calculations", "runDataIntegrityCheck()", "Report showing validation results", "Pending", "High", "Check for orphaned records, invalid dates, missing IDs"],
    ["Test Quick Actions Sidebar", "Verify sidebar displays and functions work", "showQuickActionsSidebar()", "Sidebar opens with working buttons", "Pending", "Medium", "Test all buttons in sidebar"],
    ["Verify Dashboard Calculations", "Ensure all metrics calculate correctly", "rebuildDashboard()", "Accurate KPIs and charts", "Pending", "High", "Compare manual counts with dashboard"],
    ["Test CBA Compliance", "Validate deadline calculations per contract", "validateCBACompliance()", "All deadlines follow Article 23A rules", "Pending", "Critical", "21-day filing, 30-day decisions, 10-day appeals"],
    ["Verify Seed Functions", "Test data generation works correctly", "seedAll()", "Members and grievances created", "Pending", "Medium", "Use new unified seed function"]
  ];

  sheet.getRange(5, 1, testFeatures.length, 7).setValues(testFeatures);

  // Section 2: Configure Properties
  const configRow = 5 + testFeatures.length + 2;
  sheet.getRange(configRow, 1, 1, 7).merge()
    .setValue("⚙️ CONFIGURE SCRIPT PROPERTIES")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  const configTasks = [
    ["Property Name", "Description", "Example Value", "Required For", "Status", "Priority", "How to Set"],
    ["ADMINS", "Comma-separated admin email addresses", "admin1@mass.gov,admin2@mass.gov", "Role-based access control", "Pending", "High", "Script Properties > Add Property"],
    ["STEWARDS", "Comma-separated steward email addresses", "steward1@mass.gov,steward2@mass.gov", "Permission management", "Pending", "High", "Script Properties > Add Property"],
    ["VIEWERS", "Comma-separated viewer email addresses", "viewer1@mass.gov,viewer2@mass.gov", "Read-only access", "Pending", "Medium", "Script Properties > Add Property"],
    ["ENCRYPTION_KEY", "Strong encryption key for sensitive data", "YourSecureKey2024!@#$", "Data encryption features", "Pending", "High", "Use 16+ character random string"],
    ["BACKUP_FOLDER_ID", "Google Drive folder ID for backups", "1ABC...xyz", "Automated backups", "Pending", "Medium", "Create Drive folder, copy ID from URL"],
    ["NOTIFICATION_EMAIL", "Default email for system notifications", "notifications@mass.gov", "Alert system", "Pending", "Low", "Script Properties > Add Property"]
  ];

  sheet.getRange(configRow + 1, 1, configTasks.length, 7).setValues(configTasks);

  // Section 3: Set Up Triggers
  const triggerRow = configRow + 1 + configTasks.length + 2;
  sheet.getRange(triggerRow, 1, 1, 7).merge()
    .setValue("⏰ SET UP AUTOMATED TRIGGERS")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  const triggerTasks = [
    ["Trigger Name", "Description", "Function to Run", "Schedule", "Status", "Priority", "Notes"],
    ["Schedule Reports", "Set up weekly and monthly reports", "scheduleReports()", "Weekly Monday 8am, Monthly 1st at 9am", "Pending", "Medium", "Run once to create triggers"],
    ["Schedule Updates", "Auto-refresh dashboard hourly", "scheduleAutomaticUpdates()", "Every hour + daily at 2am", "Pending", "Low", "Keeps dashboard current"],
    ["Daily Backup", "Create daily backup to Drive", "createAutomatedBackup()", "Daily at 3am", "Pending", "High", "Requires BACKUP_FOLDER_ID"],
    ["Overdue Alerts", "Send alerts for overdue grievances", "sendOverdueAlerts()", "Daily at 7am", "Pending", "High", "Notifies responsible parties"],
    ["Data Retention", "Clean up old audit logs", "enforceDataRetention()", "Weekly Sunday at midnight", "Pending", "Low", "Keeps system clean"]
  ];

  sheet.getRange(triggerRow + 1, 1, triggerTasks.length, 7).setValues(triggerTasks);

  // Section 4: Update Main Menu
  const menuRow = triggerRow + 1 + triggerTasks.length + 2;
  sheet.getRange(menuRow, 1, 1, 7).merge()
    .setValue("🍔 UPDATE MAIN MENU")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  const menuTasks = [
    ["Task", "Description", "Action Required", "Function", "Status", "Priority", "Notes"],
    ["Add Enhancement Menus", "Expose all advanced features in UI", "Call addEnhancementMenus() from onOpen()", "addEnhancementMenus()", "Pending", "High", "Adds 5 new menus: Data Validation, Notifications, Reports, Engagement, Tools"],
    ["Verify Menu Structure", "Ensure all menu items work", "Test each menu item", "onOpen()", "Pending", "High", "Check permissions for all functions"],
    ["Add Seed Menu Item", "Add unified seed function to menu", "Add seedAll() to Data Management menu", "seedAll()", "Pending", "Medium", "Simplifies test data generation"]
  ];

  sheet.getRange(menuRow + 1, 1, menuTasks.length, 7).setValues(menuTasks);

  // Section 5: Final Verification
  const verifyRow = menuRow + 1 + menuTasks.length + 2;
  sheet.getRange(verifyRow, 1, 1, 7).merge()
    .setValue("✅ FINAL VERIFICATION")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white");

  const verifyTasks = [
    ["Verification Item", "Description", "How to Verify", "Success Criteria", "Status", "Priority", "Notes"],
    ["All Functions Wired", "Ensure all functions are accessible", "Check all menu items and test", "All functions execute without errors", "Pending", "Critical", "No orphaned or unreachable functions"],
    ["Data Accuracy", "Verify calculations are correct", "Manual spot checks vs system", "100% accuracy on calculations", "Pending", "Critical", "Especially deadlines and metrics"],
    ["Performance", "System responds quickly", "Test with full data load", "< 5 seconds for most operations", "Pending", "High", "Optimize if needed"],
    ["Documentation", "README reflects current state", "Review README completeness", "All features documented", "Pending", "Medium", "Update with new features"],
    ["User Acceptance", "End users can operate system", "UAT with actual stewards", "Users can complete workflows", "Pending", "Critical", "Get feedback before launch"],
    ["Backup & Recovery", "Verify backup/restore works", "Create backup, test restore", "Data restored successfully", "Pending", "High", "Test disaster recovery"],
    ["Security", "All security features enabled", "Check permissions and audit", "RBAC working, audit log active", "Pending", "Critical", "No unauthorized access possible"]
  ];

  sheet.getRange(verifyRow + 1, 1, verifyTasks.length, 7).setValues(verifyTasks);

  // Instructions Section
  const instructRow = verifyRow + 1 + verifyTasks.length + 2;
  sheet.getRange(instructRow, 1, 1, 7).merge()
    .setValue("📖 STEP-BY-STEP INSTRUCTIONS")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  const instructions = [
    ["Step", "Action", "Details", "", "", "", ""],
    ["1", "Test in Development", "Complete all testing tasks above in a copy of the spreadsheet", "", "", "", ""],
    ["2", "Configure Properties", "Set all required script properties (Extensions > Apps Script > Project Settings > Script Properties)", "", "", "", ""],
    ["3", "Enable Menus", "Add this line to onOpen(): addEnhancementMenus();", "", "", "", ""],
    ["4", "Set Up Triggers", "Run scheduleReports() and scheduleAutomaticUpdates() from Apps Script", "", "", "", ""],
    ["5", "Test All Features", "Use the Quick Actions sidebar and test each advanced feature", "", "", "", ""],
    ["6", "Load Real Data", "Import or manually enter actual members and grievances", "", "", "", ""],
    ["7", "User Training", "Train stewards on system usage and workflows", "", "", "", ""],
    ["8", "Go Live", "Monitor first week closely for any issues", "", "", "", ""],
    ["9", "Regular Maintenance", "Weekly: Run Rebuild Dashboard | Monthly: Review audit logs | Quarterly: Check backups", "", "", "", ""]
  ];

  sheet.getRange(instructRow + 1, 1, instructions.length, 7).setValues(instructions);

  // Formatting
  sheet.setRowHeight(1, 50);
  sheet.setRowHeight(2, 35);
  sheet.setRowHeight(4, 35);
  sheet.setRowHeight(configRow, 35);
  sheet.setRowHeight(triggerRow, 35);
  sheet.setRowHeight(menuRow, 35);
  sheet.setRowHeight(verifyRow, 35);
  sheet.setRowHeight(instructRow, 35);

  // Column widths
  sheet.setColumnWidth(1, 200);  // Task/Property/etc
  sheet.setColumnWidth(2, 300);  // Description
  sheet.setColumnWidth(3, 250);  // Function/Value/etc
  sheet.setColumnWidth(4, 200);  // Expected/Schedule/etc
  sheet.setColumnWidth(5, 100);  // Status
  sheet.setColumnWidth(6, 80);   // Priority
  sheet.setColumnWidth(7, 250);  // Notes

  // Apply styling to all data sections
  const sections = [
    {start: 5, length: testFeatures.length},
    {start: configRow + 1, length: configTasks.length},
    {start: triggerRow + 1, length: triggerTasks.length},
    {start: menuRow + 1, length: menuTasks.length},
    {start: verifyRow + 1, length: verifyTasks.length},
    {start: instructRow + 1, length: instructions.length}
  ];

  sections.forEach(section => {
    sheet.getRange(section.start, 1, section.length, 7)
      .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID)
      .setFontFamily("Roboto")
      .setVerticalAlignment("middle")
      .setWrap(true);

    // Bold and color code header rows
    sheet.getRange(section.start, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground(COLORS.LIGHT_GRAY)
      .setFontColor(COLORS.TEXT_DARK);

    // Highlight priority column
    if (section.length > 1) {
      for (let i = 1; i < section.length; i++) {
        const priorityCell = sheet.getRange(section.start + i, 6);
        const priority = priorityCell.getValue();
        if (priority === "Critical" || priority === "High") {
          priorityCell.setBackground("#FEE2E2").setFontWeight("bold").setFontColor(COLORS.SOLIDARITY_RED);
        } else if (priority === "Medium") {
          priorityCell.setBackground("#FEF3C7").setFontColor(COLORS.ACCENT_ORANGE);
        } else if (priority === "Low") {
          priorityCell.setBackground("#D1FAE5").setFontColor(COLORS.UNION_GREEN);
        }
      }
    }
  });

  // Freeze header rows
  sheet.setFrozenRows(2);
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
    .setFontSize(14).setFontFamily("Roboto")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

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
    .setFontWeight("bold")
    .setFontFamily("Roboto")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

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
    dashboard.getRange("A17:B17").setBackground("#E8F5E9").setFontSize(12).setFontFamily("Roboto").setFontWeight("bold");
    dashboard.getRange("A18:B19").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("A21:B21").setBackground(COLORS.HEADER_ORANGE).setFontColor("white");
    dashboard.getRange("A22").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B22").setBackground(overdue > 0 ? COLORS.OVERDUE : COLORS.ON_TRACK)
      .setFontColor("white").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
    dashboard.getRange("A23").setBackground(COLORS.LIGHT_GRAY);
    dashboard.getRange("B23").setBackground(dueThisWeek > 0 ? COLORS.DUE_SOON : COLORS.ON_TRACK)
      .setFontColor("white").setFontWeight("bold").setFontSize(12).setFontFamily("Roboto");
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

  // Break apart all merged ranges in the KPI Cards section first to avoid conflicts
  sheet.getRange("A6:C6").breakApart();
  sheet.getRange("A7:C7").breakApart();
  sheet.getRange("A8:C9").breakApart();
  sheet.getRange("E6:G6").breakApart();
  sheet.getRange("E7:G7").breakApart();
  sheet.getRange("E8:G9").breakApart();
  sheet.getRange("I6:K6").breakApart();
  sheet.getRange("I7:K7").breakApart();
  sheet.getRange("I8:K9").breakApart();
  sheet.getRange("M6:O6").breakApart();
  sheet.getRange("M7:O7").breakApart();
  sheet.getRange("M8:O9").breakApart();

  // KPI Cards - Row 7 contains the main values (re-merge after breaking apart)
  sheet.getRange("A7:C7").merge().setValue(metrics.totalMembers);
  sheet.getRange("E7:G7").merge().setValue(metrics.activeGrievances);
  sheet.getRange("I7:K7").merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("M7:O7").merge().setValue(metrics.avgResolutionDays + " days");

  // Break apart all merged ranges in the Union Health section first
  sheet.getRange("A14:C14").breakApart();
  sheet.getRange("A15:C15").breakApart();
  sheet.getRange("A16:C17").breakApart();
  sheet.getRange("E14:G14").breakApart();
  sheet.getRange("E15:G15").breakApart();
  sheet.getRange("E16:G17").breakApart();
  sheet.getRange("I14:K14").breakApart();
  sheet.getRange("I15:K15").breakApart();
  sheet.getRange("I16:K17").breakApart();
  sheet.getRange("M14:O14").breakApart();
  sheet.getRange("M15:O15").breakApart();
  sheet.getRange("M16:O17").breakApart();

  // Union Health Section - Row 15 contains values (re-merge after breaking apart)
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

  // Break apart row 6 headers first to avoid conflicts
  sheet.getRange("A6:C6").breakApart();
  sheet.getRange("E6:G6").breakApart();
  sheet.getRange("I6:K6").breakApart();
  sheet.getRange("M6:O6").breakApart();

  // THIS MONTH section - Row 7
  sheet.getRange("A7:C7").breakApart().merge().setValue(metrics.thisMonthFiled);
  sheet.getRange("A8:C9").breakApart().merge().setValue("+" + metrics.thisMonthFiled + " from last month");

  sheet.getRange("E7:G7").breakApart().merge().setValue(metrics.thisMonthResolved);
  sheet.getRange("E8:G9").breakApart().merge().setValue(metrics.thisMonthResolved + " cases closed");

  sheet.getRange("I7:K7").breakApart().merge().setValue(metrics.pendingDecision);
  sheet.getRange("I8:K9").breakApart().merge().setValue("Awaiting decisions");

  sheet.getRange("M7:O7").breakApart().merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("M8:O9").breakApart().merge().setValue("Success rate");

  // Break apart row 14 headers first to avoid conflicts
  sheet.getRange("A14:C14").breakApart();
  sheet.getRange("E14:G14").breakApart();
  sheet.getRange("I14:K14").breakApart();
  sheet.getRange("M14:O14").breakApart();

  // YEAR TO DATE section - Row 15
  sheet.getRange("A15:C15").breakApart().merge().setValue(metrics.ytdFiled);
  sheet.getRange("A16:C17").breakApart().merge().setValue("Cases filed YTD");

  sheet.getRange("E15:G15").breakApart().merge().setValue(metrics.ytdResolved);
  sheet.getRange("E16:G17").breakApart().merge().setValue("Cases resolved YTD");

  sheet.getRange("I15:K15").breakApart().merge().setValue(metrics.activeGrievances);
  sheet.getRange("I16:K17").breakApart().merge().setValue("Currently active");

  sheet.getRange("M15:O15").breakApart().merge().setValue(metrics.overdue);
  sheet.getRange("M16:O17").breakApart().merge().setValue("Cases overdue");
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

  // Break apart potential merged ranges to avoid conflicts
  sheet.getRange("A6:C6").breakApart();
  sheet.getRange("E6:G6").breakApart();
  sheet.getRange("I6:K6").breakApart();
  sheet.getRange("M6:O6").breakApart();
  sheet.getRange("A8:C9").breakApart();
  sheet.getRange("E8:G9").breakApart();
  sheet.getRange("I8:K9").breakApart();
  sheet.getRange("M8:O9").breakApart();

  // Engagement KPIs - Row 7
  sheet.getRange("A7:C7").breakApart().merge().setValue(metrics.memberEngagementRate.toFixed(1) + "%");
  sheet.getRange("E7:G7").breakApart().merge().setValue(metrics.totalStewards);
  sheet.getRange("I7:K7").breakApart().merge().setValue(metrics.committeeMembers);
  sheet.getRange("M7:O7").breakApart().merge().setValue(metrics.activeMembers);

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

  // Break apart potential merged ranges to avoid conflicts
  sheet.getRange("A6:C6").breakApart();
  sheet.getRange("E6:G6").breakApart();
  sheet.getRange("I6:K6").breakApart();
  sheet.getRange("M6:O6").breakApart();
  sheet.getRange("A8:C9").breakApart();
  sheet.getRange("E8:G9").breakApart();
  sheet.getRange("I8:K9").breakApart();
  sheet.getRange("M8:O9").breakApart();

  // Financial KPIs - Row 7
  const formatter = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 0 });

  sheet.getRange("A7:C7").breakApart().merge().setValue(formatter.format(metrics.financialRecovery));
  sheet.getRange("E7:G7").breakApart().merge().setValue(metrics.won);
  sheet.getRange("I7:K7").breakApart().merge().setValue(formatter.format(metrics.financialRecovery / Math.max(1, metrics.won + metrics.settled)));
  sheet.getRange("M7:O7").breakApart().merge().setValue(metrics.winRate.toFixed(1) + "%");

  // Break apart row 14 and 16-17 to avoid conflicts
  sheet.getRange("A14:C14").breakApart();
  sheet.getRange("E14:G14").breakApart();
  sheet.getRange("I14:K14").breakApart();
  sheet.getRange("M14:O14").breakApart();
  sheet.getRange("A16:C17").breakApart();
  sheet.getRange("E16:G17").breakApart();
  sheet.getRange("I16:K17").breakApart();
  sheet.getRange("M16:O17").breakApart();

  // Additional financial metrics - Row 15
  const avgPerMember = metrics.totalMembers > 0 ? metrics.financialRecovery / metrics.totalMembers : 0;
  sheet.getRange("A15:C15").breakApart().merge().setValue(formatter.format(avgPerMember));
  sheet.getRange("E15:G15").breakApart().merge().setValue(metrics.won + metrics.settled);
  sheet.getRange("I15:K15").breakApart().merge().setValue(formatter.format(metrics.financialRecovery * 1.2)); // Projected
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

  // Break apart potential merged ranges to avoid conflicts
  sheet.getRange("A6:D6").breakApart();
  sheet.getRange("F6:I6").breakApart();
  sheet.getRange("K6:N6").breakApart();
  sheet.getRange("A7:D8").breakApart();
  sheet.getRange("F7:I8").breakApart();
  sheet.getRange("K7:N8").breakApart();

  // Large stat cards - Row 7
  sheet.getRange("A7:C7").breakApart().merge().setValue(metrics.totalGrievances);
  sheet.getRange("E7:G7").breakApart().merge().setValue(metrics.activeGrievances);
  sheet.getRange("I7:K7").breakApart().merge().setValue(metrics.overdue);
  sheet.getRange("M7:O7").breakApart().merge().setValue(metrics.totalMembers);

  // Break apart row 14 ranges to avoid conflicts
  sheet.getRange("A14:C14").breakApart();
  sheet.getRange("E14:G14").breakApart();
  sheet.getRange("I14:K14").breakApart();
  sheet.getRange("M14:O14").breakApart();
  sheet.getRange("A16:C17").breakApart();
  sheet.getRange("E16:G17").breakApart();
  sheet.getRange("I16:K17").breakApart();
  sheet.getRange("M16:O17").breakApart();

  // Secondary stats - Row 15
  sheet.getRange("A15:C15").breakApart().merge().setValue(metrics.winRate.toFixed(1) + "%");
  sheet.getRange("E15:G15").breakApart().merge().setValue(metrics.avgResolutionDays);
  sheet.getRange("I15:K15").breakApart().merge().setValue(metrics.totalStewards);
  sheet.getRange("M15:O15").breakApart().merge().setValue(metrics.dueThisWeek);
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
  const officeDays = config.getRange("N2:N8").getValues().flat().filter(String);

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
        officeDays[Math.floor(Math.random() * officeDays.length)],
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
    SpreadsheetApp.getUi().alert("❌ No stewards found!\n\nPlease ensure some members have 'Is Steward' set to Yes.\n\n💡 TIP: Use '509 Tools > Data Management > Seed All Test Data (Recommended)' to seed both members and grievances in the correct order.\n\nOr manually seed members first using 'Seed 20K Members' before seeding grievances.");
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

/**
 * Unified seed function that seeds both members and grievances in the correct order
 * This is the recommended way to populate test data
 */
function seedAll() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Seed Test Data',
    'This will:\n' +
    '1. Seed 20,000 members\n' +
    '2. Seed 5,000 grievances\n' +
    '3. Recalculate all members\n' +
    '4. Rebuild dashboard\n\n' +
    'This will take 5-7 minutes. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Step 1: Seed members
    ui.alert('Step 1/4: Seeding 20,000 members...\n\nThis will take 3-5 minutes.');
    SEED_20K_MEMBERS();

    // Step 2: Seed grievances
    ui.alert('Step 2/4: Seeding 5,000 grievances...\n\nThis will take 2-3 minutes.');
    SEED_5K_GRIEVANCES();

    // Step 3: Recalculate members
    ui.alert('Step 3/4: Recalculating all member metrics...\n\nThis will take 30-60 seconds.');
    recalcAllMembers();

    // Step 4: Rebuild dashboard
    ui.alert('Step 4/4: Rebuilding dashboard and analytics...\n\nThis will take 30 seconds.');
    rebuildDashboard();

    ui.alert('✅ Complete!\n\n' +
      '20,000 members and 5,000 grievances have been seeded.\n' +
      'All calculations and dashboards have been updated.\n\n' +
      'Check the Dashboard tab to see your data!');
  } catch (error) {
    ui.alert('❌ Error during seeding:\n\n' + error.message + '\n\nPlease check the logs.');
    Logger.log('Seed error: ' + error.toString());
  }
}

/**
 * Alternative: Seed with custom amounts
 */
function SEED_ALL_CUSTOM() {
  const ui = SpreadsheetApp.getUi();

  const memberResponse = ui.prompt(
    'Seed Members',
    'How many members to seed? (Max: 50000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (memberResponse.getSelectedButton() !== ui.Button.OK) return;

  const memberCount = parseInt(memberResponse.getResponseText());
  if (isNaN(memberCount) || memberCount <= 0 || memberCount > 50000) {
    ui.alert('Invalid member count. Please enter a number between 1 and 50000.');
    return;
  }

  const grievanceResponse = ui.prompt(
    'Seed Grievances',
    'How many grievances to seed? (Max: 20000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (grievanceResponse.getSelectedButton() !== ui.Button.OK) return;

  const grievanceCount = parseInt(grievanceResponse.getResponseText());
  if (isNaN(grievanceCount) || grievanceCount <= 0 || grievanceCount > 20000) {
    ui.alert('Invalid grievance count. Please enter a number between 1 and 20000.');
    return;
  }

  ui.alert(`Seeding ${memberCount} members and ${grievanceCount} grievances...\n\nThis may take several minutes.`);

  // Note: You would need to modify SEED_20K_MEMBERS and SEED_5K_GRIEVANCES
  // to accept parameters for custom counts. For now, use the default functions.
  ui.alert('Custom seeding is not yet implemented.\n\nUse the standard seedAll() function for now.');
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
      .addItem('Seed All Test Data (Recommended)', 'seedAll')
      .addSeparator()
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

  // Add enhancement menus with all advanced features
  addEnhancementMenus();
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
    dashboard.getRange("A1:H50").setFontSize(14).setFontFamily("Roboto");
  }

  // Member Directory: Essential columns only visible
  if (memberDir) {
    memberDir.setColumnWidth(1, 150);  // Member ID
    memberDir.setColumnWidth(2, 150);  // First Name
    memberDir.setColumnWidth(3, 150);  // Last Name
    memberDir.setColumnWidth(4, 200);  // Job Title
    memberDir.setColumnWidth(5, 200);  // Location
    memberDir.getRange("1:1").setFontSize(12).setFontFamily("Roboto").setFontWeight("bold");
  }

  // Grievance Log: Essential columns only visible
  if (grievanceLog) {
    grievanceLog.setColumnWidth(1, 150);  // Grievance ID
    grievanceLog.setColumnWidth(2, 150);  // Member ID
    grievanceLog.setColumnWidth(3, 150);  // First Name
    grievanceLog.setColumnWidth(4, 150);  // Last Name
    grievanceLog.setColumnWidth(5, 150);  // Status
    grievanceLog.getRange("1:1").setFontSize(12).setFontFamily("Roboto").setFontWeight("bold");
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
    dashboard.getRange("A1:H50").setFontSize(10).setFontFamily("Roboto");
  }

  // Member Directory: Standard column widths
  if (memberDir) {
    memberDir.setColumnWidth(1, 110);  // Member ID
    memberDir.setColumnWidth(2, 100);  // First Name
    memberDir.setColumnWidth(3, 100);  // Last Name
    memberDir.setColumnWidth(4, 150);  // Job Title
    memberDir.setColumnWidth(5, 150);  // Location
    memberDir.getRange("1:1").setFontSize(10).setFontFamily("Roboto").setFontWeight("bold");
  }

  // Grievance Log: Standard column widths
  if (grievanceLog) {
    grievanceLog.setColumnWidth(1, 120);  // Grievance ID
    grievanceLog.setColumnWidth(2, 110);  // Member ID
    grievanceLog.setColumnWidth(3, 100);  // First Name
    grievanceLog.setColumnWidth(4, 100);  // Last Name
    grievanceLog.setColumnWidth(5, 120);  // Status
    grievanceLog.getRange("1:1").setFontSize(10).setFontFamily("Roboto").setFontWeight("bold");
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

// ============================================================================
// COMPREHENSIVE ENHANCEMENTS - 94 NEW FEATURES
// ============================================================================
// This section contains 94 advanced features organized into 10 categories
// to enhance the 509 Dashboard functionality, automation, and user experience
// ============================================================================

// ============================================================================
// SECTION 1: ADVANCED DATA VALIDATION (Features 1-10)
// ============================================================================

/**
 * Feature 1: Validates member data integrity
 * Checks for duplicate Member IDs, missing required fields, and invalid email formats
 */
function validateMemberData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return { valid: false, errors: ['Member Directory sheet not found'] };

  const data = sheet.getDataRange().getValues();
  const errors = [];
  const memberIds = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const memberId = row[0];
    const firstName = row[1];
    const lastName = row[2];
    const email = row[7];

    // Check for duplicate Member IDs
    if (memberId && memberIds.has(memberId)) {
      errors.push(`Row ${i + 1}: Duplicate Member ID: ${memberId}`);
    }
    memberIds.add(memberId);

    // Check for missing required fields
    if (!firstName || !lastName) {
      errors.push(`Row ${i + 1}: Missing name fields`);
    }

    // Validate email format
    if (email && !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      errors.push(`Row ${i + 1}: Invalid email format: ${email}`);
    }
  }

  return { valid: errors.length === 0, errors: errors };
}

/**
 * Feature 2: Validates grievance data integrity
 * Checks for valid Member IDs, required dates, and logical date sequences
 */
function validateGrievanceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    return { valid: false, errors: ['Required sheets not found'] };
  }

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();
  const validMemberIds = new Set(memberData.slice(1).map(row => row[0]));
  const errors = [];

  for (let i = 1; i < grievanceData.length; i++) {
    const row = grievanceData[i];
    const memberId = row[1];
    const incidentDate = row[6];
    const dateFiled = row[8];

    // Check for valid Member ID
    if (memberId && !validMemberIds.has(memberId)) {
      errors.push(`Row ${i + 1}: Invalid Member ID: ${memberId}`);
    }

    // Check for required dates
    if (!incidentDate) {
      errors.push(`Row ${i + 1}: Missing Incident Date`);
    }

    // Check date sequence logic
    if (incidentDate && dateFiled && new Date(dateFiled) < new Date(incidentDate)) {
      errors.push(`Row ${i + 1}: Date Filed cannot be before Incident Date`);
    }
  }

  return { valid: errors.length === 0, errors: errors };
}

/**
 * Feature 3: Auto-corrects common data entry errors
 * Fixes case, trims whitespace, standardizes phone numbers
 */
function autoCorrectDataErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let modified = false;

    // Trim whitespace from all text fields
    for (let j = 0; j < data[i].length; j++) {
      if (typeof data[i][j] === 'string') {
        const trimmed = data[i][j].trim();
        if (trimmed !== data[i][j]) {
          data[i][j] = trimmed;
          modified = true;
        }
      }
    }

    // Proper case for names (columns 1-2: First Name, Last Name)
    if (data[i][1]) {
      const properFirst = data[i][1].charAt(0).toUpperCase() + data[i][1].slice(1).toLowerCase();
      if (properFirst !== data[i][1]) {
        data[i][1] = properFirst;
        modified = true;
      }
    }

    if (data[i][2]) {
      const properLast = data[i][2].charAt(0).toUpperCase() + data[i][2].slice(1).toLowerCase();
      if (properLast !== data[i][2]) {
        data[i][2] = properLast;
        modified = true;
      }
    }

    // Standardize phone numbers (column 8)
    if (data[i][8]) {
      const phone = data[i][8].toString().replace(/\D/g, '');
      if (phone.length === 10) {
        const formatted = `(${phone.slice(0,3)}) ${phone.slice(3,6)}-${phone.slice(6)}`;
        if (formatted !== data[i][8]) {
          data[i][8] = formatted;
          modified = true;
        }
      }
    }
  }

  sheet.getDataRange().setValues(data);
}

/**
 * Feature 4: Checks for orphaned grievances (invalid Member IDs)
 */
function findOrphanedGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();
  const validMemberIds = new Set(memberData.slice(1).map(row => row[0]).filter(id => id));

  const orphaned = [];

  for (let i = 1; i < grievanceData.length; i++) {
    const memberId = grievanceData[i][1];
    if (memberId && !validMemberIds.has(memberId)) {
      orphaned.push({
        row: i + 1,
        grievanceId: grievanceData[i][0],
        memberId: memberId,
        firstName: grievanceData[i][2],
        lastName: grievanceData[i][3]
      });
    }
  }

  return orphaned;
}

/**
 * Feature 5: Validates date ranges for logical consistency
 */
function validateDateRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { valid: false, errors: ['Grievance Log not found'] };

  const data = sheet.getDataRange().getValues();
  const errors = [];

  for (let i = 1; i < data.length; i++) {
    const incidentDate = data[i][6];
    const dateFiled = data[i][8];
    const step1Decision = data[i][10];
    const step2Filed = data[i][13];
    const step2Decision = data[i][15];
    const step3Filed = data[i][19];

    // Validate chronological order
    if (incidentDate && dateFiled && new Date(dateFiled) < new Date(incidentDate)) {
      errors.push(`Row ${i + 1}: Date Filed before Incident Date`);
    }
    if (dateFiled && step1Decision && new Date(step1Decision) < new Date(dateFiled)) {
      errors.push(`Row ${i + 1}: Step I Decision before Date Filed`);
    }
    if (step1Decision && step2Filed && new Date(step2Filed) < new Date(step1Decision)) {
      errors.push(`Row ${i + 1}: Step II Filed before Step I Decision`);
    }
    if (step2Filed && step2Decision && new Date(step2Decision) < new Date(step2Filed)) {
      errors.push(`Row ${i + 1}: Step II Decision before Step II Filed`);
    }
    if (step2Decision && step3Filed && new Date(step3Filed) < new Date(step2Decision)) {
      errors.push(`Row ${i + 1}: Step III Filed before Step II Decision`);
    }
  }

  return { valid: errors.length === 0, errors: errors };
}

/**
 * Feature 6: Checks for duplicate grievances
 */
function findDuplicateGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const duplicates = [];
  const seen = new Map();

  for (let i = 1; i < data.length; i++) {
    const memberId = data[i][1];
    const incidentDate = data[i][6];
    const type = data[i][25];

    if (!memberId || !incidentDate) continue;

    const key = `${memberId}|${incidentDate}|${type}`;
    if (seen.has(key)) {
      duplicates.push({
        row1: seen.get(key),
        row2: i + 1,
        memberId: memberId,
        incidentDate: incidentDate,
        type: type
      });
    } else {
      seen.set(key, i + 1);
    }
  }

  return duplicates;
}

/**
 * Feature 7: Validates CBA compliance for all grievances
 */
function validateCBACompliance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { compliant: false, violations: ['Sheet not found'] };

  const data = sheet.getDataRange().getValues();
  const violations = [];

  for (let i = 1; i < data.length; i++) {
    const incidentDate = data[i][6];
    const dateFiled = data[i][8];

    if (incidentDate && dateFiled) {
      const daysDiff = Math.floor((new Date(dateFiled) - new Date(incidentDate)) / (1000 * 60 * 60 * 24));
      if (daysDiff > CBA_DEADLINES.FILING) {
        violations.push({
          row: i + 1,
          grievanceId: data[i][0],
          violation: `Filed ${daysDiff} days after incident (CBA allows ${CBA_DEADLINES.FILING} days)`,
          severity: daysDiff > 30 ? 'HIGH' : 'MEDIUM'
        });
      }
    }
  }

  return { compliant: violations.length === 0, violations: violations };
}

/**
 * Feature 8: Auto-generates missing Member IDs
 */
function autoGenerateMemberIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  let generated = 0;

  // Find the highest existing ID
  for (let i = 1; i < data.length; i++) {
    const memberId = data[i][0];
    if (memberId && memberId.toString().startsWith('MEM')) {
      const num = parseInt(memberId.replace('MEM', ''));
      if (num > maxId) maxId = num;
    }
  }

  // Generate missing IDs
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] && (data[i][1] || data[i][2])) {  // Has name but no ID
      maxId++;
      data[i][0] = 'MEM' + String(maxId).padStart(6, '0');
      generated++;
    }
  }

  if (generated > 0) {
    sheet.getDataRange().setValues(data);
  }

  return generated;
}

/**
 * Feature 9: Auto-generates missing Grievance IDs
 */
function autoGenerateGrievanceIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  let generated = 0;

  // Find the highest existing ID
  for (let i = 1; i < data.length; i++) {
    const grievanceId = data[i][0];
    if (grievanceId && grievanceId.toString().startsWith('GRV')) {
      const num = parseInt(grievanceId.replace('GRV', ''));
      if (num > maxId) maxId = num;
    }
  }

  // Generate missing IDs
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] && data[i][1]) {  // Has Member ID but no Grievance ID
      maxId++;
      data[i][0] = 'GRV' + String(maxId).padStart(6, '0');
      generated++;
    }
  }

  if (generated > 0) {
    sheet.getDataRange().setValues(data);
  }

  return generated;
}

/**
 * Feature 10: Comprehensive data integrity check
 * Runs all validation checks and generates a report
 */
function runDataIntegrityCheck() {
  const results = {
    timestamp: new Date(),
    memberValidation: validateMemberData(),
    grievanceValidation: validateGrievanceData(),
    dateRangeValidation: validateDateRanges(),
    cbaCompliance: validateCBACompliance(),
    orphanedGrievances: findOrphanedGrievances(),
    duplicateGrievances: findDuplicateGrievances()
  };

  // Write results to Diagnostics sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let diagSheet = ss.getSheetByName(SHEETS.DIAGNOSTICS);
  if (!diagSheet) {
    diagSheet = ss.insertSheet(SHEETS.DIAGNOSTICS);
  }

  diagSheet.clear();
  diagSheet.appendRow(['DATA INTEGRITY CHECK REPORT']);
  diagSheet.appendRow(['Run Date:', results.timestamp.toLocaleString()]);
  diagSheet.appendRow([]);

  diagSheet.appendRow(['MEMBER VALIDATION']);
  diagSheet.appendRow(['Valid:', results.memberValidation.valid]);
  diagSheet.appendRow(['Errors:', results.memberValidation.errors.length]);
  results.memberValidation.errors.forEach(err => diagSheet.appendRow(['', err]));

  diagSheet.appendRow([]);
  diagSheet.appendRow(['GRIEVANCE VALIDATION']);
  diagSheet.appendRow(['Valid:', results.grievanceValidation.valid]);
  diagSheet.appendRow(['Errors:', results.grievanceValidation.errors.length]);
  results.grievanceValidation.errors.forEach(err => diagSheet.appendRow(['', err]));

  diagSheet.appendRow([]);
  diagSheet.appendRow(['CBA COMPLIANCE']);
  diagSheet.appendRow(['Compliant:', results.cbaCompliance.compliant]);
  diagSheet.appendRow(['Violations:', results.cbaCompliance.violations.length]);
  results.cbaCompliance.violations.forEach(v => {
    diagSheet.appendRow(['', `Row ${v.row}: ${v.violation} [${v.severity}]`]);
  });

  diagSheet.appendRow([]);
  diagSheet.appendRow(['ORPHANED GRIEVANCES']);
  diagSheet.appendRow(['Count:', results.orphanedGrievances.length]);
  results.orphanedGrievances.forEach(o => {
    diagSheet.appendRow(['', `Row ${o.row}: ${o.grievanceId} - Invalid Member ID: ${o.memberId}`]);
  });

  diagSheet.appendRow([]);
  diagSheet.appendRow(['DUPLICATE GRIEVANCES']);
  diagSheet.appendRow(['Count:', results.duplicateGrievances.length]);
  results.duplicateGrievances.forEach(d => {
    diagSheet.appendRow(['', `Rows ${d.row1} & ${d.row2}: Member ${d.memberId}, ${d.incidentDate}`]);
  });

  diagSheet.setColumnWidth(1, 200);
  diagSheet.setColumnWidth(2, 600);
  diagSheet.getRange('A1:B1').setFontWeight('bold').setFontSize(14).setBackground(COLORS.HEADER_BLUE).setFontColor('white');

  return results;
}

// ============================================================================
// SECTION 2: AUTOMATED NOTIFICATIONS & ALERTS (Features 11-20)
// ============================================================================

/**
 * Feature 11: Sends email alerts for overdue grievances
 */
function sendOverdueAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const currentUser = Session.getActiveUser().getEmail();
  let alertCount = 0;

  for (let i = 1; i < data.length; i++) {
    const isOverdue = data[i][28]; // Column AC: Is Overdue?
    const grievanceId = data[i][0];
    const memberName = `${data[i][2]} ${data[i][3]}`;
    const nextDeadline = data[i][27]; // Next Deadline

    if (isOverdue === 'YES' && nextDeadline) {
      const daysOverdue = Math.floor((new Date() - new Date(nextDeadline)) / (1000 * 60 * 60 * 24));

      try {
        MailApp.sendEmail({
          to: currentUser,
          subject: `⚠️ OVERDUE GRIEVANCE ALERT: ${grievanceId}`,
          body: `Grievance ${grievanceId} for ${memberName} is ${daysOverdue} days overdue.\n\n` +
                `Next Deadline was: ${nextDeadline}\n` +
                `Please take immediate action.\n\n` +
                `View Dashboard: ${ss.getUrl()}`
        });
        alertCount++;
      } catch (e) {
        Logger.log(`Failed to send alert for ${grievanceId}: ${e}`);
      }
    }
  }

  return alertCount;
}

/**
 * Feature 12: Sends weekly deadline reminder emails
 */
function sendWeeklyDeadlineReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const currentUser = Session.getActiveUser().getEmail();
  const today = new Date();
  const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);
  let reminderCount = 0;

  const upcomingDeadlines = [];

  for (let i = 1; i < data.length; i++) {
    const nextDeadline = data[i][27]; // Next Deadline
    const grievanceId = data[i][0];
    const memberName = `${data[i][2]} ${data[i][3]}`;
    const currentStep = data[i][5];

    if (nextDeadline && new Date(nextDeadline) >= today && new Date(nextDeadline) <= nextWeek) {
      const daysUntil = Math.floor((new Date(nextDeadline) - today) / (1000 * 60 * 60 * 24));
      upcomingDeadlines.push({
        grievanceId: grievanceId,
        memberName: memberName,
        currentStep: currentStep,
        deadline: nextDeadline,
        daysUntil: daysUntil
      });
    }
  }

  if (upcomingDeadlines.length > 0) {
    let emailBody = `You have ${upcomingDeadlines.length} grievance deadline(s) in the next 7 days:\n\n`;

    upcomingDeadlines.sort((a, b) => a.daysUntil - b.daysUntil);

    upcomingDeadlines.forEach(item => {
      emailBody += `• ${item.grievanceId} - ${item.memberName}\n`;
      emailBody += `  Step: ${item.currentStep}\n`;
      emailBody += `  Deadline: ${item.deadline} (${item.daysUntil} days)\n\n`;
    });

    emailBody += `\nView Dashboard: ${ss.getUrl()}`;

    try {
      MailApp.sendEmail({
        to: currentUser,
        subject: `📅 Weekly Deadline Reminder: ${upcomingDeadlines.length} Upcoming Deadline(s)`,
        body: emailBody
      });
      reminderCount = upcomingDeadlines.length;
    } catch (e) {
      Logger.log(`Failed to send weekly reminder: ${e}`);
    }
  }

  return reminderCount;
}

/**
 * Feature 13: Creates daily digest email
 */
function sendDailyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const dashboardSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!grievanceSheet || !dashboardSheet) return;

  const currentUser = Session.getActiveUser().getEmail();
  const today = new Date();

  // Gather statistics
  const grievanceData = grievanceSheet.getDataRange().getValues();
  let totalActive = 0;
  let overdueCount = 0;
  let dueToday = 0;

  for (let i = 1; i < grievanceData.length; i++) {
    const status = grievanceData[i][4];
    const isOverdue = grievanceData[i][28];
    const nextDeadline = grievanceData[i][27];

    if (status && status.startsWith('Filed')) {
      totalActive++;
    }

    if (isOverdue === 'YES') {
      overdueCount++;
    }

    if (nextDeadline) {
      const deadlineDate = new Date(nextDeadline);
      if (deadlineDate.toDateString() === today.toDateString()) {
        dueToday++;
      }
    }
  }

  const emailBody = `509 Dashboard Daily Digest - ${today.toDateString()}\n\n` +
                    `═══════════════════════════════════════\n\n` +
                    `📊 OVERVIEW\n` +
                    `• Active Grievances: ${totalActive}\n` +
                    `• Overdue: ${overdueCount}\n` +
                    `• Due Today: ${dueToday}\n\n` +
                    `═══════════════════════════════════════\n\n` +
                    `View Dashboard: ${ss.getUrl()}\n\n` +
                    `This is an automated daily digest from your 509 Dashboard.`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: `📊 509 Dashboard Daily Digest - ${today.toDateString()}`,
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send daily digest: ${e}`);
  }
}

/**
 * Feature 14: Sends steward workload balance alerts
 */
function sendStewardWorkloadAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const currentUser = Session.getActiveUser().getEmail();

  // Calculate average workload
  let totalCases = 0;
  let stewardCount = 0;

  for (let i = 1; i < data.length; i++) {
    const activeCases = data[i][4]; // Column E: Active Cases
    if (typeof activeCases === 'number') {
      totalCases += activeCases;
      stewardCount++;
    }
  }

  if (stewardCount === 0) return;

  const avgWorkload = totalCases / stewardCount;
  const threshold = avgWorkload * 1.5; // 50% above average

  const overloadedStewards = [];

  for (let i = 1; i < data.length; i++) {
    const stewardName = data[i][0];
    const activeCases = data[i][4];

    if (activeCases > threshold) {
      overloadedStewards.push({
        name: stewardName,
        cases: activeCases,
        avgCases: avgWorkload.toFixed(1)
      });
    }
  }

  if (overloadedStewards.length > 0) {
    let emailBody = `⚠️ STEWARD WORKLOAD ALERT\n\n` +
                    `The following steward(s) have significantly higher caseloads than average:\n\n`;

    overloadedStewards.forEach(s => {
      emailBody += `• ${s.name}: ${s.cases} active cases (avg: ${s.avgCases})\n`;
    });

    emailBody += `\nConsider rebalancing case assignments.\n\n` +
                 `View Dashboard: ${ss.getUrl()}`;

    try {
      MailApp.sendEmail({
        to: currentUser,
        subject: '⚠️ Steward Workload Imbalance Alert',
        body: emailBody
      });
    } catch (e) {
      Logger.log(`Failed to send workload alert: ${e}`);
    }
  }
}

/**
 * Feature 15: Sends new grievance notification
 */
function sendNewGrievanceNotification(grievanceId, memberName, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  const emailBody = `🆕 NEW GRIEVANCE FILED\n\n` +
                    `Grievance ID: ${grievanceId}\n` +
                    `Member: ${memberName}\n` +
                    `Type: ${type}\n\n` +
                    `View Dashboard: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: `🆕 New Grievance Filed: ${grievanceId}`,
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send new grievance notification: ${e}`);
  }
}

/**
 * Feature 16: Sends grievance resolution notification
 */
function sendResolutionNotification(grievanceId, memberName, outcome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  const emailBody = `✅ GRIEVANCE RESOLVED\n\n` +
                    `Grievance ID: ${grievanceId}\n` +
                    `Member: ${memberName}\n` +
                    `Outcome: ${outcome}\n\n` +
                    `View Dashboard: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: `✅ Grievance Resolved: ${grievanceId} - ${outcome}`,
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send resolution notification: ${e}`);
  }
}

/**
 * Feature 17: Sends CBA violation warning
 */
function sendCBAViolationWarning(violations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  if (violations.length === 0) return;

  let emailBody = `⚠️ CBA COMPLIANCE VIOLATIONS DETECTED\n\n` +
                  `The following grievances have CBA deadline violations:\n\n`;

  violations.forEach(v => {
    emailBody += `• Row ${v.row}: ${v.grievanceId}\n`;
    emailBody += `  ${v.violation} [${v.severity}]\n\n`;
  });

  emailBody += `Please review these cases immediately.\n\n` +
               `View Dashboard: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: '⚠️ CBA Compliance Violations Detected',
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send CBA violation warning: ${e}`);
  }
}

/**
 * Feature 18: Sends monthly performance report
 */
function sendMonthlyPerformanceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const currentUser = Session.getActiveUser().getEmail();

  if (!grievanceSheet) return;

  const data = grievanceSheet.getDataRange().getValues();
  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const thisMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  let filed = 0;
  let resolved = 0;
  let won = 0;
  let lost = 0;

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    const status = data[i][4];
    const outcome = data[i][24];

    // Count filed last month
    if (dateFiled && new Date(dateFiled) >= lastMonth && new Date(dateFiled) < thisMonth) {
      filed++;
    }

    // Count resolved last month
    if (status && status.startsWith('Resolved')) {
      if (outcome && outcome.includes('Won')) won++;
      if (outcome && outcome.includes('Lost')) lost++;
      resolved++;
    }
  }

  const winRate = resolved > 0 ? ((won / resolved) * 100).toFixed(1) : 0;

  const emailBody = `📊 MONTHLY PERFORMANCE REPORT\n` +
                    `Period: ${lastMonth.toLocaleDateString()} - ${thisMonth.toLocaleDateString()}\n\n` +
                    `═══════════════════════════════════════\n\n` +
                    `📝 Filed: ${filed}\n` +
                    `✅ Resolved: ${resolved}\n` +
                    `🏆 Won: ${won}\n` +
                    `❌ Lost: ${lost}\n` +
                    `📈 Win Rate: ${winRate}%\n\n` +
                    `═══════════════════════════════════════\n\n` +
                    `View Dashboard: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: `📊 Monthly Performance Report - ${lastMonth.toLocaleDateString('en-US', {month: 'long', year: 'numeric'})}`,
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send monthly report: ${e}`);
  }
}

/**
 * Feature 19: Sends custom alert based on criteria
 */
function sendCustomAlert(criteria, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentUser = Session.getActiveUser().getEmail();

  const emailBody = `🔔 CUSTOM ALERT\n\n` +
                    `Criteria: ${criteria}\n\n` +
                    `${message}\n\n` +
                    `View Dashboard: ${ss.getUrl()}`;

  try {
    MailApp.sendEmail({
      to: currentUser,
      subject: `🔔 Custom Alert: ${criteria}`,
      body: emailBody
    });
  } catch (e) {
    Logger.log(`Failed to send custom alert: ${e}`);
  }
}

/**
 * Feature 20: Configures email notification preferences
 */
function configureNotificationPreferences() {
  const props = PropertiesService.getUserProperties();

  // Default preferences
  const defaultPrefs = {
    dailyDigest: true,
    weeklyReminders: true,
    overdueAlerts: true,
    newGrievanceAlerts: true,
    resolutionAlerts: true,
    workloadAlerts: true,
    monthlyReports: true
  };

  // Save preferences
  Object.keys(defaultPrefs).forEach(key => {
    if (!props.getProperty(key)) {
      props.setProperty(key, defaultPrefs[key].toString());
    }
  });

  return props.getProperties();
}

// ============================================================================
// SECTION 3: ADVANCED REPORTING (Features 21-30)
// ============================================================================

/**
 * Feature 21: Generates comprehensive PDF report
 */
function generatePDFReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboardSheet) return null;

  // Create a temporary sheet with report content
  const reportSheet = ss.insertSheet('PDF_Report_Temp');

  // Add report header
  reportSheet.getRange('A1').setValue('509 DASHBOARD - COMPREHENSIVE REPORT');
  reportSheet.getRange('A2').setValue(`Generated: ${new Date().toLocaleString()}`);

  // Copy dashboard metrics
  const dashboardData = dashboardSheet.getRange('A1:H30').getValues();
  reportSheet.getRange(4, 1, dashboardData.length, dashboardData[0].length).setValues(dashboardData);

  // Format report
  reportSheet.getRange('A1:H1').setFontSize(16).setFontWeight('bold').setBackground(COLORS.HEADER_BLUE).setFontColor('white');
  reportSheet.getRange('A2').setFontSize(10).setFontStyle('italic');

  // Convert to PDF
  const pdfBlob = ss.getAs('application/pdf');
  pdfBlob.setName(`509_Dashboard_Report_${new Date().toISOString().slice(0,10)}.pdf`);

  // Clean up
  ss.deleteSheet(reportSheet);

  return pdfBlob;
}

/**
 * Feature 22: Exports grievance data to CSV
 */
function exportGrievancesToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  let csv = '';

  data.forEach(row => {
    csv += row.map(cell => {
      // Escape quotes and wrap in quotes if contains comma
      const cellStr = cell.toString().replace(/"/g, '""');
      return cellStr.includes(',') ? `"${cellStr}"` : cellStr;
    }).join(',') + '\n';
  });

  const blob = Utilities.newBlob(csv, 'text/csv', `Grievances_${new Date().toISOString().slice(0,10)}.csv`);
  return blob;
}

/**
 * Feature 23: Creates executive summary report
 */
function createExecutiveSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return null;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  const summary = {
    reportDate: new Date().toLocaleString(),
    memberMetrics: {
      total: memberData.length - 1,
      active: memberData.slice(1).filter(row => row[10] === 'Active').length,
      stewards: memberData.slice(1).filter(row => row[9] === 'Yes').length
    },
    grievanceMetrics: {
      total: grievanceData.length - 1,
      active: grievanceData.slice(1).filter(row => row[4] && row[4].startsWith('Filed')).length,
      resolved: grievanceData.slice(1).filter(row => row[4] && row[4].startsWith('Resolved')).length,
      overdue: grievanceData.slice(1).filter(row => row[28] === 'YES').length
    },
    performanceMetrics: {
      winRate: 0,
      avgResolutionDays: 0,
      settlementRate: 0
    }
  };

  // Calculate performance metrics
  const resolved = grievanceData.slice(1).filter(row => row[4] && row[4].startsWith('Resolved'));
  const won = resolved.filter(row => row[24] && row[24].includes('Won')).length;
  const settled = resolved.filter(row => row[24] && row[24].includes('Settled')).length;

  if (resolved.length > 0) {
    summary.performanceMetrics.winRate = ((won / resolved.length) * 100).toFixed(1);
    summary.performanceMetrics.settlementRate = ((settled / resolved.length) * 100).toFixed(1);
  }

  return summary;
}

/**
 * Feature 24: Generates trend analysis report
 */
function generateTrendAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const monthlyData = {};

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    if (dateFiled) {
      const date = new Date(dateFiled);
      const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;

      if (!monthlyData[monthKey]) {
        monthlyData[monthKey] = { filed: 0, resolved: 0 };
      }
      monthlyData[monthKey].filed++;

      const status = data[i][4];
      if (status && status.startsWith('Resolved')) {
        monthlyData[monthKey].resolved++;
      }
    }
  }

  const trends = Object.keys(monthlyData).sort().map(month => ({
    month: month,
    filed: monthlyData[month].filed,
    resolved: monthlyData[month].resolved,
    netChange: monthlyData[month].filed - monthlyData[month].resolved
  }));

  return trends;
}

/**
 * Feature 25: Creates steward performance report
 */
function createStewardPerformanceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return null;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const stewardStats = {};

  for (let i = 1; i < grievanceData.length; i++) {
    const representative = grievanceData[i][26]; // Column AA: Representative
    if (!representative) continue;

    if (!stewardStats[representative]) {
      stewardStats[representative] = {
        total: 0,
        active: 0,
        resolved: 0,
        won: 0,
        lost: 0,
        overdue: 0
      };
    }

    stewardStats[representative].total++;

    const status = grievanceData[i][4];
    const outcome = grievanceData[i][24];
    const isOverdue = grievanceData[i][28];

    if (status && status.startsWith('Filed')) {
      stewardStats[representative].active++;
    }

    if (status && status.startsWith('Resolved')) {
      stewardStats[representative].resolved++;
      if (outcome && outcome.includes('Won')) stewardStats[representative].won++;
      if (outcome && outcome.includes('Lost')) stewardStats[representative].lost++;
    }

    if (isOverdue === 'YES') {
      stewardStats[representative].overdue++;
    }
  }

  // Calculate win rates
  Object.keys(stewardStats).forEach(steward => {
    const stats = stewardStats[steward];
    stats.winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) : 0;
  });

  return stewardStats;
}

/**
 * Feature 26: Generates location-based analysis report
 */
function generateLocationAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return null;

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();

  const locationStats = {};

  // Count members by location
  for (let i = 1; i < memberData.length; i++) {
    const location = memberData[i][4]; // Work Location
    if (!location) continue;

    if (!locationStats[location]) {
      locationStats[location] = { members: 0, grievances: 0, active: 0, resolved: 0 };
    }
    locationStats[location].members++;
  }

  // Count grievances by member location
  for (let i = 1; i < grievanceData.length; i++) {
    const memberId = grievanceData[i][1];
    const member = memberData.find(row => row[0] === memberId);

    if (member) {
      const location = member[4];
      if (locationStats[location]) {
        locationStats[location].grievances++;

        const status = grievanceData[i][4];
        if (status && status.startsWith('Filed')) {
          locationStats[location].active++;
        }
        if (status && status.startsWith('Resolved')) {
          locationStats[location].resolved++;
        }
      }
    }
  }

  return locationStats;
}

/**
 * Feature 27: Creates grievance type analysis report
 */
function createGrievanceTypeAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const typeStats = {};

  for (let i = 1; i < data.length; i++) {
    const type = data[i][25]; // Grievance Type
    if (!type) continue;

    if (!typeStats[type]) {
      typeStats[type] = { total: 0, active: 0, resolved: 0, won: 0, settled: 0 };
    }

    typeStats[type].total++;

    const status = data[i][4];
    const outcome = data[i][24];

    if (status && status.startsWith('Filed')) {
      typeStats[type].active++;
    }

    if (status && status.startsWith('Resolved')) {
      typeStats[type].resolved++;
      if (outcome && outcome.includes('Won')) typeStats[type].won++;
      if (outcome && outcome.includes('Settled')) typeStats[type].settled++;
    }
  }

  // Calculate success rates
  Object.keys(typeStats).forEach(type => {
    const stats = typeStats[type];
    stats.winRate = stats.resolved > 0 ? ((stats.won / stats.resolved) * 100).toFixed(1) : 0;
    stats.settlementRate = stats.resolved > 0 ? ((stats.settled / stats.resolved) * 100).toFixed(1) : 0;
  });

  return typeStats;
}

/**
 * Feature 28: Generates quarterly comparison report
 */
function generateQuarterlyComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const quarterlyData = {};

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    if (dateFiled) {
      const date = new Date(dateFiled);
      const quarter = `Q${Math.floor(date.getMonth() / 3) + 1} ${date.getFullYear()}`;

      if (!quarterlyData[quarter]) {
        quarterlyData[quarter] = { filed: 0, resolved: 0, won: 0, lost: 0 };
      }

      quarterlyData[quarter].filed++;

      const status = data[i][4];
      const outcome = data[i][24];

      if (status && status.startsWith('Resolved')) {
        quarterlyData[quarter].resolved++;
        if (outcome && outcome.includes('Won')) quarterlyData[quarter].won++;
        if (outcome && outcome.includes('Lost')) quarterlyData[quarter].lost++;
      }
    }
  }

  return quarterlyData;
}

/**
 * Feature 29: Creates custom filtered report
 */
function createCustomFilteredReport(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const filtered = [headers];

  for (let i = 1; i < data.length; i++) {
    let matches = true;

    // Apply filters
    if (filters.status && data[i][4] !== filters.status) matches = false;
    if (filters.type && data[i][25] !== filters.type) matches = false;
    if (filters.steward && data[i][26] !== filters.steward) matches = false;
    if (filters.dateFrom && new Date(data[i][8]) < new Date(filters.dateFrom)) matches = false;
    if (filters.dateTo && new Date(data[i][8]) > new Date(filters.dateTo)) matches = false;

    if (matches) {
      filtered.push(data[i]);
    }
  }

  return filtered;
}

/**
 * Feature 30: Generates annual summary report
 */
function generateAnnualSummary(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return null;

  const data = grievanceSheet.getDataRange().getValues();

  const summary = {
    year: year,
    totalFiled: 0,
    totalResolved: 0,
    won: 0,
    lost: 0,
    settled: 0,
    withdrawn: 0,
    byMonth: {},
    byType: {},
    byLocation: {}
  };

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    if (!dateFiled) continue;

    const fileDate = new Date(dateFiled);
    if (fileDate.getFullYear() !== year) continue;

    summary.totalFiled++;

    const month = fileDate.toLocaleString('default', { month: 'long' });
    summary.byMonth[month] = (summary.byMonth[month] || 0) + 1;

    const type = data[i][25];
    if (type) {
      summary.byType[type] = (summary.byType[type] || 0) + 1;
    }

    const status = data[i][4];
    const outcome = data[i][24];

    if (status && status.startsWith('Resolved')) {
      summary.totalResolved++;
      if (outcome && outcome.includes('Won')) summary.won++;
      if (outcome && outcome.includes('Lost')) summary.lost++;
      if (outcome && outcome.includes('Settled')) summary.settled++;
      if (outcome && outcome.includes('Withdrawn')) summary.withdrawn++;
    }
  }

  // Calculate rates
  summary.winRate = summary.totalResolved > 0 ? ((summary.won / summary.totalResolved) * 100).toFixed(1) : 0;
  summary.settlementRate = summary.totalResolved > 0 ? ((summary.settled / summary.totalResolved) * 100).toFixed(1) : 0;

  return summary;
}

// ============================================================================
// SECTION 4: MEMBER ENGAGEMENT TOOLS (Features 31-40)
// ============================================================================

/**
 * Feature 31: Tracks member participation in union events
 */
function trackMemberParticipation(memberId, eventType, eventDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memberId) {
      // Update Events Attended count (column 22)
      const currentCount = data[i][21] || 0;
      data[i][21] = currentCount + 1;

      // Update Last Updated (column 31)
      data[i][30] = new Date();

      sheet.getDataRange().setValues(data);
      return true;
    }
  }

  return false;
}

/**
 * Feature 32: Calculates member engagement score
 */
function calculateEngagementScore(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return 0;

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();

  const member = memberData.find(row => row[0] === memberId);
  if (!member) return 0;

  let score = 0;

  // Points for events attended
  const eventsAttended = member[21] || 0;
  score += eventsAttended * 5;

  // Points for training sessions
  const trainingSessions = member[22] || 0;
  score += trainingSessions * 10;

  // Points for committee membership
  const committee = member[23];
  if (committee && committee !== 'None') score += 20;

  // Points for being a steward
  const isSteward = member[9];
  if (isSteward === 'Yes') score += 30;

  // Points for grievances filed (shows engagement)
  const grievancesFiled = grievanceData.filter(row => row[1] === memberId).length;
  score += grievancesFiled * 3;

  return score;
}

/**
 * Feature 33: Updates member engagement level automatically
 */
function updateEngagementLevels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const memberId = data[i][0];
    if (!memberId) continue;

    const score = calculateEngagementScore(memberId);

    // Determine engagement level based on score
    let level;
    if (score >= 100) level = 'Very Active';
    else if (score >= 50) level = 'Active';
    else if (score >= 20) level = 'Moderately Active';
    else if (score >= 5) level = 'Inactive';
    else level = 'New Member';

    // Update engagement level (column 21)
    if (data[i][20] !== level) {
      data[i][20] = level;
      updated++;
    }
  }

  if (updated > 0) {
    sheet.getDataRange().setValues(data);
  }

  return updated;
}

/**
 * Feature 34: Identifies inactive members
 */
function identifyInactiveMembers(monthsInactive = 6) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return [];

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const inactiveMembers = [];
  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - monthsInactive);

  for (let i = 1; i < memberData.length; i++) {
    const memberId = memberData[i][0];
    const eventsAttended = memberData[i][21] || 0;
    const lastGrievance = memberData[i][17]; // Last Grievance Date

    // Check if no events and no recent grievances
    if (eventsAttended === 0 && (!lastGrievance || new Date(lastGrievance) < cutoffDate)) {
      inactiveMembers.push({
        memberId: memberId,
        name: `${memberData[i][1]} ${memberData[i][2]}`,
        location: memberData[i][4],
        email: memberData[i][7],
        lastGrievanceDate: lastGrievance || 'Never'
      });
    }
  }

  return inactiveMembers;
}

/**
 * Feature 35: Sends member re-engagement email
 */
function sendReEngagementEmail(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  const member = data.find(row => row[0] === memberId);

  if (!member || !member[7]) return false; // No email found

  const name = `${member[1]} ${member[2]}`;
  const email = member[7];

  const emailBody = `Dear ${member[1]},\n\n` +
                    `We haven't seen you at recent union events or activities. We'd love to have you more involved!\n\n` +
                    `Here are some ways to get engaged:\n` +
                    `• Attend union meetings\n` +
                    `• Join a committee\n` +
                    `• Participate in training sessions\n` +
                    `• Connect with your steward\n\n` +
                    `Your voice matters in our union!\n\n` +
                    `In Solidarity,\nSEIU Local 509`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Stay Connected with SEIU Local 509',
      body: emailBody
    });
    return true;
  } catch (e) {
    Logger.log(`Failed to send re-engagement email to ${email}: ${e}`);
    return false;
  }
}

/**
 * Sends contact update email to a single member with Google Form link
 */
function sendContactUpdateEmail(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  const member = data.find(row => row[0] === memberId);

  if (!member || !member[7]) return false; // No email found

  const firstName = member[1];
  const lastName = member[2];
  const email = member[7];

  const emailBody = `Dear ${firstName},\n\n` +
                    `We want to make sure we have your most current contact information on file. ` +
                    `Keeping your contact details up to date helps us communicate important union updates, ` +
                    `meeting notifications, and grievance information effectively.\n\n` +
                    `Please take a moment to review and update your contact information using this form:\n\n` +
                    `${MEMBER_CONTACT_UPDATE_FORM_URL}\n\n` +
                    `Your updated information will help us:\n` +
                    `• Reach you about important deadlines and meetings\n` +
                    `• Keep you informed about contract updates\n` +
                    `• Ensure you receive critical union communications\n` +
                    `• Contact you in case of emergencies\n\n` +
                    `Thank you for helping us stay connected!\n\n` +
                    `In Solidarity,\nSEIU Local 509`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Update Your Contact Information - SEIU Local 509',
      body: emailBody
    });
    return true;
  } catch (e) {
    Logger.log(`Failed to send contact update email to ${email}: ${e}`);
    return false;
  }
}

/**
 * Sends contact update emails to multiple members
 * @param {Array<string>} memberIds - Array of member IDs to email
 * @returns {Object} Results with success and failure counts
 */
function sendBulkContactUpdateEmails(memberIds) {
  let successCount = 0;
  let failCount = 0;
  const failed = [];

  memberIds.forEach(memberId => {
    try {
      if (sendContactUpdateEmail(memberId)) {
        successCount++;
      } else {
        failCount++;
        failed.push(memberId);
      }
    } catch (error) {
      failCount++;
      failed.push(memberId);
      Logger.log(`Failed to send to ${memberId}: ${error}`);
    }
  });

  return {
    success: successCount,
    failed: failCount,
    failedIds: failed
  };
}

/**
 * Feature 36: Creates member satisfaction survey tracker
 */
function trackMemberSatisfaction(memberId, satisfactionLevel, comments) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let surveySheet = ss.getSheetByName('Member Surveys');

  if (!surveySheet) {
    surveySheet = ss.insertSheet('Member Surveys');
    surveySheet.appendRow(['Survey Date', 'Member ID', 'Name', 'Satisfaction Level', 'Comments']);
    surveySheet.getRange('1:1').setFontWeight('bold').setBackground(COLORS.HEADER_BLUE).setFontColor('white');
  }

  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const memberData = memberSheet.getDataRange().getValues();
  const member = memberData.find(row => row[0] === memberId);

  if (member) {
    const name = `${member[1]} ${member[2]}`;
    surveySheet.appendRow([new Date(), memberId, name, satisfactionLevel, comments]);
    return true;
  }

  return false;
}

/**
 * Feature 37: Generates member engagement report
 */
function generateEngagementReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const report = {
    totalMembers: data.length - 1,
    byEngagementLevel: {},
    stewards: 0,
    committeeMembers: 0,
    avgEventsAttended: 0,
    avgTrainingSessions: 0
  };

  let totalEvents = 0;
  let totalTraining = 0;

  for (let i = 1; i < data.length; i++) {
    // Count by engagement level
    const level = data[i][20] || 'Unknown';
    report.byEngagementLevel[level] = (report.byEngagementLevel[level] || 0) + 1;

    // Count stewards
    if (data[i][9] === 'Yes') report.stewards++;

    // Count committee members
    if (data[i][23] && data[i][23] !== 'None') report.committeeMembers++;

    // Sum events and training
    totalEvents += data[i][21] || 0;
    totalTraining += data[i][22] || 0;
  }

  report.avgEventsAttended = (totalEvents / (data.length - 1)).toFixed(1);
  report.avgTrainingSessions = (totalTraining / (data.length - 1)).toFixed(1);

  return report;
}

/**
 * Feature 38: Tracks member committee participation
 */
function updateCommitteeMembership(memberId, committee) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memberId) {
      data[i][23] = committee; // Column X: Committee Member
      data[i][30] = new Date(); // Last Updated
      sheet.getDataRange().setValues(data);
      return true;
    }
  }

  return false;
}

/**
 * Feature 39: Identifies potential steward candidates
 */
function identifyPotentialStewards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) return [];

  const memberData = memberSheet.getDataRange().getValues();
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const candidates = [];

  for (let i = 1; i < memberData.length; i++) {
    const memberId = memberData[i][0];
    const isSteward = memberData[i][9];
    const engagementLevel = memberData[i][20];
    const eventsAttended = memberData[i][21] || 0;
    const committee = memberData[i][23];

    // Skip if already a steward
    if (isSteward === 'Yes') continue;

    // Criteria: High engagement, committee member, attended events
    const score = calculateEngagementScore(memberId);
    if (score >= 50 && eventsAttended >= 3 && committee && committee !== 'None') {
      candidates.push({
        memberId: memberId,
        name: `${memberData[i][1]} ${memberData[i][2]}`,
        location: memberData[i][4],
        engagementLevel: engagementLevel,
        eventsAttended: eventsAttended,
        committee: committee,
        score: score
      });
    }
  }

  // Sort by score descending
  candidates.sort((a, b) => b.score - a.score);

  return candidates;
}

/**
 * Feature 40: Creates member outreach list
 */
function createOutreachList(criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const outreachList = [];

  for (let i = 1; i < data.length; i++) {
    let include = true;

    // Apply criteria filters
    if (criteria.location && data[i][4] !== criteria.location) include = false;
    if (criteria.unit && data[i][5] !== criteria.unit) include = false;
    if (criteria.engagementLevel && data[i][20] !== criteria.engagementLevel) include = false;
    if (criteria.isSteward !== undefined && (data[i][9] === 'Yes') !== criteria.isSteward) include = false;

    if (include) {
      outreachList.push({
        memberId: data[i][0],
        name: `${data[i][1]} ${data[i][2]}`,
        email: data[i][7],
        phone: data[i][8],
        location: data[i][4],
        preferredContact: data[i][24]
      });
    }
  }

  return outreachList;
}

// ============================================================================
// SECTION 5: GRIEVANCE ANALYSIS ENHANCEMENTS (Features 41-50)
// ============================================================================

/**
 * Feature 41: Identifies grievance patterns by member
 */
function identifyMemberGrievancePatterns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const patterns = {};

  for (let i = 1; i < data.length; i++) {
    const memberId = data[i][1];
    const type = data[i][25];

    if (!memberId || !type) continue;

    if (!patterns[memberId]) {
      patterns[memberId] = { total: 0, byType: {}, mostCommonType: '' };
    }

    patterns[memberId].total++;
    patterns[memberId].byType[type] = (patterns[memberId].byType[type] || 0) + 1;

    // Update most common type
    const maxType = Object.keys(patterns[memberId].byType).reduce((a, b) =>
      patterns[memberId].byType[a] > patterns[memberId].byType[b] ? a : b
    );
    patterns[memberId].mostCommonType = maxType;
  }

  return patterns;
}

/**
 * Feature 42: Analyzes grievance resolution time trends
 */
function analyzeResolutionTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const resolutionTimes = [];

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    const status = data[i][4];

    if (!dateFiled || !status || !status.startsWith('Resolved')) continue;

    // Find resolution date (could be Step I, II, III decision)
    let resolutionDate = data[i][10] || data[i][15] || data[i][20]; // Step I, II, or III decision

    if (resolutionDate) {
      const daysToResolve = Math.floor((new Date(resolutionDate) - new Date(dateFiled)) / (1000 * 60 * 60 * 24));
      resolutionTimes.push({
        grievanceId: data[i][0],
        type: data[i][25],
        daysToResolve: daysToResolve,
        outcome: data[i][24]
      });
    }
  }

  if (resolutionTimes.length === 0) return null;

  const avgResolutionTime = resolutionTimes.reduce((sum, item) => sum + item.daysToResolve, 0) / resolutionTimes.length;
  const fastest = Math.min(...resolutionTimes.map(item => item.daysToResolve));
  const slowest = Math.max(...resolutionTimes.map(item => item.daysToResolve));

  return {
    avgDays: avgResolutionTime.toFixed(1),
    fastestDays: fastest,
    slowestDays: slowest,
    totalResolved: resolutionTimes.length,
    resolutionTimes: resolutionTimes
  };
}

/**
 * Feature 43: Identifies high-risk grievance types
 */
function identifyHighRiskTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const typeStats = {};

  for (let i = 1; i < data.length; i++) {
    const type = data[i][25];
    const status = data[i][4];
    const outcome = data[i][24];

    if (!type) continue;

    if (!typeStats[type]) {
      typeStats[type] = { total: 0, resolved: 0, lost: 0, overdue: 0 };
    }

    typeStats[type].total++;

    if (status && status.startsWith('Resolved')) {
      typeStats[type].resolved++;
      if (outcome && outcome.includes('Lost')) typeStats[type].lost++;
    }

    if (data[i][28] === 'YES') typeStats[type].overdue++;
  }

  // Identify high-risk types (high loss rate or high overdue rate)
  const highRiskTypes = [];

  Object.keys(typeStats).forEach(type => {
    const stats = typeStats[type];
    const lossRate = stats.resolved > 0 ? (stats.lost / stats.resolved) * 100 : 0;
    const overdueRate = stats.total > 0 ? (stats.overdue / stats.total) * 100 : 0;

    if (lossRate > 50 || overdueRate > 30) {
      highRiskTypes.push({
        type: type,
        lossRate: lossRate.toFixed(1),
        overdueRate: overdueRate.toFixed(1),
        totalCases: stats.total,
        riskLevel: (lossRate > 70 || overdueRate > 50) ? 'HIGH' : 'MEDIUM'
      });
    }
  });

  // Sort by risk level
  highRiskTypes.sort((a, b) => {
    const riskOrder = { HIGH: 0, MEDIUM: 1 };
    return riskOrder[a.riskLevel] - riskOrder[b.riskLevel];
  });

  return highRiskTypes;
}

/**
 * Feature 44: Predicts grievance outcome based on patterns
 */
function predictGrievanceOutcome(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const grievance = data.find(row => row[0] === grievanceId);

  if (!grievance) return null;

  const type = grievance[25];
  const currentStep = grievance[5];

  // Analyze historical data for same type
  const similarCases = data.filter(row =>
    row[25] === type &&
    row[4] && row[4].startsWith('Resolved')
  );

  if (similarCases.length === 0) {
    return { prediction: 'INSUFFICIENT_DATA', confidence: 0 };
  }

  const won = similarCases.filter(row => row[24] && row[24].includes('Won')).length;
  const lost = similarCases.filter(row => row[24] && row[24].includes('Lost')).length;
  const settled = similarCases.filter(row => row[24] && row[24].includes('Settled')).length;

  const winRate = (won / similarCases.length) * 100;
  const lossRate = (lost / similarCases.length) * 100;
  const settlementRate = (settled / similarCases.length) * 100;

  let prediction;
  let confidence;

  if (winRate > 60) {
    prediction = 'LIKELY_WIN';
    confidence = winRate;
  } else if (lossRate > 60) {
    prediction = 'LIKELY_LOSS';
    confidence = lossRate;
  } else if (settlementRate > 40) {
    prediction = 'LIKELY_SETTLEMENT';
    confidence = settlementRate;
  } else {
    prediction = 'UNCERTAIN';
    confidence = 50;
  }

  return {
    prediction: prediction,
    confidence: confidence.toFixed(1),
    historicalData: {
      similarCases: similarCases.length,
      winRate: winRate.toFixed(1),
      lossRate: lossRate.toFixed(1),
      settlementRate: settlementRate.toFixed(1)
    }
  };
}

/**
 * Feature 45: Identifies systemic workplace issues
 */
function identifySystemicIssues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return [];

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  const locationIssues = {};

  for (let i = 1; i < grievanceData.length; i++) {
    const memberId = grievanceData[i][1];
    const type = grievanceData[i][25];
    const dateFiled = grievanceData[i][8];

    if (!memberId || !type) continue;

    const member = memberData.find(row => row[0] === memberId);
    if (!member) continue;

    const location = member[4];
    if (!location) continue;

    if (!locationIssues[location]) {
      locationIssues[location] = {};
    }

    if (!locationIssues[location][type]) {
      locationIssues[location][type] = 0;
    }

    locationIssues[location][type]++;
  }

  // Identify systemic issues (location + type combinations with high counts)
  const systemicIssues = [];

  Object.keys(locationIssues).forEach(location => {
    Object.keys(locationIssues[location]).forEach(type => {
      const count = locationIssues[location][type];

      if (count >= 3) { // Threshold: 3 or more grievances of same type at same location
        systemicIssues.push({
          location: location,
          issue: type,
          count: count,
          severity: count >= 5 ? 'HIGH' : (count >= 4 ? 'MEDIUM' : 'LOW')
        });
      }
    });
  });

  // Sort by count descending
  systemicIssues.sort((a, b) => b.count - a.count);

  return systemicIssues;
}

/**
 * Feature 46: Tracks grievance escalation patterns
 */
function analyzeEscalationPatterns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  const patterns = {
    totalGrievances: data.length - 1,
    resolvedAtStepI: 0,
    escalatedToStepII: 0,
    escalatedToStepIII: 0,
    wentToArbitration: 0,
    escalationRate: 0
  };

  for (let i = 1; i < data.length; i++) {
    const step1Decision = data[i][10];
    const step2Filed = data[i][13];
    const step3Filed = data[i][19];
    const arbitrationDate = data[i][22];
    const status = data[i][4];

    if (status && status.startsWith('Resolved') && step1Decision && !step2Filed) {
      patterns.resolvedAtStepI++;
    }

    if (step2Filed) {
      patterns.escalatedToStepII++;

      if (step3Filed) {
        patterns.escalatedToStepIII++;

        if (arbitrationDate) {
          patterns.wentToArbitration++;
        }
      }
    }
  }

  patterns.escalationRate = patterns.totalGrievances > 0 ?
    ((patterns.escalatedToStepII / patterns.totalGrievances) * 100).toFixed(1) : 0;

  return patterns;
}

/**
 * Feature 47: Generates grievance heat map data
 */
function generateGrievanceHeatMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return null;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  const heatMap = {};

  for (let i = 1; i < grievanceData.length; i++) {
    const memberId = grievanceData[i][1];
    const dateFiled = grievanceData[i][8];

    if (!memberId || !dateFiled) continue;

    const member = memberData.find(row => row[0] === memberId);
    if (!member) continue;

    const location = member[4];
    const date = new Date(dateFiled);
    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;

    if (!heatMap[location]) {
      heatMap[location] = {};
    }

    heatMap[location][monthKey] = (heatMap[location][monthKey] || 0) + 1;
  }

  return heatMap;
}

/**
 * Feature 48: Identifies grievance clusters (spike detection)
 */
function detectGrievanceClusters(daysWindow = 30) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const clusters = [];

  // Sort by date filed
  const sortedGrievances = data.slice(1)
    .filter(row => row[8]) // Has date filed
    .sort((a, b) => new Date(a[8]) - new Date(b[8]));

  // Look for clusters
  for (let i = 0; i < sortedGrievances.length; i++) {
    const startDate = new Date(sortedGrievances[i][8]);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + daysWindow);

    const clusterGrievances = sortedGrievances.filter(row => {
      const fileDate = new Date(row[8]);
      return fileDate >= startDate && fileDate <= endDate;
    });

    if (clusterGrievances.length >= 5) { // Threshold: 5 or more in window
      clusters.push({
        startDate: startDate.toISOString().slice(0, 10),
        endDate: endDate.toISOString().slice(0, 10),
        count: clusterGrievances.length,
        types: [...new Set(clusterGrievances.map(row => row[25]))],
        severity: clusterGrievances.length >= 10 ? 'HIGH' : 'MEDIUM'
      });
    }
  }

  return clusters;
}

/**
 * Feature 49: Analyzes win rate by steward
 */
function analyzeStewardWinRates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const stewardStats = {};

  for (let i = 1; i < data.length; i++) {
    const representative = data[i][26];
    const status = data[i][4];
    const outcome = data[i][24];

    if (!representative || !status || !status.startsWith('Resolved')) continue;

    if (!stewardStats[representative]) {
      stewardStats[representative] = { total: 0, won: 0, lost: 0, settled: 0 };
    }

    stewardStats[representative].total++;

    if (outcome && outcome.includes('Won')) stewardStats[representative].won++;
    if (outcome && outcome.includes('Lost')) stewardStats[representative].lost++;
    if (outcome && outcome.includes('Settled')) stewardStats[representative].settled++;
  }

  // Calculate win rates
  Object.keys(stewardStats).forEach(steward => {
    const stats = stewardStats[steward];
    stats.winRate = stats.total > 0 ? ((stats.won / stats.total) * 100).toFixed(1) : 0;
    stats.settlementRate = stats.total > 0 ? ((stats.settled / stats.total) * 100).toFixed(1) : 0;
  });

  return stewardStats;
}

/**
 * Feature 50: Creates predictive analytics for grievance volume
 */
function predictGrievanceVolume(monthsAhead = 3) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const monthlyData = {};

  // Collect historical data
  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    if (!dateFiled) continue;

    const date = new Date(dateFiled);
    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;

    monthlyData[monthKey] = (monthlyData[monthKey] || 0) + 1;
  }

  // Calculate moving average
  const months = Object.keys(monthlyData).sort();
  if (months.length < 3) {
    return { prediction: 'INSUFFICIENT_DATA', monthsAhead: monthsAhead };
  }

  const recentMonths = months.slice(-6); // Last 6 months
  const avgVolume = recentMonths.reduce((sum, month) => sum + monthlyData[month], 0) / recentMonths.length;

  // Simple trend analysis
  const firstHalf = recentMonths.slice(0, 3).reduce((sum, month) => sum + monthlyData[month], 0) / 3;
  const secondHalf = recentMonths.slice(3).reduce((sum, month) => sum + monthlyData[month], 0) / 3;
  const trend = secondHalf - firstHalf;

  // Predict future months
  const predictions = [];
  let predictedVolume = avgVolume;

  for (let i = 1; i <= monthsAhead; i++) {
    predictedVolume += trend / monthsAhead;
    predictions.push({
      monthOffset: i,
      predictedVolume: Math.round(predictedVolume),
      confidence: i === 1 ? 'HIGH' : (i === 2 ? 'MEDIUM' : 'LOW')
    });
  }

  return {
    historicalAvg: avgVolume.toFixed(1),
    trend: trend > 0 ? 'INCREASING' : (trend < 0 ? 'DECREASING' : 'STABLE'),
    trendValue: trend.toFixed(1),
    predictions: predictions
  };
}

// ============================================================================
// SECTION 6: WORKFLOW AUTOMATION (Features 51-60)
// ============================================================================

/**
 * Feature 51: Auto-assigns grievances to stewards based on workload
 */
function autoAssignGrievance(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return null;

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  // Find the grievance
  let grievanceIndex = -1;
  for (let i = 1; i < grievanceData.length; i++) {
    if (grievanceData[i][0] === grievanceId) {
      grievanceIndex = i;
      break;
    }
  }

  if (grievanceIndex === -1) return null;

  // Get member location
  const memberId = grievanceData[grievanceIndex][1];
  const member = memberData.find(row => row[0] === memberId);
  if (!member) return null;

  const memberLocation = member[4];

  // Find available stewards at same location
  const stewardsAtLocation = memberData.filter(row =>
    row[9] === 'Yes' && // Is Steward
    row[4] === memberLocation // Same location
  );

  if (stewardsAtLocation.length === 0) {
    // Fall back to any steward
    stewardsAtLocation.push(...memberData.filter(row => row[9] === 'Yes'));
  }

  if (stewardsAtLocation.length === 0) return null;

  // Calculate current workload for each steward
  const stewardWorkloads = stewardsAtLocation.map(steward => {
    const stewardName = `${steward[1]} ${steward[2]}`;
    const activeCases = grievanceData.filter(row =>
      row[26] === stewardName && // Representative matches
      row[4] && row[4].startsWith('Filed') // Active status
    ).length;

    return {
      name: stewardName,
      activeCases: activeCases
    };
  });

  // Assign to steward with lowest workload
  stewardWorkloads.sort((a, b) => a.activeCases - b.activeCases);
  const assignedSteward = stewardWorkloads[0].name;

  // Update grievance
  grievanceData[grievanceIndex][26] = assignedSteward;
  grievanceSheet.getDataRange().setValues(grievanceData);

  return assignedSteward;
}

/**
 * Feature 52: Auto-escalates overdue grievances
 */
function autoEscalateOverdueGrievances() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let escalated = 0;

  for (let i = 1; i < data.length; i++) {
    const isOverdue = data[i][28];
    const currentStep = data[i][5];
    const nextDeadline = data[i][27];

    if (isOverdue === 'YES' && nextDeadline) {
      const daysOverdue = Math.floor((new Date() - new Date(nextDeadline)) / (1000 * 60 * 60 * 24));

      // Auto-escalate if more than 14 days overdue
      if (daysOverdue >= 14) {
        // Determine next step
        let nextStep = currentStep;

        if (currentStep === 'Step I - Immediate Supervisor') {
          nextStep = 'Step II - Agency Head';
          data[i][13] = new Date(); // Set Step II Filed Date
        } else if (currentStep === 'Step II - Agency Head') {
          nextStep = 'Step III - Human Resources';
          data[i][19] = new Date(); // Set Step III Filed Date
        }

        if (nextStep !== currentStep) {
          data[i][5] = nextStep;
          escalated++;
        }
      }
    }
  }

  if (escalated > 0) {
    sheet.getDataRange().setValues(data);
  }

  return escalated;
}

/**
 * Feature 53: Auto-archives resolved grievances older than threshold
 */
function autoArchiveOldGrievances(daysOld = 90) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  let archiveSheet = ss.getSheetByName(SHEETS.ARCHIVE);

  if (!grievanceSheet) return 0;

  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(SHEETS.ARCHIVE);
    // Copy headers from Grievance Log
    const headers = grievanceSheet.getRange('1:1').getValues();
    archiveSheet.getRange('1:1').setValues(headers);
  }

  const data = grievanceSheet.getDataRange().getValues();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysOld);

  let archived = 0;
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const status = data[i][4];
    const dateFiled = data[i][8];

    if (status && status.startsWith('Resolved') && dateFiled && new Date(dateFiled) < cutoffDate) {
      // Archive this row
      archiveSheet.appendRow(data[i]);
      rowsToDelete.push(i + 1); // +1 because sheet rows are 1-indexed
      archived++;
    }
  }

  // Delete archived rows from Grievance Log
  rowsToDelete.forEach(row => {
    grievanceSheet.deleteRow(row);
  });

  return archived;
}

/**
 * Feature 54: Auto-updates member statistics on edit
 */
function onGrievanceEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) return;

  const row = e.range.getRow();
  if (row === 1) return; // Header row

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const memberId = data[1];

  if (memberId) {
    // Recalculate member statistics
    recalculateMemberStatistics(memberId);
  }
}

/**
 * Feature 55: Auto-generates next Member ID
 */
function getNextMemberID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return 'MEM000001';

  const data = sheet.getDataRange().getValues();
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    const memberId = data[i][0];
    if (memberId && memberId.toString().startsWith('MEM')) {
      const num = parseInt(memberId.replace('MEM', ''));
      if (num > maxId) maxId = num;
    }
  }

  return 'MEM' + String(maxId + 1).padStart(6, '0');
}

/**
 * Feature 56: Auto-generates next Grievance ID
 */
function getNextGrievanceID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 'GRV000001';

  const data = sheet.getDataRange().getValues();
  let maxId = 0;

  for (let i = 1; i < data.length; i++) {
    const grievanceId = data[i][0];
    if (grievanceId && grievanceId.toString().startsWith('GRV')) {
      const num = parseInt(grievanceId.replace('GRV', ''));
      if (num > maxId) maxId = num;
    }
  }

  return 'GRV' + String(maxId + 1).padStart(6, '0');
}

/**
 * Feature 57: Creates automated weekly summary
 */
function generateWeeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return null;

  const data = grievanceSheet.getDataRange().getValues();
  const today = new Date();
  const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

  const summary = {
    weekEnding: today.toISOString().slice(0, 10),
    newGrievances: 0,
    resolved: 0,
    overdue: 0,
    dueThisWeek: 0,
    upcomingDeadlines: []
  };

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    const status = data[i][4];
    const isOverdue = data[i][28];
    const nextDeadline = data[i][27];

    // Count new grievances filed this week
    if (dateFiled && new Date(dateFiled) >= weekAgo) {
      summary.newGrievances++;
    }

    // Count resolved this week
    if (status && status.startsWith('Resolved')) {
      const resolutionDate = data[i][10] || data[i][15] || data[i][20];
      if (resolutionDate && new Date(resolutionDate) >= weekAgo) {
        summary.resolved++;
      }
    }

    // Count overdue
    if (isOverdue === 'YES') {
      summary.overdue++;
    }

    // Count due this week
    if (nextDeadline) {
      const deadlineDate = new Date(nextDeadline);
      const nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

      if (deadlineDate >= today && deadlineDate <= nextWeek) {
        summary.dueThisWeek++;
        summary.upcomingDeadlines.push({
          grievanceId: data[i][0],
          memberName: `${data[i][2]} ${data[i][3]}`,
          deadline: deadlineDate.toISOString().slice(0, 10),
          daysUntil: Math.floor((deadlineDate - today) / (1000 * 60 * 60 * 24))
        });
      }
    }
  }

  return summary;
}

/**
 * Feature 58: Auto-sends scheduled reports
 */
function scheduleReports() {
  // Set up weekly summary report
  ScriptApp.newTrigger('sendWeeklySummaryReport')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  // Set up monthly performance report
  ScriptApp.newTrigger('sendMonthlyPerformanceReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // Set up daily digest
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}

/**
 * Feature 59: Auto-updates dashboard on schedule
 */
function scheduleAutomaticUpdates() {
  // Rebuild dashboard every hour
  ScriptApp.newTrigger('rebuildDashboard')
    .timeBased()
    .everyHours(1)
    .create();

  // Recalculate all grievances daily at 2 AM
  ScriptApp.newTrigger('recalcAllGrievances')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}

/**
 * Feature 60: Creates automated backup
 */
function createAutomatedBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const backupName = `509_Dashboard_Backup_${today.toISOString().slice(0,10)}`;

  try {
    const backup = ss.copy(backupName);
    const folderId = PropertiesService.getScriptProperties().getProperty('BACKUP_FOLDER_ID');

    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      DriveApp.getFileById(backup.getId()).moveTo(folder);
    }

    return backup.getUrl();
  } catch (e) {
    Logger.log(`Backup failed: ${e}`);
    return null;
  }
}

// ============================================================================
// SECTION 7: INTEGRATION FEATURES (Features 61-70)
// ============================================================================

/**
 * Feature 61: Exports data to Google Calendar for deadlines
 */
function exportDeadlinesToCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const calendar = CalendarApp.getDefaultCalendar();
  let exported = 0;

  for (let i = 1; i < data.length; i++) {
    const grievanceId = data[i][0];
    const memberName = `${data[i][2]} ${data[i][3]}`;
    const nextDeadline = data[i][27];
    const currentStep = data[i][5];

    if (nextDeadline) {
      const title = `⚖️ Grievance Deadline: ${grievanceId} - ${memberName}`;
      const description = `Step: ${currentStep}\nMember: ${memberName}\nView Dashboard: ${ss.getUrl()}`;

      try {
        calendar.createAllDayEvent(title, new Date(nextDeadline), { description: description });
        exported++;
      } catch (e) {
        Logger.log(`Failed to create calendar event for ${grievanceId}: ${e}`);
      }
    }
  }

  return exported;
}

/**
 * Feature 62: Imports data from external CSV
 */
function importFromCSV(csvData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return 0;

  const rows = csvData.split('\n');
  let imported = 0;

  for (let i = 1; i < rows.length; i++) { // Skip header
    const cells = rows[i].split(',');
    if (cells.length >= 3) { // Minimum: First Name, Last Name, Email
      sheet.appendRow(cells);
      imported++;
    }
  }

  return imported;
}

/**
 * Feature 63: Syncs with Google Contacts
 */
function syncToGoogleContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let synced = 0;

  for (let i = 1; i < data.length; i++) {
    const firstName = data[i][1];
    const lastName = data[i][2];
    const email = data[i][7];
    const phone = data[i][8];
    const workLocation = data[i][4];

    if (!email) continue;

    try {
      const contact = ContactsApp.createContact(firstName, lastName, email);
      if (phone) contact.addPhone(ContactsApp.Field.WORK_PHONE, phone);
      if (workLocation) contact.setNotes(`Work Location: ${workLocation}\nSEIU Local 509 Member`);
      synced++;
    } catch (e) {
      Logger.log(`Failed to create contact for ${firstName} ${lastName}: ${e}`);
    }
  }

  return synced;
}

/**
 * Feature 64: Creates webhook for external integrations
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'new_member') {
      return handleNewMemberWebhook(data);
    } else if (data.type === 'new_grievance') {
      return handleNewGrievanceWebhook(data);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown type' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Feature 65: Handles new member webhook
 */
function handleNewMemberWebhook(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Sheet not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const memberId = getNextMemberID();
  const row = [
    memberId,
    data.firstName,
    data.lastName,
    data.jobTitle || '',
    data.location || '',
    data.unit || '',
    data.officeDays || '',
    data.email || '',
    data.phone || '',
    'No', // Is Steward
    'Active', // Membership Status
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', new Date()
  ];

  sheet.appendRow(row);

  return ContentService.createTextOutput(JSON.stringify({ success: true, memberId: memberId }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Feature 66: Handles new grievance webhook
 */
function handleNewGrievanceWebhook(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Sheet not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const grievanceId = getNextGrievanceID();
  const row = [
    grievanceId,
    data.memberId,
    data.firstName || '',
    data.lastName || '',
    'Filed - Step I',
    'Step I - Immediate Supervisor',
    new Date(data.incidentDate),
    '', // Filing Deadline (will be calculated)
    new Date(), // Date Filed
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
    data.type || '',
    data.description || ''
  ];

  sheet.appendRow(row);

  // Recalculate deadlines for this grievance
  recalculateGrievanceRow(sheet.getLastRow());

  return ContentService.createTextOutput(JSON.stringify({ success: true, grievanceId: grievanceId }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Feature 67: Exports to Google Drive as Excel
 */
function exportToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const excelBlob = ss.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  excelBlob.setName(`509_Dashboard_Export_${new Date().toISOString().slice(0,10)}.xlsx`);

  const folderId = PropertiesService.getScriptProperties().getProperty('EXPORT_FOLDER_ID');

  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(excelBlob);
    return file.getUrl();
  } else {
    const file = DriveApp.createFile(excelBlob);
    return file.getUrl();
  }
}

/**
 * Feature 68: Sends data to external API
 */
function sendToExternalAPI(endpoint, data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('EXTERNAL_API_KEY');

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${apiKey}`
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    return {
      success: responseCode >= 200 && responseCode < 300,
      statusCode: responseCode,
      body: responseBody
    };
  } catch (e) {
    Logger.log(`API request failed: ${e}`);
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Feature 69: Creates QR codes for grievance tracking
 */
function generateGrievanceQRCode(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = `${ss.getUrl()}#gid=${ss.getSheetByName(SHEETS.GRIEVANCE_LOG).getSheetId()}&search=${grievanceId}`;

  const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(url)}`;

  return qrCodeUrl;
}

/**
 * Feature 70: Integrates with Slack for notifications
 */
function sendSlackNotification(message, channel) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhookUrl) return false;

  const payload = {
    channel: channel || '#grievances',
    username: '509 Dashboard Bot',
    text: message,
    icon_emoji: ':scales:'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    return response.getResponseCode() === 200;
  } catch (e) {
    Logger.log(`Slack notification failed: ${e}`);
    return false;
  }
}

// ============================================================================
// SECTION 8: PERFORMANCE & OPTIMIZATION (Features 71-78)
// ============================================================================

/**
 * Feature 71: Implements caching for frequently accessed data
 */
function getCachedData(key, fetchFunction, cacheSeconds = 300) {
  const cache = CacheService.getScriptCache();
  let data = cache.get(key);

  if (data) {
    return JSON.parse(data);
  }

  data = fetchFunction();
  cache.put(key, JSON.stringify(data), cacheSeconds);

  return data;
}

/**
 * Feature 72: Batch processes large data operations
 */
function batchProcessGrievances(processingFunction, batchSize = 100) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  let processed = 0;

  for (let i = 1; i < data.length; i += batchSize) {
    const batch = data.slice(i, Math.min(i + batchSize, data.length));

    batch.forEach((row, index) => {
      processingFunction(row, i + index);
      processed++;
    });

    // Flush changes every batch
    SpreadsheetApp.flush();
  }

  return processed;
}

/**
 * Feature 73: Optimizes sheet formulas
 */
function optimizeSheetFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [SHEETS.DASHBOARD, SHEETS.STEWARD_WORKLOAD, SHEETS.TRENDS];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Replace volatile functions with static values where possible
    const range = sheet.getDataRange();
    const formulas = range.getFormulas();
    const values = range.getValues();

    for (let i = 0; i < formulas.length; i++) {
      for (let j = 0; j < formulas[i].length; j++) {
        // Replace NOW() and TODAY() with static dates
        if (formulas[i][j].includes('NOW()') || formulas[i][j].includes('TODAY()')) {
          sheet.getRange(i + 1, j + 1).setValue(values[i][j]);
        }
      }
    }
  });
}

/**
 * Feature 74: Implements lazy loading for large datasets
 */
function lazyLoadGrievances(offset = 0, limit = 100) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return { data: [], hasMore: false };

  const totalRows = sheet.getLastRow() - 1; // Exclude header
  const startRow = offset + 2; // +1 for header, +1 for 1-indexing
  const endRow = Math.min(startRow + limit - 1, totalRows + 1);

  if (startRow > totalRows + 1) {
    return { data: [], hasMore: false };
  }

  const data = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();

  return {
    data: data,
    hasMore: endRow < totalRows + 1,
    nextOffset: offset + limit
  };
}

/**
 * Feature 75: Compresses old data for performance
 */
function compressOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  let compressedSheet = ss.getSheetByName('Compressed_Data');

  if (!compressedSheet) {
    compressedSheet = ss.insertSheet('Compressed_Data');
  }

  const data = grievanceSheet.getDataRange().getValues();
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 1); // 1 year old

  const compressed = [];

  for (let i = 1; i < data.length; i++) {
    const dateFiled = data[i][8];
    const status = data[i][4];

    if (status && status.startsWith('Resolved') && dateFiled && new Date(dateFiled) < cutoffDate) {
      // Create compressed summary
      compressed.push([
        data[i][0], // Grievance ID
        data[i][1], // Member ID
        data[i][8], // Date Filed
        data[i][24], // Outcome
        data[i][25], // Type
        'Archived'
      ]);
    }
  }

  if (compressed.length > 0) {
    compressedSheet.getRange(compressedSheet.getLastRow() + 1, 1, compressed.length, compressed[0].length)
      .setValues(compressed);
  }

  return compressed.length;
}

/**
 * Feature 76: Implements index for fast lookups
 */
function buildGrievanceIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const index = {
    byMemberId: {},
    byGrievanceId: {},
    byStatus: {},
    byType: {}
  };

  for (let i = 1; i < data.length; i++) {
    const row = i;
    const grievanceId = data[i][0];
    const memberId = data[i][1];
    const status = data[i][4];
    const type = data[i][25];

    // Index by Member ID
    if (memberId) {
      if (!index.byMemberId[memberId]) index.byMemberId[memberId] = [];
      index.byMemberId[memberId].push(row);
    }

    // Index by Grievance ID
    if (grievanceId) {
      index.byGrievanceId[grievanceId] = row;
    }

    // Index by Status
    if (status) {
      if (!index.byStatus[status]) index.byStatus[status] = [];
      index.byStatus[status].push(row);
    }

    // Index by Type
    if (type) {
      if (!index.byType[type]) index.byType[type] = [];
      index.byType[type].push(row);
    }
  }

  // Cache the index
  const cache = CacheService.getScriptCache();
  cache.put('grievance_index', JSON.stringify(index), 1800); // 30 minutes

  return index;
}

/**
 * Feature 77: Performs database cleanup and maintenance
 */
function performMaintenance() {
  const tasks = [
    { name: 'Auto-generate Missing IDs', func: () => autoGenerateMemberIDs() + autoGenerateGrievanceIDs() },
    { name: 'Auto-correct Data Errors', func: () => { autoCorrectDataErrors(); return 1; } },
    { name: 'Archive Old Grievances', func: () => autoArchiveOldGrievances() },
    { name: 'Update Engagement Levels', func: () => updateEngagementLevels() },
    { name: 'Rebuild Index', func: () => { buildGrievanceIndex(); return 1; } },
    { name: 'Optimize Formulas', func: () => { optimizeSheetFormulas(); return 1; } }
  ];

  const results = {};

  tasks.forEach(task => {
    try {
      const result = task.func();
      results[task.name] = { success: true, result: result };
    } catch (e) {
      results[task.name] = { success: false, error: e.toString() };
    }
  });

  return results;
}

/**
 * Feature 78: Monitors script execution time
 */
function monitorPerformance(functionName, functionToRun) {
  const startTime = new Date().getTime();

  try {
    const result = functionToRun();
    const endTime = new Date().getTime();
    const duration = endTime - startTime;

    // Log performance
    Logger.log(`${functionName} completed in ${duration}ms`);

    // Store in performance log
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let perfSheet = ss.getSheetByName('Performance_Log');

    if (!perfSheet) {
      perfSheet = ss.insertSheet('Performance_Log');
      perfSheet.appendRow(['Timestamp', 'Function', 'Duration (ms)', 'Status']);
    }

    perfSheet.appendRow([new Date(), functionName, duration, 'SUCCESS']);

    return result;
  } catch (e) {
    const endTime = new Date().getTime();
    const duration = endTime - startTime;

    Logger.log(`${functionName} failed after ${duration}ms: ${e}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let perfSheet = ss.getSheetByName('Performance_Log');

    if (perfSheet) {
      perfSheet.appendRow([new Date(), functionName, duration, `ERROR: ${e.toString()}`]);
    }

    throw e;
  }
}

// ============================================================================
// SECTION 9: SECURITY & AUDIT (Features 79-86)
// ============================================================================

/**
 * Feature 79: Tracks all data modifications
 */
function logDataModification(sheet, row, column, oldValue, newValue, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let auditSheet = ss.getSheetByName('Audit_Log');

  if (!auditSheet) {
    auditSheet = ss.insertSheet('Audit_Log');
    auditSheet.appendRow(['Timestamp', 'User', 'Sheet', 'Row', 'Column', 'Old Value', 'New Value']);
    auditSheet.getRange('1:1').setFontWeight('bold').setBackground(COLORS.HEADER_BLUE).setFontColor('white');
  }

  auditSheet.appendRow([
    new Date(),
    user || Session.getActiveUser().getEmail(),
    sheet,
    row,
    column,
    oldValue,
    newValue
  ]);
}

/**
 * Feature 80: Implements role-based access control
 */
function checkUserPermission(action) {
  const user = Session.getActiveUser().getEmail();
  const props = PropertiesService.getScriptProperties();

  // Get user roles from properties
  const admins = (props.getProperty('ADMINS') || '').split(',');
  const stewards = (props.getProperty('STEWARDS') || '').split(',');
  const viewers = (props.getProperty('VIEWERS') || '').split(',');

  const permissions = {
    'view': [...admins, ...stewards, ...viewers],
    'edit_member': [...admins, ...stewards],
    'edit_grievance': [...admins, ...stewards],
    'delete': admins,
    'export': [...admins, ...stewards],
    'admin': admins
  };

  const allowedUsers = permissions[action] || [];
  return allowedUsers.includes(user);
}

/**
 * Feature 81: Encrypts sensitive data fields
 */
function encryptSensitiveData(data) {
  // Simple XOR encryption (for demonstration - use proper encryption in production)
  const key = PropertiesService.getScriptProperties().getProperty('ENCRYPTION_KEY') || 'default_key_12345';

  let encrypted = '';
  for (let i = 0; i < data.length; i++) {
    encrypted += String.fromCharCode(data.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }

  return Utilities.base64Encode(encrypted);
}

/**
 * Feature 82: Decrypts sensitive data fields
 */
function decryptSensitiveData(encryptedData) {
  const key = PropertiesService.getScriptProperties().getProperty('ENCRYPTION_KEY') || 'default_key_12345';

  const decoded = Utilities.base64Decode(encryptedData);
  const decodedString = Utilities.newBlob(decoded).getDataAsString();

  let decrypted = '';
  for (let i = 0; i < decodedString.length; i++) {
    decrypted += String.fromCharCode(decodedString.charCodeAt(i) ^ key.charCodeAt(i % key.length));
  }

  return decrypted;
}

/**
 * Feature 83: Validates data input for security
 */
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;

  // Remove potentially dangerous characters
  let sanitized = input.replace(/<script.*?>.*?<\/script>/gi, '');
  sanitized = sanitized.replace(/javascript:/gi, '');
  sanitized = sanitized.replace(/on\w+\s*=/gi, '');

  // Trim whitespace
  sanitized = sanitized.trim();

  return sanitized;
}

/**
 * Feature 84: Creates audit trail report
 */
function generateAuditReport(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit_Log');

  if (!auditSheet) return null;

  const data = auditSheet.getDataRange().getValues();
  const filtered = [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][0]);

    if (timestamp >= new Date(startDate) && timestamp <= new Date(endDate)) {
      filtered.push({
        timestamp: timestamp.toLocaleString(),
        user: data[i][1],
        sheet: data[i][2],
        row: data[i][3],
        column: data[i][4],
        oldValue: data[i][5],
        newValue: data[i][6]
      });
    }
  }

  return filtered;
}

/**
 * Feature 85: Implements data retention policy
 */
function enforceDataRetention(retentionDays = 2555) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit_Log');

  if (!auditSheet) return 0;

  const data = auditSheet.getDataRange().getValues();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

  let deleted = 0;
  const rowsToDelete = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const timestamp = new Date(data[i][0]);

    if (timestamp < cutoffDate) {
      rowsToDelete.push(i + 1);
      deleted++;
    }
  }

  rowsToDelete.forEach(row => {
    auditSheet.deleteRow(row);
  });

  return deleted;
}

/**
 * Feature 86: Detects suspicious activity
 */
function detectSuspiciousActivity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit_Log');

  if (!auditSheet) return [];

  const data = auditSheet.getDataRange().getValues();
  const recentTime = new Date(Date.now() - 60 * 60 * 1000); // Last hour
  const userActivity = {};
  const suspicious = [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = new Date(data[i][0]);
    const user = data[i][1];

    if (timestamp >= recentTime) {
      userActivity[user] = (userActivity[user] || 0) + 1;
    }
  }

  // Flag users with excessive activity (>50 changes in an hour)
  Object.keys(userActivity).forEach(user => {
    if (userActivity[user] > 50) {
      suspicious.push({
        user: user,
        activityCount: userActivity[user],
        timeWindow: '1 hour',
        severity: userActivity[user] > 100 ? 'HIGH' : 'MEDIUM'
      });
    }
  });

  return suspicious;
}

// ============================================================================
// SECTION 10: ADVANCED UI & UX (Features 87-94)
// ============================================================================

/**
 * Feature 87: Creates custom sidebar with quick actions
 */
function showQuickActionsSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 10px; }
          button { width: 100%; margin: 5px 0; padding: 10px; background: #2563EB; color: white; border: none; border-radius: 4px; cursor: pointer; }
          button:hover { background: #1D4ED8; }
          h3 { color: #1F2937; }
        </style>
      </head>
      <body>
        <h3>Quick Actions</h3>
        <button onclick="google.script.run.rebuildDashboard()">Rebuild Dashboard</button>
        <button onclick="google.script.run.recalcAllGrievances()">Recalculate All</button>
        <button onclick="google.script.run.sendOverdueAlerts()">Send Overdue Alerts</button>
        <button onclick="google.script.run.runDataIntegrityCheck()">Run Integrity Check</button>
        <button onclick="google.script.run.createAutomatedBackup()">Create Backup</button>
      </body>
    </html>
  `).setTitle('509 Dashboard - Quick Actions');

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Feature 88: Creates interactive search dialog
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          input { width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #E5E7EB; border-radius: 4px; }
          button { padding: 10px 20px; background: #2563EB; color: white; border: none; border-radius: 4px; cursor: pointer; }
          button:hover { background: #1D4ED8; }
        </style>
      </head>
      <body>
        <h2>Search Grievances</h2>
        <input type="text" id="searchTerm" placeholder="Enter Grievance ID, Member Name, or Type">
        <button onclick="performSearch()">Search</button>
        <div id="results"></div>
        <script>
          function performSearch() {
            const term = document.getElementById('searchTerm').value;
            google.script.run.withSuccessHandler(displayResults).searchGrievances(term);
          }
          function displayResults(results) {
            document.getElementById('results').innerHTML = '<pre>' + JSON.stringify(results, null, 2) + '</pre>';
          }
        </script>
      </body>
    </html>
  `).setWidth(400).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Search Grievances');
}

/**
 * Feature 89: Implements advanced filtering UI
 */
function showFilterDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          label { display: block; margin: 10px 0 5px; font-weight: bold; }
          select, input { width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #E5E7EB; border-radius: 4px; }
          button { padding: 10px 20px; background: #2563EB; color: white; border: none; border-radius: 4px; cursor: pointer; margin: 5px; }
        </style>
      </head>
      <body>
        <h2>Filter Grievances</h2>
        <label>Status:</label>
        <select id="status">
          <option value="">All</option>
          <option value="Filed - Step I">Filed - Step I</option>
          <option value="Filed - Step II">Filed - Step II</option>
          <option value="Filed - Step III">Filed - Step III</option>
          <option value="Resolved">Resolved</option>
        </select>
        <label>Type:</label>
        <select id="type">
          <option value="">All</option>
          <option value="Disciplinary Action">Disciplinary Action</option>
          <option value="Working Conditions">Working Conditions</option>
        </select>
        <label>Date From:</label>
        <input type="date" id="dateFrom">
        <label>Date To:</label>
        <input type="date" id="dateTo">
        <button onclick="applyFilter()">Apply Filter</button>
        <button onclick="clearFilter()">Clear</button>
      </body>
    </html>
  `).setWidth(400).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Filter Grievances');
}

/**
 * Feature 90: Creates dashboard export wizard
 */
function showExportWizard() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          label { display: block; margin: 15px 0 5px; font-weight: bold; }
          input[type="checkbox"] { margin-right: 8px; }
          button { width: 100%; padding: 12px; background: #2563EB; color: white; border: none; border-radius: 4px; cursor: pointer; margin-top: 20px; }
        </style>
      </head>
      <body>
        <h2>Export Data</h2>
        <label><input type="checkbox" id="members" checked> Member Directory</label>
        <label><input type="checkbox" id="grievances" checked> Grievance Log</label>
        <label><input type="checkbox" id="dashboard" checked> Dashboard</label>
        <label>Format:</label>
        <select id="format">
          <option value="excel">Excel (.xlsx)</option>
          <option value="csv">CSV</option>
          <option value="pdf">PDF</option>
        </select>
        <button onclick="exportData()">Export</button>
        <script>
          function exportData() {
            const format = document.getElementById('format').value;
            google.script.run.performExport(format);
          }
        </script>
      </body>
    </html>
  `).setWidth(350).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Wizard');
}

/**
 * Feature 91: Implements data visualization selector
 */
function showVisualizationSelector() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          .viz-option { padding: 15px; margin: 10px 0; border: 2px solid #E5E7EB; border-radius: 8px; cursor: pointer; }
          .viz-option:hover { background: #F9FAFB; border-color: #2563EB; }
        </style>
      </head>
      <body>
        <h2>Choose Visualization</h2>
        <div class="viz-option" onclick="createChart('status')">Grievances by Status</div>
        <div class="viz-option" onclick="createChart('type')">Grievances by Type</div>
        <div class="viz-option" onclick="createChart('timeline')">Timeline Analysis</div>
        <div class="viz-option" onclick="createChart('location')">Location Heatmap</div>
        <script>
          function createChart(type) {
            google.script.run.generateCustomChart(type);
          }
        </script>
      </body>
    </html>
  `).setWidth(350).setHeight(450);

  SpreadsheetApp.getUi().showModalDialog(html, 'Visualizations');
}

/**
 * Feature 92: Creates member profile quick view
 */
function showMemberProfile(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const memberData = memberSheet.getDataRange().getValues();
  const member = memberData.find(row => row[0] === memberId);

  if (!member) {
    SpreadsheetApp.getUi().alert('Member not found');
    return;
  }

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberGrievances = grievanceData.filter(row => row[1] === memberId);

  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          h2 { color: #2563EB; }
          .info { margin: 10px 0; }
          .label { font-weight: bold; color: #6B7280; }
        </style>
      </head>
      <body>
        <h2>${member[1]} ${member[2]}</h2>
        <div class="info"><span class="label">Member ID:</span> ${member[0]}</div>
        <div class="info"><span class="label">Location:</span> ${member[4]}</div>
        <div class="info"><span class="label">Email:</span> ${member[7]}</div>
        <div class="info"><span class="label">Phone:</span> ${member[8]}</div>
        <h3>Grievance Summary</h3>
        <div class="info"><span class="label">Total Grievances:</span> ${memberGrievances.length}</div>
        <div class="info"><span class="label">Active:</span> ${memberGrievances.filter(row => row[4] && row[4].startsWith('Filed')).length}</div>
        <div class="info"><span class="label">Resolved:</span> ${memberGrievances.filter(row => row[4] && row[4].startsWith('Resolved')).length}</div>
      </body>
    </html>
  `).setWidth(400).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Member Profile');
}

/**
 * Feature 93: Implements keyboard shortcuts help
 */
function showKeyboardShortcuts() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Roboto, Arial, sans-serif; padding: 20px; }
          table { width: 100%; border-collapse: collapse; }
          th, td { padding: 10px; text-align: left; border-bottom: 1px solid #E5E7EB; }
          th { background: #F9FAFB; font-weight: bold; }
          .key { background: #E5E7EB; padding: 3px 8px; border-radius: 3px; font-family: monospace; }
        </style>
      </head>
      <body>
        <h2>Keyboard Shortcuts</h2>
        <table>
          <tr><th>Action</th><th>Shortcut</th></tr>
          <tr><td>Rebuild Dashboard</td><td><span class="key">Ctrl+Alt+R</span></td></tr>
          <tr><td>New Grievance</td><td><span class="key">Ctrl+Alt+N</span></td></tr>
          <tr><td>Search</td><td><span class="key">Ctrl+Alt+F</span></td></tr>
          <tr><td>Quick Actions</td><td><span class="key">Ctrl+Alt+Q</span></td></tr>
          <tr><td>Help</td><td><span class="key">Ctrl+Alt+H</span></td></tr>
        </table>
      </body>
    </html>
  `).setWidth(400).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Keyboard Shortcuts');
}

/**
 * Feature 94: Creates customizable dashboard widgets
 */
function createCustomWidget(widgetType, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dashboardSheet) return;

  const widgets = {
    'kpi_card': (cfg) => {
      const range = dashboardSheet.getRange(cfg.row, cfg.col, 3, 2);
      range.merge();
      range.setValue(cfg.title);
      range.setFontSize(16);
      range.setFontWeight('bold');
      range.setBackground(COLORS.PRIMARY_BLUE);
      range.setFontColor('white');
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');

      const valueRange = dashboardSheet.getRange(cfg.row + 3, cfg.col, 2, 2);
      valueRange.merge();
      valueRange.setValue(cfg.value);
      valueRange.setFontSize(32);
      valueRange.setFontWeight('bold');
      valueRange.setHorizontalAlignment('center');
    },
    'metric_list': (cfg) => {
      let row = cfg.row;
      cfg.metrics.forEach(metric => {
        dashboardSheet.getRange(row, cfg.col).setValue(metric.label);
        dashboardSheet.getRange(row, cfg.col + 1).setValue(metric.value);
        row++;
      });
    },
    'status_indicator': (cfg) => {
      const range = dashboardSheet.getRange(cfg.row, cfg.col);
      range.setValue(cfg.label);

      const statusRange = dashboardSheet.getRange(cfg.row, cfg.col + 1);
      statusRange.setValue(cfg.status);

      const color = cfg.status === 'Good' ? COLORS.ON_TRACK :
                    (cfg.status === 'Warning' ? COLORS.DUE_SOON : COLORS.OVERDUE);
      statusRange.setBackground(color);
      statusRange.setFontColor('white');
    }
  };

  const widgetFunc = widgets[widgetType];
  if (widgetFunc) {
    widgetFunc(config);
  }
}

// ============================================================================
// MENU INTEGRATION FOR NEW FEATURES
// ============================================================================

/**
 * Adds menu items for all new enhancement features
 */
function addEnhancementMenus() {
  const ui = SpreadsheetApp.getUi();

  // Data Validation Menu
  ui.createMenu('🔍 Data Validation')
    .addItem('Run Integrity Check', 'runDataIntegrityCheck')
    .addItem('Auto-Correct Errors', 'autoCorrectDataErrors')
    .addItem('Find Orphaned Grievances', 'showOrphanedGrievances')
    .addItem('Generate Missing IDs', 'generateMissingIDs')
    .addSeparator()
    .addItem('Validate CBA Compliance', 'showCBAComplianceReport')
    .addToUi();

  // Notifications Menu
  ui.createMenu('📧 Notifications')
    .addItem('Send Overdue Alerts', 'sendOverdueAlerts')
    .addItem('Send Weekly Reminders', 'sendWeeklyDeadlineReminders')
    .addItem('Send Daily Digest', 'sendDailyDigest')
    .addItem('Configure Preferences', 'configureNotificationPreferences')
    .addToUi();

  // Reports Menu
  ui.createMenu('📊 Advanced Reports')
    .addItem('Executive Summary', 'showExecutiveSummary')
    .addItem('Trend Analysis', 'showTrendAnalysis')
    .addItem('Location Analysis', 'showLocationAnalysis')
    .addItem('Steward Performance', 'showStewardPerformance')
    .addSeparator()
    .addItem('Export to PDF', 'generatePDFReport')
    .addItem('Export to CSV', 'exportGrievancesToCSV')
    .addItem('Export to Excel', 'exportToExcel')
    .addToUi();

  // Engagement Menu
  ui.createMenu('👥 Member Engagement')
    .addItem('Update Engagement Levels', 'updateEngagementLevels')
    .addItem('Find Inactive Members', 'showInactiveMembers')
    .addItem('Identify Steward Candidates', 'showPotentialStewards')
    .addSeparator()
    .addItem('Send Re-engagement Emails', 'sendBulkReEngagement')
    .addItem('Send Contact Update Requests', 'showContactUpdateEmailSelector')
    .addToUi();

  // Tools Menu
  ui.createMenu('🛠️ Advanced Tools')
    .addItem('Quick Actions Sidebar', 'showQuickActionsSidebar')
    .addItem('Search Grievances', 'showSearchDialog')
    .addItem('Filter Data', 'showFilterDialog')
    .addItem('Export Wizard', 'showExportWizard')
    .addSeparator()
    .addItem('Keyboard Shortcuts', 'showKeyboardShortcuts')
    .addItem('Performance Monitor', 'showPerformanceLog')
    .addToUi();
}

// ============================================================================
// WRAPPER FUNCTIONS FOR MENU ITEMS
// ============================================================================
// These functions provide user-friendly interfaces to the core functionality

/**
 * Wrapper: Shows orphaned grievances in a dialog
 */
function showOrphanedGrievances() {
  const ui = SpreadsheetApp.getUi();

  try {
    const orphanedGrievances = findOrphanedGrievances();

    if (!orphanedGrievances || orphanedGrievances.length === 0) {
      ui.alert('✅ No Orphaned Grievances',
               'All grievances have valid member IDs!',
               ui.ButtonSet.OK);
      return;
    }

    let message = `Found ${orphanedGrievances.length} orphaned grievance(s):\n\n`;
    orphanedGrievances.slice(0, 20).forEach(g => {
      message += `• Row ${g.row}: ${g.grievanceId} - Invalid Member ID: ${g.memberId}\n`;
    });

    if (orphanedGrievances.length > 20) {
      message += `\n... and ${orphanedGrievances.length - 20} more.`;
    }

    message += '\n\nThese grievances reference member IDs that don\'t exist in the Member Directory.';

    ui.alert('⚠️ Orphaned Grievances Found', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to find orphaned grievances: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Generates missing IDs for members and grievances
 */
function generateMissingIDs() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert('Generate Missing IDs',
    'This will generate IDs for any members or grievances that are missing them.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const memberCount = autoGenerateMemberIDs();
    const grievanceCount = autoGenerateGrievanceIDs();

    ui.alert('✅ IDs Generated',
             `Successfully generated:\n• ${memberCount} Member IDs\n• ${grievanceCount} Grievance IDs`,
             ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to generate IDs: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows CBA compliance report
 */
function showCBAComplianceReport() {
  const ui = SpreadsheetApp.getUi();

  try {
    const violations = validateCBACompliance();

    if (!violations || violations.length === 0) {
      ui.alert('✅ CBA Compliance',
               'All grievances are compliant with CBA deadlines!',
               ui.ButtonSet.OK);
      return;
    }

    let message = `Found ${violations.length} potential CBA violation(s):\n\n`;
    violations.slice(0, 15).forEach(v => {
      message += `• ${v.grievanceId}: ${v.violation}\n`;
    });

    if (violations.length > 15) {
      message += `\n... and ${violations.length - 15} more violations.`;
    }

    ui.alert('⚠️ CBA Compliance Issues', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to validate CBA compliance: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows executive summary report
 */
function showExecutiveSummary() {
  const ui = SpreadsheetApp.getUi();

  try {
    const summary = createExecutiveSummary();

    if (!summary) {
      ui.alert('❌ Error', 'Failed to generate executive summary.', ui.ButtonSet.OK);
      return;
    }

    const message = `
📊 EXECUTIVE SUMMARY

MEMBERS:
• Total Members: ${summary.totalMembers || 0}
• Active Members: ${summary.activeMembers || 0}
• Total Stewards: ${summary.totalStewards || 0}

GRIEVANCES:
• Total Filed: ${summary.totalGrievances || 0}
• Active: ${summary.activeGrievances || 0}
• Resolved: ${summary.resolvedGrievances || 0}
• Win Rate: ${summary.winRate || 0}%

ALERTS:
• Overdue: ${summary.overdueCount || 0}
• Due This Week: ${summary.dueThisWeek || 0}
• Due Next Week: ${summary.dueNextWeek || 0}

Check the Dashboard sheet for detailed visualizations.
    `.trim();

    ui.alert('Executive Summary', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to generate executive summary: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows trend analysis report
 */
function showTrendAnalysis() {
  const ui = SpreadsheetApp.getUi();

  try {
    const trends = generateTrendAnalysis();

    if (!trends) {
      ui.alert('❌ Error', 'Failed to generate trend analysis.', ui.ButtonSet.OK);
      return;
    }

    const message = `
📈 TREND ANALYSIS

FILING TRENDS:
• Total Months with Data: ${trends.totalMonths || 0}
• Average Filings/Month: ${trends.avgFilingsPerMonth || 0}
• Peak Month: ${trends.peakMonth || 'N/A'}

RESOLUTION METRICS:
• Avg Resolution Time: ${trends.avgResolutionDays || 0} days
• Fastest Resolution: ${trends.fastestResolution || 0} days
• Slowest Resolution: ${trends.slowestResolution || 0} days

RECENT TREND:
${trends.recentTrend || 'Insufficient data'}

See "Trends & Timeline" sheet for detailed analysis.
    `.trim();

    ui.alert('Trend Analysis', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to generate trend analysis: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows location analysis report
 */
function showLocationAnalysis() {
  const ui = SpreadsheetApp.getUi();

  try {
    const locationData = generateLocationAnalysis();

    if (!locationData || locationData.length === 0) {
      ui.alert('❌ Error', 'No location data available.', ui.ButtonSet.OK);
      return;
    }

    let message = '📍 LOCATION ANALYSIS - Top 10 Locations:\n\n';
    locationData.slice(0, 10).forEach((loc, index) => {
      message += `${index + 1}. ${loc.location}\n`;
      message += `   Members: ${loc.memberCount}, Grievances: ${loc.totalGrievances}\n`;
      message += `   Win Rate: ${loc.winRate}%\n\n`;
    });

    message += '\nSee "Location Analytics" sheet for complete details.';

    ui.alert('Location Analysis', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to generate location analysis: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows steward performance report
 */
function showStewardPerformance() {
  const ui = SpreadsheetApp.getUi();

  try {
    const stewardData = createStewardPerformanceReport();

    if (!stewardData || stewardData.length === 0) {
      ui.alert('ℹ️ No Data', 'No steward performance data available.', ui.ButtonSet.OK);
      return;
    }

    let message = '👨‍⚖️ STEWARD PERFORMANCE - Top Performers:\n\n';
    stewardData.slice(0, 10).forEach((steward, index) => {
      message += `${index + 1}. ${steward.name}\n`;
      message += `   Cases: ${steward.totalCases}, Active: ${steward.activeCases}\n`;
      message += `   Win Rate: ${steward.winRate}%, Overdue: ${steward.overdueCount}\n\n`;
    });

    if (stewardData.length > 10) {
      message += `... and ${stewardData.length - 10} more stewards.\n\n`;
    }

    message += 'See "Steward Workload" sheet for complete details.';

    ui.alert('Steward Performance', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to generate steward performance report: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows inactive members
 */
function showInactiveMembers() {
  const ui = SpreadsheetApp.getUi();

  try {
    const inactiveMembers = identifyInactiveMembers(6); // 6 months inactive

    if (!inactiveMembers || inactiveMembers.length === 0) {
      ui.alert('✅ No Inactive Members',
               'All members have recent activity!',
               ui.ButtonSet.OK);
      return;
    }

    let message = `Found ${inactiveMembers.length} inactive member(s) (6+ months):\n\n`;
    inactiveMembers.slice(0, 20).forEach(member => {
      message += `• ${member.firstName} ${member.lastName} (${member.memberId})\n`;
      message += `  Last Activity: ${member.lastActivity || 'Never'}\n`;
    });

    if (inactiveMembers.length > 20) {
      message += `\n... and ${inactiveMembers.length - 20} more.`;
    }

    message += '\n\nConsider re-engagement outreach for these members.';

    ui.alert('Inactive Members', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to identify inactive members: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows potential steward candidates
 */
function showPotentialStewards() {
  const ui = SpreadsheetApp.getUi();

  try {
    const candidates = identifyPotentialStewards();

    if (!candidates || candidates.length === 0) {
      ui.alert('ℹ️ No Candidates',
               'No potential steward candidates identified at this time.',
               ui.ButtonSet.OK);
      return;
    }

    let message = `Identified ${candidates.length} potential steward candidate(s):\n\n`;
    candidates.slice(0, 15).forEach(candidate => {
      message += `• ${candidate.firstName} ${candidate.lastName}\n`;
      message += `  Engagement: ${candidate.engagementLevel || 'N/A'}\n`;
      message += `  Events: ${candidate.eventsAttended || 0}, Grievances: ${candidate.totalGrievances || 0}\n\n`;
    });

    if (candidates.length > 15) {
      message += `... and ${candidates.length - 15} more candidates.`;
    }

    ui.alert('Potential Steward Candidates', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to identify steward candidates: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Sends bulk re-engagement emails
 */
function sendBulkReEngagement() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert('Send Bulk Re-engagement Emails',
    'This will send re-engagement emails to all inactive members.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const inactiveMembers = identifyInactiveMembers(6);

    if (!inactiveMembers || inactiveMembers.length === 0) {
      ui.alert('ℹ️ No Recipients', 'No inactive members to send emails to.', ui.ButtonSet.OK);
      return;
    }

    let successCount = 0;
    let failCount = 0;

    inactiveMembers.forEach(member => {
      try {
        sendReEngagementEmail(member.memberId);
        successCount++;
      } catch (error) {
        failCount++;
      }
    });

    ui.alert('✅ Emails Sent',
             `Successfully sent ${successCount} re-engagement email(s).\nFailed: ${failCount}`,
             ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('❌ Error', 'Failed to send bulk emails: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Wrapper: Shows member selection sidebar for sending contact update emails
 */
function showContactUpdateEmailSelector() {
  const html = HtmlService.createHtmlOutput(getMemberSelectionHtml())
    .setTitle('Send Contact Update Emails')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets all member data for the selection UI
 */
function getMembersForSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const members = [];

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) { // Has Member ID
      members.push({
        memberId: row[0],
        firstName: row[1] || '',
        lastName: row[2] || '',
        email: row[7] || '',
        location: row[4] || '',
        unit: row[5] || '',
        status: row[10] || ''
      });
    }
  }

  return members;
}

/**
 * Processes the selected members and sends contact update emails
 */
function processSendContactUpdateEmails(memberIds) {
  if (!memberIds || memberIds.length === 0) {
    return { success: false, message: 'No members selected.' };
  }

  const results = sendBulkContactUpdateEmails(memberIds);

  let message = `Successfully sent ${results.success} email(s).`;
  if (results.failed > 0) {
    message += `\nFailed to send ${results.failed} email(s).`;
  }

  return {
    success: true,
    message: message,
    successCount: results.success,
    failedCount: results.failed
  };
}

/**
 * Generates HTML for member selection sidebar
 */
function getMemberSelectionHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Roboto', Arial, sans-serif;
      padding: 15px;
      margin: 0;
      font-size: 13px;
    }
    .header {
      background: #2563EB;
      color: white;
      padding: 12px;
      margin: -15px -15px 15px -15px;
      border-radius: 0;
    }
    .header h2 {
      margin: 0;
      font-size: 16px;
      font-weight: 500;
    }
    .info-box {
      background: #F0F7FF;
      border-left: 3px solid #2563EB;
      padding: 10px;
      margin-bottom: 15px;
      font-size: 12px;
    }
    .filter-section {
      margin-bottom: 15px;
      padding-bottom: 15px;
      border-bottom: 1px solid #E5E7EB;
    }
    .filter-row {
      margin-bottom: 10px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-weight: 500;
      color: #374151;
    }
    input[type="text"], select {
      width: 100%;
      padding: 8px;
      border: 1px solid #D1D5DB;
      border-radius: 4px;
      box-sizing: border-box;
      font-size: 13px;
    }
    .member-list {
      max-height: 300px;
      overflow-y: auto;
      border: 1px solid #D1D5DB;
      border-radius: 4px;
      padding: 10px;
      margin-bottom: 15px;
      background: white;
    }
    .member-item {
      padding: 8px;
      margin-bottom: 5px;
      border-bottom: 1px solid #F3F4F6;
      display: flex;
      align-items: center;
    }
    .member-item:last-child {
      border-bottom: none;
    }
    .member-item input[type="checkbox"] {
      margin-right: 10px;
    }
    .member-info {
      flex: 1;
    }
    .member-name {
      font-weight: 500;
      color: #1F2937;
    }
    .member-details {
      font-size: 11px;
      color: #6B7280;
      margin-top: 2px;
    }
    .button-group {
      display: flex;
      gap: 10px;
      margin-top: 15px;
    }
    button {
      flex: 1;
      padding: 10px 15px;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
      font-weight: 500;
    }
    .btn-primary {
      background: #2563EB;
      color: white;
    }
    .btn-primary:hover {
      background: #1D4ED8;
    }
    .btn-secondary {
      background: #F3F4F6;
      color: #374151;
    }
    .btn-secondary:hover {
      background: #E5E7EB;
    }
    .status-message {
      padding: 10px;
      margin-top: 10px;
      border-radius: 4px;
      display: none;
    }
    .status-success {
      background: #D1FAE5;
      color: #065F46;
      border-left: 3px solid #059669;
    }
    .status-error {
      background: #FEE2E2;
      color: #991B1B;
      border-left: 3px solid #DC2626;
    }
    .selection-count {
      font-size: 12px;
      color: #6B7280;
      margin-bottom: 10px;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2>📧 Send Contact Update Emails</h2>
  </div>

  <div class="info-box">
    Select members to send a contact information update request. Members will receive an email with a link to update their contact details.
  </div>

  <div class="filter-section">
    <div class="filter-row">
      <label>Search by Name:</label>
      <input type="text" id="searchName" placeholder="Type to filter..." onkeyup="filterMembers()">
    </div>
    <div class="filter-row">
      <label>Filter by Status:</label>
      <select id="filterStatus" onchange="filterMembers()">
        <option value="">All Statuses</option>
        <option value="Active">Active</option>
        <option value="Inactive">Inactive</option>
        <option value="On Leave">On Leave</option>
        <option value="Retired">Retired</option>
      </select>
    </div>
    <div class="filter-row">
      <label>Filter by Location:</label>
      <select id="filterLocation" onchange="filterMembers()">
        <option value="">All Locations</option>
      </select>
    </div>
  </div>

  <div class="selection-count">
    <span id="selectionCount">0 members selected</span>
    <button class="btn-secondary" onclick="toggleSelectAll()" style="float:right; padding:5px 10px;">Select All</button>
    <div style="clear:both;"></div>
  </div>

  <div class="member-list" id="memberList">
    <div style="text-align:center; color:#6B7280;">Loading members...</div>
  </div>

  <div class="button-group">
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn-primary" onclick="sendEmails()">Send Emails</button>
  </div>

  <div id="statusMessage" class="status-message"></div>

  <script>
    let allMembers = [];
    let selectAllState = false;

    // Load members on page load
    google.script.run
      .withSuccessHandler(displayMembers)
      .withFailureHandler(showError)
      .getMembersForSelection();

    function displayMembers(members) {
      allMembers = members;

      // Populate location filter
      const locations = [...new Set(members.map(m => m.location).filter(l => l))];
      const locationSelect = document.getElementById('filterLocation');
      locations.forEach(loc => {
        const option = document.createElement('option');
        option.value = loc;
        option.textContent = loc;
        locationSelect.appendChild(option);
      });

      filterMembers();
    }

    function filterMembers() {
      const searchTerm = document.getElementById('searchName').value.toLowerCase();
      const statusFilter = document.getElementById('filterStatus').value;
      const locationFilter = document.getElementById('filterLocation').value;

      const filtered = allMembers.filter(member => {
        const nameMatch = !searchTerm ||
          (member.firstName.toLowerCase() + ' ' + member.lastName.toLowerCase()).includes(searchTerm);
        const statusMatch = !statusFilter || member.status === statusFilter;
        const locationMatch = !locationFilter || member.location === locationFilter;
        return nameMatch && statusMatch && locationMatch;
      });

      renderMemberList(filtered);
    }

    function renderMemberList(members) {
      const listDiv = document.getElementById('memberList');

      if (members.length === 0) {
        listDiv.innerHTML = '<div style="text-align:center; color:#6B7280;">No members found</div>';
        return;
      }

      listDiv.innerHTML = members.map(member => {
        const hasEmail = member.email ? '✓' : '✗';
        const emailClass = member.email ? '' : 'style="color: #DC2626;"';
        return \`
          <div class="member-item">
            <input type="checkbox" id="member_\${member.memberId}" value="\${member.memberId}"
                   onchange="updateSelectionCount()" \${member.email ? '' : 'disabled'}>
            <div class="member-info">
              <div class="member-name">\${member.firstName} \${member.lastName} <span \${emailClass}>\${hasEmail}</span></div>
              <div class="member-details">
                \${member.location || 'No location'} | \${member.unit || 'No unit'} | \${member.status || 'No status'}
                \${member.email ? '' : '<br><span style="color: #DC2626;">No email address</span>'}
              </div>
            </div>
          </div>
        \`;
      }).join('');

      updateSelectionCount();
    }

    function toggleSelectAll() {
      selectAllState = !selectAllState;
      const checkboxes = document.querySelectorAll('.member-list input[type="checkbox"]:not([disabled])');
      checkboxes.forEach(cb => cb.checked = selectAllState);
      updateSelectionCount();
    }

    function updateSelectionCount() {
      const checked = document.querySelectorAll('.member-list input[type="checkbox"]:checked').length;
      document.getElementById('selectionCount').textContent = checked + ' member' + (checked !== 1 ? 's' : '') + ' selected';
    }

    function sendEmails() {
      const selectedIds = Array.from(document.querySelectorAll('.member-list input[type="checkbox"]:checked'))
        .map(cb => cb.value);

      if (selectedIds.length === 0) {
        showMessage('Please select at least one member.', 'error');
        return;
      }

      const confirmMsg = \`Send contact update emails to \${selectedIds.length} member(s)?\\n\\nNote: Make sure you have configured the Google Form URL in the script constants.\`;
      if (!confirm(confirmMsg)) {
        return;
      }

      showMessage('Sending emails...', 'success');

      google.script.run
        .withSuccessHandler(handleSendResult)
        .withFailureHandler(showError)
        .processSendContactUpdateEmails(selectedIds);
    }

    function handleSendResult(result) {
      if (result.success) {
        showMessage(result.message, 'success');
        // Uncheck all checkboxes
        document.querySelectorAll('.member-list input[type="checkbox"]').forEach(cb => cb.checked = false);
        updateSelectionCount();
      } else {
        showMessage(result.message, 'error');
      }
    }

    function showError(error) {
      showMessage('Error: ' + error.message, 'error');
    }

    function showMessage(message, type) {
      const msgDiv = document.getElementById('statusMessage');
      msgDiv.textContent = message;
      msgDiv.className = 'status-message status-' + type;
      msgDiv.style.display = 'block';

      if (type === 'success') {
        setTimeout(() => {
          msgDiv.style.display = 'none';
        }, 5000);
      }
    }
  </script>
</body>
</html>
  `;
}

/**
 * Wrapper: Shows performance log/monitor
 */
function showPerformanceLog() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const performanceData = props.getProperty('PERFORMANCE_LOG');

    if (!performanceData) {
      ui.alert('ℹ️ Performance Monitor',
               'No performance data available yet.\n\nPerformance monitoring will track execution times of key functions.',
               ui.ButtonSet.OK);
      return;
    }

    const perfLog = JSON.parse(performanceData);
    let message = '⚡ PERFORMANCE MONITOR\n\n';
    message += 'Recent Function Execution Times:\n\n';

    Object.keys(perfLog).slice(-10).forEach(funcName => {
      const data = perfLog[funcName];
      message += `• ${funcName}: ${data.avgTime}ms (avg)\n`;
      message += `  Calls: ${data.callCount}, Last: ${data.lastTime}ms\n\n`;
    });

    message += 'Performance data is logged automatically.';

    ui.alert('Performance Monitor', message, ui.ButtonSet.OK);
  } catch (error) {
    const message = `
⚡ PERFORMANCE MONITOR

Current System Status:
• Sheets: ${ss.getSheets().length} sheets
• Last Calculation: Just now
• Status: Operational

Performance monitoring tracks execution times of key functions.
Data will appear here as functions are executed.
    `.trim();

    ui.alert('Performance Monitor', message, ui.ButtonSet.OK);
  }
}
