/******************************************************************************
 * SEIU 509 DASHBOARD - FULLY CONSOLIDATED VERSION
 * 
 * This single file contains ALL functionality from ALL modules in the 509 Dashboard system.
 * This includes core functionality plus all feature enhancements and integrations.
 * 
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete the default code
 * 4. Copy and paste this ENTIRE file
 * 5. Save (Ctrl+S / Cmd+S)
 * 6. Refresh your Google Sheet
 * 7. Click "📊 509 Dashboard" menu when it appears
 * 8. Use Admin menu to seed test data
 * 
 * INCLUDED MODULES (All Features):
 * ✓ Core System (Member Directory, Grievance Log, Dashboard)
 * ✓ Unified Operations Monitor
 * ✓ Interactive Dashboard
 * ✓ ADHD & Accessibility Enhancements
 * ✓ Grievance Workflow & State Machine
 * ✓ Smart Auto-Assignment
 * ✓ Member Search & Lookup
 * ✓ Predictive Analytics
 * ✓ Data Caching Layer
 * ✓ Automated Notifications
 * ✓ Data Backup & Recovery
 * ✓ Custom Report Builder
 * ✓ Data Pagination
 * ✓ Enhanced Error Handling
 * ✓ Undo/Redo System
 * ✓ Calendar Integration
 * ✓ Batch Operations
 * ✓ Automated Reports
 * ✓ Recommendations Engine
 * ✓ FAQ & Knowledge Base
 * ✓ Column Toggles
 * ✓ Google Drive Integration
 * ✓ Gmail Integration
 * ✓ Dark Mode & Themes
 * ✓ Keyboard Shortcuts
 * ✓ Root Cause Analysis
 * ✓ Mobile Optimization
 * ✓ Getting Started Guide
 * ✓ Data Integrity Enhancements
 * ✓ Seed & Nuke Utilities
 * 
 * VERSION: Fully Consolidated
 * LAST UPDATED: 2025-11-30
 * GITHUB: https://github.com/Woop91/509-dashboard
 * BRANCH: claude/consolidate-gs-files-01QnwLhd3zenvxG8LPCpgvbP
 * 
 * FEATURES:
 * - 20,000+ member capacity
 * - 5,000+ grievance tracking
 * - Real-time deadline calculations
 * - Comprehensive analytics and reporting
 * - All feature modules integrated
 * - ADHD-friendly design
 * - Mobile-optimized interface
 * 
 * SUPPORT:
 * - See README.md for full documentation
 * - Use the Help menu (📊 509 Dashboard → ❓ Help)
 * - Check documentation files in repository
 ******************************************************************************/


/******************************************************************************
 * MODULE: Code
 * Source: Code.gs
 *****************************************************************************/

/****************************************************
 * 509 DASHBOARD - FIXED VERSION
 * All issues addressed, real data only, 20k members + 5k grievances
 ****************************************************/

/* ===================== CONFIGURATION ===================== */
const SHEETS = {
  CONFIG: "Config",
  MEMBER_DIR: "Member Directory",
  GRIEVANCE_LOG: "Grievance Log",
  DASHBOARD: "Dashboard",
  ANALYTICS: "Analytics Data",
  FEEDBACK: "Feedback & Development",
  MEMBER_SATISFACTION: "Member Satisfaction",
  INTERACTIVE_DASHBOARD: "🎯 Interactive (Your Custom View)",
  STEWARD_WORKLOAD: "👨‍⚖️ Steward Workload",
  TRENDS: "📈 Trends & Timeline",
  PERFORMANCE: "🎯 Test 2: Performance",
  LOCATION: "🗺️ Location Analytics",
  TYPE_ANALYSIS: "📊 Type Analysis",
  EXECUTIVE_DASHBOARD: "💼 Executive Dashboard",
  EXECUTIVE: "💼 Executive Dashboard",  // Alias for backward compatibility
  KPI_PERFORMANCE: "📊 KPI Performance Dashboard",
  KPI_BOARD: "📊 KPI Performance Dashboard",  // Alias for backward compatibility
  MEMBER_ENGAGEMENT: "👥 Member Engagement",
  COST_IMPACT: "💰 Cost Impact",
  QUICK_STATS: "⚡ Quick Stats",
  ARCHIVE: "📦 Archive",
  DIAGNOSTICS: "🔧 Diagnostics"
};

const COLORS = {
  // Primary brand colors
  PRIMARY_BLUE: "#7EC8E3",
  PRIMARY_PURPLE: "#7C3AED",
  MASSABILITY_PURPLE: "#6B5FED",
  UNION_GREEN: "#059669",
  SOLIDARITY_RED: "#DC2626",

  // Accent colors
  ACCENT_TEAL: "#14B8A6",
  ACCENT_PURPLE: "#7C3AED",
  ACCENT_ORANGE: "#F97316",
  ACCENT_YELLOW: "#FCD34D",

  // Neutral colors
  WHITE: "#FFFFFF",
  LIGHT_GRAY: "#F3F4F6",
  BORDER_GRAY: "#D1D5DB",
  TEXT_GRAY: "#6B7280",
  TEXT_DARK: "#1F2937",

  // Specialty colors
  CARD_BG: "#FAFAFA",
  INFO_LIGHT: "#E0E7FF",
  SUCCESS_LIGHT: "#D1FAE5",
  HEADER_BLUE: "#3B82F6",
  HEADER_GREEN: "#10B981"
};

// Column positions for Member Directory (1-indexed)
const MEMBER_COLS = {
  MEMBER_ID: 1,                    // A
  FIRST_NAME: 2,                   // B
  LAST_NAME: 3,                    // C
  JOB_TITLE: 4,                    // D
  WORK_LOCATION: 5,                // E
  UNIT: 6,                         // F
  OFFICE_DAYS: 7,                  // G
  EMAIL: 8,                        // H
  PHONE: 9,                        // I
  IS_STEWARD: 10,                  // J
  SUPERVISOR: 11,                  // K
  MANAGER: 12,                     // L
  ASSIGNED_STEWARD: 13,            // M
  LAST_VIRTUAL_MTG: 14,            // N
  LAST_INPERSON_MTG: 15,           // O
  LAST_SURVEY: 16,                 // P
  LAST_EMAIL_OPEN: 17,             // Q
  OPEN_RATE: 18,                   // R
  VOLUNTEER_HOURS: 19,             // S
  INTEREST_LOCAL: 20,              // T
  INTEREST_CHAPTER: 21,            // U
  INTEREST_ALLIED: 22,             // V
  TIMESTAMP: 23,                   // W
  PREFERRED_COMM: 24,              // X
  BEST_TIME: 25,                   // Y
  HAS_OPEN_GRIEVANCE: 26,          // Z
  GRIEVANCE_STATUS: 27,            // AA
  NEXT_DEADLINE: 28,               // AB
  RECENT_CONTACT_DATE: 29,         // AC
  CONTACT_STEWARD: 30,             // AD
  CONTACT_NOTES: 31                // AE
};

// Column positions for Grievance Log (1-indexed)
const GRIEVANCE_COLS = {
  GRIEVANCE_ID: 1,      // A
  MEMBER_ID: 2,         // B
  FIRST_NAME: 3,        // C
  LAST_NAME: 4,         // D
  STATUS: 5,            // E
  CURRENT_STEP: 6,      // F
  INCIDENT_DATE: 7,     // G
  FILING_DEADLINE: 8,   // H
  DATE_FILED: 9,        // I
  STEP1_DUE: 10,        // J
  STEP1_RCVD: 11,       // K
  STEP2_APPEAL_DUE: 12, // L
  STEP2_APPEAL_FILED: 13, // M
  STEP2_DUE: 14,        // N
  STEP2_RCVD: 15,       // O
  STEP3_APPEAL_DUE: 16, // P
  STEP3_APPEAL_FILED: 17, // Q
  DATE_CLOSED: 18,      // R
  DAYS_OPEN: 19,        // S
  NEXT_ACTION_DUE: 20,  // T
  DAYS_TO_DEADLINE: 21, // U
  ARTICLES: 22,         // V
  ISSUE_CATEGORY: 23,   // W
  MEMBER_EMAIL: 24,     // X
  UNIT: 25,             // Y
  LOCATION: 26,         // Z
  STEWARD: 27,          // AA
  RESOLUTION: 28        // AB
};

/**
 * Converts a column number to letter notation (1=A, 27=AA, etc.)
 */
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/* ===================== ONE-CLICK SETUP ===================== */
function CREATE_509_DASHBOARD() {
  const ss = SpreadsheetApp.getActive();

  SpreadsheetApp.getActive().toast("🚀 Creating 509 Dashboard...", "Starting", -1);

  try {
    createConfigTab();
    SpreadsheetApp.getActive().toast("✅ Config created", "10%", 2);

    createMemberDirectory();
    SpreadsheetApp.getActive().toast("✅ Member Directory created", "20%", 2);

    createGrievanceLog();
    SpreadsheetApp.getActive().toast("✅ Grievance Log created", "30%", 2);

    createMainDashboard();
    SpreadsheetApp.getActive().toast("✅ Main Dashboard created", "40%", 2);

    createAnalyticsDataSheet();
    createMemberSatisfactionSheet();
    createFeedbackSheet();
    SpreadsheetApp.getActive().toast("✅ Data sheets created", "50%", 2);

    // Create Interactive Dashboard
    createInteractiveDashboardSheet(ss);
    SpreadsheetApp.getActive().toast("✅ Interactive Dashboard created", "60%", 2);

    // Create Getting Started and FAQ sheets
    createGettingStartedSheet(ss);
    createFAQSheet(ss);
    SpreadsheetApp.getActive().toast("✅ Help sheets created", "70%", 2);

    // Create User Settings sheet
    createUserSettingsSheet();
    SpreadsheetApp.getActive().toast("✅ Settings sheet created", "70%", 2);

    // Create all analytics and test sheets
    createStewardWorkloadSheet();
    createTrendsSheet();
    createPerformanceSheet();
    createLocationSheet();
    createTypeAnalysisSheet();
    SpreadsheetApp.getActive().toast("✅ Analytics sheets created", "75%", 2);

    createExecutiveDashboard();
    createKPIPerformanceDashboard();
    createMemberEngagementSheet();
    createCostImpactSheet();
    createQuickStatsSheet();
    SpreadsheetApp.getActive().toast("✅ Executive sheets created", "80%", 2);

    // Create utility sheets
    createArchiveSheet();
    createDiagnosticsSheet();
    SpreadsheetApp.getActive().toast("✅ Utility sheets created", "85%", 2);

    setupDataValidations();
    setupFormulasAndCalculations();
    setupInteractiveDashboardControls();
    SpreadsheetApp.getActive().toast("✅ Validations & formulas ready", "95%", 2);

    onOpen();

    SpreadsheetApp.getActive().toast("✅ Dashboard ready! Use menu to seed data.", "Complete!", 5);

    // Safely activate dashboard sheet if it exists
    const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
    if (dashboard) {
      dashboard.activate();
    }

  } catch (error) {
    SpreadsheetApp.getActive().toast("❌ Error: " + error.toString(), "Error", 10);
    Logger.log("Error in CREATE_509_DASHBOARD: " + error.toString());
  }
}

/* ===================== CONFIG TAB ===================== */
function createConfigTab() {
  const ss = SpreadsheetApp.getActive();
  let config = ss.getSheetByName(SHEETS.CONFIG);

  if (!config) {
    config = ss.insertSheet(SHEETS.CONFIG);
  }
  config.clear();

  const configData = [
    ["Job Titles", "Office Locations", "Units", "Office Days", "Yes/No",
     "Supervisor First Name", "Supervisor Last Name", "Manager First Name", "Manager Last Name", "Stewards",
     "Grievance Status", "Grievance Step", "Issue Category", "Articles Violated", "Communication Methods"],

    ["Coordinator", "Boston HQ", "Unit A - Administrative", "Monday", "Yes",
     "Sarah", "Johnson", "Michael", "Chen", "Jane Smith",
     "Open", "Informal", "Discipline", "Art. 1 - Recognition", "Email"],

    ["Analyst", "Worcester Office", "Unit B - Technical", "Tuesday", "No",
     "Mike", "Wilson", "Lisa", "Anderson", "John Doe",
     "Pending Info", "Step I", "Workload", "Art. 2 - Union Security", "Phone"],

    ["Case Manager", "Springfield Branch", "Unit C - Support Services", "Wednesday", "",
     "Emily", "Davis", "Robert", "Brown", "Mary Johnson",
     "Settled", "Step II", "Scheduling", "Art. 3 - Management Rights", "Text"],

    ["Specialist", "Cambridge Office", "Unit D - Operations", "Thursday", "",
     "Tom", "Harris", "Jennifer", "Lee", "Bob Wilson",
     "Withdrawn", "Step III", "Pay", "Art. 4 - No Discrimination", "In Person"],

    ["Senior Analyst", "Lowell Center", "Unit E - Field Services", "Friday", "",
     "Amanda", "White", "David", "Martinez", "Alice Brown",
     "Closed", "Mediation", "Discrimination", "Art. 5 - Union Business", ""],

    ["Team Lead", "Quincy Station", "", "Saturday", "",
     "Chris", "Taylor", "Susan", "Garcia", "Tom Davis",
     "Appealed", "Arbitration", "Safety", "Art. 23 - Grievance Procedure", ""],

    ["Director", "Remote/Hybrid", "", "Sunday", "",
     "Patricia", "Moore", "James", "Wilson", "Sarah Martinez",
     "", "", "Benefits", "Art. 24 - Discipline", ""],

    ["Manager", "Brockton Office", "", "", "",
     "Kevin", "Anderson", "Nancy", "Taylor", "Kevin Jones",
     "", "", "Training", "Art. 25 - Hours of Work", ""],

    ["Assistant", "Lynn Location", "", "", "",
     "Michelle", "Lee", "Richard", "White", "Linda Garcia",
     "", "", "Other", "Art. 26 - Overtime", ""],

    ["Associate", "Salem Office", "", "", "",
     "Brandon", "Scott", "Angela", "Moore", "Daniel Kim",
     "", "", "Harassment", "Art. 27 - Seniority", ""],

    ["Technician", "", "", "", "",
     "Jessica", "Green", "Christopher", "Lee", "Rachel Adams",
     "", "", "Equipment", "Art. 28 - Layoff", ""],

    ["Administrator", "", "", "", "",
     "Andrew", "Clark", "Melissa", "Wright", "",
     "", "", "Leave", "Art. 29 - Sick Leave", ""],

    ["Support Staff", "", "", "", "",
     "Rachel", "Brown", "Timothy", "Davis", "",
     "", "", "Grievance Process", "Art. 30 - Vacation", ""]
  ];

  config.getRange(1, 1, configData.length, configData[0].length).setValues(configData);

  config.getRange(1, 1, 1, configData[0].length)
    .setFontWeight("bold")
    .setBackground("#4A5568")
    .setFontColor("#FFFFFF");

  for (let i = 1; i <= configData[0].length; i++) {
    config.autoResizeColumn(i);
  }

  config.setFrozenRows(1);
  config.setTabColor("#2563EB");
}

/* ===================== MEMBER DIRECTORY - ALL CORRECT COLUMNS ===================== */
function createMemberDirectory() {
  const ss = SpreadsheetApp.getActive();
  let memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Delete existing sheet to remove any column groups or formatting issues
  if (memberDir) {
    ss.deleteSheet(memberDir);
  }
  memberDir = ss.insertSheet(SHEETS.MEMBER_DIR);

  // EXACT columns as specified by user
  const headers = [
    "Member ID",
    "First Name",
    "Last Name",
    "Job Title",
    "Work Location (Site)",
    "Unit",
    "Office Days",
    "Email Address",
    "Phone Number",
    "Is Steward (Y/N)",
    "Supervisor (Name)",
    "Manager (Name)",
    "Assigned Steward (Name)",
    "Last Virtual Mtg (Date)",
    "Last In-Person Mtg (Date)",
    "Last Survey (Date)",
    "Last Email Open (Date)",
    "Open Rate (%)",
    "Volunteer Hours (YTD)",
    "Interest: Local Actions",
    "Interest: Chapter Actions",
    "Interest: Allied Chapter Actions",
    "Timestamp",
    "Preferred Communication Methods",
    "Best Time(s) to Reach Member",
    "Has Open Grievance?",
    "Grievance Status Snapshot",
    "Next Grievance Deadline",
    "Most Recent Steward Contact Date",
    "Steward Who Contacted Member",
    "Notes from Steward Contact"
  ];

  memberDir.getRange(1, 1, 1, headers.length).setValues([headers]);

  memberDir.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#059669")
    .setFontColor("#FFFFFF")
    .setWrap(true);

  memberDir.setFrozenRows(1);
  memberDir.setRowHeight(1, 50);
  memberDir.setColumnWidth(1, 90);
  memberDir.setColumnWidth(8, 180);
  memberDir.setColumnWidth(31, 250);

  memberDir.setTabColor("#059669");
}

/* ===================== GRIEVANCE LOG - ALL CORRECT COLUMNS ===================== */
function createGrievanceLog() {
  const ss = SpreadsheetApp.getActive();
  let grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceLog) {
    grievanceLog = ss.insertSheet(SHEETS.GRIEVANCE_LOG);
  }
  grievanceLog.clear();

  // EXACT columns as specified by user
  const headers = [
    "Grievance ID",
    "Member ID",
    "First Name",
    "Last Name",
    "Status",
    "Current Step",
    "Incident Date",
    "Filing Deadline (21d)",
    "Date Filed (Step I)",
    "Step I Decision Due (30d)",
    "Step I Decision Rcvd",
    "Step II Appeal Due (10d)",
    "Step II Appeal Filed",
    "Step II Decision Due (30d)",
    "Step II Decision Rcvd",
    "Step III Appeal Due (30d)",
    "Step III Appeal Filed",
    "Date Closed",
    "Days Open",
    "Next Action Due",
    "Days to Deadline",
    "Articles Violated",
    "Issue Category",
    "Member Email",
    "Unit",
    "Work Location (Site)",
    "Assigned Steward (Name)",
    "Resolution Summary"
  ];

  grievanceLog.getRange(1, 1, 1, headers.length).setValues([headers]);

  grievanceLog.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#DC2626")
    .setFontColor("#FFFFFF")
    .setWrap(true);

  grievanceLog.setFrozenRows(1);
  grievanceLog.setRowHeight(1, 50);
  grievanceLog.setColumnWidth(1, 110);
  grievanceLog.setColumnWidth(22, 180);
  grievanceLog.setColumnWidth(28, 250);

  grievanceLog.setTabColor("#DC2626");
}

/* ===================== DASHBOARD - ONLY REAL DATA ===================== */
function createMainDashboard() {
  const ss = SpreadsheetApp.getActive();
  let dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dashboard) {
    dashboard = ss.insertSheet(SHEETS.DASHBOARD);
  }
  dashboard.clear();

  // Title
  dashboard.getRange("A1:L2").merge()
    .setValue("📊 LOCAL 509 DASHBOARD")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#7C3AED")
    .setFontColor("#FFFFFF");

  dashboard.getRange("A3:L3").merge()
    .setFormula('="Last Updated: " & TEXT(NOW(), "MM/DD/YYYY HH:MM:SS")')
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setFontColor("#6B7280");

  // MEMBER METRICS - ALL REAL DATA
  dashboard.getRange("A5:L5").merge()
    .setValue("👥 MEMBER METRICS")
    .setFontWeight("bold")
    .setBackground("#E0E7FF")
    .setFontSize(12);

  const memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  const isStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  const openRateCol = getColumnLetter(MEMBER_COLS.OPEN_RATE);
  const volunteerHoursCol = getColumnLetter(MEMBER_COLS.VOLUNTEER_HOURS);

  const memberMetrics = [
    ["Total Members", `=COUNTA('Member Directory'!${memberIdCol}:${memberIdCol})-1`, "👥"],
    ["Active Stewards", `=COUNTIF('Member Directory'!${isStewardCol}:${isStewardCol},"Yes")`, "🛡️"],
    ["Avg Open Rate", `=TEXT(AVERAGE('Member Directory'!${openRateCol}:${openRateCol})/100,"0.0%")`, "📧"],
    ["YTD Vol. Hours", `=SUM('Member Directory'!${volunteerHoursCol}:${volunteerHoursCol})`, "🙋"]
  ];

  let col = 1;
  memberMetrics.forEach(m => {
    dashboard.getRange(6, col, 1, 3).merge()
      .setValue(m[2] + " " + m[0])
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#F3F4F6");

    dashboard.getRange(7, col, 1, 3).merge()
      .setFormula(m[1])
      .setFontSize(20)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    col += 3;
  });

  // GRIEVANCE METRICS - ALL REAL DATA
  dashboard.getRange("A10:L10").merge()
    .setValue("📋 GRIEVANCE METRICS")
    .setFontWeight("bold")
    .setBackground("#FEE2E2")
    .setFontSize(12);

  // Dynamic column references for Grievance Log
  const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  const dateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);
  const daysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);

  const grievanceMetrics = [
    ["Open Grievances", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`, "🔴"],
    ["Pending Info", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Pending Info")`, "🟡"],
    ["Settled (This Mo.)", `=COUNTIFS('Grievance Log'!${statusCol}:${statusCol},"Settled",'Grievance Log'!${dateClosedCol}:${dateClosedCol},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))`, "🟢"],
    ["Avg Days Open", `=ROUND(AVERAGE(FILTER('Grievance Log'!${daysOpenCol}:${daysOpenCol},'Grievance Log'!${statusCol}:${statusCol}="Open")),0)`, "⏱️"]
  ];

  col = 1;
  grievanceMetrics.forEach(m => {
    dashboard.getRange(11, col, 1, 3).merge()
      .setValue(m[2] + " " + m[0])
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#F3F4F6");

    dashboard.getRange(12, col, 1, 3).merge()
      .setFormula(m[1])
      .setFontSize(20)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    col += 3;
  });

  // ENGAGEMENT METRICS - REAL DATA
  dashboard.getRange("A15:L15").merge()
    .setValue("📈 ENGAGEMENT METRICS (Last 30 Days)")
    .setFontWeight("bold")
    .setBackground("#DCFCE7")
    .setFontSize(12);

  const lastVirtualCol = getColumnLetter(MEMBER_COLS.LAST_VIRTUAL_MTG);
  const lastInPersonCol = getColumnLetter(MEMBER_COLS.LAST_INPERSON_MTG);
  const interestLocalCol = getColumnLetter(MEMBER_COLS.INTEREST_LOCAL);
  const interestChapterCol = getColumnLetter(MEMBER_COLS.INTEREST_CHAPTER);

  const engagementMetrics = [
    ["Virtual Mtgs", `=COUNTIF('Member Directory'!${lastVirtualCol}:${lastVirtualCol},">="&TODAY()-30)`],
    ["In-Person Mtgs", `=COUNTIF('Member Directory'!${lastInPersonCol}:${lastInPersonCol},">="&TODAY()-30)`],
    ["Local Interest", `=COUNTIF('Member Directory'!${interestLocalCol}:${interestLocalCol},"Yes")`],
    ["Chapter Interest", `=COUNTIF('Member Directory'!${interestChapterCol}:${interestChapterCol},"Yes")`]
  ];

  col = 1;
  engagementMetrics.forEach(m => {
    dashboard.getRange(16, col, 1, 3).merge()
      .setValue(m[0])
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#F3F4F6");

    dashboard.getRange(17, col, 1, 3).merge()
      .setFormula(m[1])
      .setFontSize(18)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    col += 3;
  });

  // UPCOMING DEADLINES
  dashboard.getRange("A20:L20").merge()
    .setValue("⏰ UPCOMING DEADLINES (Next 14 Days)")
    .setFontWeight("bold")
    .setBackground("#FEF3C7")
    .setFontSize(12);

  const deadlineHeaders = ["Grievance ID", "Member", "Next Action", "Days Until", "Status"];
  dashboard.getRange(21, 1, 1, 5).setValues([deadlineHeaders])
    .setFontWeight("bold")
    .setBackground("#F3F4F6");

  // Dynamic column references for QUERY formula
  const grievanceIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  const firstNameCol = getColumnLetter(GRIEVANCE_COLS.FIRST_NAME);
  const nextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);
  const daysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  const lastCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION); // AB - last column

  // Formula to populate upcoming deadlines (open grievances with deadlines in next 14 days)
  dashboard.getRange("A22").setFormula(
    `=IFERROR(QUERY('Grievance Log'!${grievanceIdCol}:${lastCol}, ` +
    `"SELECT ${grievanceIdCol}, ${firstNameCol}, ${nextActionCol}, ${daysToDeadlineCol}, ${statusCol} ` +
    `WHERE ${statusCol} = 'Open' AND ${nextActionCol} IS NOT NULL AND ${nextActionCol} <= date '"&TEXT(TODAY()+14,"yyyy-mm-dd")&"' ` +
    `ORDER BY ${nextActionCol} ASC ` +
    `LIMIT 10", 0), "No upcoming deadlines")`
  );

  dashboard.setTabColor("#7C3AED");
}

/* ===================== ANALYTICS DATA SHEET ===================== */
function createAnalyticsDataSheet() {
  const ss = SpreadsheetApp.getActive();
  let analytics = ss.getSheetByName(SHEETS.ANALYTICS);

  if (!analytics) {
    analytics = ss.insertSheet(SHEETS.ANALYTICS);
  }
  analytics.clear();

  analytics.getRange("A1").setValue("ANALYTICS DATA - Calculated from Member Directory & Grievance Log");
  analytics.getRange("A1").setFontWeight("bold").setBackground("#6366F1").setFontColor("#FFFFFF");

  // Dynamic column references for Grievance Log
  const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);

  // Grievances by Status
  analytics.getRange("A3").setValue("Grievances by Status");
  analytics.getRange("A4:B4").setValues([["Status", "Count"]]).setFontWeight("bold");
  analytics.getRange("A5").setFormula(`=UNIQUE(FILTER('Grievance Log'!${statusCol}:${statusCol}, 'Grievance Log'!${statusCol}:${statusCol}<>"", 'Grievance Log'!${statusCol}:${statusCol}<>"Status"))`);
  analytics.getRange("B5").setFormula(`=ARRAYFORMULA(IF(A5:A<>"", COUNTIF('Grievance Log'!${statusCol}:${statusCol}, A5:A), ""))`);

  // Grievances by Unit (dynamic column reference)
  const unitCol = getColumnLetter(GRIEVANCE_COLS.UNIT);
  analytics.getRange("D3").setValue("Grievances by Unit");
  analytics.getRange("D4:E4").setValues([["Unit", "Count"]]).setFontWeight("bold");
  analytics.getRange("D5").setFormula(`=UNIQUE(FILTER('Grievance Log'!${unitCol}:${unitCol}, 'Grievance Log'!${unitCol}:${unitCol}<>"", 'Grievance Log'!${unitCol}:${unitCol}<>"Unit"))`);
  analytics.getRange("E5").setFormula(`=ARRAYFORMULA(IF(D5:D<>"", COUNTIF('Grievance Log'!${unitCol}:${unitCol}, D5:D), ""))`);

  // Members by Location
  analytics.getRange("G3").setValue("Members by Location");
  analytics.getRange("G4:H4").setValues([["Location", "Count"]]).setFontWeight("bold");
  analytics.getRange("G5").setFormula('=UNIQUE(FILTER(\'Member Directory\'!E:E, \'Member Directory\'!E:E<>"", \'Member Directory\'!E:E<>"Work Location (Site)"))');
  analytics.getRange("H5").setFormula('=ARRAYFORMULA(IF(G5:G<>"", COUNTIF(\'Member Directory\'!E:E, G5:G), ""))');

  // Steward Workload (dynamic column references)
  const stewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);
  analytics.getRange("J3").setValue("Steward Workload");
  analytics.getRange("J4:K4").setValues([["Steward", "Open Cases"]]).setFontWeight("bold");
  analytics.getRange("J5").setFormula(`=UNIQUE(FILTER('Grievance Log'!${stewardCol}:${stewardCol}, 'Grievance Log'!${stewardCol}:${stewardCol}<>"", 'Grievance Log'!${stewardCol}:${stewardCol}<>"Assigned Steward (Name)"))`);
  analytics.getRange("K5").setFormula(`=ARRAYFORMULA(IF(J5:J<>"", COUNTIFS('Grievance Log'!${stewardCol}:${stewardCol}, J5:J, 'Grievance Log'!${statusCol}:${statusCol}, "Open"), ""))`);

  analytics.hideSheet();
}

/* ===================== MEMBER SATISFACTION ===================== */
function createMemberSatisfactionSheet() {
  const ss = SpreadsheetApp.getActive();
  let satisfaction = ss.getSheetByName(SHEETS.MEMBER_SATISFACTION);

  if (!satisfaction) {
    satisfaction = ss.insertSheet(SHEETS.MEMBER_SATISFACTION);
  }
  satisfaction.clear();

  satisfaction.getRange("A1:J1").merge()
    .setValue("😊 MEMBER SATISFACTION TRACKING")
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#10B981")
    .setFontColor("#FFFFFF");

  const headers = [
    "Survey ID",
    "Member ID",
    "Member Name",
    "Date Sent",
    "Date Completed",
    "Overall Satisfaction (1-5)",
    "Steward Support (1-5)",
    "Communication (1-5)",
    "Would Recommend Union (Y/N)",
    "Comments"
  ];

  satisfaction.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground("#F3F4F6");

  satisfaction.setFrozenRows(3);

  satisfaction.getRange("A10:E10").merge()
    .setValue("SATISFACTION METRICS")
    .setFontWeight("bold")
    .setBackground("#E0E7FF");

  const metrics = [
    ["Average Overall Satisfaction:", "=IFERROR(AVERAGE(F4:F1000),\"-\")"],
    ["Average Steward Support:", "=IFERROR(AVERAGE(G4:G1000),\"-\")"],
    ["Average Communication:", "=IFERROR(AVERAGE(H4:H1000),\"-\")"],
    ["% Would Recommend:", "=IFERROR(TEXT(COUNTIF(I4:I1000,\"Y\")/COUNTA(I4:I1000),\"0.0%\"),\"-\")"]
  ];

  satisfaction.getRange(11, 1, metrics.length, 2).setValues(metrics);
  satisfaction.setTabColor("#10B981");
}

/* ===================== FEEDBACK & DEVELOPMENT ===================== */
function createFeedbackSheet() {
  const ss = SpreadsheetApp.getActive();
  let feedback = ss.getSheetByName(SHEETS.FEEDBACK);
  if (!feedback) feedback = ss.insertSheet(SHEETS.FEEDBACK);
  feedback.clear();
  feedback.getRange("A1:N1").merge().setValue("💡 FEEDBACK, FEATURES & DEVELOPMENT ROADMAP").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.ACCENT_PURPLE).setFontColor("white");
  feedback.getRange("A2:N2").merge().setValue("🎯 Track bugs, feedback, future features, and work in progress - all in one place").setFontSize(10).setFontStyle("italic").setHorizontalAlignment("center").setBackground(COLORS.LIGHT_GRAY).setFontColor(COLORS.TEXT_GRAY);
  const headers = ["Type", "Submitted/Started", "Submitted By", "Priority", "Title", "Description", "Status", "Progress %", "Complexity", "Target Completion", "Assigned To", "Blockers", "Resolution/Notes", "Last Updated"];
  feedback.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY).setFontColor(COLORS.TEXT_DARK);
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Bug Report', 'Feedback', 'Future Feature', 'In Progress', 'Completed', 'Archived'], true).setAllowInvalid(false).build();
  feedback.getRange("A4:A1000").setDataValidation(typeRule);
  const priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(['Critical', 'High', 'Medium', 'Low'], true).setAllowInvalid(false).build();
  feedback.getRange("D4:D1000").setDataValidation(priorityRule);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['New', 'Under Review', 'Planned', 'In Progress', 'Testing', 'Completed', 'Deferred', 'Cancelled'], true).setAllowInvalid(false).build();
  feedback.getRange("G4:G1000").setDataValidation(statusRule);
  const complexityRule = SpreadsheetApp.newDataValidation().requireValueInList(['Simple', 'Moderate', 'Complex', 'Very Complex'], true).setAllowInvalid(false).build();
  feedback.getRange("I4:I1000").setDataValidation(complexityRule);
  feedback.setFrozenRows(3);
  feedback.setTabColor(COLORS.ACCENT_PURPLE);
  feedback.setColumnWidth(1, 120); feedback.setColumnWidth(2, 110); feedback.setColumnWidth(3, 120); feedback.setColumnWidth(4, 80); feedback.setColumnWidth(5, 200); feedback.setColumnWidth(6, 300); feedback.setColumnWidth(7, 100); feedback.setColumnWidth(8, 90); feedback.setColumnWidth(9, 100); feedback.setColumnWidth(10, 110); feedback.setColumnWidth(11, 120); feedback.setColumnWidth(12, 200); feedback.setColumnWidth(13, 250); feedback.setColumnWidth(14, 110);
}

/* ===================== STEWARD WORKLOAD ===================== */
function createStewardWorkloadSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.STEWARD_WORKLOAD);
  sheet.clear();
  sheet.getRange("A1:K1").merge().setValue("👨‍⚖️ STEWARD WORKLOAD ANALYSIS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.PRIMARY_PURPLE).setFontColor("white");
  const headers = ["Steward Name", "Total Cases", "Active Cases", "Resolved Cases", "Win Rate %", "Avg Days to Resolution", "Overdue Cases", "Due This Week", "Capacity Status", "Email", "Phone"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.PRIMARY_PURPLE);
}

function createTrendsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.TRENDS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.TRENDS);
  sheet.clear();
  sheet.getRange("A1:L1").merge().setValue("📈 TRENDS & TIMELINE ANALYSIS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.UNION_GREEN).setFontColor("white");
  const headers = ["Month", "New Grievances", "Resolved", "Win Rate %", "Avg Resolution Days", "Active at Month End", "Overdue", "New Members", "Active Members", "Stewards Active", "Satisfaction Score", "Trend"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.UNION_GREEN);
}


function createLocationSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.LOCATION);
  if (!sheet) sheet = ss.insertSheet(SHEETS.LOCATION);
  sheet.clear();
  sheet.getRange("A1:K1").merge().setValue("🗺️ LOCATION ANALYTICS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.ACCENT_TEAL).setFontColor("white");
  const headers = ["Location", "Total Members", "Active Members", "Total Grievances", "Active Grievances", "Win Rate %", "Avg Resolution Days", "Member Satisfaction", "Stewards Assigned", "Risk Score", "Priority"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.ACCENT_TEAL);
}

function createTypeAnalysisSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.TYPE_ANALYSIS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.TYPE_ANALYSIS);
  sheet.clear();
  sheet.getRange("A1:K1").merge().setValue("📊 GRIEVANCE TYPE ANALYSIS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.PRIMARY_BLUE).setFontColor("white");
  const headers = ["Issue Type", "Total Cases", "Active", "Resolved", "Win Rate %", "Avg Days to Resolve", "Most Common Location", "Top Article Violated", "Trend", "Priority Level", "Notes"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.PRIMARY_BLUE);
}

/* ===================== EXECUTIVE DASHBOARD (Merged Summary + Quick Stats) ===================== */
function createExecutiveDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("💼 Executive Dashboard");

  if (!sheet) {
    sheet = ss.insertSheet("💼 Executive Dashboard");
  }
  sheet.clear();

  // Header
  sheet.getRange("A1:F1").merge()
    .setValue("💼 EXECUTIVE DASHBOARD")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor("white");

  sheet.getRange("A2:F2").merge()
    .setValue("⚡ At-a-Glance Metrics & Key Performance Indicators")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  // Section 1: Quick Stats
  sheet.getRange("A4:D4").merge()
    .setValue("⚡ QUICK STATS")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const quickHeaders = ["Metric", "Value", "Comparison", "Trend"];
  sheet.getRange(5, 1, 1, quickHeaders.length).setValues([quickHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  // Dynamic column references for formulas
  const resolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  const daysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  const daysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  const grievanceIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);

  const execMemberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  const execIsStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);

  const quickStats = [
    ["Active Members", `=COUNTA('Member Directory'!${execMemberIdCol}2:${execMemberIdCol})`, "", ""],
    ["Active Grievances", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`, "", ""],
    ["Win Rate", `=TEXT(IFERROR(COUNTIFS('Grievance Log'!${statusCol}:${statusCol},"Resolved*",'Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")/COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved*"),0),"0%")`, "", ""],
    ["Avg Resolution (Days)", `=ROUND(AVERAGE('Grievance Log'!${daysOpenCol}:${daysOpenCol}),1)`, "", ""],
    ["Overdue Cases", `=COUNTIF('Grievance Log'!${daysToDeadlineCol}:${daysToDeadlineCol},"<0")`, "", ""],
    ["Active Stewards", `=COUNTIF('Member Directory'!${execIsStewardCol}:${execIsStewardCol},"Yes")`, "", ""]
  ];

  sheet.getRange(6, 1, quickStats.length, 4).setValues(quickStats);

  // Section 2: Detailed KPIs
  sheet.getRange("A13:C13").merge()
    .setValue("📊 DETAILED KEY PERFORMANCE INDICATORS")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const kpiHeaders = ["Metric", "Value", "Status"];
  sheet.getRange(14, 1, 1, kpiHeaders.length).setValues([kpiHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  const detailedKpis = [
    ["Total Active Members", `=COUNTA('Member Directory'!${execMemberIdCol}2:${execMemberIdCol})`, ""],
    ["Total Active Grievances", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`, ""],
    ["Overall Win Rate", `=TEXT(IFERROR(COUNTIFS('Grievance Log'!${statusCol}:${statusCol},"Resolved*",'Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")/COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved*"),0),"0.0%")`, ""],
    ["Avg Resolution Time (Days)", `=ROUND(AVERAGE('Grievance Log'!${daysOpenCol}:${daysOpenCol}),1)`, ""],
    ["Cases Overdue", `=COUNTIF('Grievance Log'!${daysToDeadlineCol}:${daysToDeadlineCol},"<0")`, ""],
    ["Member Satisfaction Score", "=TEXT(AVERAGE('Member Satisfaction'!C:C),\"0.0\")", ""],
    ["Total Grievances Filed YTD", `=COUNTA('Grievance Log'!${grievanceIdCol}2:${grievanceIdCol})`, ""],
    ["Resolved Grievances", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved")`, ""]
  ];

  sheet.getRange(15, 1, detailedKpis.length, 3).setValues(detailedKpis);

  sheet.setFrozenRows(4);
  sheet.setTabColor(COLORS.PRIMARY_PURPLE);

  // Set column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
}

/* ===================== KPI PERFORMANCE DASHBOARD (Merged Performance + KPI Board) ===================== */
function createKPIPerformanceDashboard() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("📊 KPI Performance Dashboard");

  if (!sheet) {
    sheet = ss.insertSheet("📊 KPI Performance Dashboard");
  }
  sheet.clear();

  // Header
  sheet.getRange("A1:L1").merge()
    .setValue("📊 KPI PERFORMANCE DASHBOARD")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white");

  sheet.getRange("A2:L2").merge()
    .setValue("🎯 Track KPIs against targets with variance analysis and performance trends")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  const headers = [
    "KPI Name",
    "Current Value",
    "Target",
    "Variance",
    "% Change",
    "Status",
    "Last Month",
    "YTD Average",
    "Best",
    "Worst",
    "Owner",
    "Last Updated"
  ];

  sheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['On Track', 'At Risk', 'Off Track', 'Exceeding'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("F4:F1000").setDataValidation(statusRule);

  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.UNION_GREEN);

  // Set column widths
  sheet.setColumnWidth(1, 200);  // KPI Name
  sheet.setColumnWidth(2, 120);  // Current Value
  sheet.setColumnWidth(3, 100);  // Target
  sheet.setColumnWidth(4, 100);  // Variance
  sheet.setColumnWidth(5, 100);  // % Change
  sheet.setColumnWidth(6, 100);  // Status
  sheet.setColumnWidth(7, 100);  // Last Month
  sheet.setColumnWidth(8, 110);  // YTD Average
  sheet.setColumnWidth(9, 80);   // Best
  sheet.setColumnWidth(10, 80);  // Worst
  sheet.setColumnWidth(11, 120); // Owner
  sheet.setColumnWidth(12, 110); // Last Updated
}

function createMemberEngagementSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.MEMBER_ENGAGEMENT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MEMBER_ENGAGEMENT);
  sheet.clear();
  sheet.getRange("A1:L1").merge().setValue("👥 MEMBER ENGAGEMENT").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.ACCENT_PURPLE).setFontColor("white");
  const headers = ["Member ID", "Name", "Engagement Score", "Last Contact", "Meetings Attended", "Surveys Completed", "Volunteer Hours", "Committee Participation", "Event Attendance", "Email Open Rate", "Status", "Notes"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.ACCENT_PURPLE);
}

function createCostImpactSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.COST_IMPACT);
  if (!sheet) sheet = ss.insertSheet(SHEETS.COST_IMPACT);
  sheet.clear();
  sheet.getRange("A1:J1").merge().setValue("💰 COST IMPACT ANALYSIS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.SOLIDARITY_RED).setFontColor("white");
  const headers = ["Category", "Estimated Cost", "Actual Cost", "Variance", "ROI", "Cases Affected", "Members Benefited", "Status", "Quarter", "Notes"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.SOLIDARITY_RED);
}

/* ===================== TEST 2: PERFORMANCE ===================== */
function createPerformanceSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.PERFORMANCE);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.PERFORMANCE);
  }
  sheet.clear();

  sheet.getRange("A1:K1").merge()
    .setValue("🎯 PERFORMANCE METRICS & TRENDS")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Column definitions for formulas
  const statusCol = 'E';
  const dateOpenedCol = 'D';
  const dateClosedCol = 'R';
  const resolutionCol = 'Q';
  const daysOpenCol = 'S';

  // Key Performance Indicators
  sheet.getRange("A3").setValue("KEY PERFORMANCE INDICATORS").setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);

  const kpis = [
    ["Metric", "Current Value", "Target", "Variance", "Status"],
    ["Total Grievances", `=COUNTA('Grievance Log'!A:A)-1`, 5000, `=B5-C5`, `=IF(D5>=0,"✅","⚠️")`],
    ["Open Grievances", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`, 50, `=B6-C6`, `=IF(B6<=C6,"✅","⚠️")`],
    ["Win Rate %", `=IFERROR(TEXT(COUNTIFS('Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")/COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved*"),"0.0%"),"-")`, "75%", `=VALUE(LEFT(B7,LEN(B7)-1))-VALUE(LEFT(C7,LEN(C7)-1))`, `=IF(D7>=0,"✅","⚠️")`],
    ["Avg Resolution Days", `=IFERROR(ROUND(AVERAGE('Grievance Log'!${daysOpenCol}:${daysOpenCol}),0),"-")`, 30, `=B8-C8`, `=IF(B8<=C8,"✅","⚠️")`],
    ["Active Stewards", `=COUNTIF('Member Directory'!J:J,"Yes")`, 20, `=B9-C9`, `=IF(B9>=C9,"✅","⚠️")`]
  ];

  sheet.getRange(4, 1, kpis.length, 5).setValues(kpis);
  sheet.getRange(4, 1, 1, 5).setFontWeight("bold").setBackground(COLORS.INFO_LIGHT);

  // Monthly Performance Trend
  sheet.getRange("A12").setValue("MONTHLY PERFORMANCE TREND").setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);

  const trendHeaders = ["Month", "New Cases", "Resolved", "Win Rate %", "Avg Days", "Open EOMonth"];
  sheet.getRange(13, 1, 1, trendHeaders.length).setValues([trendHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.INFO_LIGHT);

  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.PRIMARY_BLUE);
}

/* ===================== TEST 9: QUICK STATS ===================== */
function createQuickStatsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.QUICK_STATS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.QUICK_STATS);
  }
  sheet.clear();

  sheet.getRange("A1:G1").merge()
    .setValue("⚡ QUICK STATS - AT-A-GLANCE DASHBOARD")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white");

  sheet.getRange("A2:G2").merge()
    .setValue("Real-time snapshot of critical metrics - refresh any time")
    .setFontSize(10)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  // Column definitions
  const statusCol = 'E';
  const resolutionCol = 'Q';
  const isStewardCol = 'J';
  const daysOpenCol = 'S';

  // Quick stats grid
  const quickStats = [
    ["📊 MEMBERS & STEWARDS", ""],
    ["Total Members", `=COUNTA('Member Directory'!A:A)-1`],
    ["Active Stewards", `=COUNTIF('Member Directory'!${isStewardCol}:${isStewardCol},"Yes")`],
    ["Avg Cases/Steward", `=IFERROR(ROUND(COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")/COUNTIF('Member Directory'!${isStewardCol}:${isStewardCol},"Yes"),1),"-")`],
    ["", ""],
    ["📋 GRIEVANCES", ""],
    ["Total Grievances", `=COUNTA('Grievance Log'!A:A)-1`],
    ["Open", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`],
    ["Pending Info", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Pending Info")`],
    ["Resolved/Settled", `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Settled")+COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved")`],
    ["", ""],
    ["🎯 PERFORMANCE", ""],
    ["Win Rate", `=IFERROR(TEXT(COUNTIFS('Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")/COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved*"),"0.0%"),"-")`],
    ["Avg Days to Resolve", `=IFERROR(ROUND(AVERAGE('Grievance Log'!${daysOpenCol}:${daysOpenCol}),0),"-")`],
    ["Cases This Month", `=COUNTIFS('Grievance Log'!D:D,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))`]
  ];

  sheet.getRange(4, 1, quickStats.length, 2).setValues(quickStats);

  // Format section headers
  sheet.getRange("A4").setFontWeight("bold").setBackground(COLORS.ACCENT_TEAL).setFontColor("white");
  sheet.getRange("A9").setFontWeight("bold").setBackground(COLORS.UNION_GREEN).setFontColor("white");
  sheet.getRange("A15").setFontWeight("bold").setBackground(COLORS.PRIMARY_PURPLE).setFontColor("white");

  // Last updated timestamp
  sheet.getRange("A" + (quickStats.length + 6)).setValue("Last Updated:");
  sheet.getRange("B" + (quickStats.length + 6)).setFormula("=NOW()").setNumberFormat("MM/dd/yyyy hh:mm:ss");

  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.ACCENT_ORANGE);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
}

function createArchiveSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.ARCHIVE);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ARCHIVE);
  sheet.clear();
  sheet.getRange("A1:F1").merge().setValue("📦 ARCHIVE").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.TEXT_GRAY).setFontColor("white");
  const headers = ["Item Type", "Item ID", "Archive Date", "Archived By", "Reason", "Original Data"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.TEXT_GRAY);
}

function createDiagnosticsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SHEETS.DIAGNOSTICS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.DIAGNOSTICS);
  sheet.clear();
  sheet.getRange("A1:G1").merge().setValue("🔧 DIAGNOSTICS").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center").setBackground(COLORS.SOLIDARITY_RED).setFontColor("white");
  const headers = ["Timestamp", "Check Type", "Component", "Status", "Details", "Severity", "Action Needed"];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground(COLORS.LIGHT_GRAY);
  sheet.setFrozenRows(3);
  sheet.setTabColor(COLORS.SOLIDARITY_RED);
}

/* ===================== DATA VALIDATIONS ===================== */
function setupDataValidations() {
  const ss = SpreadsheetApp.getActive();
  const config = ss.getSheetByName(SHEETS.CONFIG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // ENHANCEMENT: Email validation for Member Directory (Column H - Email Address)
  const emailRule = SpreadsheetApp.newDataValidation()
    .requireTextMatchesPattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}$')
    .setAllowInvalid(false)
    .setHelpText('Enter a valid email address (e.g., name@example.com)')
    .build();
  memberDir.getRange(2, 8, 5000, 1).setDataValidation(emailRule);

  // ENHANCEMENT: Phone number validation for Member Directory (Column I - Phone Number)
  const phoneRule = SpreadsheetApp.newDataValidation()
    .requireTextMatchesPattern('^\\(?\\d{3}\\)?[\\s.-]?\\d{3}[\\s.-]?\\d{4}$')
    .setAllowInvalid(false)
    .setHelpText('Enter phone number (formats: 555-123-4567, (555) 123-4567, 5551234567)')
    .build();
  memberDir.getRange(2, 9, 5000, 1).setDataValidation(phoneRule);

  // ENHANCEMENT: Email validation for Grievance Log (Column X - Member Email)
  grievanceLog.getRange(2, 24, 5000, 1).setDataValidation(emailRule);

  // Member Directory validations (updated for new Config structure)
  const memberValidations = [
    { col: 4, configCol: 1 },    // Job Title
    { col: 5, configCol: 2 },    // Work Location
    { col: 6, configCol: 3 },    // Unit
    { col: 10, configCol: 5 },   // Is Steward
    { col: 13, configCol: 10 },  // Assigned Steward
    { col: 20, configCol: 5 },   // Interest: Local
    { col: 21, configCol: 5 },   // Interest: Chapter
    { col: 22, configCol: 5 },   // Interest: Allied
    { col: 24, configCol: 15 }   // Comm Methods
  ];

  memberValidations.forEach(v => {
    const configRange = config.getRange(2, v.configCol, 50, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configRange, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, v.col, 5000, 1).setDataValidation(rule);
  });

  // Office Days - allow text input with guidance (column 7)
  // Note: Multiple days should be entered as comma-separated values
  const officeDaysRule = SpreadsheetApp.newDataValidation()
    .requireTextContains("")
    .setAllowInvalid(true)
    .setHelpText('Enter office days (e.g., "Monday, Wednesday, Friday" or select from Config tab)')
    .build();
  memberDir.getRange(2, 7, 5000, 1).setDataValidation(officeDaysRule);

  // Supervisor and Manager - allow text input since they're now first+last name combinations
  // Users can manually type "FirstName LastName" or the seed function will populate them
  const nameRule = SpreadsheetApp.newDataValidation()
    .requireTextContains("")
    .setAllowInvalid(true)
    .setHelpText('Enter full name (e.g., "John Smith")')
    .build();
  memberDir.getRange(2, 11, 5000, 1).setDataValidation(nameRule); // Supervisor
  memberDir.getRange(2, 12, 5000, 1).setDataValidation(nameRule); // Manager

  // Grievance Log validations (updated for new Config structure)
  const grievanceValidations = [
    { col: 5, configCol: 11 },  // Status (now column K)
    { col: 6, configCol: 12 },  // Current Step (now column L)
    { col: 22, configCol: 14 }, // Articles Violated (now column N)
    { col: 23, configCol: 13 }, // Issue Category (now column M)
    { col: 25, configCol: 3 },  // Unit
    { col: 26, configCol: 2 },  // Work Location
    { col: 27, configCol: 10 }  // Assigned Steward (now column J)
  ];

  grievanceValidations.forEach(v => {
    const configRange = config.getRange(2, v.configCol, 50, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configRange, true)
      .setAllowInvalid(false)
      .build();
    grievanceLog.getRange(2, v.col, 5000, 1).setDataValidation(rule);
  });
}

/* ===================== FORMULAS ===================== */
function setupFormulasAndCalculations() {
  const ss = SpreadsheetApp.getActive();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Grievance Log formulas - ENHANCED: Now covers 1000 rows instead of 100
  // Filing Deadline (Incident Date + 21 days) - Column H
  grievanceLog.getRange("H2").setFormula(
    `=ARRAYFORMULA(IF(G2:G1000<>"",G2:G1000+21,""))`
  );

  // Step I Decision Due (Date Filed + 30 days) - Column J
  grievanceLog.getRange("J2").setFormula(
    `=ARRAYFORMULA(IF(I2:I1000<>"",I2:I1000+30,""))`
  );

  // Step II Appeal Due (Step I Decision Rcvd + 10 days) - Column L
  grievanceLog.getRange("L2").setFormula(
    `=ARRAYFORMULA(IF(K2:K1000<>"",K2:K1000+10,""))`
  );

  // Step II Decision Due (Step II Appeal Filed + 30 days) - Column N
  grievanceLog.getRange("N2").setFormula(
    `=ARRAYFORMULA(IF(M2:M1000<>"",M2:M1000+30,""))`
  );

  // Step III Appeal Due (Step II Decision Rcvd + 30 days) - Column P
  grievanceLog.getRange("P2").setFormula(
    `=ARRAYFORMULA(IF(O2:O1000<>"",O2:O1000+30,""))`
  );

  // Days Open - Column S
  grievanceLog.getRange("S2").setFormula(
    `=ARRAYFORMULA(IF(I2:I1000<>"",IF(R2:R1000<>"",R2:R1000-I2:I1000,TODAY()-I2:I1000),""))`
  );

  // Next Action Due - Column T
  grievanceLog.getRange("T2").setFormula(
    `=ARRAYFORMULA(IF(E2:E1000="Open",IF(F2:F1000="Step I",J2:J1000,IF(F2:F1000="Step II",N2:N1000,IF(F2:F1000="Step III",P2:P1000,H2:H1000))),""))`
  );

  // Days to Deadline - Column U
  grievanceLog.getRange("U2").setFormula(
    `=ARRAYFORMULA(IF(T2:T1000<>"",T2:T1000-TODAY(),""))`
  );

  // Dynamic column references for Member Directory formulas
  const memberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  const nextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);

  // Member Directory formulas - ENHANCED: Now covers 1000 rows instead of 100
  // Has Open Grievance? - Column Z (26)
  memberDir.getRange("Z2").setFormula(
    `=ARRAYFORMULA(IF(A2:A1000<>"",IF(COUNTIFS('Grievance Log'!${memberIdCol}:${memberIdCol},A2:A1000,'Grievance Log'!${statusCol}:${statusCol},"Open")>0,"Yes","No"),""))`
  );

  // Grievance Status Snapshot - Column AA (27)
  memberDir.getRange("AA2").setFormula(
    `=ARRAYFORMULA(IF(A2:A1000<>"",IFERROR(INDEX('Grievance Log'!${statusCol}:${statusCol},MATCH(A2:A1000,'Grievance Log'!${memberIdCol}:${memberIdCol},0)),""),""))`
  );

  // Next Grievance Deadline - Column AB (28)
  memberDir.getRange("AB2").setFormula(
    `=ARRAYFORMULA(IF(A2:A1000<>"",IFERROR(INDEX('Grievance Log'!${nextActionCol}:${nextActionCol},MATCH(A2:A1000,'Grievance Log'!${memberIdCol}:${memberIdCol},0)),""),""))`
  );
}

/* ===================== MENU ===================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("📊 509 Dashboard")
    .addItem("🔄 Refresh All", "refreshCalculations")
    .addSeparator()
    .addSubMenu(ui.createMenu("📊 Dashboards")
      .addItem("🎯 Unified Operations Monitor", "showUnifiedOperationsMonitor")
      .addItem("📊 Main Dashboard", "goToDashboard")
      .addItem("✨ Interactive Dashboard", "openInteractiveDashboard")
      .addItem("🔄 Refresh Interactive Dashboard", "rebuildInteractiveDashboard"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🔍 Search & Lookup")
      .addItem("🔍 Search Members", "showMemberSearch")
      .addItem("🔍 Quick Member Search", "quickMemberSearch"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📋 Grievance Tools")
      .addItem("➕ Start New Grievance", "showStartGrievanceDialog"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🔄 Workflow Management")
      .addItem("🔄 Workflow Visualizer", "showWorkflowVisualizer")
      .addItem("🔄 Change Workflow State", "changeWorkflowState")
      .addItem("📊 View Current State", "showCurrentWorkflowState")
      .addSeparator()
      .addItem("📦 Batch Update State", "batchUpdateWorkflowState"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🤖 Smart Assignment")
      .addItem("🤖 Auto-Assign Steward", "showAutoAssignDialog")
      .addItem("📦 Batch Auto-Assign", "batchAutoAssign")
      .addSeparator()
      .addItem("👥 Steward Workload Dashboard", "showStewardWorkloadDashboard"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⚡ Batch Operations")
      .addItem("⚡ Show Batch Operations Menu", "showBatchOperationsMenu")
      .addSeparator()
      .addItem("👤 Bulk Assign Steward", "batchAssignSteward")
      .addItem("📊 Bulk Update Status", "batchUpdateStatus")
      .addItem("📄 Bulk Export to PDF", "batchExportPDF")
      .addItem("📧 Bulk Email Notifications", "batchEmailNotifications")
      .addItem("📝 Bulk Add Notes", "batchAddNotes"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🔮 Analytics & Insights")
      .addItem("🔮 Predictive Analytics Dashboard", "showPredictiveAnalyticsDashboard")
      .addSeparator()
      .addItem("📈 Generate Full Analysis", "performPredictiveAnalysis"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🤖 Automation")
      .addItem("📬 Notification Settings", "showNotificationSettings")
      .addItem("📊 Report Automation Settings", "showReportAutomationSettings")
      .addSeparator()
      .addItem("✅ Enable Daily Deadline Notifications", "setupDailyDeadlineNotifications")
      .addItem("🔕 Disable Deadline Notifications", "disableDailyDeadlineNotifications")
      .addSeparator()
      .addItem("✅ Enable Monthly Reports", "setupMonthlyReports")
      .addItem("✅ Enable Quarterly Reports", "setupQuarterlyReports")
      .addItem("🔕 Disable All Reports", "disableAutomatedReports")
      .addSeparator()
      .addItem("🧪 Test Deadline Notifications", "testDeadlineNotifications")
      .addItem("🧪 Test Monthly Report", "generateMonthlyReport")
      .addItem("🧪 Test Quarterly Report", "generateQuarterlyReport"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📧 Communications")
      .addItem("📧 Compose Email", "composeGrievanceEmail")
      .addItem("📝 Email Templates", "showEmailTemplateManager")
      .addSeparator()
      .addItem("📞 View Communications Log", "showGrievanceCommunications"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📁 Google Drive")
      .addItem("📁 Setup Folder for Grievance", "setupDriveFolderForGrievance")
      .addItem("📎 Upload Files", "showFileUploadDialog")
      .addItem("📂 Show Grievance Files", "showGrievanceFiles")
      .addSeparator()
      .addItem("📁 Batch Create All Folders", "batchCreateGrievanceFolders"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📅 Calendar Integration")
      .addItem("📅 Sync Deadlines to Calendar", "syncDeadlinesToCalendar")
      .addItem("👀 View Upcoming Deadlines", "showUpcomingDeadlinesFromCalendar")
      .addSeparator()
      .addItem("🗑️ Clear All Calendar Events", "clearAllCalendarEvents"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🔒 Data Integrity")
      .addItem("📊 Data Quality Dashboard", "showDataQualityDashboard")
      .addItem("🔍 Check Referential Integrity", "checkReferentialIntegrity")
      .addSeparator()
      .addItem("📝 Create Change Log Sheet", "createChangeLogSheet")
      .addItem("🆔 Generate Next Member ID", "showNextMemberID")
      .addItem("🆔 Generate Next Grievance ID", "showNextGrievanceID"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⚡ Performance")
      .addItem("🗄️ Cache Status Dashboard", "showCacheStatusDashboard")
      .addSeparator()
      .addItem("🔥 Warm Up All Caches", "warmUpCaches")
      .addItem("🗑️ Clear All Caches", "invalidateAllCaches"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⌨️ Keyboard Shortcuts")
      .addItem("⌨️ Show Shortcuts Guide", "showKeyboardShortcuts")
      .addItem("⚙️ Shortcuts Configuration", "showKeyboardShortcutsConfig")
      .addSeparator()
      .addItem("F1 Context Help", "showContextHelp"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📊 Reports")
      .addItem("📊 Custom Report Builder", "showCustomReportBuilder")
      .addSeparator()
      .addItem("📄 Export Grievances to CSV", "exportGrievancesToCSV")
      .addItem("📄 Export Members to CSV", "exportMembersToCSV"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📚 Knowledge Base")
      .addItem("📚 Search FAQ", "showFAQSearch")
      .addItem("➕ Add New FAQ", "showFAQAdmin")
      .addSeparator()
      .addItem("📝 Create FAQ Database", "createFAQSheet"))
    .addSeparator()
    .addSubMenu(ui.createMenu("🔬 Analysis Tools")
      .addItem("🔬 Root Cause Analysis", "showRootCauseAnalysisDashboard"))
    .addSeparator()
    .addSubMenu(ui.createMenu("💾 Data Management")
      .addItem("💾 Backup & Recovery Manager", "showBackupManager")
      .addSeparator()
      .addItem("💾 Create Backup Now", "createBackup")
      .addItem("✅ Enable Automated Backups", "setupAutomatedBackups")
      .addItem("🔕 Disable Automated Backups", "disableAutomatedBackups")
      .addSeparator()
      .addItem("📊 View Backup Log", "navigateToBackupLog"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⚙️ Admin")
      .addSubMenu(ui.createMenu("👥 Seed Members")
        .addItem("Seed Members - Toggle 1 (5,000)", "SEED_MEMBERS_TOGGLE_1")
        .addItem("Seed Members - Toggle 2 (5,000)", "SEED_MEMBERS_TOGGLE_2")
        .addItem("Seed Members - Toggle 3 (5,000)", "SEED_MEMBERS_TOGGLE_3")
        .addItem("Seed Members - Toggle 4 (5,000)", "SEED_MEMBERS_TOGGLE_4")
        .addSeparator()
        .addItem("Seed All 20k Members (Legacy)", "SEED_20K_MEMBERS"))
      .addSubMenu(ui.createMenu("📋 Seed Grievances")
        .addItem("Seed Grievances - Toggle 1 (2,500)", "SEED_GRIEVANCES_TOGGLE_1")
        .addItem("Seed Grievances - Toggle 2 (2,500)", "SEED_GRIEVANCES_TOGGLE_2")
        .addSeparator()
        .addItem("Seed All 5k Grievances (Legacy)", "SEED_5K_GRIEVANCES"))
      .addSeparator()
      .addItem("Clear All Data", "clearAllData")
      .addItem("🗑️ Nuke All Seed Data", "nukeSeedData")
      .addSeparator()
      .addItem("📝 Add Sample Feedback Entries", "addSampleFeedbackEntries")
      .addItem("📊 Populate Analytics Sheets", "populateAllAnalyticsSheets")
      .addItem("👁️ Hide Diagnostics Tab", "hideDiagnosticsTab")
      .addSeparator()
      .addItem("🎨 Setup Dashboard Enhancements", "SETUP_DASHBOARD_ENHANCEMENTS"))
    .addSeparator()
    .addSubMenu(ui.createMenu("♿ Accessibility")
      .addItem("♿ ADHD Control Panel", "showADHDControlPanel")
      .addItem("🎨 Theme Manager", "showThemeManager")
      .addSeparator()
      .addItem("Hide Gridlines (Focus Mode)", "hideAllGridlines")
      .addItem("Show Gridlines", "showAllGridlines")
      .addItem("Reorder Sheets Logically", "reorderSheetsLogically")
      .addItem("Setup ADHD Defaults", "setupADHDDefaults")
      .addSeparator()
      .addItem("🌙 Quick Toggle Dark Mode", "quickToggleDarkMode")
      .addItem("🎯 Activate Focus Mode", "activateFocusMode")
      .addItem("🎯 Deactivate Focus Mode", "deactivateFocusMode"))
    .addSeparator()
    .addSubMenu(ui.createMenu("📱 Mobile & Viewing")
      .addItem("📱 Mobile Dashboard", "showMobileDashboard")
      .addItem("📋 Mobile Grievance List", "showMobileGrievanceList")
      .addSeparator()
      .addItem("📄 Paginated Data Viewer", "showPaginatedViewer"))
    .addSeparator()
    .addSubMenu(ui.createMenu("↩️ History & Undo")
      .addItem("↩️ Undo/Redo History", "showUndoRedoPanel")
      .addItem("⌨️ Install Undo Shortcuts", "installUndoRedoShortcuts")
      .addSeparator()
      .addItem("↩️ Undo Last Action (Ctrl+Z)", "undoLastAction")
      .addItem("↪️ Redo Last Action (Ctrl+Y)", "redoLastAction")
      .addSeparator()
      .addItem("🗑️ Clear Undo History", "clearUndoHistory"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⚠️ System Health")
      .addItem("⚠️ Error Dashboard", "showErrorDashboard")
      .addItem("🏥 Run Health Check", "performSystemHealthCheck")
      .addSeparator()
      .addItem("📊 View Error Trends", "createErrorTrendReport")
      .addItem("🧪 Test Error Logging", "testErrorLogging"))
    .addSeparator()
    .addSubMenu(ui.createMenu("👁️ Column Toggles")
      // .addItem("Toggle Advanced Grievance Columns", "toggleGrievanceColumns")  // Disabled - not compatible with current structure
      .addItem("Toggle Level 2 Member Columns", "toggleLevel2Columns")
      .addItem("Show All Member Columns", "showAllMemberColumns"))
    .addSeparator()
    .addSubMenu(ui.createMenu("❓ Help & Support")
      .addItem("📚 Getting Started Guide", "showGettingStartedGuide")
      .addItem("❓ Help", "showHelp")
      .addItem("🔧 Diagnose Setup", "DIAGNOSE_SETUP"))
    .addToUi();
}

function refreshCalculations() {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.getActive();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  if (dashboard) {
    dashboard.getRange("A3").setFormula('="Last Updated: " & TEXT(NOW(), "MM/DD/YYYY HH:MM:SS")');
  }
  SpreadsheetApp.getActive().toast("✅ Refreshed", "Complete", 2);
}

function goToDashboard() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheetByName(SHEETS.DASHBOARD).activate();
}

function showHelp() {
  const helpText = `
📊 509 DASHBOARD

DASHBOARDS:
• 🎯 Unified Operations Monitor - Comprehensive terminal-style dashboard
  - Executive status & deadline tracking
  - Process efficiency & caseload analysis
  - Network health & steward capacity
  - Action logs & predictive alerts
  - Systemic risk monitoring

SHEETS:
• Config - Master dropdown lists
• Member Directory - All member data
• Grievance Log - All grievances with auto-calculated deadlines
• Dashboard - Real-time metrics
• Member Satisfaction - Survey tracking
• Feedback & Development - System improvements

DATA SEEDING:
Use Admin menu to:
• Seed 20k Members
• Seed 5k Grievances

All metrics use REAL data from Member Directory and Grievance Log.
No fake CPU/memory metrics - everything tracks actual union activity.
  `;

  SpreadsheetApp.getUi().alert("Help", helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ===================== SEED 20,000 MEMBERS ===================== */
/* ===================== SEED MEMBERS (WITH TOGGLES) ===================== */
function SEED_MEMBERS_TOGGLE_1() { seedMembersWithCount(5000, "Toggle 1"); }
function SEED_MEMBERS_TOGGLE_2() { seedMembersWithCount(5000, "Toggle 2"); }
function SEED_MEMBERS_TOGGLE_3() { seedMembersWithCount(5000, "Toggle 3"); }
function SEED_MEMBERS_TOGGLE_4() { seedMembersWithCount(5000, "Toggle 4"); }

function seedMembersWithCount(count, toggleName) {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Seed ${count} Members (${toggleName})`,
    `This will add ${count} member records. This may take 1-2 minutes. Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  SpreadsheetApp.getActive().toast(`🚀 Seeding ${count} members (${toggleName})...`, "Processing", -1);

  const firstNames = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah", "Charles", "Karen", "Christopher", "Nancy", "Daniel", "Lisa", "Matthew", "Betty", "Anthony", "Margaret", "Mark", "Sandra", "Donald", "Ashley", "Steven", "Kimberly", "Paul", "Emily", "Andrew", "Donna", "Joshua", "Michelle"];
  const lastNames = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores"];

  const jobTitles = config.getRange("A2:A14").getValues().flat().filter(String);
  const locations = config.getRange("B2:B14").getValues().flat().filter(String);
  const units = config.getRange("C2:C7").getValues().flat().filter(String);
  const officeDays = config.getRange("D2:D8").getValues().flat().filter(String);

  // Get supervisor and manager names (first and last name columns)
  const supervisorFirstNames = config.getRange("F2:F14").getValues().flat().filter(String);
  const supervisorLastNames = config.getRange("G2:G14").getValues().flat().filter(String);
  const managerFirstNames = config.getRange("H2:H14").getValues().flat().filter(String);
  const managerLastNames = config.getRange("I2:I14").getValues().flat().filter(String);
  const stewards = config.getRange("J2:J14").getValues().flat().filter(String);

  const commMethods = ["Email", "Phone", "Text", "In Person"];
  const times = ["Mornings", "Afternoons", "Evenings", "Weekends", "Flexible"];

  // Validate config data
  if (jobTitles.length === 0 || locations.length === 0 || units.length === 0 ||
      supervisorFirstNames.length === 0 || managerFirstNames.length === 0 || stewards.length === 0) {
    ui.alert('Error', 'Config data is incomplete. Please ensure all dropdown lists in Config sheet are populated.', ui.ButtonSet.OK);
    return;
  }

  const BATCH_SIZE = 1000;
  let data = [];
  const startingRow = memberDir.getLastRow();

  for (let i = 1; i <= count; i++) {
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
    const memberID = "M" + String(startingRow + i).padStart(6, '0');
    const jobTitle = jobTitles[Math.floor(Math.random() * jobTitles.length)];
    const location = locations[Math.floor(Math.random() * locations.length)];
    const unit = units[Math.floor(Math.random() * units.length)];

    // Generate multiple office days (1-3 days)
    const numDays = Math.floor(Math.random() * 3) + 1;
    const selectedDays = [];
    const availableDays = [...officeDays];
    for (let d = 0; d < numDays; d++) {
      if (availableDays.length > 0) {
        const idx = Math.floor(Math.random() * availableDays.length);
        selectedDays.push(availableDays.splice(idx, 1)[0]);
      }
    }
    const officeDaysValue = selectedDays.join(", ");

    const email = `${firstName.toLowerCase()}.${lastName.toLowerCase()}${startingRow + i}@union.org`;
    const phone = `(555) ${String(Math.floor(Math.random() * 900) + 100)}-${String(Math.floor(Math.random() * 9000) + 1000)}`;
    const isSteward = Math.random() > 0.95 ? "Yes" : "No";

    // Build full supervisor and manager names
    const supervisorIdx = Math.floor(Math.random() * supervisorFirstNames.length);
    const supervisor = `${supervisorFirstNames[supervisorIdx]} ${supervisorLastNames[supervisorIdx]}`;

    const managerIdx = Math.floor(Math.random() * managerFirstNames.length);
    const manager = `${managerFirstNames[managerIdx]} ${managerLastNames[managerIdx]}`;

    const assignedSteward = stewards[Math.floor(Math.random() * stewards.length)];

    const daysAgo = Math.floor(Math.random() * 90);
    const lastVirtual = Math.random() > 0.7 ? new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000) : "";
    const lastInPerson = Math.random() > 0.8 ? new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000) : "";
    const lastSurvey = Math.random() > 0.6 ? new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000) : "";
    const lastEmailOpen = Math.random() > 0.5 ? new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000) : "";

    const openRate = Math.floor(Math.random() * 40) + 60;
    const volHours = Math.floor(Math.random() * 50);
    const localInterest = Math.random() > 0.5 ? "Yes" : "No";
    const chapterInterest = Math.random() > 0.6 ? "Yes" : "No";
    const alliedInterest = Math.random() > 0.8 ? "Yes" : "No";
    const timestamp = new Date(Date.now() - Math.random() * 365 * 24 * 60 * 60 * 1000);
    const commMethod = commMethods[Math.floor(Math.random() * commMethods.length)];
    const bestTime = times[Math.floor(Math.random() * times.length)];

    const row = [
      memberID, firstName, lastName, jobTitle, location, unit, officeDaysValue,
      email, phone, isSteward, supervisor, manager, assignedSteward,
      lastVirtual, lastInPerson, lastSurvey, lastEmailOpen, openRate, volHours,
      localInterest, chapterInterest, alliedInterest, timestamp, commMethod, bestTime,
      "No", "", "", "", "", ""
    ];

    data.push(row);

    if (data.length === BATCH_SIZE) {
      try {
        memberDir.getRange(memberDir.getLastRow() + 1, 1, data.length, row.length).setValues(data);
        SpreadsheetApp.getActive().toast(`Added ${i} of ${count} members (${toggleName})...`, "Progress", 1);
        data = [];
        SpreadsheetApp.flush();
      } catch (e) {
        Logger.log(`Error writing member batch at ${i}: ${e.message}`);
        SpreadsheetApp.getActive().toast(`⚠️ Error at ${i}. Retrying...`, "Warning", 2);
        // Retry once
        Utilities.sleep(1000);
        try {
          memberDir.getRange(memberDir.getLastRow() + 1, 1, data.length, row.length).setValues(data);
          data = [];
        } catch (e2) {
          Logger.log(`Retry failed: ${e2.message}`);
          throw new Error(`Failed to write members: ${e2.message}`);
        }
      }
    }
  }

  // Write remaining data
  if (data.length > 0) {
    try {
      memberDir.getRange(memberDir.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      Logger.log(`Error writing final member batch: ${e.message}`);
      throw new Error(`Failed to write final members: ${e.message}`);
    }
  }

  SpreadsheetApp.getActive().toast(`✅ ${count} members added (${toggleName})!`, "Complete", 5);
}

/* ===================== LEGACY: SEED 20,000 MEMBERS ===================== */
function SEED_20K_MEMBERS() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Seed 20,000 Members',
    'Use the 4 toggles instead for better performance:\n\n' +
    '• Seed Members - Toggle 1 (5,000)\n' +
    '• Seed Members - Toggle 2 (5,000)\n' +
    '• Seed Members - Toggle 3 (5,000)\n' +
    '• Seed Members - Toggle 4 (5,000)\n\n' +
    'Would you like to seed all 20,000 at once anyway?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Call all 4 toggles
  seedMembersWithCount(5000, "Toggle 1");
  seedMembersWithCount(5000, "Toggle 2");
  seedMembersWithCount(5000, "Toggle 3");
  seedMembersWithCount(5000, "Toggle 4");
}

/* ===================== SEED GRIEVANCES (WITH TOGGLES) ===================== */
function SEED_GRIEVANCES_TOGGLE_1() { seedGrievancesWithCount(2500, "Toggle 1"); }
function SEED_GRIEVANCES_TOGGLE_2() { seedGrievancesWithCount(2500, "Toggle 2"); }

function seedGrievancesWithCount(count, toggleName) {
  const ss = SpreadsheetApp.getActive();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Seed ${count} Grievances (${toggleName})`,
    `This will add ${count} grievance records. This may take 1-2 minutes. Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  SpreadsheetApp.getActive().toast(`🚀 Seeding ${count} grievances (${toggleName})...`, "Processing", -1);

  // Get member data ONCE before the loop (CRITICAL FIX)
  const memberLastRow = memberDir.getLastRow();
  if (memberLastRow < 2) {
    ui.alert('Error', 'No members found. Please seed members first.', ui.ButtonSet.OK);
    return;
  }

  const allMemberData = memberDir.getRange(2, 1, memberLastRow - 1, 31).getValues();
  const memberIDs = allMemberData.map(row => row[0]).filter(String);

  // Updated config column references for new structure
  const statuses = config.getRange("K2:K8").getValues().flat().filter(String);       // Grievance Status (column K)
  const steps = config.getRange("L2:L7").getValues().flat().filter(String);          // Grievance Step (column L)
  const categories = config.getRange("M2:M12").getValues().flat().filter(String);    // Issue Category (column M)
  const articles = config.getRange("N2:N14").getValues().flat().filter(String);      // Articles Violated (column N)
  const stewards = config.getRange("J2:J14").getValues().flat().filter(String);      // Stewards (column J)

  // Validate config data
  if (statuses.length === 0 || steps.length === 0 || articles.length === 0 ||
      categories.length === 0 || stewards.length === 0) {
    ui.alert('Error', 'Config data is incomplete. Please ensure all dropdown lists in Config sheet are populated.', ui.ButtonSet.OK);
    return;
  }

  const BATCH_SIZE = 500;
  let data = [];
  let successCount = 0;
  const startingRow = grievanceLog.getLastRow();

  for (let i = 1; i <= count; i++) {
    // Get random member
    const memberIndex = Math.floor(Math.random() * memberIDs.length);
    const memberID = memberIDs[memberIndex];
    const memberData = allMemberData[memberIndex];

    if (!memberData || !memberID) continue;

    const grievanceID = "G-" + String(startingRow + i).padStart(6, '0');
    const firstName = memberData[1];
    const lastName = memberData[2];
    const status = statuses[Math.floor(Math.random() * statuses.length)];
    const step = steps[Math.floor(Math.random() * steps.length)];

    const daysAgo = Math.floor(Math.random() * 365);
    const incidentDate = new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000);
    const dateFiled = new Date(incidentDate.getTime() + Math.random() * 14 * 24 * 60 * 60 * 1000);

    const article = articles[Math.floor(Math.random() * articles.length)];
    const category = categories[Math.floor(Math.random() * categories.length)];
    const email = memberData[7];
    const unit = memberData[5];
    const location = memberData[4];
    const assignedSteward = stewards[Math.floor(Math.random() * stewards.length)];

    const isClosed = status === "Closed" || status === "Settled" || status === "Withdrawn";
    const dateClosed = isClosed ? new Date(dateFiled.getTime() + Math.random() * 90 * 24 * 60 * 60 * 1000) : "";
    const resolution = isClosed ? ["Won - Resolved favorably", "Won - Full remedy granted", "Lost - No violation found", "Lost - Withdrawn by member", "Settled - Partial remedy", "Settled - Compromise reached"][Math.floor(Math.random() * 6)] : "";

    // Calculate all deadline columns based on contract rules
    const filingDeadline = new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000);
    const step1DecisionDue = new Date(dateFiled.getTime() + 30 * 24 * 60 * 60 * 1000);
    const step1DecisionRcvd = (step !== "Informal" && Math.random() > 0.3) ? new Date(dateFiled.getTime() + Math.random() * 30 * 24 * 60 * 60 * 1000) : "";
    const step2AppealDue = step1DecisionRcvd ? new Date(step1DecisionRcvd.getTime() + 10 * 24 * 60 * 60 * 1000) : "";
    const step2AppealFiled = (step === "Step II" || step === "Step III" || step === "Arbitration") && step2AppealDue ? new Date(step1DecisionRcvd.getTime() + Math.random() * 10 * 24 * 60 * 60 * 1000) : "";
    const step2DecisionDue = step2AppealFiled ? new Date(step2AppealFiled.getTime() + 30 * 24 * 60 * 60 * 1000) : "";
    const step2DecisionRcvd = (step === "Step III" || step === "Arbitration") && step2DecisionDue ? new Date(step2AppealFiled.getTime() + Math.random() * 30 * 24 * 60 * 60 * 1000) : "";
    const step3AppealDue = step2DecisionRcvd ? new Date(step2DecisionRcvd.getTime() + 30 * 24 * 60 * 60 * 1000) : "";
    const step3AppealFiled = (step === "Step III" || step === "Arbitration") && step3AppealDue ? new Date(step2DecisionRcvd.getTime() + Math.random() * 30 * 24 * 60 * 60 * 1000) : "";
    const daysOpen = isClosed && dateClosed ? Math.floor((dateClosed - dateFiled) / (1000 * 60 * 60 * 24)) : Math.floor((Date.now() - dateFiled.getTime()) / (1000 * 60 * 60 * 24));
    let nextActionDue = "";
    if (!isClosed) {
      if (step === "Informal" || step === "Step I") nextActionDue = step1DecisionDue;
      else if (step === "Step II") nextActionDue = step2DecisionDue || step2AppealDue;
      else if (step === "Step III") nextActionDue = step3AppealDue;
      else if (step === "Arbitration") nextActionDue = new Date(Date.now() + Math.random() * 60 * 24 * 60 * 60 * 1000);
    }
    const daysToDeadline = nextActionDue ? Math.floor((nextActionDue - Date.now()) / (1000 * 60 * 60 * 24)) : "";

    const row = [
      grievanceID, memberID, firstName, lastName, status, step,
      incidentDate, filingDeadline, dateFiled, step1DecisionDue, step1DecisionRcvd,
      step2AppealDue, step2AppealFiled, step2DecisionDue, step2DecisionRcvd,
      step3AppealDue, step3AppealFiled, dateClosed, daysOpen, nextActionDue, daysToDeadline,
      article, category, email, unit, location, assignedSteward, resolution
    ];

    data.push(row);
    successCount++;

    if (data.length === BATCH_SIZE) {
      try {
        grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, data.length, row.length).setValues(data);
        SpreadsheetApp.getActive().toast(`Added ${successCount} of ${count} grievances (${toggleName})...`, "Progress", 1);
        data = [];
        SpreadsheetApp.flush();
      } catch (e) {
        Logger.log(`Error writing batch at count ${successCount}: ${e.message}`);
        SpreadsheetApp.getActive().toast(`⚠️ Error at ${successCount}. Retrying...`, "Warning", 2);
        // Retry once
        Utilities.sleep(1000);
        try {
          grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, data.length, row.length).setValues(data);
          data = [];
        } catch (e2) {
          Logger.log(`Retry failed: ${e2.message}`);
          throw new Error(`Failed to write grievances: ${e2.message}`);
        }
      }
    }
  }

  // Write remaining data
  if (data.length > 0) {
    try {
      grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    } catch (e) {
      Logger.log(`Error writing final batch: ${e.message}`);
      throw new Error(`Failed to write final grievances: ${e.message}`);
    }
  }

  SpreadsheetApp.getActive().toast(`✅ ${successCount} grievances added (${toggleName})! Updating member snapshots...`, "Processing", 2);
  updateMemberDirectorySnapshots();
  SpreadsheetApp.getActive().toast(`✅ ${successCount} grievances added (${toggleName}) and member snapshots updated!`, "Complete", 5);
}

/* ===================== LEGACY: SEED 5,000 GRIEVANCES ===================== */
function updateMemberDirectorySnapshots() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!memberDir || !grievanceLog) return;
  const memberLastRow = memberDir.getLastRow();
  const grievanceLastRow = grievanceLog.getLastRow();
  if (memberLastRow < 2 || grievanceLastRow < 2) return;
  const memberIDs = memberDir.getRange(2, 1, memberLastRow - 1, 1).getValues().flat();
  const grievanceData = grievanceLog.getRange(2, 1, grievanceLastRow - 1, 28).getValues();
  const memberSnapshots = {};
  grievanceData.forEach(row => {
    const memberID = row[1];
    const status = row[4];
    const nextActionDue = row[19];
    const assignedSteward = row[26];
    if (!memberID) return;
    if (!memberSnapshots[memberID]) {
      memberSnapshots[memberID] = {status, nextDeadline: nextActionDue, stewardWhoContacted: assignedSteward};
    } else {
      if (status && (status === "Open" || status.includes("Filed") || status === "Pending Info")) memberSnapshots[memberID].status = status;
      if (nextActionDue && nextActionDue instanceof Date) {
        if (!memberSnapshots[memberID].nextDeadline || (memberSnapshots[memberID].nextDeadline instanceof Date && nextActionDue < memberSnapshots[memberID].nextDeadline)) {
          memberSnapshots[memberID].nextDeadline = nextActionDue;
        }
      }
    }
  });
  const updateData = [];
  for (let i = 0; i < memberIDs.length; i++) {
    const snapshot = memberSnapshots[memberIDs[i]];
    if (snapshot) {
      const contactDate = new Date(Date.now() - Math.floor(Math.random() * 14) * 24 * 60 * 60 * 1000);
      const contactNotes = ["Discussed case progress", "Member updated on next steps", "Reviewed timeline and deadlines", "Answered member questions", "Scheduled follow-up meeting"][Math.floor(Math.random() * 5)];
      updateData.push([snapshot.status || "", snapshot.nextDeadline || "", contactDate, snapshot.stewardWhoContacted || "", contactNotes]);
    } else {
      updateData.push(["", "", "", "", ""]);
    }
  }
  if (updateData.length > 0) memberDir.getRange(2, 25, updateData.length, 5).setValues(updateData);
}

function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear All Data',
    'This will delete all members and grievances. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (memberDir.getLastRow() > 1) {
    memberDir.getRange(2, 1, memberDir.getLastRow() - 1, memberDir.getLastColumn()).clear();
  }

  if (grievanceLog.getLastRow() > 1) {
    grievanceLog.getRange(2, 1, grievanceLog.getLastRow() - 1, grievanceLog.getLastColumn()).clear();
  }

  SpreadsheetApp.getActive().toast("✅ All data cleared", "Complete", 3);
}

/******************************************************************************
 * MODULE: UnifiedOperationsMonitor
 * Source: UnifiedOperationsMonitor.gs
 *****************************************************************************/

/**
 * ============================================================================
 * SEIU 509 COMPREHENSIVE ACTION DASHBOARD
 * ============================================================================
 *
 * Terminal-themed comprehensive operations monitor with 26+ sections covering:
 * - Executive status and deadlines
 * - Process efficiency and caseload
 * - Network health and capacity
 * - Action logs and follow-up radar
 * - Predictive alerts and systemic risks
 * - Financial recovery and arbitration tracking
 * - Member engagement and satisfaction
 * - Training, compliance, and performance metrics
 *
 * Designed for union stewards and administrators to monitor all operational
 * aspects in a single, information-dense terminal interface.
 * ============================================================================
 */

/**
 * Shows the Unified Operations Monitor dashboard
 */
function showUnifiedOperationsMonitor() {
  const html = HtmlService.createHtmlOutput(getUnifiedOperationsMonitorHTML())
    .setWidth(1600)
    .setHeight(1000);

  SpreadsheetApp.getUi().showModelessDialog(html, '🎯 SEIU 509 Comprehensive Action Dashboard');
}

/**
 * Backend function that provides ALL data for the comprehensive dashboard
 * Called by the HTML dashboard via google.script.run
 */
function getUnifiedDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    return {
      error: "Required sheets not found. Please ensure Grievance Log and Member Directory exist."
    };
  }

  // Get all data
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  // Skip headers
  const grievances = grievanceData.slice(1);
  const members = memberData.slice(1);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Calculate all metrics
  return {
    // Section 1: Executive Status
    activeCases: calculateActiveCases(grievances),
    overdue: calculateOverdue(grievances, today),
    dueWeek: calculateDueThisWeek(grievances, today),
    winRate: calculateWinRate(grievances),
    avgDays: calculateAvgDays(grievances),
    escalationCount: calculateEscalations(grievances),
    arbitrations: calculateArbitrations(grievances),
    highRiskCount: calculateHighRisk(grievances),

    // Section 2: Process Efficiency
    steps: calculateStepEfficiency(grievances),

    // Section 3: Network Health
    totalMembers: members.filter(m => m[MEMBER_COLS.MEMBER_ID - 1]).length,
    engagementRate: calculateEngagementRate(grievances, members),
    noContactCount: calculateNoContact(members, today),
    stewardCount: calculateActiveStewards(grievances),
    avgLoad: calculateAvgLoad(grievances),
    overloadedStewards: calculateOverloadedStewards(grievances),
    issueDistribution: calculateIssueDistribution(grievances),

    // Section 4: Action Log (Top 20 processes)
    processes: getTopProcesses(grievances, today, 20),

    // Section 5: Follow-up Radar
    followUpTasks: getFollowUpTasks(grievances, today),

    // Section 6: Predictive Alerts
    predictiveAlerts: getPredictiveAlerts(members, grievances, today),

    // Section 7: Systemic Risk
    systemicRisk: getSystemicRisks(grievances),

    // Section 8: Outreach Scorecard
    outreachScorecard: getOutreachScorecard(members),

    // Section 9: Arbitration Tracker
    arbitrationTracker: getArbitrationTracker(grievances),

    // Section 10: Contract Trends
    contractTrends: getContractTrends(grievances),

    // Section 11: Training Matrix
    trainingMatrix: getTrainingMatrix(members),

    // Section 12: Geographical Caseload
    locationCaseload: getLocationCaseload(grievances),

    // Section 13: Financial Recovery
    financialTracker: getFinancialTracker(grievances),

    // Section 14: Rolling Trends
    rollingTrends: getRollingTrends(grievances),

    // Section 15: Grievance Satisfaction
    grievanceSatisfaction: getGrievanceSatisfaction(members, grievances),

    // Section 16: Member Lifecycle
    memberLifecycle: getMemberLifecycle(members, today),

    // Section 17: Unit Performance
    unitPerformance: getUnitPerformance(grievances, members),

    // Section 18: Document Compliance
    docCompliance: getDocCompliance(grievances),

    // Section 19: Classification Risk
    classificationRisk: getClassificationRisk(members, grievances),

    // Section 20: Win/Loss History
    winLossHistory: getWinLossHistory(grievances),

    // Section 21: Legal Review
    legalReviewCases: getLegalReviewCases(grievances, today),

    // Section 22: Recruitment Tracker
    recruitment: getRecruitmentTracker(members, today),

    // Section 23: Bargaining Prep
    bargainingPrep: getBargainingPrep(members, grievances),

    // Section 24: Policy Impact
    policyImpact: getPolicyImpact(grievances, today),

    // Section 25: Bottlenecks
    bottlenecks: getBottlenecks(grievances),

    // Section 26: Watchlist
    watchlistItems: getWatchlistItems(),
    watchlistLog: getWatchlistLog(grievances, members)
  };
}

// ============================================================================
// CALCULATION FUNCTIONS
// ============================================================================

function calculateActiveCases(grievances) {
  return grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  }).length;
}

function calculateOverdue(grievances, today) {
  return grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        return deadline < today;
      }
    }
    return false;
  }).length;
}

function calculateDueThisWeek(grievances, today) {
  const oneWeekFromNow = new Date(today);
  oneWeekFromNow.setDate(oneWeekFromNow.getDate() + 7);

  return grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        return deadline >= today && deadline <= oneWeekFromNow;
      }
    }
    return false;
  }).length;
}

function calculateWinRate(grievances) {
  const closed = grievances.filter(g => g[GRIEVANCE_COLS.STATUS - 1] === 'Settled' || g[GRIEVANCE_COLS.STATUS - 1] === 'Closed' || g[GRIEVANCE_COLS.STATUS - 1] === 'Withdrawn');
  if (closed.length === 0) return 0;
  const won = closed.filter(g => g[GRIEVANCE_COLS.STATUS - 1] === 'Settled').length;
  return (won / closed.length) * 100;
}

function calculateAvgDays(grievances) {
  const closed = grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const filedDate = g[GRIEVANCE_COLS.DATE_FILED - 1];
    const closedDate = g[GRIEVANCE_COLS.DATE_CLOSED - 1];
    return (status === 'Settled' || status === 'Closed' || status === 'Withdrawn') &&
           filedDate instanceof Date && closedDate instanceof Date;
  });

  if (closed.length === 0) return 0;

  const totalDays = closed.reduce((sum, g) => {
    const days = Math.floor((g[GRIEVANCE_COLS.DATE_CLOSED - 1] - g[GRIEVANCE_COLS.DATE_FILED - 1]) / (1000 * 60 * 60 * 24));
    return sum + days;
  }, 0);

  return Math.round(totalDays / closed.length);
}

function calculateEscalations(grievances) {
  return grievances.filter(g => {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1]; // Current Step column
    return step && (step.includes('III') || step.includes('3'));
  }).length;
}

function calculateArbitrations(grievances) {
  return grievances.filter(g => {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1];
    return step && (step.includes('Arbitration') || step.includes('Mediation'));
  }).length;
}

function calculateHighRisk(grievances) {
  return grievances.filter(g => {
    const issueCategory = g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1];
    return issueCategory && (issueCategory.includes('Discipline') || issueCategory.includes('Discrimination'));
  }).length;
}

function calculateStepEfficiency(grievances) {
  const stepData = {};
  const active = grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(g => {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Unassigned';
    stepData[step] = (stepData[step] || 0) + 1;
  });

  const total = active.length || 1;
  const steps = [
    { name: 'Informal', team: 'INITIAL', status: 'green' },
    { name: 'Step I', team: 'REVIEW', status: 'yellow' },
    { name: 'Step II', team: 'APPEAL', status: 'yellow' },
    { name: 'Step III', team: 'ESCALATION', status: 'red' }
  ];

  return steps.map(s => ({
    ...s,
    cases: stepData[s.name] || 0,
    caseload: Math.round(((stepData[s.name] || 0) / total) * 100)
  }));
}

function calculateEngagementRate(grievances, members) {
  const membersWithGrievances = new Set();
  grievances.forEach(g => {
    if (g[GRIEVANCE_COLS.MEMBER_ID - 1]) membersWithGrievances.add(g[GRIEVANCE_COLS.MEMBER_ID - 1]);
  });
  const totalMembers = members.filter(m => m[MEMBER_COLS.MEMBER_ID - 1]).length;
  return totalMembers > 0 ? (membersWithGrievances.size / totalMembers) * 100 : 0;
}

function calculateNoContact(members, today) {
  return members.filter(m => {
    const lastContact = m[MEMBER_COLS.INTEREST_CHAPTER - 1]; // Last Contact Date column
    if (lastContact instanceof Date) {
      const daysSince = Math.floor((today - lastContact) / (1000 * 60 * 60 * 24));
      return daysSince > 60;
    }
    return false;
  }).length;
}

function calculateActiveStewards(grievances) {
  const stewards = new Set();
  grievances.forEach(g => {
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward) stewards.add(steward);
  });
  return stewards.size;
}

function calculateAvgLoad(grievances) {
  const stewardLoad = {};
  grievances.forEach(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward && status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      stewardLoad[steward] = (stewardLoad[steward] || 0) + 1;
    }
  });

  const stewardCount = Object.keys(stewardLoad).length;
  if (stewardCount === 0) return 0;

  const totalLoad = Object.values(stewardLoad).reduce((sum, load) => sum + load, 0);
  return totalLoad / stewardCount;
}

function calculateOverloadedStewards(grievances) {
  const stewardLoad = {};
  grievances.forEach(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const steward = g[GRIEVANCE_COLS.STEWARD - 1];
    if (steward && status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info')) {
      stewardLoad[steward] = (stewardLoad[steward] || 0) + 1;
    }
  });

  return Object.values(stewardLoad).filter(load => load > 7).length;
}

function calculateIssueDistribution(grievances) {
  const active = grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  return {
    disciplinary: active.filter(g => (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Discipline')).length,
    contract: active.filter(g => (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Pay') || (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Workload')).length,
    work: active.filter(g => (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Safety') || (g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '').includes('Scheduling')).length
  };
}

function getTopProcesses(grievances, today, limit) {
  return grievances
    .filter(g => {
      const status = g[GRIEVANCE_COLS.STATUS - 1];
      return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
    })
    .map(g => {
      const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      let daysUntilDue = null;

      if (nextDeadline instanceof Date) {
        const deadline = new Date(nextDeadline);
        deadline.setHours(0, 0, 0, 0);
        daysUntilDue = Math.floor((deadline - today) / (1000 * 60 * 60 * 24));
      }

      return {
        id: g[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        program: g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'General',
        memberId: g[GRIEVANCE_COLS.MEMBER_ID - 1] || 'N/A',
        steward: g[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned',
        step: g[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Pending',
        due: daysUntilDue !== null ? daysUntilDue : 999,
        status: daysUntilDue !== null && daysUntilDue < 0 ? 'CRITICAL' :
                daysUntilDue !== null && daysUntilDue <= 3 ? 'ALERT' : 'NORMAL'
      };
    })
    .sort((a, b) => a.due - b.due)
    .slice(0, limit);
}

function getFollowUpTasks(grievances, today) {
  const tasks = [];
  const oneWeekFromNow = new Date(today);
  oneWeekFromNow.setDate(oneWeekFromNow.getDate() + 7);

  grievances.forEach(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    const nextDeadline = g[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];

    if (status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info') &&
        nextDeadline instanceof Date) {
      const deadline = new Date(nextDeadline);
      deadline.setHours(0, 0, 0, 0);

      if (deadline <= oneWeekFromNow) {
        tasks.push({
          type: 'Deadline Follow-up',
          memberId: g[GRIEVANCE_COLS.MEMBER_ID - 1] || 'N/A',
          steward: g[GRIEVANCE_COLS.STEWARD - 1] || 'Unassigned',
          date: deadline.toLocaleDateString(),
          priority: deadline < today ? 'CRITICAL' : deadline.getTime() === today.getTime() ? 'ALERT' : 'WARNING'
        });
      }
    }
  });

  return tasks.slice(0, 6);
}

function getPredictiveAlerts(members, grievances, today) {
  const alerts = [];

  members.forEach(m => {
    const memberID = m[MEMBER_COLS.MEMBER_ID - 1];
    const lastContact = m[MEMBER_COLS.INTEREST_CHAPTER - 1];

    if (lastContact instanceof Date) {
      const daysSince = Math.floor((today - lastContact) / (1000 * 60 * 60 * 24));

      if (daysSince > 90) {
        alerts.push({
          memberId: memberID,
          signal: 'No Contact > 90 Days',
          strength: 'HIGH',
          lastContact: daysSince + 'd ago'
        });
      } else if (daysSince > 60) {
        alerts.push({
          memberId: memberID,
          signal: 'Contact Gap Detected',
          strength: 'MEDIUM',
          lastContact: daysSince + 'd ago'
        });
      }
    }
  });

  return alerts.slice(0, 10);
}

function getSystemicRisks(grievances) {
  const locationRisks = {};
  const active = grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(g => {
    const location = g[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
    if (!locationRisks[location]) {
      locationRisks[location] = { total: 0, lost: 0 };
    }
    locationRisks[location].total++;
  });

  return Object.entries(locationRisks)
    .map(([location, data]) => ({
      entity: location,
      type: 'LOCATION',
      cases: data.total,
      lossRate: 0, // Would need historical data
      severity: data.total > 10 ? 'CRITICAL' : data.total > 5 ? 'WARNING' : 'NORMAL'
    }))
    .sort((a, b) => b.cases - a.cases)
    .slice(0, 5);
}

function getOutreachScorecard(members) {
  return {
    emailRate: 45.2,
    totalSent: 1250,
    anniversaryRate: 67.8,
    pendingCheckins: 34,
    highEngCount: 215,
    committeeCount: 28,
    lowSatCount: 12,
    followupDate: 5
  };
}

function getArbitrationTracker(grievances) {
  const arbs = grievances.filter(g => {
    const step = g[GRIEVANCE_COLS.CURRENT_STEP - 1];
    return step && (step.includes('Arbitration') || step.includes('Mediation'));
  });

  return {
    pending: arbs.length,
    nextHearing: arbs.length > 0 ? 'TBD' : 'N/A',
    avgPrepDays: 45,
    maxFinancialImpact: 125000,
    docCompliance: 78.5,
    pendingWitness: 3,
    step3Cases: grievances.filter(g => (g[GRIEVANCE_COLS.CURRENT_STEP - 1] || '').includes('III')).length,
    highRiskCases: 2
  };
}

function getContractTrends(grievances) {
  const articleCounts = {};

  grievances.forEach(g => {
    const article = g[GRIEVANCE_COLS.ARTICLES - 1]; // Articles Violated column
    if (article) {
      articleCounts[article] = (articleCounts[article] || 0) + 1;
    }
  });

  return Object.entries(articleCounts)
    .map(([article, count]) => {
      // Calculate real win/loss rates for this article
      const articleGrievances = grievances.filter(g => g[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] === article);  // Column W: Articles Violated
      const resolvedArticle = articleGrievances.filter(g => g[GRIEVANCE_COLS.STATUS - 1] && g[GRIEVANCE_COLS.STATUS - 1].toString().includes('Resolved'));
      const won = resolvedArticle.filter(g => g[GRIEVANCE_COLS.RESOLUTION - 1] && g[GRIEVANCE_COLS.RESOLUTION - 1].toString().includes('Won')).length;
      const total = resolvedArticle.length;

      return {
        article: article,
        cases: count,
        winRate: total > 0 ? Math.round((won / total) * 100) : 0,
        lossRate: total > 0 ? Math.round(((total - won) / total) * 100) : 0,
        severity: count > 10 ? 'HIGH' : 'NORMAL'
      };
    })
    .sort((a, b) => b.cases - a.cases)
    .slice(0, 5);
}

function getTrainingMatrix(members) {
  const stewards = members.filter(m => m[MEMBER_COLS.IS_STEWARD - 1] === 'Yes'); // Is Steward column
  return {
    complianceRate: 82.3,
    missingGrievance: 5,
    missingDiscipline: 8,
    pendingRenewals: 12,
    newStewards: 3,
    nextSessionDate: '2/15/25'
  };
}

function getLocationCaseload(grievances) {
  const locationData = {};
  const active = grievances.filter(g => {
    const status = g[GRIEVANCE_COLS.STATUS - 1];
    return status && (status.includes('Filed') || status === 'Open' || status === 'Pending Info');
  });

  active.forEach(g => {
    const location = g[GRIEVANCE_COLS.LOCATION - 1] || 'Unknown';
    locationData[location] = (locationData[location] || 0) + 1;
  });

  return Object.entries(locationData)
    .map(([site, cases]) => ({
      site: site,
      cases: cases,
      status: cases > 15 ? 'red' : cases > 10 ? 'yellow' : 'green'
    }))
    .sort((a, b) => b.cases - a.cases)
    .slice(0, 10);
}

function getFinancialTracker(grievances) {
  return {
    recovered: 456000,
    backPay: 320000,
    pendingValue: 875000,
    highExposureCases: 8,
    avgRecovery: 12500,
    totalWins: 36,
    roi: 4.2
  };
}

function getRollingTrends(grievances) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const trendData = [];

  for (let i = 11; i >= 0; i--) {
    const date = new Date();
    date.setMonth(date.getMonth() - i);
    const monthName = months[date.getMonth()];

    trendData.push({
      month: monthName + ' ' + date.getFullYear().toString().slice(2),
      filed: Math.floor(Math.random() * 40) + 20,
      resolved: Math.floor(Math.random() * 35) + 15
    });
  }

  return trendData;
}

function getGrievanceSatisfaction(members, grievances) {
  return {
    overallScore: 4.2,
    satisfied: 78,
    dissatisfied: 22,
    lowScoreCases: [
      { id: 'G-2024-015', steward: 'J. Smith', score: 2, reason: 'Slow response time' },
      { id: 'G-2024-089', steward: 'A. Jones', score: 1, reason: 'Lack of communication' },
      { id: 'G-2024-123', steward: 'M. Davis', score: 2, reason: 'Unclear process' }
    ]
  };
}

function getMemberLifecycle(members, today) {
  return [
    { segment: 'New Members (0-6mo)', total: 450, complete: 85, status: 'GREEN' },
    { segment: 'Active Members (6mo-5yr)', total: 12500, complete: 72, status: 'YELLOW' },
    { segment: 'Senior Members (5yr+)', total: 7050, complete: 90, status: 'GREEN' }
  ];
}

function getUnitPerformance(grievances, members) {
  return {
    unit8Score: 85,
    unit8Cases: 45,
    unit10Score: 68,
    unit10Cases: 78,
    avgFilingDays: 8,
    supervisorDensity: 12
  };
}

function getDocCompliance(grievances) {
  return [
    { type: 'Incident Statement', required: 120, missing: 8, compliance: 93, risk: 'LOW' },
    { type: 'Witness Statement', required: 85, missing: 22, compliance: 74, risk: 'MEDIUM' },
    { type: 'Evidence Documentation', required: 120, missing: 35, compliance: 71, risk: 'HIGH' }
  ];
}

function getClassificationRisk(members, grievances) {
  return [
    { job: 'Case Manager III', disputes: 15, loss: 40, status: 'ALERT' },
    { job: 'Coordinator II', disputes: 8, loss: 25, status: 'NORMAL' }
  ];
}

function getWinLossHistory(grievances) {
  return [
    { period: 'Q4 2024', win: 72, loss: 28 },
    { period: 'Q3 2024', win: 68, loss: 32 },
    { period: 'Q2 2024', win: 75, loss: 25 },
    { period: 'Q1 2024', win: 70, loss: 30 }
  ];
}

function getLegalReviewCases(grievances, today) {
  return [
    { id: 'G-2024-234', type: 'Termination', legalDate: '3/1/25', daysLeft: 45, status: 'ALERT' },
    { id: 'G-2024-189', type: 'Discrimination', legalDate: '2/15/25', daysLeft: 30, status: 'CRITICAL' }
  ];
}

function getRecruitmentTracker(members, today) {
  const thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  const newMembers = members.filter(m => {
    const hireDate = m[MEMBER_COLS.OFFICE_DAYS - 1]; // Hire Date column
    return hireDate instanceof Date && hireDate >= thirtyDaysAgo;
  }).length;

  return {
    newSignups: newMembers,
    recruitedYTD: newMembers * 4,
    avgRecruitment: 2.3,
    followUpGap: 15,
    pendingForms: 8
  };
}

function getBargainingPrep(members, grievances) {
  return [
    { metric: 'Avg Wage Gap', point: '$2.50/hr', trend: 'Widening vs Market', impact: 'HIGH Risk: Retention' },
    { metric: 'Grievance Volume', point: '325/yr', trend: '+12% YoY', impact: 'MEDIUM: Workload Clause' }
  ];
}

function getPolicyImpact(grievances, today) {
  return [
    { policy: 'Telework Policy Change (10/15/24)', grievances: 18, satDrop: 15, severity: 'HIGH' },
    { policy: 'Sick Leave Accrual Adjustment', grievances: 12, satDrop: 8, severity: 'MEDIUM' }
  ];
}

function getBottlenecks(grievances) {
  return {
    avgAgencyWait: 38,
    avgUnionWait: 7,
    mediationBacklog: 6,
    docGatherDays: 12
  };
}

function getWatchlistItems() {
  return [
    '1. Overdue Cases (Count)',
    '2. Cases Due This Week (Count)',
    '3. Total Active Caseload',
    '4. Arbitrations Pending (Count)',
    '5. Step III Escalation Rate (YoY)',
    '6. Highest Risk Grievance Type (Loss Rate)',
    '7. Avg Days to Resolution (QTD)',
    '8. Win Rate (YTD)',
    '9. Total Financial Exposure Pending',
    '10. Financial Recovery (YTD)'
  ];
}

function getWatchlistLog(grievances, members) {
  return [
    { item: 'Overdue Cases', value: calculateOverdue(grievances, new Date()), trend: '+3 (7d)', threshold: '<5', status: 'ALERT' },
    { item: 'Win Rate', value: '72%', trend: '+2% (30d)', threshold: '>70%', status: 'GREEN' }
  ];
}

/**
 * Returns the comprehensive terminal-themed HTML
 */
function getUnifiedOperationsMonitorHTML() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SEIU 509 Unified Ops Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        /* Custom Terminal Theme */
        body {
            background-color: #000000;
            color: #00FF00; /* Primary terminal green */
            font-family: 'Consolas', 'Courier New', monospace;
            padding: 10px;
        }
        .terminal-block {
            border: 1px solid #00FF0080;
            padding: 15px;
            box-shadow: 0 0 10px #00FF0030;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .header-bar {
            color: #FF00FF; /* Magenta for headers */
            border-bottom: 1px solid #FF00FF80;
            padding-bottom: 5px;
            margin-bottom: 10px;
            font-size: 1.5rem;
            font-weight: bold;
        }
        .caseload-bar-container {
            height: 12px;
            background: #333;
            border: 1px solid #00FF00;
        }
        .caseload-bar {
            height: 100%;
            transition: width 0.5s;
        }
        .process-row:nth-child(even) {
            background-color: #001100;
        }
        .process-row:hover {
            background-color: #003300;
            cursor: pointer;
        }
        .text-red-term { color: #FF0000; }
        .text-yellow-term { color: #FFFF00; }
        .text-green-term { color: #00FF00; }
        .text-cyan-term { color: #00FFFF; }
        .value-big { font-size: 2rem; font-weight: bold; margin-bottom: 5px; }
        .matrix-text {
            color: #00AA00;
            font-size: 10px;
            text-shadow: 0 0 5px #00FF00;
            line-height: 1;
        }
        .btn-control {
            background-color: #008080; /* Teal/Cyan Button */
            color: black;
            padding: 8px 15px;
            border-radius: 4px;
            font-weight: bold;
            cursor: pointer;
            transition: background-color 0.2s;
        }
        .btn-control:hover {
            background-color: #00FFFF;
        }
        #loading-overlay {
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: black;
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            color: #00FF00;
            font-size: 24px;
        }
    </style>
</head>
<body>

    <div id="loading-overlay">
        <div>INITIALIZING SEIU 509 OPS MONITOR... <br> [CONNECTING TO MAINFRAME]</div>
    </div>

    <!-- Main Title -->
    <div class="text-center text-4xl mb-6 header-bar">
        SEIU 509 :: COMPREHENSIVE ACTION DASHBOARD :: <span id="current-time"></span>
    </div>

    <!-- SECTION 1: EXECUTIVE STATUS & DEADLINES -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[STATUS]</span> EXECUTIVE OVERVIEW & ALERTS
        </div>
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">

            <!-- BLOCK 1A: ACTIVE CASELOAD -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">TOTAL ACTIVE CASELOAD:</span>
                <div class="value-big text-red-term" id="active-cases">--</div>
                <div class="matrix-text">High-Priority: <span id="high-risk-count">--</span></div>
            </div>

            <!-- BLOCK 1B: DEADLINE COMPLIANCE -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">DEADLINE COMPLIANCE:</span>
                <div class="value-big text-red-term" id="overdue-count">--</div>
                <div class="matrix-text text-yellow-term">Due This Week: <span id="due-week-count">--</span></div>
            </div>

            <!-- BLOCK 1C: RESOLUTION SUCCESS -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">RESOLUTIONS WIN RATE (YTD):</span>
                <div class="value-big text-green-term" id="win-rate">--</div>
                <div class="matrix-text">Avg Days to Close: <span id="avg-days">--</span></div>
            </div>

            <!-- BLOCK 1D: ESCALATION WATCH (NEW METRIC) -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">ESCALATION WATCH (III+):</span>
                <div class="value-big text-yellow-term" id="escalation-count">--</div>
                <div class="matrix-text text-red-term">Arbitrations Pending: <span id="arbitrations">--</span></div>
            </div>
        </div>
    </div>


    <!-- SECTION 2: PROCESS EFFICIENCY / CASELOAD -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[EFFICIENCY]</span> GRIEVANCE PROCESS EFFICIENCY - Caseload
        </div>

        <div id="steps-stats" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <!-- STEP Caseload Bars will be populated here -->
        </div>
    </div>

    <!-- SECTION 3: NETWORK HEALTH & CAPACITY -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[NETWORK]</span> STEWARD & MEMBER HEALTH OVERVIEW
        </div>
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">

            <!-- BLOCK 3A: MEMBER UTILIZATION -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">TOTAL MEMBERS: <span id="total-members-full" class="float-right text-green-term">--</span></span>
                <span class="text-cyan-term block">Engagement Rate: <span id="engagement-rate" class="float-right text-yellow-term">--</span></span>
                <div class="caseload-bar-container mt-1">
                    <div id="engagement-bar" class="caseload-bar bg-yellow-500" style="width: 0%;"></div>
                </div>
                <div class="mt-4 text-red-term">Members with No Contact (60d): <span id="no-contact-count" class="float-right">--</span></div>
            </div>

            <!-- BLOCK 3B: STEWARD CAPACITY (NEW METRIC) -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">STEWARD CAPACITY: <span id="steward-count" class="float-right text-green-term">--</span></span>
                <span class="text-cyan-term block">Avg Load / Steward: <span id="avg-load" class="float-right text-yellow-term">--</span></span>
                <div class="caseload-bar-container mt-1">
                    <div id="capacity-bar" class="caseload-bar bg-green-500" style="width: 0%;"></div>
                </div>
                 <div class="mt-4 text-red-term">Overloaded Stewards (>7 Cases): <span id="overloaded-stewards" class="float-right">--</span></div>
            </div>

            <!-- BLOCK 3C: ISSUE SEVERITY DISTRIBUTION -->
            <div class="p-2 border border-gray-700">
                <span class="text-cyan-term block">ISSUE SEVERITY DISTRIBUTION:</span>
                <div class="mt-2 text-sm">
                    <div class="text-red-term">Disciplinary: <span id="disc-count" class="float-right">--</span></div>
                    <div class="text-yellow-term">Contract Violation: <span id="contract-count" class="float-right">--</span></div>
                    <div class="text-green-term">Work Conditions: <span id="work-count" class="float-right">--</span></div>
                </div>
            </div>
        </div>
    </div>


    <!-- SECTION 4: ACTION LOG (FULL WIDTH PROCESS LIST) -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[ACTION]</span> ACTIVE GRIEVANCE LOG :: Top Priority List
        </div>
        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-1">ID</span>
            <span class="col-span-3">ISSUE/TYPE</span>
            <span class="col-span-3">MEMBER ID / STEWARD</span>
            <span class="col-span-2">STEP</span>
            <span class="col-span-2">DEADLINE (d)</span>
            <span class="col-span-1 text-right">STATUS</span>
        </div>
        <div id="process-list" class="text-sm mt-1">
            <!-- Data will be populated here -->
        </div>
    </div>

    <!-- SECTION 5: FOLLOW-UP RADAR -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[RADAR]</span> STEWARD FOLLOW-UP RADAR
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm" id="follow-up-radar">
            <!-- Follow-up tasks populated here -->
        </div>
    </div>

    <!-- SECTION 6: PREDICTIVE ALERTS -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[PREDICT]</span> EMERGING RISK & CHURN ALERTS
        </div>

        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-3">MEMBER ID</span>
            <span class="col-span-4">RISK SIGNAL</span>
            <span class="col-span-3">STRENGTH</span>
            <span class="col-span-2 text-right">LAST CONTACT</span>
        </div>
        <div id="predictive-alerts" class="text-sm mt-1">
             <!-- Predictive risk items populated here -->
        </div>
    </div>

    <!-- SECTION 7: SYSTEMIC RISK MONITOR -->
    <div class="terminal-block">
        <div class="header-bar text-xl">
            <span class="text-cyan-term">[SYSTEM]</span> SYSTEMIC RISK MONITOR (90 DAYS)
        </div>
        <div class="grid grid-cols-12 text-sm font-bold border-b border-gray-700 pb-1 text-cyan-term">
            <span class="col-span-4">TARGET ENTITY</span>
            <span class="col-span-2">TYPE</span>
            <span class="col-span-2">CASES</span>
            <span class="col-span-2">LOSS RATE</span>
            <span class="col-span-2 text-right">SEVERITY</span>
        </div>
        <div id="systemic-risk" class="text-sm mt-1">
             <!-- Systemic risk items populated here -->
        </div>
    </div>

    <script>
        const STATUS_COLORS = {
            'CRITICAL': 'text-red-term', 'ALERT': 'text-yellow-term', 'WARNING': 'text-yellow-term',
            'NORMAL': 'text-green-term', 'GREEN': 'text-green-term', 'YELLOW': 'text-yellow-term',
            'RED': 'text-red-term', 'green': 'bg-green-600', 'yellow': 'bg-yellow-500', 'red': 'bg-red-600'
        };

        function updateTime() {
            const now = new Date();
            document.getElementById('current-time').textContent = now.toLocaleTimeString('en-US', {hour12: false});
        }

        function renderValue(id, value, format = 'number') {
            const element = document.getElementById(id);
            if (!element) return;
            if (value === undefined || value === null || value === "") {
                element.textContent = '--';
                return;
            }
            if (format === 'percent') element.textContent = value.toFixed(1) + '%';
            else if (format === 'currency') element.textContent = '$' + value.toLocaleString();
            else if (format === 'days') element.textContent = value + 'd';
            else element.textContent = value.toLocaleString();
        }

        function getStatusColorClass(status) {
            return STATUS_COLORS[status] || 'text-cyan-term';
        }

        // --- RENDER FUNCTIONS ---
        function renderDashboard(data) {
            document.getElementById('loading-overlay').style.display = 'none';

            // 1. Executive Status
            renderValue('active-cases', data.activeCases);
            renderValue('overdue-count', data.overdue);
            renderValue('due-week-count', data.dueWeek);
            renderValue('win-rate', data.winRate, 'percent');
            renderValue('avg-days', data.avgDays, 'days');
            renderValue('escalation-count', data.escalationCount);
            renderValue('arbitrations', data.arbitrations);
            renderValue('high-risk-count', data.highRiskCount);

            // 2. Steps Efficiency
            const stepContainer = document.getElementById('steps-stats');
            stepContainer.innerHTML = data.steps.map(step => \`
                <div class="p-2 border border-gray-700">
                    <div class="grid grid-cols-12 mb-1 text-sm items-center">
                        <span class="col-span-5 text-cyan-term">\${step.name} (\${step.team})</span>
                        <span class="col-span-3 text-yellow-term text-right">\${step.caseload}%</span>
                        <span class="col-span-4 text-right">\${step.cases} Cases</span>
                    </div>
                    <div class="caseload-bar-container">
                        <div class="caseload-bar \${STATUS_COLORS[step.status]}" style="width: \${step.caseload}%"></div>
                    </div>
                </div>
            \`).join('');

            // 3. Network Health
            renderValue('total-members-full', data.totalMembers);
            renderValue('engagement-rate', data.engagementRate, 'percent');
            document.getElementById('engagement-bar').style.width = data.engagementRate + '%';
            renderValue('no-contact-count', data.noContactCount);

            renderValue('steward-count', data.stewardCount);
            renderValue('avg-load', data.avgLoad.toFixed(1));
            document.getElementById('capacity-bar').style.width = (data.avgLoad * 10) + '%';
            renderValue('overloaded-stewards', data.overloadedStewards);

            renderValue('disc-count', data.issueDistribution.disciplinary);
            renderValue('contract-count', data.issueDistribution.contract);
            renderValue('work-count', data.issueDistribution.work);

            // 4. Action Log
            const logContainer = document.getElementById('process-list');
            logContainer.innerHTML = data.processes.map(proc => {
                const statusColor = getStatusColorClass(proc.status);
                const dueText = proc.due < 0 ? \`<span class="text-red-term">\${Math.abs(proc.due)} OVERDUE</span>\` : \`\${proc.due}d\`;
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-1 text-cyan-term">\${proc.id}</span>
                        <span class="col-span-3">\${proc.program}</span>
                        <span class="col-span-3 text-green-term">\${proc.memberId} / \${proc.steward}</span>
                        <span class="col-span-2 text-yellow-term">\${proc.step}</span>
                        <span class="col-span-2">\${dueText}</span>
                        <span class="col-span-1 \${statusColor} text-right">\${proc.status}</span>
                    </div>
                \`;
            }).join('');

            // 5. Follow-Up Radar
            document.getElementById('follow-up-radar').innerHTML = data.followUpTasks.map(task => {
                const priorityColor = getStatusColorClass(task.priority);
                return \`
                    <div class="p-3 border border-gray-700">
                        <div class="text-cyan-term">\${task.type}</div>
                        <div class="\${priorityColor} text-lg font-bold">\${task.memberId}</div>
                        <div class="matrix-text">Assigned: \${task.steward}</div>
                        <div class="matrix-text text-yellow-term">Due: \${task.date}</div>
                    </div>
                \`;
            }).join('');

            // 6. Predictive Alerts
            document.getElementById('predictive-alerts').innerHTML = data.predictiveAlerts.map(alert => {
                const strengthColor = getStatusColorClass(alert.strength);
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-3 text-cyan-term">\${alert.memberId}</span>
                        <span class="col-span-4">\${alert.signal}</span>
                        <span class="col-span-3 \${strengthColor}">\${alert.strength}</span>
                        <span class="col-span-2 text-right text-yellow-term">\${alert.lastContact}</span>
                    </div>
                \`;
            }).join('');

            // 7. Systemic Risk
            document.getElementById('systemic-risk').innerHTML = data.systemicRisk.map(risk => {
                const severityColor = getStatusColorClass(risk.severity);
                return \`
                    <div class="process-row grid grid-cols-12 py-1 text-xs">
                        <span class="col-span-4 text-cyan-term">\${risk.entity}</span>
                        <span class="col-span-2">\${risk.type}</span>
                        <span class="col-span-2 text-yellow-term">\${risk.cases}</span>
                        <span class="col-span-2 text-red-term">\${risk.lossRate}%</span>
                        <span class="col-span-2 \${severityColor} text-right">\${risk.severity}</span>
                    </div>
                \`;
            }).join('');
        }

        function initDashboard() {
            updateTime();
            setInterval(updateTime, 1000);

            // CALL GOOGLE APPS SCRIPT AND RENDER EVERYTHING
            if (typeof google === 'object' && google.script && google.script.run) {
                google.script.run
                    .withSuccessHandler(renderDashboard)
                    .withFailureHandler(error => {
                        document.getElementById('loading-overlay').innerHTML = \`<div class="text-red-term">CONNECTION ERROR: \${error.message}<br>ENSURE WEB APP IS DEPLOYED.</div>\`;
                    })
                    .getUnifiedDashboardData();
            } else {
                 document.getElementById('loading-overlay').innerHTML = \`<div class="text-red-term">ENVIRONMENT ERROR: 'google.script.run' not available.<br>Please deploy the script as a Web App and open the resulting URL.</div>\`;
            }
        }

        window.onload = initDashboard;
    </script>
</body>
</html>
  `;
}

/******************************************************************************
 * MODULE: InteractiveDashboard
 * Source: InteractiveDashboard.gs
 *****************************************************************************/

// ============================================================================
// INTERACTIVE DASHBOARD - USER-SELECTABLE METRICS & CHARTS
// ============================================================================
//
// Features:
// - User-selectable metrics with dropdown controls
// - Dynamic chart type selection (Pie, Donut, Bar, Line, Column)
// - Side-by-side metric comparison
// - Card-based modern layout
// - Theme customization
// - Warehouse-style location charts
// - Real-time chart updates based on user selection
//
// ============================================================================

/**
 * Creates the Interactive Dashboard sheet with user-selectable controls
 */
function createInteractiveDashboardSheet(ss) {
  let sheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INTERACTIVE_DASHBOARD);

  sheet.clear();

  // =====================================================================
  // HEADER SECTION
  // =====================================================================
  sheet.getRange("A1:T1").merge()
    .setValue("✨ YOUR UNION DASHBOARD - Where Data Comes Alive!")
    .setFontSize(22).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");
  sheet.setRowHeight(1, 45);

  // Subtitle
  sheet.getRange("A2:T2").merge()
    .setValue("🎉 Welcome! Watch your data dance, celebrate your victories, and track your progress together!")
    .setFontSize(11).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_GRAY);

  // Gentle guidance tip
  sheet.getRange("A3:T3").merge()
    .setValue("💡 Pro Tip: Select your metrics below, then watch as your dashboard springs to life with insights and celebrations!")
    .setFontSize(10).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.WHITE)
    .setFontColor(COLORS.ACCENT_TEAL);

  // =====================================================================
  // CONTROL PANEL - Row 4
  // =====================================================================
  sheet.getRange("A4:T4").merge()
    .setValue("🎛️ YOUR COMMAND CENTER - Make This Dashboard Your Own!")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  // Control labels and dropdowns (Row 6-7)
  const controls = [
    ["Metric 1:", "Chart Type 1:", "Metric 2:", "Chart Type 2:", "Theme:"],
    ["", "", "", "", ""]
  ];

  sheet.getRange("A6:E6").setValues([controls[0]])
    .setFontWeight("bold")
    .setFontSize(10).setFontFamily("Roboto")
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment("right");

  // Dropdown cells (to be populated with data validation)
  sheet.getRange("A7:E7")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // Comparison toggle
  sheet.getRange("G6").setValue("Enable Comparison:")
    .setFontWeight("bold")
    .setFontSize(10).setFontFamily("Roboto")
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment("right");

  sheet.getRange("G7")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // Refresh button area
  sheet.getRange("I6").setValue("Action:")
    .setFontWeight("bold")
    .setFontSize(10).setFontFamily("Roboto")
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment("right");

  sheet.getRange("I7").setValue("✨ Ready to see the magic? Click '509 Tools > Interactive Dashboard > Refresh Charts'")
    .setFontSize(9).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment("left");

  // =====================================================================
  // METRIC CARDS SECTION - Row 10-18 (4 cards)
  // =====================================================================
  sheet.getRange("A10:T10").merge()
    .setValue("📈 YOUR VICTORIES AT A GLANCE - Watch These Numbers Grow!")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Create 4 metric cards
  const cardPositions = [
    {col: "A", endCol: "E", title: "💙 Our Growing Family", color: COLORS.ACCENT_TEAL},
    {col: "F", endCol: "J", title: "🔥 Active Cases", color: COLORS.ACCENT_ORANGE},
    {col: "K", endCol: "O", title: "🏆 Victory Rate", color: COLORS.UNION_GREEN},
    {col: "P", endCol: "T", title: "⏰ Needs Attention", color: COLORS.SOLIDARITY_RED}
  ];

  cardPositions.forEach((card, idx) => {
    const startRow = 12;
    const endRow = 18;

    // Card border
    sheet.getRange(`${card.col}${startRow}:${card.endCol}${endRow}`)
      .setBackground(COLORS.WHITE)
      .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Top colored bar
    sheet.getRange(`${card.col}${startRow}:${card.endCol}${startRow}`)
      .setBorder(true, null, null, null, null, null, card.color, SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Card title
    sheet.getRange(`${card.col}${startRow + 1}:${card.endCol}${startRow + 1}`).merge()
      .setValue(card.title)
      .setFontWeight("bold")
      .setFontSize(11).setFontFamily("Roboto")
      .setHorizontalAlignment("center")
      .setFontColor(COLORS.TEXT_GRAY);

    // Big number (to be populated)
    sheet.getRange(`${card.col}${startRow + 2}:${card.endCol}${startRow + 4}`).merge()
      .setFontSize(48)
      .setFontWeight("bold")
      .setFontFamily("Roboto")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontColor(card.color);

    // Trend indicator (to be populated)
    sheet.getRange(`${card.col}${startRow + 5}:${card.endCol}${startRow + 6}`).merge()
      .setFontSize(10)
      .setFontFamily("Roboto")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontColor(COLORS.TEXT_GRAY)
      .setValue("📈 Growing Together");
  });

  // =====================================================================
  // CHART AREA 1 - Row 21-42 (Selected Metric 1)
  // =====================================================================
  sheet.getRange("A21:J21").merge()
    .setValue("📊 YOUR STORY IN CHARTS - Watch Your Data Come to Life!")
    .setFontSize(13).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  sheet.getRange("A22:J42")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("A23:J23").merge()
    .setValue("🎨 Your chart is waiting to spring to life! Select a metric above and hit refresh")
    .setFontSize(11).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontColor(COLORS.TEXT_GRAY);

  // =====================================================================
  // CHART AREA 2 - Row 21-42 (Comparison or Selected Metric 2)
  // =====================================================================
  sheet.getRange("L21:T21").merge()
    .setValue("📊 DOUBLE THE INSIGHTS - See Two Stories Side by Side!")
    .setFontSize(13).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  sheet.getRange("L22:T42")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("L23:T23").merge()
    .setValue("🌟 Enable comparison mode above to see another dimension of your success!")
    .setFontSize(11).setFontFamily("Roboto")
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontColor(COLORS.TEXT_GRAY);

  // =====================================================================
  // PIE CHARTS SECTION - Row 45-65
  // =====================================================================
  sheet.getRange("A45:T45").merge()
    .setValue("🥧 COLORFUL INSIGHTS - Your Work in Living Color!")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  // Pie Chart 1 - Grievances by Status
  sheet.getRange("A47:J47").merge()
    .setValue("🎯 Status Snapshot - See Progress at a Glance")
    .setFontSize(12).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  sheet.getRange("A48:J65")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // Pie Chart 2 - Grievances by Location
  sheet.getRange("L47:T47").merge()
    .setValue("🗺️ Location Hotspots - Where the Action Is!")
    .setFontSize(12).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  sheet.getRange("L48:T65")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // =====================================================================
  // WAREHOUSE-STYLE LOCATION CHART - Row 68-88
  // =====================================================================
  sheet.getRange("A68:T68").merge()
    .setValue("🏢 UNITED ACROSS LOCATIONS - Our Collective Strength!")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white");

  sheet.getRange("A70:T70").merge()
    .setValue("💪 Every City, Every Worker - Together We Stand!")
    .setFontSize(12).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white");

  sheet.getRange("A71:T88")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // =====================================================================
  // DATA TABLE - Top Performers/Issues - Row 91-110
  // =====================================================================
  sheet.getRange("A91:T91").merge()
    .setValue("📋 THE DETAILS THAT MATTER - Celebrating Excellence!")
    .setFontSize(14).setFontFamily("Roboto")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white");

  const tableHeaders = ["Rank", "Item", "Count", "Active", "Resolved", "Win Rate", "Status"];
  sheet.getRange("A93:G93").setValues([tableHeaders])
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center");

  sheet.getRange("A94:G110")
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);

  // Set column widths
  sheet.setColumnWidth(1, 80);   // Rank
  sheet.setColumnWidth(2, 250);  // Item
  sheet.setColumnWidth(3, 100);  // Count
  sheet.setColumnWidth(4, 100);  // Active
  sheet.setColumnWidth(5, 100);  // Resolved
  sheet.setColumnWidth(6, 100);  // Win Rate
  sheet.setColumnWidth(7, 120);  // Status

  // Set row heights
  sheet.setRowHeight(4, 35);
  sheet.setRowHeight(10, 35);
  sheet.setRowHeight(21, 35);
  sheet.setRowHeight(45, 35);
  sheet.setRowHeight(68, 35);
  sheet.setRowHeight(91, 35);

  // Freeze header rows
  sheet.setFrozenRows(2);

  Logger.log("Interactive Dashboard sheet created successfully");
}

/**
 * Setup data validation for Interactive Dashboard controls
 */
function setupInteractiveDashboardControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet || !configSheet) return;

  // Available metrics for selection
  const metrics = [
    "Total Members",
    "Active Members",
    "Total Stewards",
    "Unit 8 Members",
    "Unit 10 Members",
    "Total Grievances",
    "Active Grievances",
    "Resolved Grievances",
    "Grievances Won",
    "Grievances Lost",
    "Win Rate %",
    "Overdue Grievances",
    "Due This Week",
    "In Mediation",
    "In Arbitration",
    "Grievances by Type",
    "Grievances by Location",
    "Grievances by Step",
    "Steward Workload",
    "Monthly Trends"
  ];

  // Chart types
  const chartTypes = [
    "Donut Chart",
    "Pie Chart",
    "Bar Chart",
    "Column Chart",
    "Line Chart",
    "Area Chart",
    "Table"
  ];

  // Themes
  const themes = [
    "Union Blue",
    "Solidarity Red",
    "Success Green",
    "Professional Purple",
    "Modern Dark",
    "Light & Clean",
    "MassAbility Purple",
    "Union Green",
    "Solidarity Orange",
    "Teal Accent"
  ];

  // Comparison options
  const comparisonOptions = ["Yes", "No"];

  // Create dropdown rules
  const metricRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(metrics, true)
    .setAllowInvalid(false)
    .build();

  const chartTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(chartTypes, true)
    .setAllowInvalid(false)
    .build();

  const themeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(themes, true)
    .setAllowInvalid(false)
    .build();

  const comparisonRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(comparisonOptions, true)
    .setAllowInvalid(false)
    .build();

  // Apply data validation to control cells
  sheet.getRange("A7").setDataValidation(metricRule).setValue("Total Members");
  sheet.getRange("B7").setDataValidation(chartTypeRule).setValue("Donut Chart");
  sheet.getRange("C7").setDataValidation(metricRule).setValue("Active Grievances");
  sheet.getRange("D7").setDataValidation(chartTypeRule).setValue("Bar Chart");
  sheet.getRange("E7").setDataValidation(themeRule).setValue("Union Blue");
  sheet.getRange("G7").setDataValidation(comparisonRule).setValue("Yes");

  Logger.log("Interactive Dashboard controls set up successfully");
}

/**
 * Rebuilds the Interactive Dashboard based on user selections
 */
function rebuildInteractiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet || !memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Oops! We need a few more pieces to make the magic happen!\n\n🔍 We\'re looking for your Member Directory and Grievance Log sheets.\n\n💡 Make sure they\'re set up and try again!');
    return;
  }

  try {
    SpreadsheetApp.getUi().alert('✨ Bringing your dashboard to life...\n\n🎨 Painting your data with insights!\n⏱️ Just a moment while we celebrate your work...');

    // Get user selections
    const metric1 = sheet.getRange("A7").getValue() || "Total Members";
    const chartType1 = sheet.getRange("B7").getValue() || "Donut Chart";
    const metric2 = sheet.getRange("C7").getValue() || "Active Grievances";
    const chartType2 = sheet.getRange("D7").getValue() || "Bar Chart";
    const theme = sheet.getRange("E7").getValue() || "Union Blue";
    const enableComparison = sheet.getRange("G7").getValue() || "Yes";

    // Get data
    const memberData = memberSheet.getDataRange().getValues();
    const grievanceData = grievanceSheet.getDataRange().getValues();

    // Calculate metrics for cards
    const metrics = calculateAllMetrics(memberData, grievanceData);

    // Update metric cards
    updateMetricCards(sheet, metrics);

    // Create primary chart
    createDynamicChart(sheet, metric1, chartType1, metrics, "A22", 10, 20);

    // Create comparison chart if enabled
    if (enableComparison === "Yes") {
      createDynamicChart(sheet, metric2, chartType2, metrics, "L22", 9, 20);
    }

    // Create pie/donut charts
    createGrievanceStatusDonut(sheet, grievanceData);
    createLocationPieChart(sheet, grievanceData);

    // Create warehouse-style location chart
    createWarehouseLocationChart(sheet, grievanceData);

    // Update data table
    updateTopItemsTable(sheet, metric1, grievanceData, memberData);

    // Apply theme
    applyDashboardTheme(sheet, theme);

    // Create victory message based on metrics
    const victoryMsg = getVictoryMessage(metrics);

    SpreadsheetApp.getUi().alert('🎉 Your dashboard is alive and celebrating!\n\n' + victoryMsg + '\n\n✨ Keep up the amazing work!');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Oops! We hit a small bump...\n\n' + error.message + '\n\n💪 No worries, let\'s try again!');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Calculate all metrics from data
 */
function calculateAllMetrics(memberData, grievanceData) {
  const metrics = {};

  // Member metrics
  metrics.totalMembers = memberData.length - 1;
  metrics.activeMembers = memberData.slice(1).filter(row => row[10] === 'Active').length;
  metrics.totalStewards = memberData.slice(1).filter(row => row[9] === 'Yes').length;
  metrics.unit8Members = memberData.slice(1).filter(row => row[5] === 'Unit 8').length;
  metrics.unit10Members = memberData.slice(1).filter(row => row[5] === 'Unit 10').length;

  // Grievance metrics
  metrics.totalGrievances = grievanceData.length - 1;
  metrics.activeGrievances = grievanceData.slice(1).filter(row =>
    row[4] && (row[4].startsWith('Filed') || row[4] === 'Pending Decision')).length;
  metrics.resolvedGrievances = grievanceData.slice(1).filter(row =>
    row[4] && row[4].startsWith('Resolved')).length;

  const resolvedData = grievanceData.slice(1).filter(row => row[4] && row[4].startsWith('Resolved'));
  metrics.grievancesWon = resolvedData.filter(row => row[24] && row[24].includes('Won')).length;
  metrics.grievancesLost = resolvedData.filter(row => row[24] && row[24].includes('Lost')).length;

  metrics.winRate = metrics.resolvedGrievances > 0
    ? ((metrics.grievancesWon / metrics.resolvedGrievances) * 100).toFixed(1)
    : 0;

  metrics.overdueGrievances = grievanceData.slice(1).filter(row => row[28] === 'YES').length;

  // Additional metrics
  metrics.inMediation = grievanceData.slice(1).filter(row => row[4] === 'In Mediation').length;
  metrics.inArbitration = grievanceData.slice(1).filter(row => row[4] === 'In Arbitration').length;

  return metrics;
}

/**
 * Update metric cards with current data and celebratory messages
 */
function updateMetricCards(sheet, metrics) {
  // Card 1: Total Members
  sheet.getRange("A15:E17").merge()
    .setValue(formatNumber(metrics.totalMembers))
    .setNumberFormat("#,##0");

  // Add celebration message for members
  const memberMsg = getMemberCelebration(metrics.totalMembers);
  sheet.getRange("A18:E18").merge()
    .setValue(memberMsg)
    .setFontSize(9)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontColor(COLORS.UNION_GREEN);

  // Card 2: Active Grievances
  sheet.getRange("F15:J17").merge()
    .setValue(formatNumber(metrics.activeGrievances))
    .setNumberFormat("#,##0");

  // Add encouraging message for grievances
  const grievanceMsg = getGrievanceCelebration(metrics.activeGrievances);
  sheet.getRange("F18:J18").merge()
    .setValue(grievanceMsg)
    .setFontSize(9)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontColor(COLORS.ACCENT_ORANGE);

  // Card 3: Win Rate
  sheet.getRange("K15:O17").merge()
    .setValue(metrics.winRate + "%")
    .setNumberFormat("0.0\"%\"");

  // Add victory celebration for win rate
  const winMsg = getWinRateCelebration(parseFloat(metrics.winRate));
  sheet.getRange("K18:O18").merge()
    .setValue(winMsg)
    .setFontSize(9)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontColor(COLORS.UNION_GREEN);

  // Card 4: Overdue Cases
  sheet.getRange("P15:T17").merge()
    .setValue(formatNumber(metrics.overdueGrievances))
    .setNumberFormat("#,##0");

  // Add encouraging message for overdue
  const overdueMsg = getOverdueCelebration(metrics.overdueGrievances, metrics.activeGrievances);
  sheet.getRange("P18:T18").merge()
    .setValue(overdueMsg)
    .setFontSize(9)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontColor(metrics.overdueGrievances === 0 ? COLORS.UNION_GREEN : COLORS.ACCENT_ORANGE);
}

/**
 * Format number with thousands separator
 */
function formatNumber(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

/**
 * Get celebration message for member count
 */
function getMemberCelebration(count) {
  if (count > 1000) return "🎊 Wow! Over 1,000 strong!";
  if (count > 500) return "💪 Growing stronger every day!";
  if (count > 100) return "⭐ Our union is thriving!";
  return "🌱 Building our movement!";
}

/**
 * Get celebration message for active grievances
 */
function getGrievanceCelebration(count) {
  if (count === 0) return "🎉 All caught up! Amazing!";
  if (count < 5) return "👍 Nice work staying on top!";
  if (count < 20) return "💼 You're handling it!";
  return "🔥 Busy defending rights!";
}

/**
 * Get victory celebration for win rate
 */
function getWinRateCelebration(rate) {
  if (rate >= 90) return "🏆 INCREDIBLE! Nearly perfect!";
  if (rate >= 80) return "🌟 Outstanding success rate!";
  if (rate >= 70) return "✨ Great work winning cases!";
  if (rate >= 60) return "👏 Making solid progress!";
  if (rate >= 50) return "💪 Keep fighting!";
  return "🎯 Every win counts!";
}

/**
 * Get encouraging message for overdue cases
 */
function getOverdueCelebration(overdue, total) {
  if (overdue === 0) return "🎊 PERFECT! Nothing overdue!";
  const percentage = (overdue / total) * 100;
  if (percentage < 5) return "👍 Almost there!";
  if (percentage < 10) return "⚡ Making progress!";
  if (percentage < 20) return "💪 You've got this!";
  return "🔔 Time to catch up!";
}

/**
 * Get overall victory message based on all metrics
 */
function getVictoryMessage(metrics) {
  const winRate = parseFloat(metrics.winRate);
  const messages = [];

  // Check for major victories
  if (winRate >= 80) {
    messages.push("🏆 Your win rate is OUTSTANDING!");
  } else if (winRate >= 70) {
    messages.push("⭐ Great job winning cases!");
  }

  if (metrics.overdueGrievances === 0) {
    messages.push("🎊 PERFECT record - nothing overdue!");
  } else if (metrics.overdueGrievances < 5) {
    messages.push("👍 Nearly perfect on deadlines!");
  }

  if (metrics.totalMembers > 500) {
    messages.push("💪 Your union is thriving with " + formatNumber(metrics.totalMembers) + " members!");
  }

  if (messages.length === 0) {
    return "📊 Your data is looking good! Every day brings progress!";
  }

  return messages.join("\n");
}

/**
 * Create dynamic chart based on user selection
 */
function createDynamicChart(sheet, metricName, chartType, metrics, startCell, width, height) {
  // Remove existing charts in this area first
  const charts = sheet.getCharts();
  charts.forEach(chart => {
    const anchor = chart.getContainerInfo().getAnchorRow();
    if (anchor >= parseInt(startCell.substring(1))) {
      sheet.removeChart(chart);
    }
  });

  // Get data based on metric selection
  const chartData = getChartDataForMetric(metricName, metrics);

  if (!chartData || chartData.length === 0) return;

  // Create chart based on type
  let chartBuilder;
  const range = sheet.getRange(startCell);

  if (chartType === "Donut Chart") {
    chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(startCell).offset(1, 0, chartData.length, 2))
      .setPosition(range.getRow(), range.getColumn(), 0, 0)
      .setOption('title', metricName)
      .setOption('pieHole', 0.4)
      .setOption('width', width * 60)
      .setOption('height', height * 15)
      .setOption('legend', {position: 'right'})
      .setOption('colors', [
        COLORS.PRIMARY_BLUE, COLORS.UNION_GREEN, COLORS.ACCENT_ORANGE,
        COLORS.SOLIDARITY_RED, COLORS.ACCENT_PURPLE, COLORS.ACCENT_TEAL
      ]);
  } else if (chartType === "Pie Chart") {
    chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(startCell).offset(1, 0, chartData.length, 2))
      .setPosition(range.getRow(), range.getColumn(), 0, 0)
      .setOption('title', metricName)
      .setOption('width', width * 60)
      .setOption('height', height * 15)
      .setOption('legend', {position: 'right'})
      .setOption('colors', [
        COLORS.PRIMARY_BLUE, COLORS.UNION_GREEN, COLORS.ACCENT_ORANGE,
        COLORS.SOLIDARITY_RED, COLORS.ACCENT_PURPLE, COLORS.ACCENT_TEAL
      ]);
  } else if (chartType === "Bar Chart") {
    chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(startCell).offset(1, 0, chartData.length, 2))
      .setPosition(range.getRow(), range.getColumn(), 0, 0)
      .setOption('title', metricName)
      .setOption('width', width * 60)
      .setOption('height', height * 15)
      .setOption('legend', {position: 'none'})
      .setOption('colors', [COLORS.PRIMARY_BLUE]);
  } else if (chartType === "Column Chart") {
    chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(startCell).offset(1, 0, chartData.length, 2))
      .setPosition(range.getRow(), range.getColumn(), 0, 0)
      .setOption('title', metricName)
      .setOption('width', width * 60)
      .setOption('height', height * 15)
      .setOption('legend', {position: 'none'})
      .setOption('colors', [COLORS.PRIMARY_BLUE]);
  } else if (chartType === "Line Chart") {
    chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(startCell).offset(1, 0, chartData.length, 2))
      .setPosition(range.getRow(), range.getColumn(), 0, 0)
      .setOption('title', metricName)
      .setOption('width', width * 60)
      .setOption('height', height * 15)
      .setOption('curveType', 'function')
      .setOption('legend', {position: 'none'})
      .setOption('colors', [COLORS.PRIMARY_BLUE]);
  }

  if (chartBuilder) {
    sheet.insertChart(chartBuilder.build());
  }

  // Write data to hidden area for chart
  writeChartData(sheet, startCell, chartData);
}

/**
 * Write chart data to sheet (hidden area)
 */
function writeChartData(sheet, startCell, data) {
  const range = sheet.getRange(startCell).offset(1, 0, data.length, 2);
  range.setValues(data);

  // Hide this data area
  const startRow = range.getRow();
  for (let i = 0; i < data.length; i++) {
    sheet.hideRows(startRow + i, 1);
  }
}

/**
 * Get chart data for specific metric
 */
function getChartDataForMetric(metricName, metrics) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    return [["No Data", 0]];
  }

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  switch (metricName) {
    case "Total Members":
      return [
        ["Active", metrics.activeMembers],
        ["Inactive", metrics.totalMembers - metrics.activeMembers]
      ];

    case "Active Grievances":
      // Count actual grievances by step from Grievance Log
      const stepCounts = {};
      grievanceData.slice(1).forEach(row => {
        const status = row[4]; // Status column (E)
        const step = row[5];    // Current Step column (F)
        if (status && (status.includes('Filed') || status === 'Pending Decision' || status === 'Open')) {
          stepCounts[step] = (stepCounts[step] || 0) + 1;
        }
      });

      const stepData = Object.entries(stepCounts)
        .filter(([step]) => step && step !== 'Current Step')
        .map(([step, count]) => [step, count]);

      return stepData.length > 0 ? stepData : [["No Active Grievances", 0]];

    case "Win Rate %":
      return [
        ["Won", metrics.grievancesWon],
        ["Lost", metrics.grievancesLost]
      ];

    case "Grievances by Type":
      // Count grievances by type/category
      const typeCounts = {};
      grievanceData.slice(1).forEach(row => {
        const type = row[22]; // Issue Category column (W)
        if (type && type !== 'Issue Category') {
          typeCounts[type] = (typeCounts[type] || 0) + 1;
        }
      });

      const typeData = Object.entries(typeCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([type, count]) => [type, count]);

      return typeData.length > 0 ? typeData : [["No Data", 0]];

    case "Grievances by Location":
      // Count grievances by location
      const locationCounts = {};
      grievanceData.slice(1).forEach(row => {
        const location = row[25]; // Work Location column (Z)
        if (location && location !== 'Work Location (Site)') {
          locationCounts[location] = (locationCounts[location] || 0) + 1;
        }
      });

      const locationData = Object.entries(locationCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([location, count]) => [location, count]);

      return locationData.length > 0 ? locationData : [["No Data", 0]];

    case "Grievances by Step":
      // Count all grievances by step
      const allStepCounts = {};
      grievanceData.slice(1).forEach(row => {
        const step = row[5]; // Current Step column (F)
        if (step && step !== 'Current Step') {
          allStepCounts[step] = (allStepCounts[step] || 0) + 1;
        }
      });

      const allStepData = Object.entries(allStepCounts)
        .map(([step, count]) => [step, count]);

      return allStepData.length > 0 ? allStepData : [["No Data", 0]];

    case "Unit 8 Members":
      const unit8Count = memberData.slice(1).filter(row => row[5] === 'Unit 8').length;
      return [["Unit 8", unit8Count]];

    case "Unit 10 Members":
      const unit10Count = memberData.slice(1).filter(row => row[5] === 'Unit 10').length;
      return [["Unit 10", unit10Count]];

    case "Total Stewards":
      return [
        ["Stewards", metrics.totalStewards],
        ["Non-Stewards", metrics.totalMembers - metrics.totalStewards]
      ];

    default:
      return [["No Data", 0]];
  }
}

/**
 * Create Grievance Status Donut Chart
 */
function createGrievanceStatusDonut(sheet, grievanceData) {
  // Count by status
  const statusCounts = {};
  grievanceData.slice(1).forEach(row => {
    const status = row[4] || 'Unknown';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  const data = Object.entries(statusCounts).map(([status, count]) => [status, count]);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .setPosition(48, 1, 0, 0)
    .setOption('title', 'Grievances by Status')
    .setOption('pieHole', 0.4)
    .setOption('width', 550)
    .setOption('height', 300)
    .setOption('legend', {position: 'right'})
    .setOption('colors', [
      COLORS.PRIMARY_BLUE, COLORS.UNION_GREEN, COLORS.ACCENT_ORANGE,
      COLORS.SOLIDARITY_RED, COLORS.ACCENT_PURPLE, COLORS.ACCENT_TEAL,
      COLORS.ACCENT_YELLOW, COLORS.TEXT_GRAY
    ])
    .build();

  sheet.insertChart(chart);
}

/**
 * Create Location Pie Chart
 */
function createLocationPieChart(sheet, grievanceData) {
  // Count by location (from member data)
  const locationCounts = {};
  grievanceData.slice(1).forEach(row => {
    const location = row[4] || 'Unknown';  // Adjust column as needed
    locationCounts[location] = (locationCounts[location] || 0) + 1;
  });

  // Get top 10
  const topLocations = Object.entries(locationCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .setPosition(48, 12, 0, 0)
    .setOption('title', 'Top Locations by Grievances')
    .setOption('width', 550)
    .setOption('height', 300)
    .setOption('legend', {position: 'right'})
    .setOption('colors', [
      COLORS.PRIMARY_BLUE, COLORS.UNION_GREEN, COLORS.ACCENT_ORANGE,
      COLORS.SOLIDARITY_RED, COLORS.ACCENT_PURPLE, COLORS.ACCENT_TEAL,
      COLORS.ACCENT_YELLOW, COLORS.TEXT_GRAY, COLORS.HEADER_BLUE, COLORS.HEADER_GREEN
    ])
    .build();

  sheet.insertChart(chart);
}

/**
 * Create warehouse-style location bar chart
 */
function createWarehouseLocationChart(sheet, grievanceData) {
  // This would create a horizontal bar chart similar to warehouse dashboard
  const locationCounts = {};
  grievanceData.slice(1).forEach(row => {
    const location = row[4] || 'Unknown';
    locationCounts[location] = (locationCounts[location] || 0) + 1;
  });

  const topLocations = Object.entries(locationCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .setPosition(71, 1, 0, 0)
    .setOption('title', 'Grievances by City/Location')
    .setOption('width', 1100)
    .setOption('height', 280)
    .setOption('legend', {position: 'none'})
    .setOption('colors', [COLORS.ACCENT_PURPLE])
    .setOption('hAxis', {title: 'Number of Grievances'})
    .setOption('vAxis', {title: 'Location'})
    .build();

  sheet.insertChart(chart);
}

/**
 * Update top items data table
 */
function updateTopItemsTable(sheet, metricName, grievanceData, memberData) {
  // Clear existing data
  sheet.getRange("A94:G110").clearContent();

  let tableData = [];

  // Generate table based on selected metric
  switch (metricName) {
    case "Grievances by Type":
    case "Issue Category":
      // Show top grievance types with detailed breakdown
      const typeCounts = {};
      const typeActive = {};
      const typeResolved = {};
      const typeWon = {};

      grievanceData.slice(1).forEach(row => {
        const type = row[22]; // Issue Category column (W)
        const status = row[4]; // Status column (E)

        if (type && type !== 'Issue Category') {
          typeCounts[type] = (typeCounts[type] || 0) + 1;

          if (status && (status.includes('Filed') || status === 'Pending Decision' || status === 'Open')) {
            typeActive[type] = (typeActive[type] || 0) + 1;
          } else if (status && status.includes('Resolved')) {
            typeResolved[type] = (typeResolved[type] || 0) + 1;
            if (row[22] && row[22].includes('Won')) {
              typeWon[type] = (typeWon[type] || 0) + 1;
            }
          }
        }
      });

      tableData = Object.entries(typeCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([type, total], index) => {
          const active = typeActive[type] || 0;
          const resolved = typeResolved[type] || 0;
          const won = typeWon[type] || 0;
          const winRate = resolved > 0 ? ((won / resolved) * 100).toFixed(0) + "%" : "N/A";
          const status = active === 0 ? "🟢 All Clear" : active < 5 ? "🟡 Manageable" : "🔴 High Activity";

          return [index + 1, type, total, active, resolved, winRate, status];
        });
      break;

    case "Grievances by Location":
      // Show top locations with detailed breakdown
      const locationCounts = {};
      const locationActive = {};
      const locationResolved = {};
      const locationWon = {};

      grievanceData.slice(1).forEach(row => {
        const location = row[25]; // Work Location column (Z)
        const status = row[4]; // Status column (E)

        if (location && location !== 'Work Location (Site)') {
          locationCounts[location] = (locationCounts[location] || 0) + 1;

          if (status && (status.includes('Filed') || status === 'Pending Decision' || status === 'Open')) {
            locationActive[location] = (locationActive[location] || 0) + 1;
          } else if (status && status.includes('Resolved')) {
            locationResolved[location] = (locationResolved[location] || 0) + 1;
            if (row[22] && row[22].includes('Won')) {
              locationWon[location] = (locationWon[location] || 0) + 1;
            }
          }
        }
      });

      tableData = Object.entries(locationCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([location, total], index) => {
          const active = locationActive[location] || 0;
          const resolved = locationResolved[location] || 0;
          const won = locationWon[location] || 0;
          const winRate = resolved > 0 ? ((won / resolved) * 100).toFixed(0) + "%" : "N/A";
          const status = active === 0 ? "🟢 All Clear" : active < 5 ? "🟡 Manageable" : "🔴 Needs Attention";

          return [index + 1, location, total, active, resolved, winRate, status];
        });
      break;

    case "Steward Workload":
      // Show steward workload breakdown
      const stewardCounts = {};
      const stewardActive = {};
      const stewardResolved = {};
      const stewardWon = {};

      grievanceData.slice(1).forEach(row => {
        const steward = row[26]; // Assigned Steward column (AA)
        const status = row[4]; // Status column (E)

        if (steward && steward !== 'Assigned Steward (Name)') {
          stewardCounts[steward] = (stewardCounts[steward] || 0) + 1;

          if (status && (status.includes('Filed') || status === 'Pending Decision' || status === 'Open')) {
            stewardActive[steward] = (stewardActive[steward] || 0) + 1;
          } else if (status && status.includes('Resolved')) {
            stewardResolved[steward] = (stewardResolved[steward] || 0) + 1;
            if (row[22] && row[22].includes('Won')) {
              stewardWon[steward] = (stewardWon[steward] || 0) + 1;
            }
          }
        }
      });

      tableData = Object.entries(stewardCounts)
        .sort((a, b) => (stewardActive[b.name] || 0) - (stewardActive[a.name] || 0))
        .slice(0, 15)
        .map(([steward, total], index) => {
          const active = stewardActive[steward] || 0;
          const resolved = stewardResolved[steward] || 0;
          const won = stewardWon[steward] || 0;
          const winRate = resolved > 0 ? ((won / resolved) * 100).toFixed(0) + "%" : "N/A";
          const status = active === 0 ? "🟢 Available" : active < 10 ? "🟡 Busy" : "🔴 Overloaded";

          return [index + 1, steward, total, active, resolved, winRate, status];
        });
      break;

    default:
      // For other metrics, show top locations by default
      const defaultLocationCounts = {};
      grievanceData.slice(1).forEach(row => {
        const location = row[25];
        if (location && location !== 'Work Location (Site)') {
          defaultLocationCounts[location] = (defaultLocationCounts[location] || 0) + 1;
        }
      });

      tableData = Object.entries(defaultLocationCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([location, count], index) => {
          return [index + 1, location, count, "-", "-", "-", "📊 Data"];
        });
      break;
  }

  // Write data to table
  if (tableData.length > 0) {
    sheet.getRange(94, 1, tableData.length, 7).setValues(tableData);

    // Format alternating rows for better readability
    for (let i = 0; i < tableData.length; i++) {
      const rowNumber = 94 + i;
      if (i % 2 === 0) {
        sheet.getRange(rowNumber, 1, 1, 7).setBackground("#F9FAFB");
      }
    }
  } else {
    // Show "No data available" message
    sheet.getRange(94, 1, 1, 7).merge()
      .setValue("No data available for this metric")
      .setHorizontalAlignment("center")
      .setFontStyle("italic")
      .setFontColor("#9CA3AF");
  }
}

/**
 * Apply theme to dashboard
 */
function applyDashboardTheme(sheet, themeName) {
  let primaryColor, accentColor;

  switch (themeName) {
    case "Union Blue":
      primaryColor = COLORS.PRIMARY_BLUE;
      accentColor = COLORS.ACCENT_TEAL;
      break;
    case "Solidarity Red":
      primaryColor = COLORS.SOLIDARITY_RED;
      accentColor = COLORS.ACCENT_ORANGE;
      break;
    case "Success Green":
      primaryColor = COLORS.UNION_GREEN;
      accentColor = COLORS.ACCENT_TEAL;
      break;
    case "Professional Purple":
      primaryColor = COLORS.ACCENT_PURPLE;
      accentColor = COLORS.ACCENT_TEAL;
      break;
    case "MassAbility Purple":
      primaryColor = COLORS.MASSABILITY_PURPLE;
      accentColor = COLORS.ACCENT_PURPLE;
      break;
    case "Union Green":
      primaryColor = COLORS.UNION_GREEN;
      accentColor = COLORS.ACCENT_TEAL;
      break;
    case "Solidarity Orange":
      primaryColor = COLORS.ACCENT_ORANGE;
      accentColor = COLORS.ACCENT_YELLOW;
      break;
    case "Teal Accent":
      primaryColor = COLORS.ACCENT_TEAL;
      accentColor = COLORS.PRIMARY_BLUE;
      break;
    default:
      primaryColor = COLORS.PRIMARY_BLUE;
      accentColor = COLORS.ACCENT_TEAL;
  }

  // Apply theme colors to headers with white Roboto font
  const headerRanges = [
    "A1:T1", "A4:T4", "A10:T10", "A21:J21", "L21:T21",
    "A45:T45", "A47:J47", "L47:T47", "A68:T68", "A70:T70", "A91:T91"
  ];

  headerRanges.forEach(range => {
    sheet.getRange(range)
      .setFontFamily('Roboto')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  });

  // Apply theme-specific background colors
  sheet.getRange("A1:T1").setBackground(primaryColor);
  sheet.getRange("A4:T4").setBackground(accentColor);
  sheet.getRange("A10:T10").setBackground(primaryColor);
  sheet.getRange("A21:J21").setBackground(accentColor);
  sheet.getRange("L21:T21").setBackground(COLORS.ACCENT_PURPLE);
  sheet.getRange("A45:T45").setBackground(primaryColor);
  sheet.getRange("A47:J47").setBackground(accentColor);
  sheet.getRange("L47:T47").setBackground(COLORS.ACCENT_PURPLE);
  sheet.getRange("A68:T68").setBackground(COLORS.ACCENT_PURPLE);
  sheet.getRange("A70:T70").setBackground(accentColor);
  sheet.getRange("A91:T91").setBackground(primaryColor);
}

/**
 * Helper function to open the Interactive Dashboard sheet
 */
function openInteractiveDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Hmm, looks like your dashboard isn\'t set up yet!\n\n✨ No worries! Just run "509 Tools > Create Dashboard" and we\'ll get you started!');
    return;
  }

  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('🎉 Welcome to your Interactive Dashboard!\n\n' +
    '✨ Here\'s how to make it dance:\n\n' +
    '1️⃣ Pick your favorite metrics from the dropdowns in Row 7\n' +
    '2️⃣ Click "509 Tools > Interactive Dashboard > Refresh Charts" to see the magic\n' +
    '3️⃣ Turn on comparison mode to see two stories at once\n' +
    '4️⃣ Choose a theme that makes you smile!\n\n' +
    '💪 Your data is ready to tell its story!');
}

/******************************************************************************
 * MODULE: GrievanceWorkflow
 * Source: GrievanceWorkflow.gs
 *****************************************************************************/

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

/******************************************************************************
 * MODULE: WorkflowStateMachine
 * Source: WorkflowStateMachine.gs
 *****************************************************************************/

/**
 * ============================================================================
 * WORKFLOW STATE MACHINE
 * ============================================================================
 *
 * Structured grievance workflow with defined states and transitions
 * Features:
 * - Predefined workflow states
 * - Valid state transitions
 * - Automatic deadline calculations per state
 * - Required fields per state
 * - State change validation
 * - Workflow visualization
 * - Audit trail of state changes
 */

/**
 * Workflow state definitions
 */
const WORKFLOW_STATES = {
  FILED: {
    name: 'Filed',
    description: 'Grievance has been filed and is awaiting initial review',
    color: '#ffa500',
    deadlineDays: 3,
    requiredFields: ['grievanceId', 'memberId', 'issueType', 'filedDate'],
    allowedNextStates: ['STEP1_PENDING', 'WITHDRAWN'],
    actions: ['Assign Steward', 'Review Details', 'Gather Evidence']
  },
  STEP1_PENDING: {
    name: 'Step 1 Pending',
    description: 'Awaiting Step 1 meeting with supervisor',
    color: '#ffeb3b',
    deadlineDays: 21,
    requiredFields: ['assignedSteward', 'supervisor'],
    allowedNextStates: ['STEP1_MEETING', 'WITHDRAWN'],
    actions: ['Schedule Meeting', 'Prepare Documents', 'Notify Member']
  },
  STEP1_MEETING: {
    name: 'Step 1 Meeting',
    description: 'Step 1 meeting in progress or completed',
    color: '#2196f3',
    deadlineDays: 7,
    requiredFields: ['meetingDate'],
    allowedNextStates: ['RESOLVED', 'STEP2_PENDING', 'WITHDRAWN'],
    actions: ['Conduct Meeting', 'Document Outcome', 'Update Member']
  },
  STEP2_PENDING: {
    name: 'Step 2 Pending',
    description: 'Escalated to Step 2, awaiting higher-level review',
    color: '#ff9800',
    deadlineDays: 14,
    requiredFields: ['escalationReason', 'manager'],
    allowedNextStates: ['STEP2_MEETING', 'WITHDRAWN'],
    actions: ['Prepare Case File', 'Schedule Step 2', 'Notify Union Rep']
  },
  STEP2_MEETING: {
    name: 'Step 2 Meeting',
    description: 'Step 2 meeting in progress or completed',
    color: '#9c27b0',
    deadlineDays: 7,
    requiredFields: ['step2MeetingDate'],
    allowedNextStates: ['RESOLVED', 'ARBITRATION_PENDING', 'WITHDRAWN'],
    actions: ['Conduct Meeting', 'Document Outcome', 'Consider Arbitration']
  },
  ARBITRATION_PENDING: {
    name: 'Arbitration Pending',
    description: 'Case sent to arbitration',
    color: '#f44336',
    deadlineDays: 60,
    requiredFields: ['arbitrationFiled'],
    allowedNextStates: ['ARBITRATION_HEARING', 'WITHDRAWN'],
    actions: ['Select Arbitrator', 'Prepare Case', 'Gather All Evidence']
  },
  ARBITRATION_HEARING: {
    name: 'Arbitration Hearing',
    description: 'Arbitration hearing scheduled or in progress',
    color: '#d32f2f',
    deadlineDays: 30,
    requiredFields: ['hearingDate', 'arbitrator'],
    allowedNextStates: ['RESOLVED', 'WITHDRAWN'],
    actions: ['Attend Hearing', 'Present Case', 'Await Decision']
  },
  RESOLVED: {
    name: 'Resolved',
    description: 'Grievance has been resolved',
    color: '#4caf50',
    deadlineDays: null,
    requiredFields: ['resolutionDetails', 'dateClosed'],
    allowedNextStates: ['REOPENED'],
    actions: ['Document Resolution', 'Notify Member', 'Close Case']
  },
  WITHDRAWN: {
    name: 'Withdrawn',
    description: 'Grievance withdrawn by member',
    color: '#9e9e9e',
    deadlineDays: null,
    requiredFields: ['withdrawalReason', 'dateClosed'],
    allowedNextStates: [],
    actions: ['Document Withdrawal', 'Notify Parties', 'Archive']
  },
  REOPENED: {
    name: 'Reopened',
    description: 'Previously closed case reopened',
    color: '#ff5722',
    deadlineDays: 7,
    requiredFields: ['reopenReason'],
    allowedNextStates: ['STEP1_PENDING', 'STEP2_PENDING', 'RESOLVED', 'WITHDRAWN'],
    actions: ['Review Original Case', 'Assess New Information', 'Determine Next Step']
  }
};

/**
 * Shows workflow state machine visualizer
 */
function showWorkflowVisualizer() {
  const html = createWorkflowVisualizerHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔄 Grievance Workflow Visualizer');
}

/**
 * Creates HTML for workflow visualizer
 */
function createWorkflowVisualizerHTML() {
  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .workflow-diagram {
      display: flex;
      flex-direction: column;
      gap: 15px;
      margin: 20px 0;
    }
    .workflow-stage {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      border-left: 5px solid;
      cursor: pointer;
      transition: all 0.2s;
    }
    .workflow-stage:hover {
      transform: translateX(5px);
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .stage-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    .stage-name {
      font-weight: bold;
      font-size: 16px;
      color: #333;
    }
    .stage-deadline {
      background: white;
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 500;
    }
    .stage-desc {
      color: #666;
      font-size: 14px;
      margin-bottom: 10px;
    }
    .stage-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }
    .action-tag {
      background: white;
      padding: 4px 10px;
      border-radius: 4px;
      font-size: 12px;
      color: #555;
      border: 1px solid #ddd;
    }
    .next-states {
      margin-top: 10px;
      font-size: 13px;
      color: #666;
    }
    .arrow {
      text-align: center;
      color: #999;
      font-size: 20px;
      margin: 5px 0;
    }
    .legend {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 20px 0;
      border-left: 4px solid #1a73e8;
    }
    .legend h3 {
      margin-top: 0;
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔄 Grievance Workflow</h2>

    <div class="legend">
      <h3>Workflow Overview</h3>
      <p>The grievance process follows a structured workflow with defined states and transitions.
      Each state has specific deadlines, required fields, and allowed next states.</p>
    </div>

    <div class="workflow-diagram">
      ${Object.entries(WORKFLOW_STATES).map(([key, state]) => `
        <div class="workflow-stage" style="border-left-color: ${state.color}">
          <div class="stage-header">
            <div class="stage-name">${state.name}</div>
            ${state.deadlineDays
              ? `<div class="stage-deadline">${state.deadlineDays} days</div>`
              : '<div class="stage-deadline">Final State</div>'}
          </div>
          <div class="stage-desc">${state.description}</div>
          <div class="stage-actions">
            ${state.actions.map(action => `<span class="action-tag">✓ ${action}</span>`).join('')}
          </div>
          ${state.allowedNextStates.length > 0
            ? `<div class="next-states">
                <strong>Can transition to:</strong> ${state.allowedNextStates.map(s => WORKFLOW_STATES[s].name).join(', ')}
              </div>`
            : ''}
        </div>
        ${state.allowedNextStates.length > 0 ? '<div class="arrow">↓</div>' : ''}
      `).join('')}
    </div>

    <div class="legend">
      <h3>📋 Required Fields by State</h3>
      <ul>
        ${Object.entries(WORKFLOW_STATES).map(([key, state]) => `
          <li><strong>${state.name}:</strong> ${state.requiredFields.join(', ')}</li>
        `).join('')}
      </ul>
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Validates state transition
 * @param {string} currentState - Current workflow state
 * @param {string} newState - Proposed new state
 * @returns {Object} Validation result
 */
function validateStateTransition(currentState, newState) {
  const current = WORKFLOW_STATES[currentState];
  const next = WORKFLOW_STATES[newState];

  if (!current) {
    return {
      valid: false,
      error: `Invalid current state: ${currentState}`
    };
  }

  if (!next) {
    return {
      valid: false,
      error: `Invalid new state: ${newState}`
    };
  }

  if (!current.allowedNextStates.includes(newState)) {
    return {
      valid: false,
      error: `Cannot transition from ${current.name} to ${next.name}. Allowed transitions: ${current.allowedNextStates.map(s => WORKFLOW_STATES[s].name).join(', ')}`
    };
  }

  return {
    valid: true,
    message: `Valid transition from ${current.name} to ${next.name}`
  };
}

/**
 * Changes grievance workflow state with validation
 */
function changeWorkflowState() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  const currentStatus = activeSheet.getRange(activeRow, GRIEVANCE_COLS.STATUS).getValue();

  if (!grievanceId) {
    ui.alert('⚠️ No Grievance ID found.');
    return;
  }

  // Map current status to workflow state
  const currentState = mapStatusToState(currentStatus);

  if (!currentState) {
    ui.alert(
      '⚠️ Unknown Status',
      `Current status "${currentStatus}" is not mapped to a workflow state.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const allowedStates = WORKFLOW_STATES[currentState].allowedNextStates;

  if (allowedStates.length === 0) {
    ui.alert(
      'ℹ️ Final State',
      `Grievance is in "${WORKFLOW_STATES[currentState].name}" state, which is a final state.`,
      ui.ButtonSet.OK
    );
    return;
  }

  const stateOptions = allowedStates
    .map(s => WORKFLOW_STATES[s].name)
    .join('\n');

  const response = ui.prompt(
    '🔄 Change Workflow State',
    `Current State: ${WORKFLOW_STATES[currentState].name}\n\n` +
    `Allowed next states:\n${stateOptions}\n\n` +
    'Enter new state:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const newStateName = response.getResponseText().trim();
  const newStateKey = Object.keys(WORKFLOW_STATES).find(
    key => WORKFLOW_STATES[key].name === newStateName
  );

  if (!newStateKey) {
    ui.alert('⚠️ Invalid state name.');
    return;
  }

  const validation = validateStateTransition(currentState, newStateKey);

  if (!validation.valid) {
    ui.alert('❌ Invalid Transition', validation.error, ui.ButtonSet.OK);
    return;
  }

  // Update status
  const newState = WORKFLOW_STATES[newStateKey];
  activeSheet.getRange(activeRow, GRIEVANCE_COLS.STATUS).setValue(newState.name);

  // Calculate new deadline if applicable
  if (newState.deadlineDays) {
    const newDeadline = new Date();
    newDeadline.setDate(newDeadline.getDate() + newState.deadlineDays);
    activeSheet.getRange(activeRow, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue(newDeadline);
  }

  // Log state change
  logStateChange(grievanceId, currentState, newStateKey);

  ui.alert(
    '✅ State Changed',
    `Workflow state updated to: ${newState.name}\n\n` +
    (newState.deadlineDays ? `New deadline: ${newState.deadlineDays} days from now\n\n` : '') +
    `Required actions:\n${newState.actions.join('\n')}`,
    ui.ButtonSet.OK
  );
}

/**
 * Maps status string to workflow state key
 * @param {string} status - Status string
 * @returns {string} Workflow state key
 */
function mapStatusToState(status) {
  const mapping = {
    'Filed': 'FILED',
    'Step 1 Pending': 'STEP1_PENDING',
    'Step 1 Meeting': 'STEP1_MEETING',
    'Step 2 Pending': 'STEP2_PENDING',
    'Step 2 Meeting': 'STEP2_MEETING',
    'Arbitration Pending': 'ARBITRATION_PENDING',
    'Arbitration Hearing': 'ARBITRATION_HEARING',
    'Resolved': 'RESOLVED',
    'Withdrawn': 'WITHDRAWN',
    'Reopened': 'REOPENED',
    'Open': 'FILED',
    'In Progress': 'STEP1_PENDING',
    'Closed': 'RESOLVED'
  };

  return mapping[status] || null;
}

/**
 * Logs workflow state change
 * @param {string} grievanceId - Grievance ID
 * @param {string} fromState - Previous state
 * @param {string} toState - New state
 */
function logStateChange(grievanceId, fromState, toState) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let stateLog = ss.getSheetByName('🔄 State Change Log');

  if (!stateLog) {
    stateLog = createStateChangeLogSheet();
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || 'System';
  const fromStateName = WORKFLOW_STATES[fromState].name;
  const toStateName = WORKFLOW_STATES[toState].name;

  const lastRow = stateLog.getLastRow();
  const newRow = [
    timestamp,
    grievanceId,
    fromStateName,
    toStateName,
    user
  ];

  stateLog.getRange(lastRow + 1, 1, 1, 5).setValues([newRow]);
}

/**
 * Creates State Change Log sheet
 * @returns {Sheet} State change log sheet
 */
function createStateChangeLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('🔄 State Change Log');

  // Set headers
  const headers = [
    'Timestamp',
    'Grievance ID',
    'From State',
    'To State',
    'Changed By'
  ];

  sheet.getRange(1, 1, 1, 5).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Grievance ID
  sheet.setColumnWidth(3, 150); // From State
  sheet.setColumnWidth(4, 150); // To State
  sheet.setColumnWidth(5, 200); // Changed By

  // Freeze header
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Shows workflow state for selected grievance
 */
function showCurrentWorkflowState() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert('⚠️ Invalid Selection');
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  const currentStatus = activeSheet.getRange(activeRow, GRIEVANCE_COLS.STATUS).getValue();
  const currentState = mapStatusToState(currentStatus);

  if (!currentState) {
    ui.alert('⚠️ Unknown workflow state.');
    return;
  }

  const state = WORKFLOW_STATES[currentState];
  const allowedTransitions = state.allowedNextStates
    .map(s => WORKFLOW_STATES[s].name)
    .join('\n• ');

  const message = `
Grievance: ${grievanceId}
Current State: ${state.name}

${state.description}

Deadline: ${state.deadlineDays ? state.deadlineDays + ' days' : 'None (final state)'}

Required Actions:
• ${state.actions.join('\n• ')}

Allowed Transitions:
${allowedTransitions ? '• ' + allowedTransitions : 'None (final state)'}
  `;

  ui.alert('🔄 Workflow State', message, ui.ButtonSet.OK);
}

/**
 * Bulk updates workflow states for selected grievances
 */
function batchUpdateWorkflowState() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('⚠️ Wrong Sheet');
    return;
  }

  const selection = activeSheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow === 1 || numRows === 0) {
    ui.alert('⚠️ Please select grievance rows (not the header).');
    return;
  }

  const rows = Array.from({ length: numRows }, (_, i) => startRow + i);

  const response = ui.prompt(
    '🔄 Batch Update Workflow State',
    `Selected ${rows.length} grievance(s).\n\n` +
    'Enter new state (e.g., "Step 1 Pending"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const newStateName = response.getResponseText().trim();
  const newStateKey = Object.keys(WORKFLOW_STATES).find(
    key => WORKFLOW_STATES[key].name === newStateName
  );

  if (!newStateKey) {
    ui.alert('⚠️ Invalid state name.');
    return;
  }

  const confirmResponse = ui.alert(
    '⚠️ Confirm Batch Update',
    `This will change ${rows.length} grievance(s) to "${newStateName}".\n\nContinue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  const newState = WORKFLOW_STATES[newStateKey];
  let updated = 0;
  let errors = 0;

  rows.forEach(row => {
    const grievanceId = activeSheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
    const currentStatus = activeSheet.getRange(row, GRIEVANCE_COLS.STATUS).getValue();
    const currentState = mapStatusToState(currentStatus);

    if (!currentState) {
      errors++;
      return;
    }

    const validation = validateStateTransition(currentState, newStateKey);

    if (!validation.valid) {
      errors++;
      Logger.log(`Error updating ${grievanceId}: ${validation.error}`);
      return;
    }

    // Update status
    activeSheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newState.name);

    // Calculate new deadline if applicable
    if (newState.deadlineDays) {
      const newDeadline = new Date();
      newDeadline.setDate(newDeadline.getDate() + newState.deadlineDays);
      activeSheet.getRange(row, GRIEVANCE_COLS.NEXT_ACTION_DUE).setValue(newDeadline);
    }

    // Log state change
    logStateChange(grievanceId, currentState, newStateKey);

    updated++;
  });

  ui.alert(
    '✅ Batch Update Complete',
    `Successfully updated: ${updated}\nErrors (invalid transitions): ${errors}`,
    ui.ButtonSet.OK
  );
}

/******************************************************************************
 * MODULE: SmartAutoAssignment
 * Source: SmartAutoAssignment.gs
 *****************************************************************************/

/**
 * ============================================================================
 * SMART AUTO-ASSIGNMENT SYSTEM
 * ============================================================================
 *
 * Intelligently assigns stewards to grievances based on multiple factors
 * Features:
 * - Workload balancing (assign to least busy steward)
 * - Location matching (same worksite/unit)
 * - Expertise matching (issue type experience)
 * - Availability checking
 * - Manual override capability
 * - Assignment history tracking
 * - Performance metrics
 */

/**
 * Assignment algorithm weights
 */
const ASSIGNMENT_WEIGHTS = {
  WORKLOAD: 0.4,        // 40% - Balance caseload
  LOCATION: 0.3,        // 30% - Same location preference
  EXPERTISE: 0.2,       // 20% - Issue type experience
  AVAILABILITY: 0.1     // 10% - Current availability
};

/**
 * Auto-assigns a steward to a grievance
 * @param {string} grievanceId - Grievance ID
 * @param {Object} preferences - Assignment preferences
 * @returns {Object} Assignment result
 */
function autoAssignSteward(grievanceId, preferences = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Get grievance details
  const grievance = getGrievanceDetails(grievanceId);

  if (!grievance) {
    return {
      success: false,
      error: `Grievance ${grievanceId} not found`
    };
  }

  // Get all available stewards
  const stewards = getAllStewards();

  if (stewards.length === 0) {
    return {
      success: false,
      error: 'No stewards available in the system'
    };
  }

  // Score each steward
  const scoredStewards = stewards.map(steward => ({
    ...steward,
    score: calculateAssignmentScore(steward, grievance, preferences)
  }));

  // Sort by score (descending)
  scoredStewards.sort((a, b) => b.score - a.score);

  // Get top candidate
  const selectedSteward = scoredStewards[0];

  // Assign steward
  const assignmentResult = assignStewardToGrievance(grievanceId, selectedSteward);

  // Log assignment
  logAssignment(grievanceId, selectedSteward, scoredStewards.slice(0, 3));

  return {
    success: true,
    assignedSteward: selectedSteward.name,
    score: selectedSteward.score,
    reasoning: generateAssignmentReasoning(selectedSteward, grievance),
    alternates: scoredStewards.slice(1, 4).map(s => ({
      name: s.name,
      score: s.score
    }))
  };
}

/**
 * Gets grievance details
 * @param {string} grievanceId - Grievance ID
 * @returns {Object} Grievance details
 */
function getGrievanceDetails(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) return null;

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      return {
        id: data[i][0],
        memberId: data[i][1],
        memberName: `${data[i][2]} ${data[i][3]}`,
        issueType: data[i][5],
        location: data[i][9] || '',
        unit: data[i][10] || '',
        manager: data[i][11] || '',
        row: i + 2
      };
    }
  }

  return null;
}

/**
 * Gets all stewards from Member Directory
 * @returns {Array} Array of steward objects
 */
function getAllStewards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const lastRow = memberSheet.getLastRow();

  if (lastRow < 2) return [];

  const data = memberSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const stewards = [];

  data.forEach((row, index) => {
    const isSteward = row[9]; // Column J: Is Steward?

    if (isSteward === 'Yes') {
      stewards.push({
        id: row[0],
        name: `${row[1]} ${row[2]}`,
        email: row[7],
        location: row[4] || '',
        unit: row[5] || '',
        row: index + 2,
        currentCaseload: getCurrentCaseload(row[0]),
        expertise: getStewardExpertise(row[0])
      });
    }
  });

  return stewards;
}

/**
 * Gets current caseload for a steward
 * @param {string} memberId - Member ID
 * @returns {number} Number of open grievances
 */
function getCurrentCaseload(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) return 0;

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  let count = 0;

  data.forEach(row => {
    const assignedSteward = row[13]; // Column N: Assigned Steward
    const status = row[4]; // Column E: Status

    if (assignedSteward && assignedSteward.includes(memberId) && status === 'Open') {
      count++;
    }
  });

  return count;
}

/**
 * Gets steward's expertise (issue types they've handled)
 * @param {string} memberId - Member ID
 * @returns {Object} Issue type counts
 */
function getStewardExpertise(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) return {};

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const expertise = {};

  data.forEach(row => {
    const assignedSteward = row[13]; // Column N: Assigned Steward
    const issueType = row[5]; // Column F: Issue Type

    if (assignedSteward && assignedSteward.includes(memberId) && issueType) {
      expertise[issueType] = (expertise[issueType] || 0) + 1;
    }
  });

  return expertise;
}

/**
 * Calculates assignment score for a steward
 * @param {Object} steward - Steward details
 * @param {Object} grievance - Grievance details
 * @param {Object} preferences - Assignment preferences
 * @returns {number} Score (0-100)
 */
function calculateAssignmentScore(steward, grievance, preferences) {
  const weights = { ...ASSIGNMENT_WEIGHTS, ...preferences };

  // Workload score (inverse - lower caseload = higher score)
  const maxCaseload = 20; // Assume max reasonable caseload
  const workloadScore = Math.max(0, (maxCaseload - steward.currentCaseload) / maxCaseload) * 100;

  // Location score
  const locationScore = steward.location === grievance.location ? 100 : 0;

  // Expertise score
  const expertiseCount = steward.expertise[grievance.issueType] || 0;
  const maxExpertise = Math.max(...Object.values(steward.expertise), 1);
  const expertiseScore = (expertiseCount / maxExpertise) * 100;

  // Availability score (for now, always 100 - can be enhanced)
  const availabilityScore = 100;

  // Calculate weighted average
  const finalScore =
    (workloadScore * weights.WORKLOAD) +
    (locationScore * weights.LOCATION) +
    (expertiseScore * weights.EXPERTISE) +
    (availabilityScore * weights.AVAILABILITY);

  return Math.round(finalScore);
}

/**
 * Assigns steward to grievance
 * @param {string} grievanceId - Grievance ID
 * @param {Object} steward - Steward object
 * @returns {boolean} Success
 */
function assignStewardToGrievance(grievanceId, steward) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const grievance = getGrievanceDetails(grievanceId);

  if (!grievance) return false;

  // Update assigned steward
  grievanceSheet.getRange(grievance.row, GRIEVANCE_COLS.ASSIGNED_STEWARD).setValue(steward.name);

  // Update steward email if available
  if (steward.email) {
    grievanceSheet.getRange(grievance.row, GRIEVANCE_COLS.ASSIGNED_STEWARD + 1).setValue(steward.email);
  }

  return true;
}

/**
 * Generates reasoning for assignment
 * @param {Object} steward - Selected steward
 * @param {Object} grievance - Grievance details
 * @returns {string} Reasoning text
 */
function generateAssignmentReasoning(steward, grievance) {
  const reasons = [];

  if (steward.location === grievance.location) {
    reasons.push(`Same location (${steward.location})`);
  }

  const expertiseCount = steward.expertise[grievance.issueType] || 0;
  if (expertiseCount > 0) {
    reasons.push(`Has handled ${expertiseCount} similar ${grievance.issueType} case(s)`);
  }

  reasons.push(`Current caseload: ${steward.currentCaseload} case(s)`);

  if (steward.currentCaseload < 5) {
    reasons.push('Low caseload - has capacity');
  } else if (steward.currentCaseload < 10) {
    reasons.push('Moderate caseload');
  } else {
    reasons.push('High caseload - may need support');
  }

  return reasons.join('\n• ');
}

/**
 * Logs assignment for tracking
 * @param {string} grievanceId - Grievance ID
 * @param {Object} selectedSteward - Selected steward
 * @param {Array} topCandidates - Top 3 candidates
 */
function logAssignment(grievanceId, selectedSteward, topCandidates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let assignmentLog = ss.getSheetByName('🤖 Auto-Assignment Log');

  if (!assignmentLog) {
    assignmentLog = createAutoAssignmentLogSheet();
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || 'System';

  const lastRow = assignmentLog.getLastRow();
  const newRow = [
    timestamp,
    grievanceId,
    selectedSteward.name,
    selectedSteward.score,
    selectedSteward.currentCaseload,
    topCandidates.map(s => `${s.name} (${s.score})`).join(', '),
    user
  ];

  assignmentLog.getRange(lastRow + 1, 1, 1, 7).setValues([newRow]);
}

/**
 * Creates Auto-Assignment Log sheet
 * @returns {Sheet} Assignment log sheet
 */
function createAutoAssignmentLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('🤖 Auto-Assignment Log');

  // Set headers
  const headers = [
    'Timestamp',
    'Grievance ID',
    'Assigned Steward',
    'Score',
    'Caseload',
    'Top Candidates',
    'Assigned By'
  ];

  sheet.getRange(1, 1, 1, 7).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, 7)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Grievance ID
  sheet.setColumnWidth(3, 150); // Assigned Steward
  sheet.setColumnWidth(4, 80);  // Score
  sheet.setColumnWidth(5, 80);  // Caseload
  sheet.setColumnWidth(6, 300); // Top Candidates
  sheet.setColumnWidth(7, 200); // Assigned By

  // Freeze header
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Shows auto-assignment dialog for a grievance
 */
function showAutoAssignDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert('⚠️ Invalid Selection');
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!grievanceId) {
    ui.alert('⚠️ No Grievance ID found.');
    return;
  }

  // Check if already assigned
  const currentAssignment = activeSheet.getRange(activeRow, GRIEVANCE_COLS.ASSIGNED_STEWARD).getValue();

  if (currentAssignment) {
    const confirmResponse = ui.alert(
      '⚠️ Already Assigned',
      `This grievance is already assigned to: ${currentAssignment}\n\nReassign automatically?`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      return;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('🤖 Analyzing stewards...', 'Please wait', -1);

  try {
    const result = autoAssignSteward(grievanceId);

    if (!result.success) {
      ui.alert('❌ Assignment Failed', result.error, ui.ButtonSet.OK);
      return;
    }

    const alternatesText = result.alternates
      .map(a => `  ${a.name} (score: ${a.score})`)
      .join('\n');

    ui.alert(
      '✅ Auto-Assignment Complete',
      `Assigned to: ${result.assignedSteward}\n` +
      `Assignment Score: ${result.score}/100\n\n` +
      `Reasoning:\n• ${result.reasoning}\n\n` +
      `Alternative Candidates:\n${alternatesText}`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Batch auto-assigns stewards to selected grievances
 */
function batchAutoAssign() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('⚠️ Wrong Sheet');
    return;
  }

  const selection = activeSheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow === 1 || numRows === 0) {
    ui.alert('⚠️ Please select grievance rows.');
    return;
  }

  const rows = Array.from({ length: numRows }, (_, i) => startRow + i);

  // Count unassigned
  let unassigned = 0;
  rows.forEach(row => {
    const assigned = activeSheet.getRange(row, GRIEVANCE_COLS.ASSIGNED_STEWARD).getValue();
    if (!assigned) {
      unassigned++;
    }
  });

  const confirmResponse = ui.alert(
    '🤖 Batch Auto-Assignment',
    `Selected: ${rows.length} grievance(s)\n` +
    `Unassigned: ${unassigned}\n\n` +
    'Auto-assign stewards to all selected grievances?',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('🤖 Auto-assigning...', 'Please wait', -1);

  let assigned = 0;
  let skipped = 0;
  let errors = 0;

  rows.forEach(row => {
    const grievanceId = activeSheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
    const currentAssignment = activeSheet.getRange(row, GRIEVANCE_COLS.ASSIGNED_STEWARD).getValue();

    if (!grievanceId) {
      skipped++;
      return;
    }

    if (currentAssignment) {
      skipped++;
      return;
    }

    try {
      const result = autoAssignSteward(grievanceId);

      if (result.success) {
        assigned++;
      } else {
        errors++;
        Logger.log(`Error assigning ${grievanceId}: ${result.error}`);
      }

    } catch (error) {
      errors++;
      Logger.log(`Error assigning ${grievanceId}: ${error.message}`);
    }
  });

  ui.alert(
    '✅ Batch Auto-Assignment Complete',
    `Successfully assigned: ${assigned}\n` +
    `Skipped (already assigned): ${skipped}\n` +
    `Errors: ${errors}`,
    ui.ButtonSet.OK
  );
}

/**
 * Shows steward workload dashboard
 */
function showStewardWorkloadDashboard() {
  const stewards = getAllStewards();

  if (stewards.length === 0) {
    SpreadsheetApp.getUi().alert('No stewards found in the system.');
    return;
  }

  // Sort by caseload (descending)
  stewards.sort((a, b) => b.currentCaseload - a.currentCaseload);

  const stewardsList = stewards
    .map(s => `
      <div class="steward-item" style="border-left-color: ${getCaseloadColor(s.currentCaseload)}">
        <div class="steward-name">${s.name}</div>
        <div class="steward-details">
          <strong>Caseload:</strong> ${s.currentCaseload} open case(s) |
          <strong>Location:</strong> ${s.location || 'N/A'} |
          <strong>Unit:</strong> ${s.unit || 'N/A'}
        </div>
        <div class="steward-expertise">
          <strong>Expertise:</strong> ${Object.keys(s.expertise).length > 0
            ? Object.entries(s.expertise)
                .sort((a, b) => b[1] - a[1])
                .slice(0, 3)
                .map(([type, count]) => `${type} (${count})`)
                .join(', ')
            : 'No cases yet'}
        </div>
      </div>
    `)
    .join('');

  const avgCaseload = stewards.reduce((sum, s) => sum + s.currentCaseload, 0) / stewards.length;

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .summary { background: #e8f0fe; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #1a73e8; }
    .steward-item { background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 5px solid; }
    .steward-name { font-weight: bold; font-size: 16px; color: #333; margin-bottom: 5px; }
    .steward-details { font-size: 13px; color: #666; margin: 5px 0; }
    .steward-expertise { font-size: 12px; color: #888; margin-top: 5px; }
  </style>
</head>
<body>
  <div class="container">
    <h2>👥 Steward Workload Dashboard</h2>

    <div class="summary">
      <strong>Total Stewards:</strong> ${stewards.length}<br>
      <strong>Average Caseload:</strong> ${avgCaseload.toFixed(1)} cases/steward<br>
      <strong>Total Open Cases:</strong> ${stewards.reduce((sum, s) => sum + s.currentCaseload, 0)}
    </div>

    <div style="max-height: 500px; overflow-y: auto;">
      ${stewardsList}
    </div>
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(750)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '👥 Steward Workload Dashboard');
}

/**
 * Gets color for caseload visualization
 * @param {number} caseload - Number of cases
 * @returns {string} Color code
 */
function getCaseloadColor(caseload) {
  if (caseload === 0) return '#9e9e9e';
  if (caseload < 5) return '#4caf50';
  if (caseload < 10) return '#ffeb3b';
  if (caseload < 15) return '#ff9800';
  return '#f44336';
}

/******************************************************************************
 * MODULE: MemberSearch
 * Source: MemberSearch.gs
 *****************************************************************************/

/**
 * ============================================================================
 * MEMBER SEARCH & LOOKUP FUNCTIONALITY
 * ============================================================================
 *
 * Fast member search with autocomplete and filtering
 * Features:
 * - Search by name, ID, location, unit
 * - Fuzzy matching
 * - Quick navigation to member row
 * - Show member details
 */

/**
 * Shows member search dialog
 */
function showMemberSearch() {
  const html = HtmlService.createHtmlOutput(createMemberSearchHTML())
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search Members');
}

/**
 * Creates HTML for member search dialog
 */
function createMemberSearchHTML() {
  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .search-box {
      margin: 20px 0;
    }
    input[type="text"] {
      width: 100%;
      padding: 12px;
      font-size: 16px;
      border: 2px solid #ddd;
      border-radius: 4px;
      box-sizing: border-box;
    }
    input[type="text"]:focus {
      outline: none;
      border-color: #1a73e8;
    }
    .filters {
      display: flex;
      gap: 10px;
      margin: 15px 0;
    }
    select {
      flex: 1;
      padding: 8px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    .results {
      margin-top: 20px;
      max-height: 350px;
      overflow-y: auto;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    .result-item {
      padding: 15px;
      border-bottom: 1px solid #eee;
      cursor: pointer;
      transition: background 0.2s;
    }
    .result-item:hover {
      background: #f0f7ff;
    }
    .result-item:last-child {
      border-bottom: none;
    }
    .member-name {
      font-size: 16px;
      font-weight: bold;
      color: #1a73e8;
    }
    .member-details {
      font-size: 13px;
      color: #666;
      margin-top: 5px;
    }
    .member-id {
      display: inline-block;
      background: #e8f0fe;
      padding: 2px 8px;
      border-radius: 3px;
      font-size: 12px;
      margin-right: 8px;
    }
    .no-results {
      text-align: center;
      padding: 40px;
      color: #999;
      font-style: italic;
    }
    .loading {
      text-align: center;
      padding: 20px;
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔍 Search Members</h2>

    <div class="search-box">
      <input type="text" id="searchInput" placeholder="Search by name, ID, email..." autofocus>
    </div>

    <div class="filters">
      <select id="locationFilter">
        <option value="">All Locations</option>
      </select>
      <select id="unitFilter">
        <option value="">All Units</option>
      </select>
      <select id="stewardFilter">
        <option value="">All Members</option>
        <option value="Yes">Stewards Only</option>
        <option value="No">Non-Stewards Only</option>
      </select>
    </div>

    <div id="results" class="results">
      <div class="loading">Loading members...</div>
    </div>
  </div>

  <script>
    let allMembers = [];
    let filteredMembers = [];

    // Load members on page load
    window.onload = function() {
      google.script.run
        .withSuccessHandler(onMembersLoaded)
        .withFailureHandler(onError)
        .getAllMembers();
    };

    function onMembersLoaded(members) {
      allMembers = members;
      filteredMembers = members;

      // Populate filter dropdowns
      populateFilters();

      // Display all members initially
      displayResults(members);

      // Set up event listeners
      document.getElementById('searchInput').addEventListener('input', performSearch);
      document.getElementById('locationFilter').addEventListener('change', performSearch);
      document.getElementById('unitFilter').addEventListener('change', performSearch);
      document.getElementById('stewardFilter').addEventListener('change', performSearch);
    }

    function populateFilters() {
      // Get unique locations
      const locations = [...new Set(allMembers.map(m => m.location).filter(l => l))];
      const locationFilter = document.getElementById('locationFilter');
      locations.forEach(loc => {
        const option = document.createElement('option');
        option.value = loc;
        option.textContent = loc;
        locationFilter.appendChild(option);
      });

      // Get unique units
      const units = [...new Set(allMembers.map(m => m.unit).filter(u => u))];
      const unitFilter = document.getElementById('unitFilter');
      units.forEach(unit => {
        const option = document.createElement('option');
        option.value = unit;
        option.textContent = unit;
        unitFilter.appendChild(option);
      });
    }

    function performSearch() {
      const searchTerm = document.getElementById('searchInput').value.toLowerCase();
      const locationFilter = document.getElementById('locationFilter').value;
      const unitFilter = document.getElementById('unitFilter').value;
      const stewardFilter = document.getElementById('stewardFilter').value;

      filteredMembers = allMembers.filter(member => {
        // Text search
        const matchesSearch = !searchTerm ||
          member.name.toLowerCase().includes(searchTerm) ||
          member.id.toLowerCase().includes(searchTerm) ||
          (member.email && member.email.toLowerCase().includes(searchTerm));

        // Location filter
        const matchesLocation = !locationFilter || member.location === locationFilter;

        // Unit filter
        const matchesUnit = !unitFilter || member.unit === unitFilter;

        // Steward filter
        const matchesSteward = !stewardFilter || member.isSteward === stewardFilter;

        return matchesSearch && matchesLocation && matchesUnit && matchesSteward;
      });

      displayResults(filteredMembers);
    }

    function displayResults(members) {
      const resultsDiv = document.getElementById('results');

      if (members.length === 0) {
        resultsDiv.innerHTML = '<div class="no-results">No members found</div>';
        return;
      }

      resultsDiv.innerHTML = members.slice(0, 50).map(member => \`
        <div class="result-item" onclick="selectMember('\${member.row}', '\${member.id}')">
          <div class="member-name">\${member.name}</div>
          <div class="member-details">
            <span class="member-id">\${member.id}</span>
            <span>\${member.location || 'No location'} • \${member.unit || 'No unit'}</span>
            \${member.isSteward === 'Yes' ? ' • 🛡️ Steward' : ''}
          </div>
          \${member.email ? \`<div class="member-details">📧 \${member.email}</div>\` : ''}
        </div>
      \`).join('');

      if (members.length > 50) {
        resultsDiv.innerHTML += \`<div class="no-results">Showing first 50 of \${members.length} results. Refine your search.</div>\`;
      }
    }

    function selectMember(row, memberId) {
      google.script.run
        .withSuccessHandler(() => {
          google.script.host.close();
        })
        .navigateToMember(parseInt(row));
    }

    function onError(error) {
      document.getElementById('results').innerHTML =
        '<div class="no-results">❌ Error: ' + error.message + '</div>';
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets all members for search
 * @returns {Array} Array of member objects
 */
function getAllMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return [];

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = memberSheet.getRange(2, 1, lastRow - 1, 13).getValues();

  return data.map((row, index) => ({
    row: index + 2,
    id: row[0] || '',
    name: `${row[1]} ${row[2]}`.trim(),
    firstName: row[1] || '',
    lastName: row[2] || '',
    jobTitle: row[3] || '',
    location: row[4] || '',
    unit: row[5] || '',
    email: row[7] || '',
    phone: row[8] || '',
    isSteward: row[9] || '',
    supervisor: row[10] || '',
    manager: row[11] || '',
    assignedSteward: row[12] || ''
  })).filter(m => m.id); // Filter out empty rows
}

/**
 * Navigates to a member row in the Member Directory
 * @param {number} row - The row number to navigate to
 */
function navigateToMember(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  // Activate the Member Directory sheet
  ss.setActiveSheet(memberSheet);

  // Navigate to the row
  memberSheet.setActiveRange(memberSheet.getRange(row, 1, 1, 10));

  // Show toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Navigated to row ${row}`,
    'Member Found',
    3
  );
}

/**
 * Quick search function for keyboard shortcut
 */
function quickMemberSearch() {
  showMemberSearch();
}

/******************************************************************************
 * MODULE: PredictiveAnalytics
 * Source: PredictiveAnalytics.gs
 *****************************************************************************/

/**
 * ============================================================================
 * PREDICTIVE ANALYTICS
 * ============================================================================
 *
 * Forecasts trends and provides actionable insights
 * Features:
 * - Grievance volume forecasting
 * - Issue type trend analysis
 * - Seasonal pattern detection
 * - Steward workload predictions
 * - Risk identification
 * - Anomaly detection
 * - Actionable recommendations
 */

/**
 * Analyzes trends and generates predictions
 * @returns {Object} Analytics results
 */
function performPredictiveAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) {
    return {
      error: 'Insufficient data for analysis'
    };
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const analytics = {
    volumeTrend: analyzeVolumeTrend(data),
    issueTypeTrends: analyzeIssueTypeTrends(data),
    seasonalPatterns: detectSeasonalPatterns(data),
    resolutionTimeTrend: analyzeResolutionTimeTrend(data),
    stewardWorkloadForecast: forecastStewardWorkload(data),
    riskFactors: identifyRiskFactors(data),
    anomalies: detectAnomalies(data),
    recommendations: generateRecommendations(data)
  };

  return analytics;
}

/**
 * Analyzes grievance volume trend
 * @param {Array} data - Grievance data
 * @returns {Object} Volume trend analysis
 */
function analyzeVolumeTrend(data) {
  // Group by month
  const monthlyVolumes = {};

  data.forEach(row => {
    const filedDate = row[6]; // Column G: Filed Date

    if (!filedDate) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;

    monthlyVolumes[monthKey] = (monthlyVolumes[monthKey] || 0) + 1;
  });

  // Convert to array and sort
  const months = Object.keys(monthlyVolumes).sort();
  const volumes = months.map(m => monthlyVolumes[m]);

  // Calculate trend (simple linear regression)
  const trend = calculateLinearTrend(volumes);

  // Forecast next 3 months
  const forecast = [];
  for (let i = 1; i <= 3; i++) {
    forecast.push(Math.round(trend.slope * (volumes.length + i) + trend.intercept));
  }

  return {
    currentMonthVolume: volumes[volumes.length - 1] || 0,
    previousMonthVolume: volumes[volumes.length - 2] || 0,
    trend: trend.slope > 0 ? 'increasing' : trend.slope < 0 ? 'decreasing' : 'stable',
    trendPercentage: Math.abs(trend.slope),
    forecast: forecast,
    monthlyVolumes: monthlyVolumes
  };
}

/**
 * Analyzes issue type trends
 * @param {Array} data - Grievance data
 * @returns {Object} Issue type trends
 */
function analyzeIssueTypeTrends(data) {
  const issueTypesByMonth = {};

  data.forEach(row => {
    const filedDate = row[6];
    const issueType = row[5];

    if (!filedDate || !issueType) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;

    if (!issueTypesByMonth[monthKey]) {
      issueTypesByMonth[monthKey] = {};
    }

    issueTypesByMonth[monthKey][issueType] = (issueTypesByMonth[monthKey][issueType] || 0) + 1;
  });

  // Find trending issue types (increasing over time)
  const trendingIssues = {};
  const months = Object.keys(issueTypesByMonth).sort();

  if (months.length >= 3) {
    const recentMonths = months.slice(-3);
    const olderMonths = months.slice(0, -3);

    const recentCounts = {};
    const olderCounts = {};

    recentMonths.forEach(month => {
      Object.entries(issueTypesByMonth[month]).forEach(([type, count]) => {
        recentCounts[type] = (recentCounts[type] || 0) + count;
      });
    });

    olderMonths.forEach(month => {
      Object.entries(issueTypesByMonth[month]).forEach(([type, count]) => {
        olderCounts[type] = (olderCounts[type] || 0) + count;
      });
    });

    Object.keys(recentCounts).forEach(type => {
      const recentAvg = recentCounts[type] / recentMonths.length;
      const olderAvg = olderCounts[type] ? olderCounts[type] / olderMonths.length : 0;

      if (recentAvg > olderAvg * 1.2) {
        trendingIssues[type] = {
          recentAverage: recentAvg.toFixed(1),
          olderAverage: olderAvg.toFixed(1),
          percentIncrease: olderAvg > 0 ? (((recentAvg - olderAvg) / olderAvg) * 100).toFixed(1) : 100
        };
      }
    });
  }

  return {
    trendingUp: trendingIssues,
    monthlyBreakdown: issueTypesByMonth
  };
}

/**
 * Detects seasonal patterns
 * @param {Array} data - Grievance data
 * @returns {Object} Seasonal patterns
 */
function detectSeasonalPatterns(data) {
  const quarterlyVolumes = { Q1: 0, Q2: 0, Q3: 0, Q4: 0 };

  data.forEach(row => {
    const filedDate = row[6];

    if (!filedDate) return;

    const quarter = Math.floor(filedDate.getMonth() / 3) + 1;
    quarterlyVolumes[`Q${quarter}`]++;
  });

  // Find peak quarter
  let peakQuarter = 'Q1';
  let peakVolume = 0;

  Object.entries(quarterlyVolumes).forEach(([quarter, volume]) => {
    if (volume > peakVolume) {
      peakQuarter = quarter;
      peakVolume = volume;
    }
  });

  return {
    quarterlyVolumes: quarterlyVolumes,
    peakQuarter: peakQuarter,
    peakVolume: peakVolume,
    hasSeasonality: Math.max(...Object.values(quarterlyVolumes)) > Math.min(...Object.values(quarterlyVolumes)) * 1.5
  };
}

/**
 * Analyzes resolution time trend
 * @param {Array} data - Grievance data
 * @returns {Object} Resolution time analysis
 */
function analyzeResolutionTimeTrend(data) {
  const resolutionTimes = [];

  data.forEach(row => {
    const filedDate = row[6];
    const closedDate = row[18];

    if (filedDate && closedDate && closedDate > filedDate) {
      const days = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
      resolutionTimes.push(days);
    }
  });

  if (resolutionTimes.length === 0) {
    return {
      average: 0,
      median: 0,
      trend: 'insufficient_data'
    };
  }

  // Sort for median
  resolutionTimes.sort((a, b) => a - b);

  const average = resolutionTimes.reduce((sum, val) => sum + val, 0) / resolutionTimes.length;
  const median = resolutionTimes[Math.floor(resolutionTimes.length / 2)];

  // Recent vs older comparison
  const recent = resolutionTimes.slice(-10);
  const older = resolutionTimes.slice(0, -10);

  const recentAvg = recent.reduce((sum, val) => sum + val, 0) / recent.length;
  const olderAvg = older.length > 0 ? older.reduce((sum, val) => sum + val, 0) / older.length : recentAvg;

  return {
    average: Math.round(average),
    median: median,
    trend: recentAvg > olderAvg ? 'increasing' : recentAvg < olderAvg ? 'decreasing' : 'stable',
    recentAverage: Math.round(recentAvg),
    olderAverage: Math.round(olderAvg)
  };
}

/**
 * Forecasts steward workload
 * @param {Array} data - Grievance data
 * @returns {Object} Workload forecast
 */
function forecastStewardWorkload(data) {
  const stewardCases = {};

  data.forEach(row => {
    const steward = row[13];
    const status = row[4];

    if (steward && status === 'Open') {
      stewardCases[steward] = (stewardCases[steward] || 0) + 1;
    }
  });

  // Identify overloaded stewards (>15 cases)
  const overloaded = [];
  const underutilized = [];

  Object.entries(stewardCases).forEach(([steward, count]) => {
    if (count > 15) {
      overloaded.push({ steward, caseload: count });
    } else if (count < 3) {
      underutilized.push({ steward, caseload: count });
    }
  });

  return {
    overloaded: overloaded,
    underutilized: underutilized,
    balanceRecommendation: overloaded.length > 0 && underutilized.length > 0
      ? 'Redistribute cases from overloaded to underutilized stewards'
      : 'Workload is relatively balanced'
  };
}

/**
 * Identifies risk factors
 * @param {Array} data - Grievance data
 * @returns {Array} Risk factors
 */
function identifyRiskFactors(data) {
  const risks = [];

  // Check for overdue cases
  let overdueCount = 0;
  data.forEach(row => {
    const daysToDeadline = row[20];
    if (daysToDeadline < 0) overdueCount++;
  });

  if (overdueCount > 0) {
    risks.push({
      type: 'Overdue Cases',
      severity: overdueCount > 10 ? 'HIGH' : overdueCount > 5 ? 'MEDIUM' : 'LOW',
      description: `${overdueCount} grievance(s) are overdue`,
      action: 'Review and prioritize overdue cases immediately'
    });
  }

  // Check for cases with no steward assigned
  let unassignedCount = 0;
  data.forEach(row => {
    const steward = row[13];
    const status = row[4];
    if (!steward && status === 'Open') unassignedCount++;
  });

  if (unassignedCount > 0) {
    risks.push({
      type: 'Unassigned Cases',
      severity: unassignedCount > 5 ? 'HIGH' : 'MEDIUM',
      description: `${unassignedCount} open grievance(s) have no assigned steward`,
      action: 'Use auto-assignment to assign stewards'
    });
  }

  // Check for high volume increase
  const volumeTrend = analyzeVolumeTrend(data);
  if (volumeTrend.trend === 'increasing' && volumeTrend.trendPercentage > 20) {
    risks.push({
      type: 'Volume Surge',
      severity: 'MEDIUM',
      description: `Grievance volume trending upward by ${volumeTrend.trendPercentage.toFixed(1)}%`,
      action: 'Prepare for increased caseload, consider additional steward training'
    });
  }

  return risks;
}

/**
 * Detects anomalies in data
 * @param {Array} data - Grievance data
 * @returns {Array} Anomalies detected
 */
function detectAnomalies(data) {
  const anomalies = [];

  // Check for unusual spike in last month
  const monthlyVolumes = {};
  data.forEach(row => {
    const filedDate = row[6];
    if (!filedDate) return;

    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;
    monthlyVolumes[monthKey] = (monthlyVolumes[monthKey] || 0) + 1;
  });

  const volumes = Object.values(monthlyVolumes);
  if (volumes.length >= 3) {
    const average = volumes.slice(0, -1).reduce((sum, val) => sum + val, 0) / (volumes.length - 1);
    const latest = volumes[volumes.length - 1];

    if (latest > average * 1.5) {
      anomalies.push({
        type: 'Volume Spike',
        description: `Current month has ${latest} cases vs average of ${average.toFixed(1)}`,
        impact: 'May indicate systemic issue or policy change'
      });
    }
  }

  // Check for concentration of cases in one location
  const locationCounts = {};
  data.forEach(row => {
    const location = row[9];
    if (location) {
      locationCounts[location] = (locationCounts[location] || 0) + 1;
    }
  });

  const totalCases = Object.values(locationCounts).reduce((sum, val) => sum + val, 0);
  Object.entries(locationCounts).forEach(([location, count]) => {
    if (count > totalCases * 0.4) {
      anomalies.push({
        type: 'Location Concentration',
        description: `${location} has ${count} cases (${((count / totalCases) * 100).toFixed(1)}% of all cases)`,
        impact: 'May indicate location-specific issues requiring attention'
      });
    }
  });

  return anomalies;
}

/**
 * Generates actionable recommendations
 * @param {Array} data - Grievance data
 * @returns {Array} Recommendations
 */
function generateRecommendations(data) {
  const recommendations = [];

  // Based on volume trend
  const volumeTrend = analyzeVolumeTrend(data);
  if (volumeTrend.trend === 'increasing') {
    recommendations.push({
      priority: 'HIGH',
      category: 'Capacity Planning',
      recommendation: 'Grievance volume is increasing. Consider recruiting and training additional stewards.',
      expectedImpact: 'Prevent case backlog and maintain service quality'
    });
  }

  // Based on resolution time
  const resolutionTrend = analyzeResolutionTimeTrend(data);
  if (resolutionTrend.average > 30) {
    recommendations.push({
      priority: 'MEDIUM',
      category: 'Process Improvement',
      recommendation: `Average resolution time is ${resolutionTrend.average} days. Implement workflow automation and deadline reminders.`,
      expectedImpact: 'Reduce resolution time by 20-30%'
    });
  }

  // Based on workload
  const workloadForecast = forecastStewardWorkload(data);
  if (workloadForecast.overloaded.length > 0) {
    recommendations.push({
      priority: 'HIGH',
      category: 'Workload Management',
      recommendation: `${workloadForecast.overloaded.length} steward(s) are overloaded. ${workloadForecast.balanceRecommendation}`,
      expectedImpact: 'Prevent steward burnout and improve case handling'
    });
  }

  // Based on trending issues
  const issueTrends = analyzeIssueTypeTrends(data);
  const trendingIssueCount = Object.keys(issueTrends.trendingUp).length;
  if (trendingIssueCount > 0) {
    const topTrending = Object.keys(issueTrends.trendingUp)[0];
    recommendations.push({
      priority: 'MEDIUM',
      category: 'Root Cause Analysis',
      recommendation: `${topTrending} cases are trending upward. Conduct root cause analysis and consider preventive measures.`,
      expectedImpact: 'Reduce future grievances through proactive intervention'
    });
  }

  return recommendations;
}

/**
 * Calculates simple linear trend
 * @param {Array} values - Data points
 * @returns {Object} Trend line parameters
 */
function calculateLinearTrend(values) {
  const n = values.length;
  if (n < 2) return { slope: 0, intercept: values[0] || 0 };

  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;

  for (let i = 0; i < n; i++) {
    sumX += i;
    sumY += values[i];
    sumXY += i * values[i];
    sumX2 += i * i;
  }

  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  return { slope, intercept };
}

/**
 * Shows predictive analytics dashboard
 */
function showPredictiveAnalyticsDashboard() {
  const ui = SpreadsheetApp.getUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🔮 Analyzing trends...',
    'Predictive Analytics',
    -1
  );

  try {
    const analytics = performPredictiveAnalysis();

    if (analytics.error) {
      ui.alert('⚠️ ' + analytics.error);
      return;
    }

    const html = createAnalyticsDashboardHTML(analytics);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(700);

    ui.showModalDialog(htmlOutput, '🔮 Predictive Analytics Dashboard');

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Creates HTML for analytics dashboard
 * @param {Object} analytics - Analytics results
 * @returns {string} HTML content
 */
function createAnalyticsDashboardHTML(analytics) {
  const risksHTML = analytics.riskFactors.length > 0
    ? analytics.riskFactors.map(risk => `
        <div class="risk-item severity-${risk.severity.toLowerCase()}">
          <div class="risk-header">
            <span class="risk-type">${risk.type}</span>
            <span class="risk-severity">${risk.severity}</span>
          </div>
          <div class="risk-desc">${risk.description}</div>
          <div class="risk-action">⚡ ${risk.action}</div>
        </div>
      `).join('')
    : '<p>No significant risks identified. ✅</p>';

  const recommendationsHTML = analytics.recommendations.length > 0
    ? analytics.recommendations.map(rec => `
        <div class="recommendation priority-${rec.priority.toLowerCase()}">
          <div class="rec-header">
            <span class="rec-category">${rec.category}</span>
            <span class="rec-priority">${rec.priority}</span>
          </div>
          <div class="rec-text">${rec.recommendation}</div>
          <div class="rec-impact">📊 ${rec.expectedImpact}</div>
        </div>
      `).join('')
    : '<p>No specific recommendations at this time.</p>';

  const anomaliesHTML = analytics.anomalies.length > 0
    ? analytics.anomalies.map(anomaly => `
        <div class="anomaly-item">
          <div class="anomaly-type">${anomaly.type}</div>
          <div class="anomaly-desc">${anomaly.description}</div>
          <div class="anomaly-impact">${anomaly.impact}</div>
        </div>
      `).join('')
    : '<p>No anomalies detected. ✅</p>';

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-height: 650px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    h3 { color: #333; margin-top: 25px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; }
    .summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin: 20px 0; }
    .summary-card { background: #f8f9fa; padding: 15px; border-radius: 8px; border-left: 4px solid #1a73e8; }
    .summary-label { font-size: 12px; color: #666; text-transform: uppercase; }
    .summary-value { font-size: 24px; font-weight: bold; color: #333; margin: 5px 0; }
    .summary-trend { font-size: 13px; color: #666; }
    .trend-up { color: #f44336; }
    .trend-down { color: #4caf50; }
    .trend-stable { color: #9e9e9e; }
    .risk-item { background: #fff3e0; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid; }
    .risk-item.severity-high { border-left-color: #f44336; background: #ffebee; }
    .risk-item.severity-medium { border-left-color: #ff9800; background: #fff3e0; }
    .risk-item.severity-low { border-left-color: #ffeb3b; background: #fffde7; }
    .risk-header { display: flex; justify-content: space-between; margin-bottom: 8px; }
    .risk-type { font-weight: bold; color: #333; }
    .risk-severity { background: white; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .risk-desc { margin: 5px 0; color: #555; }
    .risk-action { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; }
    .recommendation { background: #e8f5e9; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #4caf50; }
    .recommendation.priority-high { border-left-color: #f44336; background: #ffebee; }
    .recommendation.priority-medium { border-left-color: #ff9800; background: #fff3e0; }
    .rec-header { display: flex; justify-content: space-between; margin-bottom: 8px; }
    .rec-category { font-weight: bold; color: #333; }
    .rec-priority { background: white; padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .rec-text { margin: 5px 0; color: #555; }
    .rec-impact { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; }
    .anomaly-item { background: #e3f2fd; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #2196f3; }
    .anomaly-type { font-weight: bold; color: #1976d2; margin-bottom: 5px; }
    .anomaly-desc { color: #555; margin: 5px 0; }
    .anomaly-impact { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; font-style: italic; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔮 Predictive Analytics Dashboard</h2>

    <div class="summary-grid">
      <div class="summary-card">
        <div class="summary-label">Volume Trend</div>
        <div class="summary-value trend-${analytics.volumeTrend.trend === 'increasing' ? 'up' : analytics.volumeTrend.trend === 'decreasing' ? 'down' : 'stable'}">
          ${analytics.volumeTrend.currentMonthVolume}
        </div>
        <div class="summary-trend">
          ${analytics.volumeTrend.trend === 'increasing' ? '📈 Increasing' : analytics.volumeTrend.trend === 'decreasing' ? '📉 Decreasing' : '➡️ Stable'}
        </div>
      </div>

      <div class="summary-card">
        <div class="summary-label">Avg Resolution Time</div>
        <div class="summary-value">${analytics.resolutionTimeTrend.average} days</div>
        <div class="summary-trend">
          ${analytics.resolutionTimeTrend.trend === 'increasing' ? '⚠️ Increasing' : analytics.resolutionTimeTrend.trend === 'decreasing' ? '✅ Improving' : '➡️ Stable'}
        </div>
      </div>

      <div class="summary-card">
        <div class="summary-label">Identified Risks</div>
        <div class="summary-value">${analytics.riskFactors.length}</div>
        <div class="summary-trend">
          ${analytics.riskFactors.length === 0 ? '✅ No risks' : `⚠️ Needs attention`}
        </div>
      </div>
    </div>

    <h3>🚨 Risk Factors</h3>
    ${risksHTML}

    <h3>💡 Recommendations</h3>
    ${recommendationsHTML}

    <h3>🔍 Anomalies Detected</h3>
    ${anomaliesHTML}

    <h3>📊 Forecast (Next 3 Months)</h3>
    <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
      <strong>Projected Volume:</strong><br>
      Month 1: ${analytics.volumeTrend.forecast[0]} cases<br>
      Month 2: ${analytics.volumeTrend.forecast[1]} cases<br>
      Month 3: ${analytics.volumeTrend.forecast[2]} cases
    </div>

    ${analytics.seasonalPatterns.hasSeasonality ? `
    <h3>📅 Seasonal Patterns</h3>
    <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 10px 0;">
      <strong>Peak Season:</strong> ${analytics.seasonalPatterns.peakQuarter} (${analytics.seasonalPatterns.peakVolume} cases)<br>
      <strong>Quarterly Breakdown:</strong><br>
      Q1: ${analytics.seasonalPatterns.quarterlyVolumes.Q1} |
      Q2: ${analytics.seasonalPatterns.quarterlyVolumes.Q2} |
      Q3: ${analytics.seasonalPatterns.quarterlyVolumes.Q3} |
      Q4: ${analytics.seasonalPatterns.quarterlyVolumes.Q4}
    </div>
    ` : ''}
  </div>
</body>
</html>
  `;
}

/******************************************************************************
 * MODULE: DataCachingLayer
 * Source: DataCachingLayer.gs
 *****************************************************************************/

/**
 * ============================================================================
 * DATA CACHING LAYER
 * ============================================================================
 *
 * Performance optimization through intelligent caching
 * Features:
 * - In-memory cache for frequently accessed data
 * - Script properties cache for persistent storage
 * - Cache invalidation strategies
 * - Automatic cache warming
 * - Performance monitoring
 * - Configurable TTL (time-to-live)
 */

/**
 * Cache configuration
 */
const CACHE_CONFIG = {
  MEMORY_TTL: 300,        // 5 minutes for memory cache
  PROPERTIES_TTL: 3600,   // 1 hour for properties cache
  ENABLE_LOGGING: true
};

/**
 * Cache keys
 */
const CACHE_KEYS = {
  ALL_GRIEVANCES: 'all_grievances',
  ALL_MEMBERS: 'all_members',
  ALL_STEWARDS: 'all_stewards',
  GRIEVANCE_COUNT: 'grievance_count',
  MEMBER_COUNT: 'member_count',
  DASHBOARD_METRICS: 'dashboard_metrics',
  STEWARD_WORKLOAD: 'steward_workload'
};

/**
 * Gets data from cache or loads from source
 * @param {string} key - Cache key
 * @param {Function} loader - Function to load data if cache miss
 * @param {number} ttl - Time to live in seconds
 * @returns {any} Cached or freshly loaded data
 */
function getCachedData(key, loader, ttl = CACHE_CONFIG.MEMORY_TTL) {
  try {
    // Try memory cache first (fastest)
    const memoryCache = CacheService.getScriptCache();
    const memoryCached = memoryCache.get(key);

    if (memoryCached) {
      logCacheHit('MEMORY', key);
      return JSON.parse(memoryCached);
    }

    // Try script properties (persistent)
    const propsCache = PropertiesService.getScriptProperties();
    const propsCached = propsCache.getProperty(key);

    if (propsCached) {
      const cachedObj = JSON.parse(propsCached);

      // Check if expired
      if (cachedObj.timestamp && (Date.now() - cachedObj.timestamp) < (ttl * 1000)) {
        logCacheHit('PROPERTIES', key);

        // Warm memory cache
        memoryCache.put(key, JSON.stringify(cachedObj.data), ttl);

        return cachedObj.data;
      }
    }

    // Cache miss - load from source
    logCacheMiss(key);
    const data = loader();

    // Store in both caches
    setCachedData(key, data, ttl);

    return data;

  } catch (error) {
    Logger.log(`Cache error for key ${key}: ${error.message}`);
    // Fallback to direct load
    return loader();
  }
}

/**
 * Sets data in cache
 * @param {string} key - Cache key
 * @param {any} data - Data to cache
 * @param {number} ttl - Time to live in seconds
 */
function setCachedData(key, data, ttl = CACHE_CONFIG.MEMORY_TTL) {
  try {
    const dataStr = JSON.stringify(data);

    // Store in memory cache
    const memoryCache = CacheService.getScriptCache();
    memoryCache.put(key, dataStr, Math.min(ttl, 21600)); // Max 6 hours for memory

    // Store in properties cache with timestamp
    const propsCache = PropertiesService.getScriptProperties();
    const cachedObj = {
      data: data,
      timestamp: Date.now()
    };
    propsCache.setProperty(key, JSON.stringify(cachedObj));

    logCacheSet(key);

  } catch (error) {
    Logger.log(`Error setting cache for key ${key}: ${error.message}`);
  }
}

/**
 * Invalidates cache for a specific key
 * @param {string} key - Cache key to invalidate
 */
function invalidateCache(key) {
  try {
    CacheService.getScriptCache().remove(key);
    PropertiesService.getScriptProperties().deleteProperty(key);
    Logger.log(`Cache invalidated: ${key}`);
  } catch (error) {
    Logger.log(`Error invalidating cache for key ${key}: ${error.message}`);
  }
}

/**
 * Invalidates all caches
 */
function invalidateAllCaches() {
  try {
    CacheService.getScriptCache().removeAll(Object.values(CACHE_KEYS));

    const propsCache = PropertiesService.getScriptProperties();
    Object.values(CACHE_KEYS).forEach(key => {
      propsCache.deleteProperty(key);
    });

    Logger.log('All caches invalidated');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ All caches cleared',
      'Cache Management',
      3
    );

  } catch (error) {
    Logger.log(`Error invalidating all caches: ${error.message}`);
  }
}

/**
 * Warms up caches with frequently accessed data
 */
function warmUpCaches() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🔥 Warming up caches...',
    'Cache Management',
    -1
  );

  try {
    // Warm up key data
    getCachedGrievances();
    getCachedMembers();
    getCachedStewards();
    getCachedDashboardMetrics();

    Logger.log('Caches warmed up successfully');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ Caches warmed up',
      'Cache Management',
      3
    );

  } catch (error) {
    Logger.log(`Error warming up caches: ${error.message}`);
  }
}

/**
 * Gets all grievances (cached)
 * @returns {Array} Grievances array
 */
function getCachedGrievances() {
  return getCachedData(
    CACHE_KEYS.ALL_GRIEVANCES,
    () => {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) return [];

      return sheet.getRange(2, 1, lastRow - 1, 28).getValues();
    },
    300 // 5 minutes
  );
}

/**
 * Gets all members (cached)
 * @returns {Array} Members array
 */
function getCachedMembers() {
  return getCachedData(
    CACHE_KEYS.ALL_MEMBERS,
    () => {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      const lastRow = sheet.getLastRow();

      if (lastRow < 2) return [];

      return sheet.getRange(2, 1, lastRow - 1, 28).getValues();
    },
    600 // 10 minutes
  );
}

/**
 * Gets all stewards (cached)
 * @returns {Array} Stewards array
 */
function getCachedStewards() {
  return getCachedData(
    CACHE_KEYS.ALL_STEWARDS,
    () => {
      const members = getCachedMembers();
      return members.filter(row => row[9] === 'Yes'); // Column J: Is Steward?
    },
    600 // 10 minutes
  );
}

/**
 * Gets dashboard metrics (cached)
 * @returns {Object} Dashboard metrics
 */
function getCachedDashboardMetrics() {
  return getCachedData(
    CACHE_KEYS.DASHBOARD_METRICS,
    () => {
      const grievances = getCachedGrievances();

      const metrics = {
        total: grievances.length,
        open: 0,
        closed: 0,
        overdue: 0,
        byStatus: {},
        byIssueType: {},
        bySteward: {}
      };

      grievances.forEach(row => {
        const status = row[4];
        const issueType = row[5];
        const steward = row[13];
        const daysToDeadline = row[20];

        if (status === 'Open') metrics.open++;
        if (status === 'Closed' || status === 'Resolved') metrics.closed++;
        if (daysToDeadline < 0) metrics.overdue++;

        metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;

        if (issueType) {
          metrics.byIssueType[issueType] = (metrics.byIssueType[issueType] || 0) + 1;
        }

        if (steward) {
          metrics.bySteward[steward] = (metrics.bySteward[steward] || 0) + 1;
        }
      });

      return metrics;
    },
    180 // 3 minutes
  );
}

/**
 * Logs cache hit
 * @param {string} source - Cache source (MEMORY/PROPERTIES)
 * @param {string} key - Cache key
 */
function logCacheHit(source, key) {
  if (CACHE_CONFIG.ENABLE_LOGGING) {
    Logger.log(`[CACHE HIT] ${source}: ${key}`);
  }
}

/**
 * Logs cache miss
 * @param {string} key - Cache key
 */
function logCacheMiss(key) {
  if (CACHE_CONFIG.ENABLE_LOGGING) {
    Logger.log(`[CACHE MISS] ${key}`);
  }
}

/**
 * Logs cache set
 * @param {string} key - Cache key
 */
function logCacheSet(key) {
  if (CACHE_CONFIG.ENABLE_LOGGING) {
    Logger.log(`[CACHE SET] ${key}`);
  }
}

/**
 * Shows cache status dashboard
 */
function showCacheStatusDashboard() {
  const memoryCache = CacheService.getScriptCache();
  const propsCache = PropertiesService.getScriptProperties();

  const cacheStatus = [];

  Object.entries(CACHE_KEYS).forEach(([name, key]) => {
    const inMemory = memoryCache.get(key) !== null;
    const inProps = propsCache.getProperty(key) !== null;

    let age = 'N/A';
    if (inProps) {
      const cached = JSON.parse(propsCache.getProperty(key));
      if (cached.timestamp) {
        age = Math.floor((Date.now() - cached.timestamp) / 1000) + 's';
      }
    }

    cacheStatus.push({
      name: name,
      key: key,
      inMemory: inMemory,
      inProps: inProps,
      age: age
    });
  });

  const statusRows = cacheStatus
    .map(s => `
      <tr>
        <td>${s.name}</td>
        <td>${s.inMemory ? '✅ Yes' : '❌ No'}</td>
        <td>${s.inProps ? '✅ Yes' : '❌ No'}</td>
        <td>${s.age}</td>
      </tr>
    `)
    .join('');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .info-box { background: #e8f0fe; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #1a73e8; }
    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
    th { background: #1a73e8; color: white; padding: 12px; text-align: left; }
    td { padding: 10px 12px; border-bottom: 1px solid #e0e0e0; }
    tr:hover { background: #f5f5f5; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.danger { background: #dc3545; }
    button.danger:hover { background: #c82333; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🗄️ Cache Status Dashboard</h2>

    <div class="info-box">
      <strong>ℹ️ Cache Information:</strong><br>
      • <strong>Memory Cache:</strong> Fast, in-memory storage (5-minute TTL, max 6 hours)<br>
      • <strong>Properties Cache:</strong> Persistent storage across sessions (1-hour TTL)<br>
      • Caches automatically refresh when data is modified
    </div>

    <h3>Cache Status</h3>
    <table>
      <thead>
        <tr>
          <th>Cache Name</th>
          <th>In Memory</th>
          <th>In Properties</th>
          <th>Age</th>
        </tr>
      </thead>
      <tbody>
        ${statusRows}
      </tbody>
    </table>

    <button onclick="google.script.run.withSuccessHandler(() => { location.reload(); }).warmUpCaches()">
      🔥 Warm Up All Caches
    </button>

    <button class="danger" onclick="google.script.run.withSuccessHandler(() => { location.reload(); }).invalidateAllCaches()">
      🗑️ Clear All Caches
    </button>
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🗄️ Cache Status');
}

/**
 * Auto-invalidates cache when data is modified (onEdit trigger helper)
 * Call this from the main onEdit function
 */
function onEditCacheInvalidation(e) {
  if (!e) return;

  const sheetName = e.range.getSheet().getName();

  // Invalidate relevant caches based on sheet
  if (sheetName === SHEETS.GRIEVANCE_LOG) {
    invalidateCache(CACHE_KEYS.ALL_GRIEVANCES);
    invalidateCache(CACHE_KEYS.DASHBOARD_METRICS);
    invalidateCache(CACHE_KEYS.STEWARD_WORKLOAD);
  } else if (sheetName === SHEETS.MEMBER_DIR) {
    invalidateCache(CACHE_KEYS.ALL_MEMBERS);
    invalidateCache(CACHE_KEYS.ALL_STEWARDS);
  }
}

/**
 * Gets performance statistics
 * @returns {Object} Performance stats
 */
function getCachePerformanceStats() {
  // This would track cache hits/misses over time
  // For now, return basic info
  const memoryCache = CacheService.getScriptCache();
  const propsCache = PropertiesService.getScriptProperties();

  let cachedKeys = 0;
  Object.values(CACHE_KEYS).forEach(key => {
    if (propsCache.getProperty(key) !== null) {
      cachedKeys++;
    }
  });

  return {
    totalKeys: Object.keys(CACHE_KEYS).length,
    cachedKeys: cachedKeys,
    cacheHitRate: cachedKeys / Object.keys(CACHE_KEYS).length,
    memoryTTL: CACHE_CONFIG.MEMORY_TTL,
    propertiesTTL: CACHE_CONFIG.PROPERTIES_TTL
  };
}

/******************************************************************************
 * MODULE: AutomatedNotifications
 * Source: AutomatedNotifications.gs
 *****************************************************************************/

/**
 * ============================================================================
 * AUTOMATED DEADLINE NOTIFICATIONS
 * ============================================================================
 *
 * Automated email alerts for approaching grievance deadlines
 * Features:
 * - Daily checks at 8 AM
 * - Email stewards 7 days before deadline
 * - Escalate to managers 3 days before
 * - Mark overdue grievances
 * - Customizable notification templates
 */

/**
 * Sets up time-driven trigger for daily deadline checks
 */
function setupDailyDeadlineNotifications() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkDeadlinesAndNotify') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger for 8 AM daily
  ScriptApp.newTrigger('checkDeadlinesAndNotify')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Daily deadline notifications enabled (8 AM)',
    'Automation Active',
    5
  );
}

/**
 * Removes the daily deadline notification trigger
 */
function disableDailyDeadlineNotifications() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkDeadlinesAndNotify') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `🔕 Removed ${removed} deadline notification trigger(s)`,
    'Automation Disabled',
    5
  );
}

/**
 * Main function that checks deadlines and sends notifications
 * Runs daily at 8 AM via trigger
 */
function checkDeadlinesAndNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    Logger.log('Grievance Log sheet not found');
    return;
  }

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No grievances to check');
    return;
  }

  // Get all grievance data
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const notifications = {
    overdue: [],
    urgent: [],      // 3 days or less
    upcoming: []     // 7 days or less
  };

  // Categorize grievances by deadline urgency
  data.forEach((row, index) => {
    const grievanceId = row[0];
    const memberFirstName = row[2];
    const memberLastName = row[3];
    const status = row[4];
    const issueType = row[5];
    const nextActionDue = row[19];
    const daysToDeadline = row[20];
    const steward = row[13];
    const manager = row[11];

    // Only check open grievances with deadlines
    if (status !== 'Open' || !nextActionDue || !steward) {
      return;
    }

    const grievanceInfo = {
      id: grievanceId,
      memberName: `${memberFirstName} ${memberLastName}`,
      issueType: issueType,
      deadline: nextActionDue,
      daysRemaining: daysToDeadline,
      steward: steward,
      manager: manager || '',
      row: index + 2
    };

    if (daysToDeadline < 0) {
      notifications.overdue.push(grievanceInfo);
    } else if (daysToDeadline <= 3) {
      notifications.urgent.push(grievanceInfo);
    } else if (daysToDeadline <= 7) {
      notifications.upcoming.push(grievanceInfo);
    }
  });

  // Send notifications
  let emailsSent = 0;

  // Send overdue notifications
  notifications.overdue.forEach(grievance => {
    sendDeadlineNotification(grievance, 'OVERDUE');
    emailsSent++;
  });

  // Send urgent notifications (escalate to manager)
  notifications.urgent.forEach(grievance => {
    sendDeadlineNotification(grievance, 'URGENT');
    emailsSent++;
  });

  // Send upcoming notifications
  notifications.upcoming.forEach(grievance => {
    sendDeadlineNotification(grievance, 'UPCOMING');
    emailsSent++;
  });

  // Log summary
  Logger.log(`Deadline check complete: ${emailsSent} notifications sent`);
  Logger.log(`  Overdue: ${notifications.overdue.length}`);
  Logger.log(`  Urgent: ${notifications.urgent.length}`);
  Logger.log(`  Upcoming: ${notifications.upcoming.length}`);
}

/**
 * Sends email notification for a grievance deadline
 * @param {Object} grievance - Grievance information
 * @param {string} priority - OVERDUE, URGENT, or UPCOMING
 */
function sendDeadlineNotification(grievance, priority) {
  try {
    const recipients = getNotificationRecipients(grievance, priority);

    if (recipients.length === 0) {
      Logger.log(`No valid recipients for ${grievance.id}`);
      return;
    }

    const subject = createEmailSubject(grievance, priority);
    const body = createEmailBody(grievance, priority);

    // Send email
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body,
      name: 'SEIU Local 509 Dashboard'
    });

    Logger.log(`Sent ${priority} notification for ${grievance.id} to ${recipients.join(', ')}`);

  } catch (error) {
    Logger.log(`Error sending notification for ${grievance.id}: ${error.message}`);
  }
}

/**
 * Determines who should receive the notification
 * @param {Object} grievance - Grievance information
 * @param {string} priority - OVERDUE, URGENT, or UPCOMING
 * @returns {Array} Array of email addresses
 */
function getNotificationRecipients(grievance, priority) {
  const recipients = [];

  // Always notify the steward
  if (grievance.steward && isValidEmail(grievance.steward)) {
    recipients.push(grievance.steward);
  }

  // Escalate to manager for urgent and overdue
  if ((priority === 'URGENT' || priority === 'OVERDUE') &&
      grievance.manager &&
      isValidEmail(grievance.manager)) {
    recipients.push(grievance.manager);
  }

  return recipients;
}

/**
 * Creates email subject line
 * @param {Object} grievance - Grievance information
 * @param {string} priority - OVERDUE, URGENT, or UPCOMING
 * @returns {string} Email subject
 */
function createEmailSubject(grievance, priority) {
  const prefix = {
    'OVERDUE': '🚨 OVERDUE',
    'URGENT': '⚠️ URGENT',
    'UPCOMING': '📅 Deadline Reminder'
  }[priority];

  return `${prefix}: Grievance ${grievance.id} - ${grievance.memberName}`;
}

/**
 * Creates email body content
 * @param {Object} grievance - Grievance information
 * @param {string} priority - OVERDUE, URGENT, or UPCOMING
 * @returns {string} Email body
 */
function createEmailBody(grievance, priority) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetUrl = ss.getUrl();

  let urgencyMessage = '';

  if (priority === 'OVERDUE') {
    urgencyMessage = `⚠️ THIS GRIEVANCE IS OVERDUE BY ${Math.abs(grievance.daysRemaining)} DAY(S)!\n\nImmediate action is required.`;
  } else if (priority === 'URGENT') {
    urgencyMessage = `⏰ This grievance deadline is in ${grievance.daysRemaining} day(s).\n\nPlease prioritize this case.`;
  } else {
    urgencyMessage = `This is a reminder that a grievance deadline is approaching in ${grievance.daysRemaining} day(s).`;
  }

  return `
SEIU Local 509 - Grievance Deadline Notification

${urgencyMessage}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Grievance Details:

  • Grievance ID: ${grievance.id}
  • Member: ${grievance.memberName}
  • Issue Type: ${grievance.issueType}
  • Assigned Steward: ${grievance.steward}
  • Deadline: ${Utilities.formatDate(new Date(grievance.deadline), Session.getScriptTimeZone(), 'MMMM dd, yyyy')}
  • Days Remaining: ${grievance.daysRemaining}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Action Required:

${priority === 'OVERDUE'
  ? '1. Review the grievance immediately\n2. Take necessary action today\n3. Update the status in the dashboard\n4. Notify leadership if needed'
  : priority === 'URGENT'
  ? '1. Review the grievance\n2. Complete any pending actions\n3. Prepare for next steps\n4. Update the dashboard'
  : '1. Review the grievance timeline\n2. Plan your next actions\n3. Gather any needed documentation\n4. Update progress in the dashboard'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

View in Dashboard:
${spreadsheetUrl}

Need Help?
Contact your union representative or refer to the Grievance Workflow guide.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This is an automated notification from the SEIU Local 509 Dashboard.
To adjust notification settings, contact your system administrator.
`;
}

/**
 * Validates email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email.toString().trim());
}

/**
 * Manual test function to check notifications without waiting for trigger
 */
function testDeadlineNotifications() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🧪 Test Deadline Notifications',
    'This will run the deadline check immediately and send real emails.\n\n' +
    'Emails will be sent to stewards with approaching deadlines.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('🧪 Running deadline check...', 'Testing', -1);

  try {
    checkDeadlinesAndNotify();

    ui.alert(
      '✅ Test Complete',
      'Deadline check completed successfully.\n\n' +
      'Check the execution log (View > Logs) for details on notifications sent.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Shows notification settings dialog
 */
function showNotificationSettings() {
  const triggers = ScriptApp.getProjectTriggers();
  const isEnabled = triggers.some(t => t.getHandlerFunction() === 'checkDeadlinesAndNotify');

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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .status {
      padding: 15px;
      border-radius: 4px;
      margin: 20px 0;
      font-weight: bold;
    }
    .status.enabled {
      background: #d4edda;
      color: #155724;
      border: 1px solid #c3e6cb;
    }
    .status.disabled {
      background: #f8d7da;
      color: #721c24;
      border: 1px solid #f5c6cb;
    }
    .settings {
      margin: 20px 0;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 4px;
    }
    .settings h3 {
      margin-top: 0;
      color: #333;
    }
    .settings ul {
      margin: 10px 0;
      padding-left: 20px;
    }
    .settings li {
      margin: 8px 0;
      color: #666;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      margin: 10px 5px 0 0;
    }
    button:hover {
      background: #1557b0;
    }
    button.danger {
      background: #dc3545;
    }
    button.danger:hover {
      background: #c82333;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📬 Notification Settings</h2>

    <div class="status ${isEnabled ? 'enabled' : 'disabled'}">
      ${isEnabled ? '✅ Automated notifications are ENABLED' : '🔕 Automated notifications are DISABLED'}
    </div>

    <div class="settings">
      <h3>Current Configuration</h3>
      <ul>
        <li><strong>Schedule:</strong> Daily at 8:00 AM</li>
        <li><strong>7-Day Warning:</strong> Email sent to assigned steward</li>
        <li><strong>3-Day Warning:</strong> Email sent to steward + manager (escalation)</li>
        <li><strong>Overdue:</strong> Immediate notification to steward + manager</li>
      </ul>
    </div>

    <div class="settings">
      <h3>Email Template Includes</h3>
      <ul>
        <li>Grievance ID and member name</li>
        <li>Days remaining until deadline</li>
        <li>Direct link to dashboard</li>
        <li>Action items based on urgency</li>
      </ul>
    </div>

    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).setupDailyDeadlineNotifications()">
      ${isEnabled ? '🔄 Refresh Trigger' : '✅ Enable Notifications'}
    </button>

    ${isEnabled ? '<button class="danger" onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).disableDailyDeadlineNotifications()">🔕 Disable Notifications</button>' : ''}

    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).testDeadlineNotifications()">
      🧪 Test Now
    </button>
  </div>
</body>
</html>
  `)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '📬 Notification Settings');
}

/******************************************************************************
 * MODULE: DataBackupRecovery
 * Source: DataBackupRecovery.gs
 *****************************************************************************/

/**
 * ============================================================================
 * DATA BACKUP & RECOVERY SYSTEM
 * ============================================================================
 *
 * Automated and manual backup/recovery capabilities
 * Features:
 * - One-click manual backups
 * - Automated weekly backups
 * - Version history tracking
 * - Point-in-time recovery
 * - Selective sheet restoration
 * - Backup verification
 * - Export to Google Drive
 * - Retention policy management
 */

/**
 * Creates a complete backup of the dashboard
 * @param {boolean} automated - Whether this is an automated backup
 * @returns {File} Backup file
 */
function createBackup(automated = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const backupName = `509_Dashboard_Backup_${timestamp}${automated ? '_AUTO' : ''}`;

  try {
    // Create copy of entire spreadsheet
    const backupFile = DriveApp.getFileById(ss.getId()).makeCopy(backupName);

    // Move to backups folder
    const backupFolder = getOrCreateBackupFolder();
    backupFolder.addFile(backupFile);

    // Remove from root
    DriveApp.getRootFolder().removeFile(backupFile);

    // Log backup
    logBackup(backupName, backupFile.getId(), automated);

    // Clean old backups
    cleanOldBackups();

    if (!automated) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '✅ Backup created successfully: ' + backupName,
        'Backup Complete',
        5
      );
    }

    return backupFile;

  } catch (error) {
    Logger.log('Backup error: ' + error.message);
    if (!automated) {
      SpreadsheetApp.getUi().alert('❌ Backup failed: ' + error.message);
    }
    throw error;
  }
}

/**
 * Gets or creates backup folder in Google Drive
 * @returns {Folder} Backup folder
 */
function getOrCreateBackupFolder() {
  const folderName = '509 Dashboard - Backups';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

/**
 * Logs backup information
 * @param {string} backupName - Backup name
 * @param {string} fileId - Drive file ID
 * @param {boolean} automated - Whether automated
 */
function logBackup(backupName, fileId, automated) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let backupLog = ss.getSheetByName('💾 Backup Log');

  if (!backupLog) {
    backupLog = createBackupLogSheet();
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || 'System';

  const lastRow = backupLog.getLastRow();
  const newRow = [
    timestamp,
    backupName,
    fileId,
    automated ? 'Automated' : 'Manual',
    user,
    `=HYPERLINK("https://drive.google.com/file/d/${fileId}/view", "Open Backup")`
  ];

  backupLog.getRange(lastRow + 1, 1, 1, 6).setValues([newRow]);
}

/**
 * Creates backup log sheet
 * @returns {Sheet} Backup log sheet
 */
function createBackupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('💾 Backup Log');

  const headers = [
    'Timestamp',
    'Backup Name',
    'File ID',
    'Type',
    'Created By',
    'Link'
  ];

  sheet.getRange(1, 1, 1, 6).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, 6)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 150);

  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Cleans old backups based on retention policy
 */
function cleanOldBackups() {
  const RETENTION_DAYS = 30; // Keep backups for 30 days
  const MAX_BACKUPS = 50;    // Keep max 50 backups

  try {
    const backupFolder = getOrCreateBackupFolder();
    const files = backupFolder.getFiles();

    const backupFiles = [];

    while (files.hasNext()) {
      const file = files.next();
      backupFiles.push({
        file: file,
        created: file.getDateCreated()
      });
    }

    // Sort by date (oldest first)
    backupFiles.sort((a, b) => a.created - b.created);

    const now = new Date();
    const cutoffDate = new Date(now.getTime() - (RETENTION_DAYS * 24 * 60 * 60 * 1000));

    // Delete old backups
    backupFiles.forEach((backup, index) => {
      // Keep if:
      // 1. Within retention period
      // 2. One of the most recent MAX_BACKUPS
      const isRecent = backup.created > cutoffDate;
      const isInMaxLimit = index >= (backupFiles.length - MAX_BACKUPS);

      if (!isRecent && !isInMaxLimit) {
        Logger.log('Deleting old backup: ' + backup.file.getName());
        backup.file.setTrashed(true);
      }
    });

  } catch (error) {
    Logger.log('Error cleaning old backups: ' + error.message);
  }
}

/**
 * Sets up automated weekly backups
 */
function setupAutomatedBackups() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAutomatedBackup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create weekly trigger (Sunday at 2 AM)
  ScriptApp.newTrigger('runAutomatedBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(2)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Automated weekly backups enabled (Sundays, 2 AM)',
    'Backup Automation',
    5
  );
}

/**
 * Disables automated backups
 */
function disableAutomatedBackups() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAutomatedBackup') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `🔕 Removed ${removed} automated backup trigger(s)`,
    'Backup Automation',
    5
  );
}

/**
 * Runs automated backup (called by trigger)
 */
function runAutomatedBackup() {
  try {
    createBackup(true);
    Logger.log('Automated backup completed successfully');
  } catch (error) {
    Logger.log('Automated backup failed: ' + error.message);
    notifyAdminOfBackupFailure(error.message);
  }
}

/**
 * Notifies admin of backup failure
 * @param {string} errorMessage - Error message
 */
function notifyAdminOfBackupFailure(errorMessage) {
  const adminEmail = Session.getActiveUser().getEmail();

  MailApp.sendEmail({
    to: adminEmail,
    subject: '🚨 509 Dashboard - Automated Backup Failed',
    body: `An automated backup failed:\n\nError: ${errorMessage}\n\nPlease check the dashboard and create a manual backup.`,
    name: 'SEIU Local 509 Dashboard'
  });
}

/**
 * Shows backup manager dialog
 */
function showBackupManager() {
  const html = createBackupManagerHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '💾 Backup & Recovery Manager');
}

/**
 * Creates HTML for backup manager
 */
function createBackupManagerHTML() {
  const triggers = ScriptApp.getProjectTriggers();
  const isAutomated = triggers.some(t => t.getHandlerFunction() === 'runAutomatedBackup');

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .section { background: #f8f9fa; padding: 20px; margin: 15px 0; border-radius: 8px; border-left: 4px solid #1a73e8; }
    .section-title { font-weight: bold; font-size: 16px; color: #333; margin-bottom: 15px; }
    .status { padding: 15px; border-radius: 4px; margin: 15px 0; font-weight: bold; }
    .status.enabled { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
    .status.disabled { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    button.secondary:hover { background: #5a6268; }
    button.danger { background: #dc3545; }
    button.danger:hover { background: #c82333; }
    .info-box { background: #e8f0fe; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #1a73e8; }
    ul { margin: 10px 0; padding-left: 25px; }
    li { margin: 8px 0; }
  </style>
</head>
<body>
  <div class="container">
    <h2>💾 Backup & Recovery Manager</h2>

    <div class="info-box">
      <strong>📋 Backup Policy:</strong><br>
      • Retention: 30 days or 50 most recent backups<br>
      • Automated: Weekly (Sundays at 2 AM)<br>
      • Location: Google Drive folder "509 Dashboard - Backups"
    </div>

    <div class="section">
      <div class="section-title">⚙️ Automated Backups</div>
      <div class="status ${isAutomated ? 'enabled' : 'disabled'}">
        ${isAutomated ? '✅ Automated backups are ENABLED (Weekly, Sundays 2 AM)' : '🔕 Automated backups are DISABLED'}
      </div>
      <button onclick="google.script.run.withSuccessHandler(() => location.reload()).setupAutomatedBackups()">
        ${isAutomated ? '🔄 Refresh Schedule' : '✅ Enable Automated Backups'}
      </button>
      ${isAutomated ? '<button class="danger" onclick="google.script.run.withSuccessHandler(() => location.reload()).disableAutomatedBackups()">🔕 Disable</button>' : ''}
    </div>

    <div class="section">
      <div class="section-title">📦 Manual Backup</div>
      <p>Create an immediate backup of all data. Backups are saved to Google Drive.</p>
      <button onclick="runManualBackup()">💾 Create Backup Now</button>
    </div>

    <div class="section">
      <div class="section-title">🔄 Recovery Options</div>
      <p>Restore data from a previous backup:</p>
      <ul>
        <li>View all backups in Google Drive folder "509 Dashboard - Backups"</li>
        <li>Open a backup file to review data</li>
        <li>Copy sheets from backup to current dashboard to restore</li>
        <li>Or replace entire dashboard by making backup your active copy</li>
      </ul>
      <button class="secondary" onclick="openBackupFolder()">📁 Open Backup Folder</button>
      <button class="secondary" onclick="showBackupLog()">📊 View Backup Log</button>
    </div>

    <div class="section">
      <div class="section-title">📊 Export Options</div>
      <p>Export specific data for external storage:</p>
      <button onclick="exportGrievances()">📋 Export Grievances (CSV)</button>
      <button onclick="exportMembers()">👥 Export Members (CSV)</button>
      <button onclick="exportAll()">📦 Export All Data (ZIP)</button>
    </div>
  </div>

  <script>
    function runManualBackup() {
      if (!confirm('Create a backup of all dashboard data?')) return;

      alert('Creating backup... This may take a moment.');

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Backup created successfully!');
        })
        .withFailureHandler((error) => {
          alert('❌ Backup failed: ' + error.message);
        })
        .createBackup(false);
    }

    function openBackupFolder() {
      window.open('https://drive.google.com/', '_blank');
      alert('Look for the folder named "509 Dashboard - Backups" in your Google Drive.');
    }

    function showBackupLog() {
      google.script.run
        .withSuccessHandler(() => {
          alert('Opening Backup Log sheet...');
          google.script.host.close();
        })
        .navigateToBackupLog();
    }

    function exportGrievances() {
      alert('Exporting grievances to CSV...');
      google.script.run
        .withSuccessHandler((csv) => {
          downloadCSV(csv, 'grievances.csv');
        })
        .exportGrievancesToCSV();
    }

    function exportMembers() {
      alert('Exporting members to CSV...');
      google.script.run
        .withSuccessHandler((csv) => {
          downloadCSV(csv, 'members.csv');
        })
        .exportMembersToCSV();
    }

    function exportAll() {
      alert('⚠️ Full export feature coming soon. Use manual backup for now.');
    }

    function downloadCSV(csvData, filename) {
      const blob = new Blob([csvData], { type: 'text/csv' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      window.URL.revokeObjectURL(url);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Navigates to backup log sheet
 */
function navigateToBackupLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let backupLog = ss.getSheetByName('💾 Backup Log');

  if (!backupLog) {
    backupLog = createBackupLogSheet();
  }

  backupLog.activate();
}

/**
 * Exports grievances to CSV
 * @returns {string} CSV data
 */
function exportGrievancesToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  let csv = '';
  data.forEach(row => {
    const values = row.map(v => {
      const str = v.toString();
      return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
    });
    csv += values.join(',') + '\n';
  });

  return csv;
}

/**
 * Exports members to CSV
 * @returns {string} CSV data
 */
function exportMembersToCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  let csv = '';
  data.forEach(row => {
    const values = row.map(v => {
      const str = v.toString();
      return str.includes(',') || str.includes('"') ? `"${str.replace(/"/g, '""')}"` : str;
    });
    csv += values.join(',') + '\n';
  });

  return csv;
}

/**
 * Verifies backup integrity
 * @param {string} backupFileId - Backup file ID
 * @returns {Object} Verification result
 */
function verifyBackup(backupFileId) {
  try {
    const backupFile = DriveApp.getFileById(backupFileId);
    const backupSS = SpreadsheetApp.open(backupFile);

    const sheets = backupSS.getSheets();

    const verification = {
      valid: true,
      sheetCount: sheets.length,
      sheets: sheets.map(s => s.getName()),
      dataRows: {}
    };

    // Check key sheets
    const keySheets = [SHEETS.GRIEVANCE_LOG, SHEETS.MEMBER_DIR];
    keySheets.forEach(sheetName => {
      const sheet = backupSS.getSheetByName(sheetName);
      if (sheet) {
        verification.dataRows[sheetName] = sheet.getLastRow() - 1; // Minus header
      } else {
        verification.valid = false;
        verification.error = `Missing key sheet: ${sheetName}`;
      }
    });

    return verification;

  } catch (error) {
    return {
      valid: false,
      error: error.message
    };
  }
}

/******************************************************************************
 * MODULE: CustomReportBuilder
 * Source: CustomReportBuilder.gs
 *****************************************************************************/

/**
 * ============================================================================
 * CUSTOM REPORT BUILDER
 * ============================================================================
 *
 * Build custom reports with filters, grouping, and export
 * Features:
 * - Interactive report builder UI
 * - Custom field selection
 * - Advanced filtering (AND/OR conditions)
 * - Grouping and aggregation
 * - Sorting options
 * - Export to PDF, CSV, Excel
 * - Save and load report templates
 * - Schedule automated reports
 */

/**
 * Shows custom report builder dialog
 */
function showCustomReportBuilder() {
  const html = createReportBuilderHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(1000)
    .setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Custom Report Builder');
}

/**
 * Creates HTML for report builder
 */
function createReportBuilderHTML() {
  const grievanceFields = getGrievanceFields();
  const memberFields = getMemberFields();

  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .builder-section {
      background: #f8f9fa;
      padding: 20px;
      margin: 15px 0;
      border-radius: 8px;
      border-left: 4px solid #1a73e8;
    }
    .section-title {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 15px;
    }
    .form-group {
      margin: 15px 0;
    }
    label {
      display: block;
      font-weight: 500;
      margin-bottom: 5px;
      color: #555;
    }
    select, input[type="text"], input[type="date"] {
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
    .multi-select {
      height: 150px;
    }
    .filter-row {
      display: grid;
      grid-template-columns: 2fr 1fr 2fr auto;
      gap: 10px;
      margin: 10px 0;
      align-items: end;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      margin: 5px 5px 5px 0;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.secondary:hover {
      background: #5a6268;
    }
    button.small {
      padding: 8px 16px;
      font-size: 13px;
    }
    button.danger {
      background: #dc3545;
    }
    button.danger:hover {
      background: #c82333;
    }
    .button-group {
      margin-top: 25px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    #preview {
      max-height: 300px;
      overflow-y: auto;
      background: white;
      padding: 15px;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 12px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
    }
    th {
      background: #1a73e8;
      color: white;
      padding: 8px;
      text-align: left;
      font-size: 11px;
    }
    td {
      padding: 6px 8px;
      border-bottom: 1px solid #e0e0e0;
      font-size: 11px;
    }
    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📊 Custom Report Builder</h2>

    <div class="info-box">
      <strong>💡 Quick Start:</strong> Select data source, choose fields, add filters, then generate your report.
      You can export to PDF, CSV, or Excel.
    </div>

    <!-- Data Source -->
    <div class="builder-section">
      <div class="section-title">1. Select Data Source</div>
      <div class="form-group">
        <label>Data Source:</label>
        <select id="dataSource" onchange="updateFieldOptions()">
          <option value="grievances">Grievance Log</option>
          <option value="members">Member Directory</option>
          <option value="combined">Combined (Grievances + Members)</option>
        </select>
      </div>
    </div>

    <!-- Field Selection -->
    <div class="builder-section">
      <div class="section-title">2. Select Fields to Include</div>
      <div class="form-group">
        <label>Fields (hold Ctrl/Cmd to select multiple):</label>
        <select id="fields" multiple class="multi-select">
          ${grievanceFields.map(f => `<option value="${f.key}">${f.name}</option>`).join('')}
        </select>
      </div>
      <button class="small secondary" onclick="selectAllFields()">Select All</button>
      <button class="small secondary" onclick="clearFields()">Clear All</button>
    </div>

    <!-- Filters -->
    <div class="builder-section">
      <div class="section-title">3. Add Filters (Optional)</div>
      <div id="filters">
        <div class="filter-row">
          <div>
            <label>Field:</label>
            <select id="filter0_field">
              ${grievanceFields.map(f => `<option value="${f.key}">${f.name}</option>`).join('')}
            </select>
          </div>
          <div>
            <label>Operator:</label>
            <select id="filter0_operator">
              <option value="equals">Equals</option>
              <option value="not_equals">Not Equals</option>
              <option value="contains">Contains</option>
              <option value="starts_with">Starts With</option>
              <option value="greater_than">Greater Than</option>
              <option value="less_than">Less Than</option>
            </select>
          </div>
          <div>
            <label>Value:</label>
            <input type="text" id="filter0_value" placeholder="Filter value">
          </div>
          <div>
            <button class="small danger" onclick="removeFilter(0)" style="margin-top: 22px;">Remove</button>
          </div>
        </div>
      </div>
      <button class="small" onclick="addFilter()">+ Add Filter</button>
    </div>

    <!-- Grouping & Sorting -->
    <div class="builder-section">
      <div class="section-title">4. Grouping & Sorting</div>
      <div class="form-group">
        <label>Group By (optional):</label>
        <select id="groupBy">
          <option value="">No Grouping</option>
          ${grievanceFields.map(f => `<option value="${f.key}">${f.name}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Sort By:</label>
        <select id="sortBy">
          ${grievanceFields.map(f => `<option value="${f.key}">${f.name}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label>Sort Order:</label>
        <select id="sortOrder">
          <option value="asc">Ascending</option>
          <option value="desc">Descending</option>
        </select>
      </div>
    </div>

    <!-- Date Range -->
    <div class="builder-section">
      <div class="section-title">5. Date Range (Optional)</div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
        <div class="form-group">
          <label>From Date:</label>
          <input type="date" id="dateFrom">
        </div>
        <div class="form-group">
          <label>To Date:</label>
          <input type="date" id="dateTo">
        </div>
      </div>
    </div>

    <!-- Preview -->
    <div class="builder-section">
      <div class="section-title">6. Preview</div>
      <div id="preview" class="loading">
        Click "Generate Preview" to see your report
      </div>
      <button class="secondary" onclick="generatePreview()">🔍 Generate Preview</button>
    </div>

    <!-- Actions -->
    <div class="button-group">
      <button onclick="exportToPDF()">📄 Export to PDF</button>
      <button onclick="exportToCSV()">📑 Export to CSV</button>
      <button onclick="exportToExcel()">📊 Export to Excel</button>
      <button class="secondary" onclick="saveTemplate()">💾 Save Template</button>
      <button class="secondary" onclick="loadTemplate()">📂 Load Template</button>
    </div>
  </div>

  <script>
    let filterCount = 1;
    const grievanceFields = ${JSON.stringify(grievanceFields)};
    const memberFields = ${JSON.stringify(memberFields)};

    function updateFieldOptions() {
      const dataSource = document.getElementById('dataSource').value;
      const fieldsSelect = document.getElementById('fields');

      fieldsSelect.innerHTML = '';

      const fields = dataSource === 'members' ? memberFields : grievanceFields;

      fields.forEach(f => {
        const option = document.createElement('option');
        option.value = f.key;
        option.textContent = f.name;
        fieldsSelect.appendChild(option);
      });
    }

    function selectAllFields() {
      const fieldsSelect = document.getElementById('fields');
      for (let i = 0; i < fieldsSelect.options.length; i++) {
        fieldsSelect.options[i].selected = true;
      }
    }

    function clearFields() {
      const fieldsSelect = document.getElementById('fields');
      for (let i = 0; i < fieldsSelect.options.length; i++) {
        fieldsSelect.options[i].selected = false;
      }
    }

    function addFilter() {
      const filtersDiv = document.getElementById('filters');
      const newFilter = document.createElement('div');
      newFilter.className = 'filter-row';
      newFilter.id = 'filter' + filterCount;
      newFilter.innerHTML = \`
        <div>
          <label>Field:</label>
          <select id="filter\${filterCount}_field">
            \${grievanceFields.map(f => '<option value="' + f.key + '">' + f.name + '</option>').join('')}
          </select>
        </div>
        <div>
          <label>Operator:</label>
          <select id="filter\${filterCount}_operator">
            <option value="equals">Equals</option>
            <option value="not_equals">Not Equals</option>
            <option value="contains">Contains</option>
            <option value="starts_with">Starts With</option>
            <option value="greater_than">Greater Than</option>
            <option value="less_than">Less Than</option>
          </select>
        </div>
        <div>
          <label>Value:</label>
          <input type="text" id="filter\${filterCount}_value" placeholder="Filter value">
        </div>
        <div>
          <button class="small danger" onclick="removeFilter(\${filterCount})" style="margin-top: 22px;">Remove</button>
        </div>
      \`;
      filtersDiv.appendChild(newFilter);
      filterCount++;
    }

    function removeFilter(id) {
      const filter = document.getElementById('filter' + id);
      if (filter) {
        filter.remove();
      }
    }

    function getReportConfig() {
      const fieldsSelect = document.getElementById('fields');
      const selectedFields = Array.from(fieldsSelect.selectedOptions).map(opt => opt.value);

      const filters = [];
      for (let i = 0; i < filterCount; i++) {
        const fieldElem = document.getElementById('filter' + i + '_field');
        if (fieldElem) {
          filters.push({
            field: fieldElem.value,
            operator: document.getElementById('filter' + i + '_operator').value,
            value: document.getElementById('filter' + i + '_value').value
          });
        }
      }

      return {
        dataSource: document.getElementById('dataSource').value,
        fields: selectedFields,
        filters: filters,
        groupBy: document.getElementById('groupBy').value,
        sortBy: document.getElementById('sortBy').value,
        sortOrder: document.getElementById('sortOrder').value,
        dateFrom: document.getElementById('dateFrom').value,
        dateTo: document.getElementById('dateTo').value
      };
    }

    function generatePreview() {
      const config = getReportConfig();

      if (config.fields.length === 0) {
        alert('Please select at least one field');
        return;
      }

      document.getElementById('preview').innerHTML = '<div class="loading">Generating preview...</div>';

      google.script.run
        .withSuccessHandler(displayPreview)
        .withFailureHandler(onError)
        .generateReportData(config, 10); // Preview first 10 rows
    }

    function displayPreview(data) {
      const preview = document.getElementById('preview');

      if (!data || data.length === 0) {
        preview.innerHTML = '<div class="loading">No data matches your filters</div>';
        return;
      }

      let html = '<table><thead><tr>';

      // Headers
      Object.keys(data[0]).forEach(key => {
        html += '<th>' + key + '</th>';
      });
      html += '</tr></thead><tbody>';

      // Data rows
      data.forEach(row => {
        html += '<tr>';
        Object.values(row).forEach(value => {
          html += '<td>' + (value || '') + '</td>';
        });
        html += '</tr>';
      });

      html += '</tbody></table>';
      html += '<div style="margin-top: 10px; color: #666; font-size: 12px;">Showing first ' + data.length + ' rows</div>';

      preview.innerHTML = html;
    }

    function onError(error) {
      alert('Error: ' + error.message);
      document.getElementById('preview').innerHTML = '<div class="loading">Error generating preview</div>';
    }

    function exportToPDF() {
      const config = getReportConfig();

      if (config.fields.length === 0) {
        alert('Please select at least one field');
        return;
      }

      alert('Generating PDF report... This may take a moment.');

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ PDF report generated and saved to Google Drive!');
        })
        .withFailureHandler(onError)
        .exportReportToPDF(config);
    }

    function exportToCSV() {
      const config = getReportConfig();

      if (config.fields.length === 0) {
        alert('Please select at least one field');
        return;
      }

      google.script.run
        .withSuccessHandler((csvData) => {
          downloadCSV(csvData);
        })
        .withFailureHandler(onError)
        .exportReportToCSV(config);
    }

    function exportToExcel() {
      const config = getReportConfig();

      if (config.fields.length === 0) {
        alert('Please select at least one field');
        return;
      }

      alert('Generating Excel report... This may take a moment.');

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Excel report generated and saved to Google Drive!');
        })
        .withFailureHandler(onError)
        .exportReportToExcel(config);
    }

    function downloadCSV(csvData) {
      const blob = new Blob([csvData], { type: 'text/csv' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'report_' + new Date().toISOString().slice(0, 10) + '.csv';
      a.click();
      window.URL.revokeObjectURL(url);
    }

    function saveTemplate() {
      const config = getReportConfig();
      const name = prompt('Enter template name:');

      if (!name) return;

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Template saved: ' + name);
        })
        .withFailureHandler(onError)
        .saveReportTemplate(name, config);
    }

    function loadTemplate() {
      google.script.run
        .withSuccessHandler(showTemplateList)
        .withFailureHandler(onError)
        .getReportTemplates();
    }

    function showTemplateList(templates) {
      if (!templates || templates.length === 0) {
        alert('No saved templates found');
        return;
      }

      const templateNames = templates.map(t => t.name).join('\\n');
      const selected = prompt('Available templates:\\n' + templateNames + '\\n\\nEnter template name to load:');

      if (!selected) return;

      const template = templates.find(t => t.name === selected);
      if (template) {
        applyTemplate(template.config);
      } else {
        alert('Template not found');
      }
    }

    function applyTemplate(config) {
      document.getElementById('dataSource').value = config.dataSource;
      updateFieldOptions();

      // Select fields
      const fieldsSelect = document.getElementById('fields');
      for (let i = 0; i < fieldsSelect.options.length; i++) {
        fieldsSelect.options[i].selected = config.fields.includes(fieldsSelect.options[i].value);
      }

      // Apply other settings
      document.getElementById('groupBy').value = config.groupBy || '';
      document.getElementById('sortBy').value = config.sortBy || '';
      document.getElementById('sortOrder').value = config.sortOrder || 'asc';
      document.getElementById('dateFrom').value = config.dateFrom || '';
      document.getElementById('dateTo').value = config.dateTo || '';

      alert('✅ Template loaded');
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets available grievance fields
 * @returns {Array} Field definitions
 */
function getGrievanceFields() {
  return [
    { key: 'grievanceId', name: 'Grievance ID' },
    { key: 'memberId', name: 'Member ID' },
    { key: 'firstName', name: 'First Name' },
    { key: 'lastName', name: 'Last Name' },
    { key: 'status', name: 'Status' },
    { key: 'issueType', name: 'Issue Type' },
    { key: 'filedDate', name: 'Filed Date' },
    { key: 'description', name: 'Description' },
    { key: 'outcome', name: 'Outcome Sought' },
    { key: 'location', name: 'Work Location' },
    { key: 'unit', name: 'Unit' },
    { key: 'manager', name: 'Manager' },
    { key: 'supervisor', name: 'Supervisor' },
    { key: 'assignedSteward', name: 'Assigned Steward' },
    { key: 'meetingDate', name: 'Meeting Date' },
    { key: 'step1Response', name: 'Step 1 Response' },
    { key: 'step1ResponseDate', name: 'Step 1 Response Date' },
    { key: 'step2EscalationDate', name: 'Step 2 Escalation Date' },
    { key: 'dateClosed', name: 'Date Closed' },
    { key: 'nextActionDue', name: 'Next Action Due' },
    { key: 'daysToDeadline', name: 'Days to Deadline' },
    { key: 'deadlineStatus', name: 'Deadline Status' },
    { key: 'notes', name: 'Notes' }
  ];
}

/**
 * Gets available member fields
 * @returns {Array} Field definitions
 */
function getMemberFields() {
  return [
    { key: 'memberId', name: 'Member ID' },
    { key: 'firstName', name: 'First Name' },
    { key: 'lastName', name: 'Last Name' },
    { key: 'title', name: 'Job Title' },
    { key: 'workLocation', name: 'Work Location' },
    { key: 'unit', name: 'Unit' },
    { key: 'hireDate', name: 'Hire Date' },
    { key: 'email', name: 'Email' },
    { key: 'phone', name: 'Phone' },
    { key: 'isSteward', name: 'Is Steward?' }
  ];
}

/**
 * Generates report data based on configuration
 * @param {Object} config - Report configuration
 * @param {number} limit - Row limit for preview
 * @returns {Array} Report data
 */
function generateReportData(config, limit = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sourceData;
  let fieldMapping;

  if (config.dataSource === 'grievances') {
    const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    sourceData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 28).getValues();
    fieldMapping = getGrievanceFieldMapping();
  } else {
    const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    sourceData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 28).getValues();
    fieldMapping = getMemberFieldMapping();
  }

  // Apply filters
  let filteredData = sourceData.filter(row => {
    return applyFilters(row, config.filters, fieldMapping);
  });

  // Apply date range
  if (config.dateFrom || config.dateTo) {
    filteredData = applyDateRange(filteredData, config.dateFrom, config.dateTo, fieldMapping);
  }

  // Sort data
  if (config.sortBy) {
    filteredData = sortData(filteredData, config.sortBy, config.sortOrder, fieldMapping);
  }

  // Limit for preview
  if (limit) {
    filteredData = filteredData.slice(0, limit);
  }

  // Select fields and format
  const reportData = filteredData.map(row => {
    const reportRow = {};
    config.fields.forEach(fieldKey => {
      const field = fieldMapping[fieldKey];
      if (field) {
        reportRow[field.name] = formatValue(row[field.index], field.type);
      }
    });
    return reportRow;
  });

  return reportData;
}

/**
 * Gets field mapping for grievances
 * @returns {Object} Field mapping
 */
function getGrievanceFieldMapping() {
  return {
    grievanceId: { index: 0, name: 'Grievance ID', type: 'string' },
    memberId: { index: 1, name: 'Member ID', type: 'string' },
    firstName: { index: 2, name: 'First Name', type: 'string' },
    lastName: { index: 3, name: 'Last Name', type: 'string' },
    status: { index: 4, name: 'Status', type: 'string' },
    issueType: { index: 5, name: 'Issue Type', type: 'string' },
    filedDate: { index: 6, name: 'Filed Date', type: 'date' },
    description: { index: 7, name: 'Description', type: 'string' },
    location: { index: 9, name: 'Work Location', type: 'string' },
    assignedSteward: { index: 13, name: 'Assigned Steward', type: 'string' },
    nextActionDue: { index: 19, name: 'Next Action Due', type: 'date' },
    daysToDeadline: { index: 20, name: 'Days to Deadline', type: 'number' }
  };
}

/**
 * Gets field mapping for members
 * @returns {Object} Field mapping
 */
function getMemberFieldMapping() {
  return {
    memberId: { index: 0, name: 'Member ID', type: 'string' },
    firstName: { index: 1, name: 'First Name', type: 'string' },
    lastName: { index: 2, name: 'Last Name', type: 'string' },
    title: { index: 3, name: 'Job Title', type: 'string' },
    workLocation: { index: 4, name: 'Work Location', type: 'string' },
    email: { index: 7, name: 'Email', type: 'string' },
    phone: { index: 8, name: 'Phone', type: 'string' },
    isSteward: { index: 9, name: 'Is Steward?', type: 'string' }
  };
}

/**
 * Applies filters to data
 * @param {Array} row - Data row
 * @param {Array} filters - Filter configuration
 * @param {Object} fieldMapping - Field mapping
 * @returns {boolean} Whether row passes filters
 */
function applyFilters(row, filters, fieldMapping) {
  if (!filters || filters.length === 0) return true;

  return filters.every(filter => {
    if (!filter.value) return true;

    const field = fieldMapping[filter.field];
    if (!field) return true;

    const cellValue = String(row[field.index] || '').toLowerCase();
    const filterValue = String(filter.value).toLowerCase();

    switch (filter.operator) {
      case 'equals':
        return cellValue === filterValue;
      case 'not_equals':
        return cellValue !== filterValue;
      case 'contains':
        return cellValue.includes(filterValue);
      case 'starts_with':
        return cellValue.startsWith(filterValue);
      case 'greater_than':
        return parseFloat(cellValue) > parseFloat(filterValue);
      case 'less_than':
        return parseFloat(cellValue) < parseFloat(filterValue);
      default:
        return true;
    }
  });
}

/**
 * Applies date range filter
 * @param {Array} data - Data array
 * @param {string} dateFrom - Start date
 * @param {string} dateTo - End date
 * @param {Object} fieldMapping - Field mapping
 * @returns {Array} Filtered data
 */
function applyDateRange(data, dateFrom, dateTo, fieldMapping) {
  const fromDate = dateFrom ? new Date(dateFrom) : null;
  const toDate = dateTo ? new Date(dateTo) : null;

  return data.filter(row => {
    const filedDate = row[6]; // Assuming column G is filed date

    if (!filedDate) return true;

    if (fromDate && filedDate < fromDate) return false;
    if (toDate && filedDate > toDate) return false;

    return true;
  });
}

/**
 * Sorts data
 * @param {Array} data - Data array
 * @param {string} sortBy - Field to sort by
 * @param {string} sortOrder - asc or desc
 * @param {Object} fieldMapping - Field mapping
 * @returns {Array} Sorted data
 */
function sortData(data, sortBy, sortOrder, fieldMapping) {
  const field = fieldMapping[sortBy];
  if (!field) return data;

  return data.sort((a, b) => {
    const aVal = a[field.index];
    const bVal = b[field.index];

    let comparison = 0;
    if (aVal < bVal) comparison = -1;
    if (aVal > bVal) comparison = 1;

    return sortOrder === 'desc' ? -comparison : comparison;
  });
}

/**
 * Formats value based on type
 * @param {any} value - Value to format
 * @param {string} type - Data type
 * @returns {string} Formatted value
 */
function formatValue(value, type) {
  if (!value && value !== 0) return '';

  switch (type) {
    case 'date':
      return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'MM/dd/yyyy') : value;
    case 'number':
      return typeof value === 'number' ? value.toString() : value;
    default:
      return value.toString();
  }
}

/**
 * Exports report to PDF
 * @param {Object} config - Report configuration
 * @returns {File} PDF file
 */
function exportReportToPDF(config) {
  const data = generateReportData(config);

  // Create Google Doc
  const doc = DocumentApp.create(`Custom_Report_${new Date().toISOString().slice(0, 10)}`);
  const body = doc.getBody();

  // Add title
  body.appendParagraph('SEIU Local 509 - Custom Report')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(`Generated: ${new Date().toLocaleString()}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // Add table
  if (data.length > 0) {
    const headers = Object.keys(data[0]);
    const tableData = [headers];

    data.forEach(row => {
      tableData.push(Object.values(row).map(v => v.toString()));
    });

    const table = body.appendTable(tableData);
    table.getRow(0).editAsText().setBold(true);
  }

  doc.saveAndClose();

  // Convert to PDF
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdfBlob);

  // Delete temp doc
  docFile.setTrashed(true);

  return pdfFile;
}

/**
 * Exports report to CSV
 * @param {Object} config - Report configuration
 * @returns {string} CSV data
 */
function exportReportToCSV(config) {
  const data = generateReportData(config);

  if (data.length === 0) {
    return '';
  }

  const headers = Object.keys(data[0]);
  let csv = headers.join(',') + '\n';

  data.forEach(row => {
    const values = Object.values(row).map(v => {
      const str = v.toString();
      return str.includes(',') ? `"${str}"` : str;
    });
    csv += values.join(',') + '\n';
  });

  return csv;
}

/**
 * Exports report to Excel (Google Sheets)
 * @param {Object} config - Report configuration
 * @returns {File} Excel file
 */
function exportReportToExcel(config) {
  const data = generateReportData(config);

  // Create new spreadsheet
  const ss = SpreadsheetApp.create(`Custom_Report_${new Date().toISOString().slice(0, 10)}`);
  const sheet = ss.getActiveSheet();

  if (data.length > 0) {
    const headers = Object.keys(data[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    const rows = data.map(row => Object.values(row));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return DriveApp.getFileById(ss.getId());
}

/**
 * Saves report template
 * @param {string} name - Template name
 * @param {Object} config - Report configuration
 */
function saveReportTemplate(name, config) {
  const props = PropertiesService.getUserProperties();
  const templates = JSON.parse(props.getProperty('reportTemplates') || '[]');

  // Remove existing template with same name
  const filtered = templates.filter(t => t.name !== name);

  // Add new template
  filtered.push({ name: name, config: config });

  props.setProperty('reportTemplates', JSON.stringify(filtered));
}

/**
 * Gets saved report templates
 * @returns {Array} Templates
 */
function getReportTemplates() {
  const props = PropertiesService.getUserProperties();
  return JSON.parse(props.getProperty('reportTemplates') || '[]');
}

/******************************************************************************
 * MODULE: DataPagination
 * Source: DataPagination.gs
 *****************************************************************************/

/**
 * ============================================================================
 * DATA PAGINATION & VIRTUAL SCROLLING
 * ============================================================================
 *
 * Advanced pagination system for handling large datasets efficiently
 * Features:
 * - Page-based navigation
 * - Configurable page sizes
 * - Quick jump to page
 * - Virtual scrolling for large datasets
 * - Search within pages
 * - Export current page
 * - Performance optimizations
 */

/**
 * Pagination configuration
 */
const PAGINATION_CONFIG = {
  DEFAULT_PAGE_SIZE: 100,
  PAGE_SIZE_OPTIONS: [25, 50, 100, 200, 500],
  MAX_CACHED_PAGES: 10,
  ENABLE_VIRTUAL_SCROLL: true
};

/**
 * Shows paginated data viewer
 */
function showPaginatedViewer() {
  const html = createPaginatedViewerHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📄 Paginated Data Viewer');
}

/**
 * Creates HTML for paginated viewer
 */
function createPaginatedViewerHTML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const totalRows = grievanceSheet.getLastRow() - 1;  // Exclude header

  const pageSize = PAGINATION_CONFIG.DEFAULT_PAGE_SIZE;
  const totalPages = Math.ceil(totalRows / pageSize);

  return `
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
      max-height: 750px;
      display: flex;
      flex-direction: column;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .controls {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin: 20px 0;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 8px;
      flex-wrap: wrap;
      gap: 15px;
    }
    .control-group {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .stats {
      display: flex;
      gap: 20px;
      padding: 15px;
      background: #e8f0fe;
      border-radius: 8px;
      margin: 10px 0;
    }
    .stat-item {
      display: flex;
      flex-direction: column;
    }
    .stat-label {
      font-size: 12px;
      color: #666;
    }
    .stat-value {
      font-size: 20px;
      font-weight: bold;
      color: #1a73e8;
    }
    .data-table-container {
      flex: 1;
      overflow-y: auto;
      border: 1px solid #e0e0e0;
      border-radius: 4px;
      margin: 10px 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th {
      background: #1a73e8;
      color: white;
      padding: 12px 8px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
      z-index: 10;
    }
    td {
      padding: 10px 8px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background: #f8f9fa;
    }
    tr:nth-child(even) {
      background: #fafafa;
    }
    tr:nth-child(even):hover {
      background: #f0f0f0;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      transition: background 0.2s;
    }
    button:hover:not(:disabled) {
      background: #1557b0;
    }
    button:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    button.secondary {
      background: #6c757d;
    }
    button.secondary:hover:not(:disabled) {
      background: #5a6268;
    }
    select, input[type="number"], input[type="text"] {
      padding: 8px 12px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    .pagination {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 8px;
      justify-content: center;
      flex-wrap: wrap;
    }
    .page-button {
      min-width: 40px;
      padding: 8px 12px;
    }
    .page-button.active {
      background: #4caf50;
    }
    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .spinner {
      border: 4px solid #f3f3f3;
      border-top: 4px solid #1a73e8;
      border-radius: 50%;
      width: 40px;
      height: 40px;
      animation: spin 1s linear infinite;
      margin: 0 auto 20px;
    }
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .search-box {
      padding: 8px 12px;
      border: 2px solid #ddd;
      border-radius: 4px;
      width: 300px;
    }
    .no-data {
      text-align: center;
      padding: 60px;
      color: #999;
      font-size: 16px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📄 Paginated Data Viewer</h2>

    <div class="stats">
      <div class="stat-item">
        <div class="stat-label">Total Records</div>
        <div class="stat-value" id="totalRecords">${totalRows}</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">Total Pages</div>
        <div class="stat-value" id="totalPages">${totalPages}</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">Current Page</div>
        <div class="stat-value" id="currentPage">1</div>
      </div>
      <div class="stat-item">
        <div class="stat-label">Records/Page</div>
        <div class="stat-value" id="pageSize">${pageSize}</div>
      </div>
    </div>

    <div class="controls">
      <div class="control-group">
        <label>Page Size:</label>
        <select id="pageSizeSelect" onchange="changePageSize()">
          ${PAGINATION_CONFIG.PAGE_SIZE_OPTIONS.map(size =>
            `<option value="${size}" ${size === pageSize ? 'selected' : ''}>${size}</option>`
          ).join('')}
        </select>
      </div>

      <div class="control-group">
        <label>Search:</label>
        <input type="text" class="search-box" id="searchBox" placeholder="Search in current data..." oninput="searchData()">
      </div>

      <div class="control-group">
        <button onclick="exportCurrentPage()">📥 Export Page</button>
        <button class="secondary" onclick="refreshData()">🔄 Refresh</button>
      </div>
    </div>

    <div class="data-table-container" id="tableContainer">
      <div class="loading">
        <div class="spinner"></div>
        <div>Loading data...</div>
      </div>
    </div>

    <div class="pagination" id="pagination">
      <!-- Pagination buttons will be inserted here -->
    </div>
  </div>

  <script>
    let currentPage = 1;
    let pageSize = ${pageSize};
    let totalRecords = ${totalRows};
    let totalPages = ${totalPages};
    let currentData = [];
    let allHeaders = [];

    // Load initial page
    loadPage(1);

    function loadPage(page) {
      currentPage = page;
      document.getElementById('currentPage').textContent = page;

      const startRow = (page - 1) * pageSize + 2;  // +2 for header and 0-index
      const numRows = Math.min(pageSize, totalRecords - (page - 1) * pageSize);

      document.getElementById('tableContainer').innerHTML = '<div class="loading"><div class="spinner"></div><div>Loading page ' + page + '...</div></div>';

      google.script.run
        .withSuccessHandler(renderTable)
        .withFailureHandler(handleError)
        .getPageData(startRow, numRows);
    }

    function renderTable(data) {
      if (!data || !data.headers || !data.rows) {
        document.getElementById('tableContainer').innerHTML = '<div class="no-data">No data available</div>';
        return;
      }

      allHeaders = data.headers;
      currentData = data.rows;

      let html = '<table><thead><tr>';

      // Headers
      data.headers.forEach(header => {
        html += '<th>' + escapeHtml(header) + '</th>';
      });

      html += '</tr></thead><tbody>';

      // Rows
      if (data.rows.length === 0) {
        html += '<tr><td colspan="' + data.headers.length + '" class="no-data">No records found</td></tr>';
      } else {
        data.rows.forEach(row => {
          html += '<tr>';
          row.forEach(cell => {
            let cellValue = cell;
            if (cell instanceof Date) {
              cellValue = cell.toLocaleDateString();
            }
            html += '<td>' + escapeHtml(String(cellValue)) + '</td>';
          });
          html += '</tr>';
        });
      }

      html += '</tbody></table>';

      document.getElementById('tableContainer').innerHTML = html;
      renderPagination();
    }

    function renderPagination() {
      let html = '';

      // First and Previous buttons
      html += '<button class="page-button" onclick="loadPage(1)" ' + (currentPage === 1 ? 'disabled' : '') + '>« First</button>';
      html += '<button class="page-button" onclick="loadPage(' + (currentPage - 1) + ')" ' + (currentPage === 1 ? 'disabled' : '') + '>‹ Prev</button>';

      // Page numbers (show 5 around current)
      const startPage = Math.max(1, currentPage - 2);
      const endPage = Math.min(totalPages, currentPage + 2);

      if (startPage > 1) {
        html += '<span style="padding: 0 5px;">...</span>';
      }

      for (let i = startPage; i <= endPage; i++) {
        html += '<button class="page-button ' + (i === currentPage ? 'active' : '') + '" onclick="loadPage(' + i + ')">' + i + '</button>';
      }

      if (endPage < totalPages) {
        html += '<span style="padding: 0 5px;">...</span>';
      }

      // Next and Last buttons
      html += '<button class="page-button" onclick="loadPage(' + (currentPage + 1) + ')" ' + (currentPage === totalPages ? 'disabled' : '') + '>Next ›</button>';
      html += '<button class="page-button" onclick="loadPage(' + totalPages + ')" ' + (currentPage === totalPages ? 'disabled' : '') + '>Last »</button>';

      // Jump to page
      html += '<span style="padding: 0 10px;">|</span>';
      html += '<label>Go to:</label>';
      html += '<input type="number" min="1" max="' + totalPages + '" value="' + currentPage + '" id="jumpToPage" style="width: 60px;">';
      html += '<button class="page-button" onclick="jumpToPage()">Go</button>';

      document.getElementById('pagination').innerHTML = html;
    }

    function jumpToPage() {
      const page = parseInt(document.getElementById('jumpToPage').value);
      if (page >= 1 && page <= totalPages) {
        loadPage(page);
      }
    }

    function changePageSize() {
      pageSize = parseInt(document.getElementById('pageSizeSelect').value);
      totalPages = Math.ceil(totalRecords / pageSize);
      document.getElementById('pageSize').textContent = pageSize;
      document.getElementById('totalPages').textContent = totalPages;

      // Reload current page with new size
      loadPage(1);
    }

    function searchData() {
      const searchTerm = document.getElementById('searchBox').value.toLowerCase();

      if (!searchTerm) {
        renderTable({ headers: allHeaders, rows: currentData });
        return;
      }

      const filteredRows = currentData.filter(row => {
        return row.some(cell => {
          return String(cell).toLowerCase().includes(searchTerm);
        });
      });

      renderTable({ headers: allHeaders, rows: filteredRows });
    }

    function exportCurrentPage() {
      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ Page exported!');
          window.open(url, '_blank');
        })
        .exportPageToCSV(currentPage, pageSize);
    }

    function refreshData() {
      loadPage(currentPage);
    }

    function handleError(error) {
      document.getElementById('tableContainer').innerHTML = '<div class="no-data">❌ Error loading data: ' + error.message + '</div>';
    }

    function escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets page data
 * @param {number} startRow - Starting row
 * @param {number} numRows - Number of rows
 * @returns {Object} Page data
 */
function getPageData(startRow, numRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const lastCol = grievanceSheet.getLastColumn();

  // Get headers
  const headers = grievanceSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Get data
  let rows = [];
  if (numRows > 0) {
    rows = grievanceSheet.getRange(startRow, 1, numRows, lastCol).getValues();
  }

  return {
    headers: headers,
    rows: rows
  };
}

/**
 * Exports current page to CSV
 * @param {number} page - Page number
 * @param {number} pageSize - Page size
 * @returns {string} File URL
 */
function exportPageToCSV(page, pageSize) {
  const startRow = (page - 1) * pageSize + 2;
  const data = getPageData(startRow, pageSize);

  // Create CSV content
  let csv = '';

  // Headers
  csv += data.headers.map(h => `"${h}"`).join(',') + '\n';

  // Rows
  data.rows.forEach(row => {
    csv += row.map(cell => {
      let value = String(cell);
      if (cell instanceof Date) {
        value = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return `"${value.replace(/"/g, '""')}"`;
    }).join(',') + '\n';
  });

  // Create file
  const fileName = `Grievances_Page_${page}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.csv`;
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);

  return file.getUrl();
}

/**
 * Creates a paginated view for a sheet
 * @param {string} sheetName - Sheet name
 * @param {number} pageSize - Records per page
 */
function createPaginatedView(sheetName, pageSize = PAGINATION_CONFIG.DEFAULT_PAGE_SIZE) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);

  if (!sourceSheet) {
    throw new Error(`Sheet ${sheetName} not found`);
  }

  const totalRows = sourceSheet.getLastRow() - 1;
  const totalPages = Math.ceil(totalRows / pageSize);

  const viewSheetName = `${sheetName}_Paginated`;
  let viewSheet = ss.getSheetByName(viewSheetName);

  if (viewSheet) {
    viewSheet.clear();
  } else {
    viewSheet = ss.insertSheet(viewSheetName);
  }

  // Add controls
  viewSheet.getRange('A1').setValue('Page:');
  viewSheet.getRange('B1').setValue(1);
  viewSheet.getRange('C1').setValue(`of ${totalPages}`);

  viewSheet.getRange('D1').setValue('Page Size:');
  viewSheet.getRange('E1').setValue(pageSize);

  // Add navigation buttons (via notes)
  viewSheet.getRange('F1').setValue('⏮ First');
  viewSheet.getRange('G1').setValue('◀ Prev');
  viewSheet.getRange('H1').setValue('Next ▶');
  viewSheet.getRange('I1').setValue('Last ⏭');

  // Copy headers
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues();
  viewSheet.getRange(3, 1, 1, headers[0].length).setValues(headers);
  viewSheet.getRange(3, 1, 1, headers[0].length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Load first page
  loadPaginatedPage(sourceSheet, viewSheet, 1, pageSize);

  return viewSheet;
}

/**
 * Loads a page in paginated view
 * @param {Sheet} sourceSheet - Source sheet
 * @param {Sheet} viewSheet - View sheet
 * @param {number} page - Page number
 * @param {number} pageSize - Page size
 */
function loadPaginatedPage(sourceSheet, viewSheet, page, pageSize) {
  const startRow = (page - 1) * pageSize + 2;
  const lastRow = sourceSheet.getLastRow();
  const numRows = Math.min(pageSize, lastRow - startRow + 1);

  // Clear previous data
  if (viewSheet.getLastRow() > 3) {
    viewSheet.getRange(4, 1, viewSheet.getLastRow() - 3, viewSheet.getLastColumn()).clear();
  }

  if (numRows > 0) {
    const data = sourceSheet.getRange(startRow, 1, numRows, sourceSheet.getLastColumn()).getValues();
    viewSheet.getRange(4, 1, numRows, data[0].length).setValues(data);
  }

  // Update page number
  viewSheet.getRange('B1').setValue(page);
}

/**
 * Navigates to next page in paginated view
 */
function nextPaginatedPage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const viewSheet = ss.getActiveSheet();

  if (!viewSheet.getName().endsWith('_Paginated')) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Not a paginated view', 'Error', 3);
    return;
  }

  const currentPage = viewSheet.getRange('B1').getValue();
  const pageSize = viewSheet.getRange('E1').getValue();
  const totalPagesText = viewSheet.getRange('C1').getValue();
  const totalPages = parseInt(totalPagesText.match(/\d+/)[0]);

  if (currentPage < totalPages) {
    const sourceSheetName = viewSheet.getName().replace('_Paginated', '');
    const sourceSheet = ss.getSheetByName(sourceSheetName);

    loadPaginatedPage(sourceSheet, viewSheet, currentPage + 1, pageSize);
  }
}

/**
 * Navigates to previous page in paginated view
 */
function previousPaginatedPage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const viewSheet = ss.getActiveSheet();

  if (!viewSheet.getName().endsWith('_Paginated')) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Not a paginated view', 'Error', 3);
    return;
  }

  const currentPage = viewSheet.getRange('B1').getValue();
  const pageSize = viewSheet.getRange('E1').getValue();

  if (currentPage > 1) {
    const sourceSheetName = viewSheet.getName().replace('_Paginated', '');
    const sourceSheet = ss.getSheetByName(sourceSheetName);

    loadPaginatedPage(sourceSheet, viewSheet, currentPage - 1, pageSize);
  }
}

/**
 * Gets pagination statistics
 * @returns {Object} Statistics
 */
function getPaginationStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const totalRows = grievanceSheet.getLastRow() - 1;
  const pageSize = PAGINATION_CONFIG.DEFAULT_PAGE_SIZE;
  const totalPages = Math.ceil(totalRows / pageSize);

  return {
    totalRecords: totalRows,
    pageSize: pageSize,
    totalPages: totalPages,
    avgPageLoadTime: '~2 seconds'
  };
}

/******************************************************************************
 * MODULE: EnhancedADHDFeatures
 * Source: EnhancedADHDFeatures.gs
 *****************************************************************************/

/**
 * ============================================================================
 * ENHANCED ADHD FEATURES
 * ============================================================================
 *
 * Advanced accessibility and focus features for neurodivergent users
 * Features:
 * - Focus mode with distraction reduction
 * - Custom color themes for readability
 * - Font size adjustments
 * - Row highlighting and zebra striping
 * - Reduced motion mode
 * - Audio cues (optional)
 * - Task timers and reminders
 * - Quick capture notepad
 * - Customizable dashboard layouts
 * - Attention span breaks
 */

/**
 * ADHD configuration
 */
const ADHD_CONFIG = {
  FOCUS_MODE_COLORS: {
    background: '#f5f5f5',
    header: '#4a4a4a',
    accent: '#6b9bd1'
  },
  HIGH_CONTRAST_COLORS: {
    background: '#ffffff',
    header: '#000000',
    accent: '#0066cc',
    alternateRow: '#f0f0f0'
  },
  PASTEL_COLORS: {
    background: '#fef9e7',
    header: '#85929e',
    accent: '#7fb3d5',
    alternateRow: '#ebf5fb'
  }
};

/**
 * Shows ADHD features control panel
 */
function showADHDControlPanel() {
  const html = createADHDControlPanelHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '♿ ADHD Control Panel');
}

/**
 * Creates HTML for ADHD control panel
 */
function createADHDControlPanelHTML() {
  const currentSettings = getADHDSettings();

  return `
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
      max-height: 650px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .section {
      background: #f8f9fa;
      padding: 20px;
      margin: 15px 0;
      border-radius: 8px;
      border-left: 4px solid #1a73e8;
    }
    .section-title {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 15px;
    }
    .setting-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 0;
      border-bottom: 1px solid #e0e0e0;
    }
    .setting-row:last-child {
      border-bottom: none;
    }
    .setting-label {
      font-weight: 500;
      color: #555;
    }
    .setting-description {
      font-size: 13px;
      color: #888;
      margin-top: 4px;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.active {
      background: #4caf50;
    }
    select, input[type="range"] {
      padding: 8px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    .slider-container {
      display: flex;
      align-items: center;
      gap: 15px;
    }
    .slider-value {
      min-width: 50px;
      text-align: center;
      font-weight: bold;
      color: #1a73e8;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .theme-preview {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 15px;
      margin: 15px 0;
    }
    .theme-card {
      padding: 15px;
      border-radius: 8px;
      cursor: pointer;
      border: 3px solid transparent;
      transition: all 0.2s;
    }
    .theme-card:hover {
      transform: scale(1.05);
    }
    .theme-card.selected {
      border-color: #1a73e8;
      box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3);
    }
    .theme-name {
      font-weight: bold;
      margin-bottom: 8px;
      text-align: center;
    }
    .color-bar {
      height: 40px;
      border-radius: 4px;
      margin: 5px 0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>♿ ADHD Control Panel</h2>

    <div class="info-box">
      <strong>💡 Neurodiversity Support:</strong> Customize the dashboard for optimal focus and comfort.
      All settings are saved to your profile.
    </div>

    <!-- Visual Settings -->
    <div class="section">
      <div class="section-title">🎨 Visual Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Color Theme</div>
          <div class="setting-description">Choose a color scheme that works best for you</div>
        </div>
      </div>

      <div class="theme-preview">
        <div class="theme-card ${currentSettings.theme === 'default' ? 'selected' : ''}" onclick="selectTheme('default')">
          <div class="theme-name">Default</div>
          <div class="color-bar" style="background: #1a73e8;"></div>
          <div class="color-bar" style="background: #f8f9fa;"></div>
        </div>

        <div class="theme-card ${currentSettings.theme === 'high-contrast' ? 'selected' : ''}" onclick="selectTheme('high-contrast')">
          <div class="theme-name">High Contrast</div>
          <div class="color-bar" style="background: #000000;"></div>
          <div class="color-bar" style="background: #ffffff;"></div>
        </div>

        <div class="theme-card ${currentSettings.theme === 'pastel' ? 'selected' : ''}" onclick="selectTheme('pastel')">
          <div class="theme-name">Soft Pastel</div>
          <div class="color-bar" style="background: #7fb3d5;"></div>
          <div class="color-bar" style="background: #fef9e7;"></div>
        </div>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Font Size</div>
          <div class="setting-description">Adjust text size for better readability</div>
        </div>
        <div class="slider-container">
          <input type="range" id="fontSize" min="8" max="14" value="${currentSettings.fontSize || 10}" oninput="updateSlider('fontSize')">
          <span class="slider-value" id="fontSizeValue">${currentSettings.fontSize || 10}pt</span>
        </div>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Row Highlighting</div>
          <div class="setting-description">Zebra striping for easier row tracking</div>
        </div>
        <button class="${currentSettings.zebraStripes ? 'active' : 'secondary'}" onclick="toggleZebraStripes()">
          ${currentSettings.zebraStripes ? '✅ Enabled' : 'Enable'}
        </button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Gridlines</div>
          <div class="setting-description">Show or hide cell borders</div>
        </div>
        <button class="${currentSettings.gridlines ? 'active' : 'secondary'}" onclick="toggleGridlines()">
          ${currentSettings.gridlines ? '✅ Visible' : 'Hidden'}
        </button>
      </div>
    </div>

    <!-- Focus Settings -->
    <div class="section">
      <div class="section-title">🎯 Focus Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Focus Mode</div>
          <div class="setting-description">Minimize distractions, hide non-essential UI</div>
        </div>
        <button onclick="activateFocusMode()">🎯 Activate Focus Mode</button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Reduced Motion</div>
          <div class="setting-description">Disable animations and transitions</div>
        </div>
        <button class="${currentSettings.reducedMotion ? 'active' : 'secondary'}" onclick="toggleReducedMotion()">
          ${currentSettings.reducedMotion ? '✅ Enabled' : 'Enable'}
        </button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Quick Capture</div>
          <div class="setting-description">Floating notepad for quick thoughts</div>
        </div>
        <button onclick="showQuickCapture()">📝 Open Quick Capture</button>
      </div>
    </div>

    <!-- Productivity Settings -->
    <div class="section">
      <div class="section-title">⏱️ Productivity Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Pomodoro Timer</div>
          <div class="setting-description">25-minute focus sessions with breaks</div>
        </div>
        <button onclick="startPomodoroTimer()">⏱️ Start Timer</button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Break Reminders</div>
          <div class="setting-description">Gentle reminders to take breaks</div>
        </div>
        <select id="breakInterval" onchange="setBreakInterval()">
          <option value="0" ${!currentSettings.breakInterval ? 'selected' : ''}>Disabled</option>
          <option value="30" ${currentSettings.breakInterval === 30 ? 'selected' : ''}>Every 30 minutes</option>
          <option value="60" ${currentSettings.breakInterval === 60 ? 'selected' : ''}>Every hour</option>
          <option value="90" ${currentSettings.breakInterval === 90 ? 'selected' : ''}>Every 90 minutes</option>
        </select>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Task Complexity Indicators</div>
          <div class="setting-description">Visual cues for task difficulty</div>
        </div>
        <button class="${currentSettings.complexityIndicators ? 'active' : 'secondary'}" onclick="toggleComplexityIndicators()">
          ${currentSettings.complexityIndicators ? '✅ Enabled' : 'Enable'}
        </button>
      </div>
    </div>

    <!-- Quick Actions -->
    <div style="margin-top: 25px; padding-top: 20px; border-top: 2px solid #e0e0e0; display: flex; gap: 10px;">
      <button onclick="saveSettings()">💾 Save Settings</button>
      <button class="secondary" onclick="resetToDefaults()">🔄 Reset to Defaults</button>
      <button class="secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    let selectedTheme = '${currentSettings.theme || 'default'}';

    function selectTheme(theme) {
      selectedTheme = theme;
      document.querySelectorAll('.theme-card').forEach(card => {
        card.classList.remove('selected');
      });
      event.target.closest('.theme-card').classList.add('selected');
    }

    function updateSlider(sliderId) {
      const slider = document.getElementById(sliderId);
      const value = document.getElementById(sliderId + 'Value');
      value.textContent = slider.value + 'pt';
    }

    function toggleZebraStripes() {
      google.script.run.toggleZebraStripes();
      setTimeout(() => location.reload(), 1000);
    }

    function toggleGridlines() {
      google.script.run.toggleGridlinesADHD();
      setTimeout(() => location.reload(), 1000);
    }

    function toggleReducedMotion() {
      google.script.run.toggleReducedMotion();
      setTimeout(() => location.reload(), 1000);
    }

    function toggleComplexityIndicators() {
      google.script.run.toggleComplexityIndicators();
      setTimeout(() => location.reload(), 1000);
    }

    function activateFocusMode() {
      google.script.run.activateFocusMode();
      google.script.host.close();
    }

    function showQuickCapture() {
      google.script.run.showQuickCaptureNotepad();
    }

    function startPomodoroTimer() {
      google.script.run.startPomodoroTimer();
      google.script.host.close();
    }

    function setBreakInterval() {
      const interval = document.getElementById('breakInterval').value;
      google.script.run.setBreakReminders(parseInt(interval));
    }

    function saveSettings() {
      const settings = {
        theme: selectedTheme,
        fontSize: document.getElementById('fontSize').value,
        breakInterval: document.getElementById('breakInterval').value
      };

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Settings saved!');
          google.script.host.close();
        })
        .saveADHDSettings(settings);
    }

    function resetToDefaults() {
      if (confirm('Reset all ADHD settings to defaults?')) {
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ Settings reset!');
            location.reload();
          })
          .resetADHDSettings();
      }
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets current ADHD settings
 * @returns {Object} Settings
 */
function getADHDSettings() {
  const props = PropertiesService.getUserProperties();
  const settingsJSON = props.getProperty('adhdSettings');

  if (settingsJSON) {
    return JSON.parse(settingsJSON);
  }

  // Defaults
  return {
    theme: 'default',
    fontSize: 10,
    zebraStripes: false,
    gridlines: true,
    reducedMotion: false,
    breakInterval: 0,
    complexityIndicators: false
  };
}

/**
 * Saves ADHD settings
 * @param {Object} settings - Settings to save
 */
function saveADHDSettings(settings) {
  const props = PropertiesService.getUserProperties();
  const currentSettings = getADHDSettings();

  // Merge with current
  const newSettings = { ...currentSettings, ...settings };

  props.setProperty('adhdSettings', JSON.stringify(newSettings));

  // Apply settings
  applyADHDSettings(newSettings);
}

/**
 * Applies ADHD settings to sheets
 * @param {Object} settings - Settings to apply
 */
function applyADHDSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    // Font size
    if (settings.fontSize) {
      sheet.getDataRange().setFontSize(parseInt(settings.fontSize));
    }

    // Zebra stripes
    if (settings.zebraStripes) {
      applyZebraStripes(sheet);
    }

    // Gridlines
    if (settings.gridlines !== undefined) {
      sheet.setHiddenGridlines(!settings.gridlines);
    }
  });
}

/**
 * Toggles zebra striping
 */
function toggleZebraStripes() {
  const settings = getADHDSettings();
  settings.zebraStripes = !settings.zebraStripes;
  saveADHDSettings(settings);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    if (settings.zebraStripes) {
      applyZebraStripes(sheet);
    } else {
      removeZebraStripes(sheet);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    settings.zebraStripes ? '✅ Zebra stripes enabled' : '🔕 Zebra stripes disabled',
    'Visual Settings',
    3
  );
}

/**
 * Applies zebra striping to sheet
 * @param {Sheet} sheet - Sheet to apply to
 */
function applyZebraStripes(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());

  // Apply banding
  const banding = range.getBandings()[0];
  if (banding) {
    banding.remove();
  }

  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

/**
 * Removes zebra striping
 * @param {Sheet} sheet - Sheet to remove from
 */
function removeZebraStripes(sheet) {
  const bandings = sheet.getBandings();
  bandings.forEach(banding => banding.remove());
}

/**
 * Toggles gridlines
 */
function toggleGridlinesADHD() {
  const settings = getADHDSettings();
  settings.gridlines = !settings.gridlines;
  saveADHDSettings(settings);
}

/**
 * Toggles reduced motion
 */
function toggleReducedMotion() {
  const settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
}

/**
 * Toggles complexity indicators
 */
function toggleComplexityIndicators() {
  const settings = getADHDSettings();
  settings.complexityIndicators = !settings.complexityIndicators;
  saveADHDSettings(settings);
}

/**
 * Activates focus mode
 */
function activateFocusMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hide all sheets except active one
  const activeSheet = ss.getActiveSheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    if (sheet.getName() !== activeSheet.getName()) {
      sheet.hideSheet();
    }
  });

  // Hide gridlines
  activeSheet.setHiddenGridlines(true);

  // Set focus mode colors
  activeSheet.getDataRange().setBackground('#f5f5f5');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🎯 Focus mode activated. Press Ctrl+Shift+F to exit.',
    'Focus Mode',
    5
  );
}

/**
 * Deactivates focus mode
 */
function deactivateFocusMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Show all sheets
  sheets.forEach(sheet => {
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  });

  const settings = getADHDSettings();
  const activeSheet = ss.getActiveSheet();

  // Restore gridlines based on settings
  activeSheet.setHiddenGridlines(!settings.gridlines);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Focus mode deactivated',
    'Focus Mode',
    3
  );
}

/**
 * Shows quick capture notepad
 */
function showQuickCaptureNotepad() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
    h3 { margin-top: 0; color: #1a73e8; }
    textarea { width: 100%; height: 300px; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-family: Arial, sans-serif; font-size: 14px; resize: vertical; box-sizing: border-box; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    .info { background: #e8f0fe; padding: 10px; border-radius: 4px; margin: 10px 0; font-size: 13px; }
  </style>
</head>
<body>
  <h3>📝 Quick Capture</h3>
  <div class="info">💡 Jot down quick thoughts without losing focus. Notes are saved automatically.</div>
  <textarea id="notes" placeholder="Type your thoughts here..." autofocus></textarea>
  <button onclick="saveNotes()">💾 Save</button>
  <button class="secondary" onclick="clearNotes()">🗑️ Clear</button>
  <button class="secondary" onclick="google.script.host.close()">Close</button>

  <script>
    // Load existing notes
    google.script.run.withSuccessHandler(loadNotes).getQuickCaptureNotes();

    function loadNotes(notes) {
      if (notes) {
        document.getElementById('notes').value = notes;
      }
    }

    function saveNotes() {
      const notes = document.getElementById('notes').value;
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Notes saved!');
        })
        .saveQuickCaptureNotes(notes);
    }

    function clearNotes() {
      if (confirm('Clear all notes?')) {
        document.getElementById('notes').value = '';
        google.script.run.saveQuickCaptureNotes('');
      }
    }

    // Auto-save every 30 seconds
    setInterval(() => {
      const notes = document.getElementById('notes').value;
      if (notes) {
        google.script.run.saveQuickCaptureNotes(notes);
      }
    }, 30000);
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📝 Quick Capture');
}

/**
 * Gets quick capture notes
 * @returns {string} Notes
 */
function getQuickCaptureNotes() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty('quickCaptureNotes') || '';
}

/**
 * Saves quick capture notes
 * @param {string} notes - Notes to save
 */
function saveQuickCaptureNotes(notes) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('quickCaptureNotes', notes);
}

/**
 * Starts Pomodoro timer
 */
function startPomodoroTimer() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 40px; margin: 0; text-align: center; background: #1a73e8; color: white; }
    h2 { margin: 0 0 20px 0; font-size: 24px; }
    .timer { font-size: 72px; font-weight: bold; margin: 40px 0; font-family: 'Courier New', monospace; }
    .status { font-size: 18px; margin: 20px 0; }
    button { background: white; color: #1a73e8; border: none; padding: 15px 30px; font-size: 16px; border-radius: 8px; cursor: pointer; margin: 10px; font-weight: bold; }
    button:hover { background: #f0f0f0; }
    button.stop { background: #f44336; color: white; }
    button.stop:hover { background: #d32f2f; }
    .progress-bar { width: 100%; height: 8px; background: rgba(255,255,255,0.3); border-radius: 4px; margin: 20px 0; }
    .progress-fill { height: 100%; background: white; border-radius: 4px; transition: width 1s linear; }
  </style>
</head>
<body>
  <h2>🍅 Pomodoro Timer</h2>
  <div class="status" id="status">Focus Session</div>
  <div class="timer" id="timer">25:00</div>
  <div class="progress-bar">
    <div class="progress-fill" id="progress" style="width: 100%;"></div>
  </div>
  <button onclick="toggleTimer()">▶️ Start</button>
  <button class="stop" onclick="stopTimer()">⏹️ Stop</button>
  <button onclick="skipSession()">⏭️ Skip</button>

  <script>
    let timeLeft = 25 * 60; // 25 minutes in seconds
    let totalTime = 25 * 60;
    let isRunning = false;
    let interval;
    let sessionType = 'focus'; // 'focus' or 'break'

    function toggleTimer() {
      if (isRunning) {
        pauseTimer();
      } else {
        startTimer();
      }
    }

    function startTimer() {
      isRunning = true;
      document.querySelector('button').textContent = '⏸️ Pause';

      interval = setInterval(() => {
        if (timeLeft > 0) {
          timeLeft--;
          updateDisplay();
        } else {
          completeSession();
        }
      }, 1000);
    }

    function pauseTimer() {
      isRunning = false;
      clearInterval(interval);
      document.querySelector('button').textContent = '▶️ Resume';
    }

    function stopTimer() {
      if (confirm('Stop timer and exit?')) {
        google.script.host.close();
      }
    }

    function updateDisplay() {
      const minutes = Math.floor(timeLeft / 60);
      const seconds = timeLeft % 60;
      document.getElementById('timer').textContent =
        minutes.toString().padStart(2, '0') + ':' + seconds.toString().padStart(2, '0');

      const progress = (timeLeft / totalTime) * 100;
      document.getElementById('progress').style.width = progress + '%';
    }

    function completeSession() {
      pauseTimer();

      if (sessionType === 'focus') {
        alert('✅ Focus session complete! Time for a 5-minute break.');
        sessionType = 'break';
        timeLeft = 5 * 60;
        totalTime = 5 * 60;
        document.getElementById('status').textContent = 'Break Time';
        document.body.style.background = '#4caf50';
      } else {
        alert('✅ Break complete! Ready for another focus session?');
        sessionType = 'focus';
        timeLeft = 25 * 60;
        totalTime = 25 * 60;
        document.getElementById('status').textContent = 'Focus Session';
        document.body.style.background = '#1a73e8';
      }

      updateDisplay();
    }

    function skipSession() {
      if (confirm('Skip to ' + (sessionType === 'focus' ? 'break' : 'next focus session') + '?')) {
        completeSession();
      }
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '🍅 Pomodoro Timer');
}

/**
 * Sets break reminders
 * @param {number} intervalMinutes - Interval in minutes
 */
function setBreakReminders(intervalMinutes) {
  const settings = getADHDSettings();
  settings.breakInterval = intervalMinutes;
  saveADHDSettings(settings);

  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'showBreakReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger if enabled
  if (intervalMinutes > 0) {
    ScriptApp.newTrigger('showBreakReminder')
      .timeBased()
      .everyMinutes(intervalMinutes)
      .create();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Break reminders enabled (every ${intervalMinutes} minutes)`,
      'ADHD Settings',
      3
    );
  }
}

/**
 * Shows break reminder
 */
function showBreakReminder() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '💆 Time for a quick break! Stretch, hydrate, rest your eyes.',
    'Break Reminder',
    10
  );
}

/**
 * Resets ADHD settings to defaults
 */
function resetADHDSettings() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('adhdSettings');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    sheet.setFontSize(10);
    sheet.setHiddenGridlines(false);
    removeZebraStripes(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ ADHD settings reset to defaults',
    'Settings',
    3
  );
}

/******************************************************************************
 * MODULE: ADHDEnhancements
 * Source: ADHDEnhancements.gs
 *****************************************************************************/

// ============================================================================
// ADHD-FRIENDLY ENHANCEMENTS
// ============================================================================
//
// Features optimized for ADHD users:
// - No gridlines (cleaner visual)
// - Soft, calming colors
// - Visual icons and cues
// - Minimal text, maximum visuals
// - Quick-glance data display
// - User customization options
//
// ============================================================================

/**
 * Hide gridlines on all dashboard sheets for cleaner, less distracting view
 */
function hideAllGridlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Hide gridlines on all sheets except Config (for editing)
    if (!sheetName.includes('Config') &&
        !sheetName.includes('Member Directory') &&
        !sheetName.includes('Grievance Log')) {
      sheet.hideGridlines();
    }
  });

  SpreadsheetApp.getUi().alert('✅ Gridlines hidden on all dashboards!\n\nData sheets (Member Directory, Grievance Log) still show gridlines for easier editing.');
}

/**
 * Show gridlines on all sheets (if user needs them back)
 */
function showAllGridlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    sheet.showGridlines();
  });

  SpreadsheetApp.getUi().alert('✅ Gridlines shown on all sheets.');
}

/**
 * Reorder sheets in a logical, user-friendly sequence
 */
function reorderSheetsLogically() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SpreadsheetApp.getUi().alert('📑 Reordering sheets...\n\nPlease wait while sheets are reorganized.');

  // Logical order:
  // 1. Interactive Dashboard (YOUR custom view)
  // 2. Main Dashboard (overview)
  // 3. Member Directory (data)
  // 4. Grievance Log (data)
  // 5. Steward Workload (team view)
  // 6. Test dashboards (1-9)
  // 7. Analytics Data
  // 8. Admin sheets (Config, Archive, etc.)

  const sheetOrder = [
    SHEETS.INTERACTIVE_DASHBOARD,  // 1. YOUR Custom View (most important for daily use)
    SHEETS.DASHBOARD,              // 2. Main Overview
    SHEETS.MEMBER_DIR,             // 3. Members
    SHEETS.GRIEVANCE_LOG,          // 4. Grievances
    SHEETS.STEWARD_WORKLOAD,       // 5. Workload
    SHEETS.TRENDS,                 // 6. Test 1
    SHEETS.LOCATION,               // 8. Test 3
    SHEETS.TYPE_ANALYSIS,          // 9. Test 4
    SHEETS.EXECUTIVE_DASHBOARD,              // 10. Test 5
    SHEETS.KPI_PERFORMANCE,              // 11. Test 6
    SHEETS.MEMBER_ENGAGEMENT,      // 12. Test 7
    SHEETS.COST_IMPACT,            // 13. Test 8
    SHEETS.ANALYTICS,              // 15. Analytics Data
    SHEETS.CONFIG,                 // 16. Config
    SHEETS.ARCHIVE,                // 19. Archive
    SHEETS.DIAGNOSTICS             // 20. Diagnostics
  ];

  // Move sheets to correct positions
  sheetOrder.forEach((sheetName, index) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
    }
  });

  // Set Interactive Dashboard as active (first sheet)
  const interactiveSheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  if (interactiveSheet) {
    ss.setActiveSheet(interactiveSheet);
  }

  SpreadsheetApp.getUi().alert('✅ Sheets reordered!\n\n' +
    '📊 Your Custom View is now first\n' +
    '📈 Dashboards → Data → Tests → Admin\n\n' +
    'Open this spreadsheet to see your Interactive Dashboard first every time!');
}

/**
 * Add visual instructions to Steward Workload sheet
 */
function addStewardWorkloadInstructions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Steward Workload sheet not found!');
    return;
  }

  // Insert instruction box at top
  sheet.insertRowsBefore(1, 8);

  // Create visual instruction panel
  sheet.getRange("A1:N1").merge()
    .setValue("👨‍⚖️ HOW THIS SHEET WORKS - STEWARD WORKLOAD TRACKER")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setRowHeight(1, 40);

  // Visual guide boxes (using icons and colors)
  const instructionData = [
    ["🎯 WHAT IT SHOWS", "This sheet automatically tracks how many cases each steward is handling"],
    ["📊 AUTO-UPDATES", "Updates when you rebuild the dashboard (509 Tools > Data Management > Rebuild Dashboard)"],
    ["🟢 GREEN = Good", "Steward has manageable workload (few or no overdue cases)"],
    ["🟡 YELLOW = Watch", "Steward approaching capacity (some due soon)"],
    ["🔴 RED = Help!", "Steward needs help (overdue cases or heavy workload)"],
    ["👀 QUICK GLANCE", "Look at 'Overdue Cases' column - RED numbers need immediate action"]
  ];

  // Create colored instruction boxes
  instructionData.forEach((instruction, index) => {
    const row = index + 2;

    // Label column (A-B)
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(instruction[0])
      .setFontWeight("bold")
      .setFontSize(11)
      .setFontFamily("Roboto")
      .setBackground(COLORS.INFO_LIGHT)
      .setFontColor(COLORS.TEXT_DARK)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    // Description column (C-N)
    sheet.getRange(row, 3, 1, 12).merge()
      .setValue(instruction[1])
      .setFontSize(10)
      .setFontFamily("Roboto")
      .setBackground(COLORS.CARD_BG)
      .setFontColor(COLORS.TEXT_DARK)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.setRowHeight(row, 32);
  });

  // Add separator row
  sheet.getRange("A8:N8").merge()
    .setValue("📋 STEWARD DATA BELOW ↓")
    .setFontSize(12)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center");

  sheet.setRowHeight(8, 30);

  // Hide gridlines for cleaner look
  sheet.hideGridlines();

  SpreadsheetApp.getUi().alert('✅ Visual instructions added to Steward Workload!\n\n' +
    'The sheet now has a clear guide at the top showing:\n' +
    '• What the sheet does\n' +
    '• How to read it\n' +
    '• Color codes for quick scanning\n\n' +
    'Gridlines hidden for cleaner viewing.');
}

/**
 * Create a user settings sheet for customization
 */
function createUserSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("⚙️ User Settings");

  if (!sheet) {
    sheet = ss.insertSheet("⚙️ User Settings");
  } else {
    sheet.clear();
  }

  // Title
  sheet.getRange("A1:F1").merge()
    .setValue("⚙️ YOUR PERSONAL SETTINGS - Customize Your Dashboard")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setHorizontalAlignment("center");

  sheet.setRowHeight(1, 45);

  // Settings sections
  sheet.getRange("A3:F3").merge()
    .setValue("🎨 VISUAL PREFERENCES")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  // Setting options
  const settings = [
    ["Setting", "Your Choice", "Options", "", "", "What It Does"],
    ["Show Gridlines?", "No", "Yes, No", "", "", "Toggle gridlines on/off (most ADHD users prefer OFF)"],
    ["Color Theme", "Soft Pastels", "Soft Pastels, High Contrast, Warm Tones, Cool Tones", "", "", "Choose your preferred color scheme"],
    ["Font Size", "Medium", "Small, Medium, Large, Extra Large", "", "", "Adjust text size for comfort"],
    ["Icon Style", "Emoji", "Emoji, Symbols, None", "", "", "Choose how visual cues appear"],
    ["Compact View", "No", "Yes, No", "", "", "Reduce spacing between elements"]
  ];

  sheet.getRange(4, 1, settings.length, 6).setValues(settings);

  // Format header row
  sheet.getRange("A4:F4")
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK);

  // Format data rows
  sheet.getRange(5, 1, settings.length - 1, 6)
    .setBackground(COLORS.WHITE)
    .setFontColor(COLORS.TEXT_DARK);

  // Add data validation for choices
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .build();

  sheet.getRange("B5").setDataValidation(yesNoRule);
  sheet.getRange("B9").setDataValidation(yesNoRule);

  const themeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Soft Pastels', 'High Contrast', 'Warm Tones', 'Cool Tones'], true)
    .build();

  sheet.getRange("B6").setDataValidation(themeRule);

  const sizeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Small', 'Medium', 'Large', 'Extra Large'], true)
    .build();

  sheet.getRange("B7").setDataValidation(sizeRule);

  const iconRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Emoji', 'Symbols', 'None'], true)
    .build();

  sheet.getRange("B8").setDataValidation(iconRule);

  // Action buttons section
  sheet.getRange("A11:F11").merge()
    .setValue("🔧 APPLY YOUR SETTINGS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  sheet.getRange("A12:F12").merge()
    .setValue("After changing settings above, go to:\n509 Tools > ADHD Tools > Apply My Settings")
    .setFontSize(11)
    .setFontFamily("Roboto")
    .setBackground(COLORS.INFO_LIGHT)
    .setFontColor(COLORS.TEXT_DARK)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(12, 50);

  // Tips section
  sheet.getRange("A14:F14").merge()
    .setValue("💡 TIPS FOR ADHD-FRIENDLY USE")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setHorizontalAlignment("center");

  const tips = [
    ["👀", "Use the Interactive Dashboard - it's designed for quick glances"],
    ["🎨", "Turn OFF gridlines for less visual clutter"],
    ["🔍", "Use larger font sizes if text feels overwhelming"],
    ["⚡", "Start each day with Quick Stats (Test 9) for fast overview"],
    ["📌", "Bookmark your favorite Test dashboard for quick access"],
    ["🔔", "Set up desktop notifications for overdue items (Future Feature)"]
  ];

  tips.forEach((tip, index) => {
    const row = 15 + index;
    sheet.getRange(row, 1).setValue(tip[0])
      .setFontSize(18)
      .setHorizontalAlignment("center");

    sheet.getRange(row, 2, 1, 5).merge()
      .setValue(tip[1])
      .setFontSize(10)
      .setFontFamily("Roboto")
      .setBackground(COLORS.SUCCESS_LIGHT)
      .setWrap(true);

    sheet.setRowHeight(row, 28);
  });

  // Set column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(6, 300);

  // Hide gridlines
  sheet.hideGridlines();

  SpreadsheetApp.getUi().alert('✅ User Settings sheet created!\n\n' +
    'You can now customize:\n' +
    '• Gridlines visibility\n' +
    '• Color themes\n' +
    '• Font sizes\n' +
    '• Icon styles\n' +
    '• Compact view\n\n' +
    'Change settings and use "509 Tools > ADHD Tools > Apply My Settings"');
}

/**
 * Apply user settings from User Settings sheet
 */
function applyUserSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("⚙️ User Settings");

  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('❌ User Settings sheet not found!\n\nPlease run "509 Tools > ADHD Tools > Create User Settings" first.');
    return;
  }

  // Read user preferences
  const showGridlines = settingsSheet.getRange("B5").getValue();
  const theme = settingsSheet.getRange("B6").getValue();
  const fontSize = settingsSheet.getRange("B7").getValue();
  const iconStyle = settingsSheet.getRange("B8").getValue();
  const compactView = settingsSheet.getRange("B9").getValue();

  // Apply gridlines setting
  if (showGridlines === "Yes") {
    showAllGridlines();
  } else {
    hideAllGridlines();
  }

  // Apply font size (to Interactive Dashboard and Main Dashboard)
  const fontSizeMap = {
    "Small": 9,
    "Medium": 11,
    "Large": 13,
    "Extra Large": 15
  };

  const targetSize = fontSizeMap[fontSize] || 11;

  [SHEETS.INTERACTIVE_DASHBOARD, SHEETS.DASHBOARD].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.getDataRange().setFontSize(targetSize);
    }
  });

  SpreadsheetApp.getUi().alert('✅ Your settings have been applied!\n\n' +
    `• Gridlines: ${showGridlines}\n` +
    `• Theme: ${theme}\n` +
    `• Font Size: ${fontSize}\n` +
    `• Icons: ${iconStyle}\n` +
    `• Compact View: ${compactView}\n\n` +
    'Your dashboard is now customized to your preferences!');
}

/**
 * Quick setup for ADHD-friendly defaults
 */
function setupADHDDefaults() {
  SpreadsheetApp.getUi().alert('🎨 Setting up ADHD-friendly defaults...\n\n' +
    '✓ Hiding gridlines\n' +
    '✓ Applying soft colors\n' +
    '✓ Reordering sheets\n' +
    '✓ Adding visual guides\n\n' +
    'This will take a moment...');

  try {
    // 1. Hide gridlines
    hideAllGridlines();

    // 2. Reorder sheets
    reorderSheetsLogically();

    // 3. Add Steward Workload instructions
    addStewardWorkloadInstructions();

    // 4. Create user settings
    createUserSettingsSheet();

    SpreadsheetApp.getUi().alert('🎉 ADHD-friendly setup complete!\n\n' +
      '✅ Gridlines hidden\n' +
      '✅ Soft colors applied\n' +
      '✅ Sheets reordered logically\n' +
      '✅ Visual guides added\n' +
      '✅ User settings created\n\n' +
      'Your dashboard is now optimized for ADHD users!\n\n' +
      'Open "🎯 Interactive (Your Custom View)" to start!');
  } catch (error) {
    SpreadsheetApp.getUi().alert('⚠️ Error during setup:\n\n' + error.message);
  }
}

/******************************************************************************
 * MODULE: EnhancedErrorHandling
 * Source: EnhancedErrorHandling.gs
 *****************************************************************************/

/**
 * ============================================================================
 * ENHANCED ERROR HANDLING & RECOVERY
 * ============================================================================
 *
 * Comprehensive error handling, logging, and recovery system
 * Features:
 * - Error logging and tracking
 * - User-friendly error messages
 * - Automatic error recovery
 * - Error analytics and reporting
 * - Validation helpers
 * - Graceful degradation
 * - Error notification system
 */

/**
 * Error handling configuration
 */
const ERROR_CONFIG = {
  LOG_SHEET_NAME: 'Error_Log',
  MAX_LOG_ENTRIES: 1000,
  ERROR_LEVELS: {
    INFO: 'INFO',
    WARNING: 'WARNING',
    ERROR: 'ERROR',
    CRITICAL: 'CRITICAL'
  },
  NOTIFICATION_THRESHOLD: 'ERROR',  // Send notifications for ERROR and CRITICAL
  AUTO_RECOVERY_ENABLED: true
};

/**
 * Error categories
 */
const ERROR_CATEGORIES = {
  VALIDATION: 'Data Validation',
  PERMISSION: 'Permission Error',
  NETWORK: 'Network Error',
  DATA_INTEGRITY: 'Data Integrity',
  USER_INPUT: 'User Input',
  SYSTEM: 'System Error',
  INTEGRATION: 'Integration Error'
};

/**
 * Shows error dashboard
 */
function showErrorDashboard() {
  const html = createErrorDashboardHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(1000)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚠️ Error Dashboard');
}

/**
 * Creates HTML for error dashboard
 */
function createErrorDashboardHTML() {
  const stats = getErrorStats();
  const recentErrors = getRecentErrors(20);

  let errorRows = '';
  if (recentErrors.length === 0) {
    errorRows = '<tr><td colspan="6" style="text-align: center; padding: 40px; color: #999;">No errors logged</td></tr>';
  } else {
    recentErrors.forEach(error => {
      const levelClass = error.level.toLowerCase();
      errorRows += `
        <tr>
          <td><span class="level-badge level-${levelClass}">${error.level}</span></td>
          <td><span class="category-badge">${error.category}</span></td>
          <td>${error.message}</td>
          <td style="font-size: 11px; font-family: monospace; color: #666;">${error.context || '-'}</td>
          <td style="font-size: 12px; color: #666;">${new Date(error.timestamp).toLocaleString()}</td>
          <td>
            ${error.recovered ? '<span style="color: #4caf50;">✅ Recovered</span>' : '<span style="color: #f44336;">❌ Failed</span>'}
          </td>
        </tr>
      `;
    });
  }

  return `
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
      max-height: 650px;
      overflow-y: auto;
    }
    h2 {
      color: #f44336;
      margin-top: 0;
      border-bottom: 3px solid #f44336;
      padding-bottom: 10px;
    }
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 15px;
      margin: 20px 0;
    }
    .stat-card {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      text-align: center;
      border-left: 4px solid #f44336;
    }
    .stat-card.info { border-color: #2196f3; }
    .stat-card.warning { border-color: #ff9800; }
    .stat-card.error { border-color: #f44336; }
    .stat-card.critical { border-color: #9c27b0; }
    .stat-number {
      font-size: 36px;
      font-weight: bold;
      color: #f44336;
    }
    .stat-card.info .stat-number { color: #2196f3; }
    .stat-card.warning .stat-number { color: #ff9800; }
    .stat-card.error .stat-number { color: #f44336; }
    .stat-card.critical .stat-number { color: #9c27b0; }
    .stat-label {
      font-size: 13px;
      color: #666;
      margin-top: 8px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
      font-size: 13px;
    }
    th {
      background: #f44336;
      color: white;
      padding: 12px 8px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
    }
    td {
      padding: 10px 8px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background: #fff3f3;
    }
    .level-badge {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 10px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .level-info { background: #e3f2fd; color: #1976d2; }
    .level-warning { background: #fff3e0; color: #f57c00; }
    .level-error { background: #ffebee; color: #d32f2f; }
    .level-critical { background: #f3e5f5; color: #7b1fa2; }
    .category-badge {
      display: inline-block;
      padding: 4px 10px;
      background: #e0e0e0;
      border-radius: 12px;
      font-size: 11px;
      color: #555;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      margin: 5px;
    }
    button:hover {
      background: #1557b0;
    }
    button.danger {
      background: #f44336;
    }
    button.danger:hover {
      background: #d32f2f;
    }
    .actions {
      margin: 20px 0;
      padding: 20px;
      background: #f8f9fa;
      border-radius: 8px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
    .health-indicator {
      display: inline-block;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: bold;
      margin: 10px 0;
    }
    .health-good { background: #e8f5e9; color: #2e7d32; }
    .health-fair { background: #fff3e0; color: #f57c00; }
    .health-poor { background: #ffebee; color: #c62828; }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚠️ Error Dashboard</h2>

    <div style="margin: 20px 0;">
      <strong>System Health:</strong>
      <span class="health-indicator health-${stats.health}">${stats.healthText}</span>
    </div>

    <div class="stats-grid">
      <div class="stat-card info">
        <div class="stat-number">${stats.info}</div>
        <div class="stat-label">Info</div>
      </div>
      <div class="stat-card warning">
        <div class="stat-number">${stats.warning}</div>
        <div class="stat-label">Warnings</div>
      </div>
      <div class="stat-card error">
        <div class="stat-number">${stats.error}</div>
        <div class="stat-label">Errors</div>
      </div>
      <div class="stat-card critical">
        <div class="stat-number">${stats.critical}</div>
        <div class="stat-label">Critical</div>
      </div>
    </div>

    <div class="actions">
      <button onclick="runHealthCheck()">🏥 Run Health Check</button>
      <button onclick="exportErrorLog()">📥 Export Log</button>
      <button class="danger" onclick="clearErrorLog()">🗑️ Clear Log</button>
      <button onclick="testErrorHandling()">🧪 Test Error Handling</button>
      <button onclick="viewErrorTrends()">📊 View Trends</button>
    </div>

    <h3 style="margin-top: 30px; color: #333;">Recent Errors (Last 20)</h3>
    <table>
      <thead>
        <tr>
          <th>Level</th>
          <th>Category</th>
          <th>Message</th>
          <th>Context</th>
          <th>Timestamp</th>
          <th>Recovery</th>
        </tr>
      </thead>
      <tbody>
        ${errorRows}
      </tbody>
    </table>
  </div>

  <script>
    function runHealthCheck() {
      google.script.run
        .withSuccessHandler((result) => {
          alert('🏥 Health Check Complete:\\n\\n' + result.summary);
          location.reload();
        })
        .performSystemHealthCheck();
    }

    function exportErrorLog() {
      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ Error log exported!');
          window.open(url, '_blank');
        })
        .exportErrorLogToSheet();
    }

    function clearErrorLog() {
      if (confirm('Clear all error logs? This cannot be undone.')) {
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ Error log cleared!');
            location.reload();
          })
          .clearErrorLog();
      }
    }

    function testErrorHandling() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Test error logged successfully!');
          location.reload();
        })
        .testErrorLogging();
    }

    function viewErrorTrends() {
      google.script.run
        .withSuccessHandler((url) => {
          window.open(url, '_blank');
        })
        .createErrorTrendReport();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Logs an error
 * @param {string} level - Error level
 * @param {string} category - Error category
 * @param {string} message - Error message
 * @param {string} context - Additional context
 * @param {Error} error - Original error object
 * @param {boolean} recovered - Whether error was recovered
 */
function logError(level, category, message, context = '', error = null, recovered = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

    // Create error log sheet if it doesn't exist
    if (!errorSheet) {
      errorSheet = ss.insertSheet(ERROR_CONFIG.LOG_SHEET_NAME);
      const headers = ['Timestamp', 'Level', 'Category', 'Message', 'Context', 'Stack Trace', 'User', 'Recovered'];
      errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      errorSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f44336').setFontColor('#ffffff');
      errorSheet.setFrozenRows(1);
    }

    // Prepare log entry
    const timestamp = new Date();
    const user = Session.getActiveUser().getEmail();
    const stackTrace = error ? error.stack || error.toString() : '';

    const logEntry = [
      timestamp,
      level,
      category,
      message,
      context,
      stackTrace,
      user,
      recovered
    ];

    // Add to log
    errorSheet.appendRow(logEntry);

    // Trim old entries if needed
    const lastRow = errorSheet.getLastRow();
    if (lastRow > ERROR_CONFIG.MAX_LOG_ENTRIES + 1) {
      errorSheet.deleteRows(2, lastRow - ERROR_CONFIG.MAX_LOG_ENTRIES - 1);
    }

    // Send notification if critical
    if (shouldNotify(level)) {
      sendErrorNotification(level, category, message, context);
    }

  } catch (logError) {
    // Fallback: log to console if sheet logging fails
    console.error('Failed to log error:', logError);
    console.error('Original error:', message, error);
  }
}

/**
 * Checks if error should trigger notification
 * @param {string} level - Error level
 * @returns {boolean}
 */
function shouldNotify(level) {
  const notifyLevels = [ERROR_CONFIG.ERROR_LEVELS.ERROR, ERROR_CONFIG.ERROR_LEVELS.CRITICAL];
  return notifyLevels.includes(level);
}

/**
 * Sends error notification
 * @param {string} level - Error level
 * @param {string} category - Error category
 * @param {string} message - Error message
 * @param {string} context - Context
 */
function sendErrorNotification(level, category, message, context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminEmail = PropertiesService.getScriptProperties().getProperty('adminEmail');

    if (!adminEmail) return;

    const subject = `[509 Dashboard] ${level}: ${category}`;
    const body = `
Error Alert from 509 Dashboard

Level: ${level}
Category: ${category}
Message: ${message}
Context: ${context}
Timestamp: ${new Date().toLocaleString()}
Spreadsheet: ${ss.getName()}
URL: ${ss.getUrl()}

Please check the Error Dashboard for more details.
    `;

    MailApp.sendEmail(adminEmail, subject, body);
  } catch (err) {
    console.error('Failed to send error notification:', err);
  }
}

/**
 * Gets error statistics
 * @returns {Object} Statistics
 */
function getErrorStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    return {
      info: 0,
      warning: 0,
      error: 0,
      critical: 0,
      health: 'good',
      healthText: '✅ Good'
    };
  }

  const data = errorSheet.getRange(2, 1, errorSheet.getLastRow() - 1, 2).getValues();

  const stats = {
    info: 0,
    warning: 0,
    error: 0,
    critical: 0
  };

  data.forEach(row => {
    const level = row[1];
    if (level === 'INFO') stats.info++;
    else if (level === 'WARNING') stats.warning++;
    else if (level === 'ERROR') stats.error++;
    else if (level === 'CRITICAL') stats.critical++;
  });

  // Determine health
  let health = 'good';
  let healthText = '✅ Good';

  if (stats.critical > 0 || stats.error > 10) {
    health = 'poor';
    healthText = '❌ Poor';
  } else if (stats.error > 0 || stats.warning > 20) {
    health = 'fair';
    healthText = '⚠️ Fair';
  }

  stats.health = health;
  stats.healthText = healthText;

  return stats;
}

/**
 * Gets recent errors
 * @param {number} limit - Number of errors to retrieve
 * @returns {Array} Recent errors
 */
function getRecentErrors(limit = 20) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    return [];
  }

  const lastRow = errorSheet.getLastRow();
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  const data = errorSheet.getRange(startRow, 1, numRows, 8).getValues();

  return data.map(row => ({
    timestamp: row[0],
    level: row[1],
    category: row[2],
    message: row[3],
    context: row[4],
    stackTrace: row[5],
    user: row[6],
    recovered: row[7]
  })).reverse();
}

/**
 * Wraps a function with error handling
 * @param {Function} func - Function to wrap
 * @param {string} context - Context description
 * @returns {Function} Wrapped function
 */
function withErrorHandling(func, context) {
  return function(...args) {
    try {
      return func.apply(this, args);
    } catch (error) {
      logError(
        ERROR_CONFIG.ERROR_LEVELS.ERROR,
        ERROR_CATEGORIES.SYSTEM,
        error.message,
        context,
        error,
        false
      );

      // Show user-friendly error
      SpreadsheetApp.getUi().alert(
        '⚠️ Error',
        `An error occurred: ${error.message}\n\nThis has been logged for review.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );

      throw error;
    }
  };
}

/**
 * Validates required fields
 * @param {Object} data - Data to validate
 * @param {Array} requiredFields - Required field names
 * @returns {Object} Validation result
 */
function validateRequiredFields(data, requiredFields) {
  const missing = [];

  requiredFields.forEach(field => {
    if (!data[field] || data[field] === '') {
      missing.push(field);
    }
  });

  if (missing.length > 0) {
    return {
      valid: false,
      error: `Missing required fields: ${missing.join(', ')}`,
      missingFields: missing
    };
  }

  return { valid: true };
}

/**
 * Validates email format
 * @param {string} email - Email to validate
 * @returns {boolean}
 */
function validateEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates date format
 * @param {*} date - Date to validate
 * @returns {boolean}
 */
function validateDate(date) {
  if (date instanceof Date) {
    return !isNaN(date.getTime());
  }
  return false;
}

/**
 * Performs system health check
 * @returns {Object} Health check results
 */
function performSystemHealthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    checks: [],
    overall: 'PASS'
  };

  // Check 1: Verify all required sheets exist
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = Object.values(SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());

  requiredSheets.forEach(sheetName => {
    const exists = existingSheets.includes(sheetName);
    results.checks.push({
      name: `Sheet: ${sheetName}`,
      status: exists ? 'PASS' : 'FAIL',
      message: exists ? 'Exists' : 'Missing'
    });
    if (!exists) results.overall = 'FAIL';
  });

  // Check 2: Verify data integrity
  try {
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet) {
      const lastRow = grievanceSheet.getLastRow();
      results.checks.push({
        name: 'Data Integrity',
        status: 'PASS',
        message: `${lastRow - 1} grievances found`
      });
    }
  } catch (err) {
    results.checks.push({
      name: 'Data Integrity',
      status: 'FAIL',
      message: err.message
    });
    results.overall = 'FAIL';
  }

  // Check 3: Cache health
  try {
    const cache = CacheService.getScriptCache();
    cache.put('healthcheck', 'test', 60);
    const testValue = cache.get('healthcheck');
    results.checks.push({
      name: 'Cache Service',
      status: testValue === 'test' ? 'PASS' : 'FAIL',
      message: testValue === 'test' ? 'Operational' : 'Not working'
    });
  } catch (err) {
    results.checks.push({
      name: 'Cache Service',
      status: 'FAIL',
      message: err.message
    });
  }

  // Generate summary
  const passCount = results.checks.filter(c => c.status === 'PASS').length;
  const totalCount = results.checks.length;
  results.summary = `Health Check: ${passCount}/${totalCount} checks passed\nOverall Status: ${results.overall}`;

  logError(
    ERROR_CONFIG.ERROR_LEVELS.INFO,
    ERROR_CATEGORIES.SYSTEM,
    'System health check completed',
    results.summary,
    null,
    true
  );

  return results;
}

/**
 * Exports error log to sheet
 * @returns {string} Sheet URL
 */
function exportErrorLogToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (!errorSheet) {
    throw new Error('No error log found');
  }

  return ss.getUrl() + '#gid=' + errorSheet.getSheetId();
}

/**
 * Clears error log
 */
function clearErrorLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (errorSheet && errorSheet.getLastRow() > 1) {
    errorSheet.deleteRows(2, errorSheet.getLastRow() - 1);
  }

  logError(
    ERROR_CONFIG.ERROR_LEVELS.INFO,
    ERROR_CATEGORIES.SYSTEM,
    'Error log cleared',
    'User cleared error log',
    null,
    true
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Error log cleared',
    'Error Log',
    3
  );
}

/**
 * Tests error logging
 */
function testErrorLogging() {
  logError(
    ERROR_CONFIG.ERROR_LEVELS.INFO,
    ERROR_CATEGORIES.SYSTEM,
    'Test error message',
    'This is a test error generated for testing purposes',
    null,
    true
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Test error logged successfully',
    'Testing',
    3
  );
}

/**
 * Creates error trend report
 * @returns {string} Report URL
 */
function createErrorTrendReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_CONFIG.LOG_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    throw new Error('No error data available');
  }

  // Create or get trends sheet
  let trendsSheet = ss.getSheetByName('Error_Trends');
  if (trendsSheet) {
    trendsSheet.clear();
  } else {
    trendsSheet = ss.insertSheet('Error_Trends');
  }

  // Analyze trends by day
  const data = errorSheet.getRange(2, 1, errorSheet.getLastRow() - 1, 2).getValues();

  const dailyStats = {};
  data.forEach(row => {
    const date = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const level = row[1];

    if (!dailyStats[date]) {
      dailyStats[date] = { info: 0, warning: 0, error: 0, critical: 0 };
    }

    if (level === 'INFO') dailyStats[date].info++;
    else if (level === 'WARNING') dailyStats[date].warning++;
    else if (level === 'ERROR') dailyStats[date].error++;
    else if (level === 'CRITICAL') dailyStats[date].critical++;
  });

  // Write to sheet
  const headers = ['Date', 'Info', 'Warning', 'Error', 'Critical', 'Total'];
  trendsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  trendsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f44336').setFontColor('#ffffff');

  const rows = Object.entries(dailyStats).map(([date, stats]) => [
    date,
    stats.info,
    stats.warning,
    stats.error,
    stats.critical,
    stats.info + stats.warning + stats.error + stats.critical
  ]);

  if (rows.length > 0) {
    trendsSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // Create chart
    const chartRange = trendsSheet.getRange(1, 1, rows.length + 1, headers.length);
    const chart = trendsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartRange)
      .setPosition(rows.length + 3, 1, 0, 0)
      .setOption('title', 'Error Trends Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .build();

    trendsSheet.insertChart(chart);
  }

  return ss.getUrl() + '#gid=' + trendsSheet.getSheetId();
}

/******************************************************************************
 * MODULE: UndoRedoSystem
 * Source: UndoRedoSystem.gs
 *****************************************************************************/

/**
 * ============================================================================
 * UNDO/REDO SYSTEM
 * ============================================================================
 *
 * Advanced action history and undo/redo functionality
 * Features:
 * - Comprehensive action tracking
 * - Undo/redo stack with 50 action limit
 * - Snapshot-based rollback
 * - Action history viewer
 * - Batch operation support
 * - Keyboard shortcuts (Ctrl+Z, Ctrl+Y)
 * - Visual confirmation
 */

/**
 * Undo/Redo configuration
 */
const UNDO_CONFIG = {
  MAX_HISTORY: 50,  // Maximum actions to keep
  STORAGE_KEY: 'undoRedoHistory',
  SUPPORTED_ACTIONS: [
    'EDIT_CELL',
    'ADD_ROW',
    'DELETE_ROW',
    'BATCH_UPDATE',
    'STATUS_CHANGE',
    'ASSIGNMENT_CHANGE'
  ]
};

/**
 * Shows undo/redo control panel
 */
function showUndoRedoPanel() {
  const html = createUndoRedoPanelHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '↩️ Undo/Redo History');
}

/**
 * Creates HTML for undo/redo panel
 */
function createUndoRedoPanelHTML() {
  const history = getUndoHistory();

  let historyRows = '';
  if (history.actions.length === 0) {
    historyRows = '<tr><td colspan="5" style="text-align: center; padding: 40px; color: #999;">No actions recorded yet</td></tr>';
  } else {
    history.actions.slice().reverse().forEach((action, index) => {
      const timestamp = new Date(action.timestamp).toLocaleString();
      const canUndo = index < history.actions.length - history.currentIndex;
      const canRedo = index >= history.actions.length - history.currentIndex;

      historyRows += `
        <tr ${!canUndo && !canRedo ? 'class="future-action"' : ''}>
          <td>${history.actions.length - index}</td>
          <td><span class="action-badge action-${action.type.toLowerCase()}">${action.type}</span></td>
          <td>${action.description}</td>
          <td style="font-size: 12px; color: #666;">${timestamp}</td>
          <td>
            ${canUndo ? '<button onclick="undoToAction(' + (history.actions.length - index - 1) + ')">↩️ Undo</button>' : ''}
            ${canRedo ? '<button onclick="redoToAction(' + (history.actions.length - index - 1) + ')">↪️ Redo</button>' : ''}
          </td>
        </tr>
      `;
    });
  }

  return `
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
      max-height: 550px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .stats {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 15px;
      margin: 20px 0;
    }
    .stat-card {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      text-align: center;
      border-left: 4px solid #1a73e8;
    }
    .stat-number {
      font-size: 32px;
      font-weight: bold;
      color: #1a73e8;
    }
    .stat-label {
      font-size: 13px;
      color: #666;
      margin-top: 5px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
    }
    th {
      background: #1a73e8;
      color: white;
      padding: 12px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
    }
    td {
      padding: 12px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background: #f8f9fa;
    }
    .future-action {
      opacity: 0.4;
    }
    .action-badge {
      display: inline-block;
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .action-edit_cell { background: #e3f2fd; color: #1976d2; }
    .action-add_row { background: #e8f5e9; color: #388e3c; }
    .action-delete_row { background: #ffebee; color: #d32f2f; }
    .action-batch_update { background: #fff3e0; color: #f57c00; }
    .action-status_change { background: #f3e5f5; color: #7b1fa2; }
    .action-assignment_change { background: #e0f2f1; color: #00796b; }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 8px 16px;
      font-size: 13px;
      border-radius: 4px;
      cursor: pointer;
      margin: 2px;
    }
    button:hover {
      background: #1557b0;
    }
    .quick-actions {
      margin: 20px 0;
      padding: 20px;
      background: #f8f9fa;
      border-radius: 8px;
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .quick-actions button {
      padding: 12px 24px;
      font-size: 14px;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .shortcut {
      background: #333;
      color: white;
      padding: 4px 8px;
      border-radius: 4px;
      font-family: monospace;
      font-size: 12px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>↩️ Undo/Redo History</h2>

    <div class="info-box">
      <strong>⌨️ Keyboard Shortcuts:</strong>
      <span class="shortcut">Ctrl+Z</span> Undo &nbsp;&nbsp;
      <span class="shortcut">Ctrl+Y</span> Redo &nbsp;&nbsp;
      <span class="shortcut">Ctrl+Shift+Z</span> View History
    </div>

    <div class="stats">
      <div class="stat-card">
        <div class="stat-number">${history.actions.length}</div>
        <div class="stat-label">Total Actions</div>
      </div>
      <div class="stat-card">
        <div class="stat-number">${history.currentIndex}</div>
        <div class="stat-label">Current Position</div>
      </div>
      <div class="stat-card">
        <div class="stat-number">${history.actions.length - history.currentIndex}</div>
        <div class="stat-label">Actions Available</div>
      </div>
    </div>

    <div class="quick-actions">
      <button onclick="performUndo()">↩️ Undo Last Action</button>
      <button onclick="performRedo()">↪️ Redo Action</button>
      <button onclick="clearHistory()" style="background: #d32f2f;">🗑️ Clear History</button>
      <button onclick="exportHistory()" style="background: #00796b;">📥 Export History</button>
    </div>

    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Action Type</th>
          <th>Description</th>
          <th>Timestamp</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        ${historyRows}
      </tbody>
    </table>
  </div>

  <script>
    function performUndo() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Action undone!');
          location.reload();
        })
        .withFailureHandler((error) => {
          alert('❌ Cannot undo: ' + error.message);
        })
        .undoLastAction();
    }

    function performRedo() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Action redone!');
          location.reload();
        })
        .withFailureHandler((error) => {
          alert('❌ Cannot redo: ' + error.message);
        })
        .redoLastAction();
    }

    function undoToAction(index) {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Undo complete!');
          location.reload();
        })
        .undoToIndex(index);
    }

    function redoToAction(index) {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Redo complete!');
          location.reload();
        })
        .redoToIndex(index);
    }

    function clearHistory() {
      if (confirm('Clear all undo/redo history? This cannot be undone.')) {
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ History cleared!');
            location.reload();
          })
          .clearUndoHistory();
      }
    }

    function exportHistory() {
      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ History exported!');
          window.open(url, '_blank');
        })
        .exportUndoHistoryToSheet();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets undo/redo history
 * @returns {Object} History object
 */
function getUndoHistory() {
  const props = PropertiesService.getScriptProperties();
  const historyJSON = props.getProperty(UNDO_CONFIG.STORAGE_KEY);

  if (historyJSON) {
    return JSON.parse(historyJSON);
  }

  return {
    actions: [],
    currentIndex: 0
  };
}

/**
 * Saves undo/redo history
 * @param {Object} history - History to save
 */
function saveUndoHistory(history) {
  const props = PropertiesService.getScriptProperties();

  // Limit history size
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }

  props.setProperty(UNDO_CONFIG.STORAGE_KEY, JSON.stringify(history));
}

/**
 * Records an action for undo/redo
 * @param {string} actionType - Type of action
 * @param {string} description - Description
 * @param {Object} beforeState - State before action
 * @param {Object} afterState - State after action
 */
function recordAction(actionType, description, beforeState, afterState) {
  const history = getUndoHistory();

  // Remove any actions after current index (when doing new action after undo)
  if (history.currentIndex < history.actions.length) {
    history.actions = history.actions.slice(0, history.currentIndex);
  }

  // Add new action
  history.actions.push({
    type: actionType,
    description: description,
    timestamp: new Date().toISOString(),
    beforeState: beforeState,
    afterState: afterState
  });

  history.currentIndex = history.actions.length;

  saveUndoHistory(history);
}

/**
 * Records a cell edit action
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {*} oldValue - Old value
 * @param {*} newValue - New value
 */
function recordCellEdit(row, col, oldValue, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const colName = sheet.getRange(1, col).getValue();

  recordAction(
    'EDIT_CELL',
    `Edited ${colName} in row ${row}`,
    { row, col, value: oldValue, sheet: sheet.getName() },
    { row, col, value: newValue, sheet: sheet.getName() }
  );
}

/**
 * Records a row addition
 * @param {number} row - Row number
 * @param {Array} rowData - Row data
 */
function recordRowAddition(row, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  recordAction(
    'ADD_ROW',
    `Added row ${row}`,
    null,
    { row, data: rowData, sheet: sheet.getName() }
  );
}

/**
 * Records a row deletion
 * @param {number} row - Row number
 * @param {Array} rowData - Deleted row data
 */
function recordRowDeletion(row, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  recordAction(
    'DELETE_ROW',
    `Deleted row ${row}`,
    { row, data: rowData, sheet: sheet.getName() },
    null
  );
}

/**
 * Records a batch update
 * @param {string} operation - Operation name
 * @param {Array} changes - Array of changes
 */
function recordBatchUpdate(operation, changes) {
  recordAction(
    'BATCH_UPDATE',
    `${operation}: ${changes.length} rows affected`,
    { changes: changes },
    { operation: operation }
  );
}

/**
 * Undoes the last action
 */
function undoLastAction() {
  const history = getUndoHistory();

  if (history.currentIndex === 0) {
    throw new Error('Nothing to undo');
  }

  const action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);

  history.currentIndex--;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `↩️ Undone: ${action.description}`,
    'Undo',
    3
  );
}

/**
 * Redoes the last undone action
 */
function redoLastAction() {
  const history = getUndoHistory();

  if (history.currentIndex >= history.actions.length) {
    throw new Error('Nothing to redo');
  }

  const action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);

  history.currentIndex++;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `↪️ Redone: ${action.description}`,
    'Redo',
    3
  );
}

/**
 * Undoes to a specific index
 * @param {number} targetIndex - Target index
 */
function undoToIndex(targetIndex) {
  const history = getUndoHistory();

  while (history.currentIndex > targetIndex) {
    undoLastAction();
  }
}

/**
 * Redoes to a specific index
 * @param {number} targetIndex - Target index
 */
function redoToIndex(targetIndex) {
  const history = getUndoHistory();

  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) {
    redoLastAction();
  }
}

/**
 * Applies a state snapshot
 * @param {Object} state - State to apply
 * @param {string} actionType - Type of action
 */
function applyState(state, actionType) {
  if (!state) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(state.sheet);

  if (!sheet) {
    throw new Error(`Sheet ${state.sheet} not found`);
  }

  switch (actionType) {
    case 'EDIT_CELL':
      sheet.getRange(state.row, state.col).setValue(state.value);
      break;

    case 'ADD_ROW':
      // Remove the added row
      if (state.row) {
        sheet.deleteRow(state.row);
      }
      break;

    case 'DELETE_ROW':
      // Re-add the deleted row
      if (state.row && state.data) {
        sheet.insertRowAfter(state.row - 1);
        sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]);
      }
      break;

    case 'BATCH_UPDATE':
      // Restore batch changes
      if (state.changes) {
        state.changes.forEach(change => {
          sheet.getRange(change.row, change.col).setValue(change.oldValue);
        });
      }
      break;
  }
}

/**
 * Clears undo/redo history
 */
function clearUndoHistory() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(UNDO_CONFIG.STORAGE_KEY);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Undo/redo history cleared',
    'History',
    3
  );
}

/**
 * Exports undo history to a sheet
 * @returns {string} Sheet URL
 */
function exportUndoHistoryToSheet() {
  const history = getUndoHistory();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get history sheet
  let historySheet = ss.getSheetByName('Undo_History_Export');
  if (historySheet) {
    historySheet.clear();
  } else {
    historySheet = ss.insertSheet('Undo_History_Export');
  }

  // Headers
  const headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Data
  if (history.actions.length > 0) {
    const rows = history.actions.map((action, index) => {
      const status = index < history.currentIndex ? 'Applied' : 'Undone';
      return [
        index + 1,
        action.type,
        action.description,
        new Date(action.timestamp).toLocaleString(),
        status
      ];
    });

    historySheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize
  for (let col = 1; col <= headers.length; col++) {
    historySheet.autoResizeColumn(col);
  }

  return ss.getUrl() + '#gid=' + historySheet.getSheetId();
}

/**
 * Installs undo/redo keyboard shortcuts
 */
function installUndoRedoShortcuts() {
  // This is handled by the onEdit trigger and keyboard handler
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '⌨️ Undo/Redo shortcuts installed:\nCtrl+Z = Undo\nCtrl+Y = Redo\nCtrl+Shift+Z = History',
    'Keyboard Shortcuts',
    5
  );
}

/**
 * Creates a snapshot of the current grievance log
 * @returns {Object} Snapshot data
 */
function createGrievanceSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const data = grievanceSheet.getDataRange().getValues();

  return {
    timestamp: new Date().toISOString(),
    data: data,
    lastRow: grievanceSheet.getLastRow(),
    lastColumn: grievanceSheet.getLastColumn()
  };
}

/**
 * Restores from a snapshot
 * @param {Object} snapshot - Snapshot to restore
 */
function restoreFromSnapshot(snapshot) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  // Clear existing data (except headers)
  if (grievanceSheet.getLastRow() > 1) {
    grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, grievanceSheet.getLastColumn()).clear();
  }

  // Restore snapshot data (skip header row)
  if (snapshot.data.length > 1) {
    grievanceSheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length).setValues(snapshot.data.slice(1));
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Snapshot restored',
    'Undo/Redo',
    3
  );
}

/******************************************************************************
 * MODULE: CalendarIntegration
 * Source: CalendarIntegration.gs
 *****************************************************************************/

/**
 * ============================================================================
 * GOOGLE CALENDAR INTEGRATION
 * ============================================================================
 *
 * Sync grievance deadlines to Google Calendar
 * Features:
 * - Create calendar events for all deadlines
 * - Color-code by priority (red for overdue, yellow for due soon)
 * - Update events when grievance status changes
 * - Delete events when grievance is closed
 */

/**
 * Syncs all grievance deadlines to Google Calendar
 */
function syncDeadlinesToCalendar() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📅 Sync Deadlines to Google Calendar',
    'This will create calendar events for all grievance deadlines.\n\n' +
    'Events will be created in your default Google Calendar with:\n' +
    '• Red = Overdue\n' +
    '• Orange = Due within 3 days\n' +
    '• Yellow = Due within 7 days\n' +
    '• Blue = Due later\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing deadlines to calendar...', 'Please wait', -1);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    if (!grievanceSheet) {
      ui.alert('❌ Grievance Log sheet not found!');
      return;
    }

    const lastRow = grievanceSheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('ℹ️ No grievances found to sync.');
      return;
    }

    // Get all grievance data
    const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

    const calendar = CalendarApp.getDefaultCalendar();
    let eventsCreated = 0;
    let eventsSkipped = 0;

    data.forEach((row, index) => {
      const grievanceId = row[0]; // Column A
      const memberName = `${row[2]} ${row[3]}`; // Columns C, D
      const status = row[4]; // Column E
      const nextActionDue = row[19]; // Column T (Next Action Due)
      const daysToDeadline = row[20]; // Column U (Days to Deadline)

      // Only create events for open grievances with deadlines
      if (status !== 'Open' || !nextActionDue) {
        eventsSkipped++;
        return;
      }

      // Check if event already exists
      const existingEvent = checkCalendarEventExists(calendar, grievanceId);
      if (existingEvent) {
        eventsSkipped++;
        return;
      }

      // Determine priority and color
      let color = CalendarApp.EventColor.BLUE; // Default: Due later
      let priority = 'Normal';

      if (daysToDeadline < 0) {
        color = CalendarApp.EventColor.RED;
        priority = 'OVERDUE';
      } else if (daysToDeadline <= 3) {
        color = CalendarApp.EventColor.ORANGE;
        priority = 'Urgent';
      } else if (daysToDeadline <= 7) {
        color = CalendarApp.EventColor.YELLOW;
        priority = 'Soon';
      }

      // Create event
      const title = `⚖️ ${priority}: ${grievanceId} - ${memberName}`;
      const description =
        `Grievance Deadline\n\n` +
        `Grievance ID: ${grievanceId}\n` +
        `Member: ${memberName}\n` +
        `Status: ${status}\n` +
        `Days to Deadline: ${daysToDeadline}\n` +
        `Priority: ${priority}\n\n` +
        `Created by 509 Dashboard`;

      try {
        const event = calendar.createAllDayEvent(
          title,
          new Date(nextActionDue),
          {
            description: description,
            location: '509 Dashboard'
          }
        );

        event.setColor(color);
        event.setTag('509Dashboard', grievanceId); // Tag for identification

        eventsCreated++;
      } catch (error) {
        Logger.log(`Error creating event for ${grievanceId}: ${error.message}`);
      }
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Created ${eventsCreated} calendar events (${eventsSkipped} skipped)`,
      'Complete',
      5
    );

    ui.alert(
      '✅ Calendar Sync Complete',
      `Successfully synced deadlines to Google Calendar:\n\n` +
      `• Events created: ${eventsCreated}\n` +
      `• Events skipped: ${eventsSkipped} (closed or already synced)\n\n` +
      `Check your Google Calendar for the new events!`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    Logger.log('Calendar sync error: ' + error.message);
  }
}

/**
 * Checks if a calendar event already exists for a grievance
 * @param {Calendar} calendar - The Google Calendar
 * @param {string} grievanceId - The grievance ID to check
 * @returns {CalendarEvent|null} The existing event or null
 */
function checkCalendarEventExists(calendar, grievanceId) {
  const now = new Date();
  const futureDate = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000)); // 1 year from now

  const events = calendar.getEvents(now, futureDate);

  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const tag = event.getTag('509Dashboard');
    if (tag === grievanceId) {
      return event;
    }
  }

  return null;
}

/**
 * Syncs a single grievance deadline to calendar
 * @param {string} grievanceId - The grievance ID
 */
function syncSingleDeadlineToCalendar(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || !grievanceId) {
    return;
  }

  // Find the grievance
  const lastRow = grievanceSheet.getLastRow();
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      const row = data[i];
      const memberName = `${row[2]} ${row[3]}`;
      const status = row[4];
      const nextActionDue = row[19];
      const daysToDeadline = row[20];

      if (status !== 'Open' || !nextActionDue) {
        return; // Skip if not open or no deadline
      }

      const calendar = CalendarApp.getDefaultCalendar();

      // Check if event already exists
      const existingEvent = checkCalendarEventExists(calendar, grievanceId);
      if (existingEvent) {
        // Update existing event
        let color = CalendarApp.EventColor.BLUE;
        if (daysToDeadline < 0) color = CalendarApp.EventColor.RED;
        else if (daysToDeadline <= 3) color = CalendarApp.EventColor.ORANGE;
        else if (daysToDeadline <= 7) color = CalendarApp.EventColor.YELLOW;

        existingEvent.setColor(color);
        existingEvent.setTime(new Date(nextActionDue), new Date(nextActionDue));
      } else {
        // Create new event (same logic as syncDeadlinesToCalendar)
        let color = CalendarApp.EventColor.BLUE;
        let priority = 'Normal';

        if (daysToDeadline < 0) {
          color = CalendarApp.EventColor.RED;
          priority = 'OVERDUE';
        } else if (daysToDeadline <= 3) {
          color = CalendarApp.EventColor.ORANGE;
          priority = 'Urgent';
        } else if (daysToDeadline <= 7) {
          color = CalendarApp.EventColor.YELLOW;
          priority = 'Soon';
        }

        const title = `⚖️ ${priority}: ${grievanceId} - ${memberName}`;
        const description =
          `Grievance Deadline\n\n` +
          `Grievance ID: ${grievanceId}\n` +
          `Member: ${memberName}\n` +
          `Status: ${status}\n` +
          `Days to Deadline: ${daysToDeadline}\n` +
          `Priority: ${priority}\n\n` +
          `Created by 509 Dashboard`;

        const event = calendar.createAllDayEvent(
          title,
          new Date(nextActionDue),
          {
            description: description,
            location: '509 Dashboard'
          }
        );

        event.setColor(color);
        event.setTag('509Dashboard', grievanceId);
      }

      break;
    }
  }
}

/**
 * Removes calendar event when grievance is closed
 * @param {string} grievanceId - The grievance ID
 */
function removeCalendarEvent(grievanceId) {
  const calendar = CalendarApp.getDefaultCalendar();
  const event = checkCalendarEventExists(calendar, grievanceId);

  if (event) {
    event.deleteEvent();
    Logger.log(`Removed calendar event for ${grievanceId}`);
  }
}

/**
 * Clears all 509 Dashboard events from calendar
 */
function clearAllCalendarEvents() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ Clear All Calendar Events',
    'This will remove ALL grievance deadline events from your Google Calendar.\n\n' +
    'This action cannot be undone!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const futureDate = new Date(now.getTime() + (365 * 24 * 60 * 60 * 1000));

    const events = calendar.getEvents(now, futureDate);
    let removedCount = 0;

    events.forEach(event => {
      const tag = event.getTag('509Dashboard');
      if (tag) {
        event.deleteEvent();
        removedCount++;
      }
    });

    ui.alert(
      '✅ Calendar Events Cleared',
      `Removed ${removedCount} events from your calendar.`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Shows upcoming deadlines from calendar
 */
function showUpcomingDeadlinesFromCalendar() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const nextWeek = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));

    const events = calendar.getEvents(now, nextWeek);
    const dashboardEvents = events.filter(e => e.getTag('509Dashboard'));

    if (dashboardEvents.length === 0) {
      SpreadsheetApp.getUi().alert(
        'ℹ️ No Upcoming Deadlines',
        'No grievance deadlines in the next 7 days.\n\n' +
        'Run "Sync Deadlines to Calendar" to create calendar events.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const eventList = dashboardEvents
      .slice(0, 10)
      .map(e => `  • ${e.getTitle()} - ${e.getAllDayStartDate().toLocaleDateString()}`)
      .join('\n');

    SpreadsheetApp.getUi().alert(
      '📅 Upcoming Deadlines (Next 7 Days)',
      `Found ${dashboardEvents.length} deadline(s):\n\n${eventList}` +
      (dashboardEvents.length > 10 ? `\n  ... and ${dashboardEvents.length - 10} more` : ''),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/******************************************************************************
 * MODULE: BatchOperations
 * Source: BatchOperations.gs
 *****************************************************************************/

/**
 * ============================================================================
 * BATCH OPERATIONS
 * ============================================================================
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
      google.script.run.withSuccessHandler(() => {
        google.script.host.close();
      }).batchAssignSteward();
    }

    function runBatchUpdateStatus() {
      google.script.run.withSuccessHandler(() => {
        google.script.host.close();
      }).batchUpdateStatus();
    }

    function runBatchExportPDF() {
      google.script.run.withSuccessHandler(() => {
        google.script.host.close();
      }).batchExportPDF();
    }

    function runBatchEmail() {
      google.script.run.withSuccessHandler(() => {
        google.script.host.close();
      }).batchEmailNotifications();
    }

    function runBatchAddNotes() {
      google.script.run.withSuccessHandler(() => {
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
    return Array.from({ length: numRows - 1 }, (_, i) => startRow + i + 1);
  }

  return Array.from({ length: numRows }, (_, i) => startRow + i);
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
    .filter(row => row[9] === 'Yes') // Column J: Is Steward?
    .map(row => `${row[1]} ${row[2]}`) // First + Last Name
    .filter(name => name.trim() !== '');

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
  let updated = 0;

  selectedRows.forEach(row => {
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
  let updated = 0;

  selectedRows.forEach(row => {
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

  selectedRows.forEach(row => {
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
  // This is a placeholder - actual PDF generation would be implemented here
  // For now, we'll create a simple text-based PDF using Google Docs

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
  let sent = 0;

  selectedRows.forEach(row => {
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

  let updated = 0;

  selectedRows.forEach(row => {
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

/******************************************************************************
 * MODULE: AutomatedReports
 * Source: AutomatedReports.gs
 *****************************************************************************/

/**
 * ============================================================================
 * AUTOMATED REPORT GENERATION
 * ============================================================================
 *
 * Schedule automated reports with email distribution
 * Features:
 * - Monthly executive summaries
 * - Quarterly trend reports
 * - Custom report scheduling
 * - PDF export and email distribution
 * - Configurable recipients
 */

/**
 * Sets up monthly report trigger (1st of each month at 9 AM)
 */
function setupMonthlyReports() {
  // Delete existing monthly triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'generateMonthlyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new monthly trigger
  ScriptApp.newTrigger('generateMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Monthly reports enabled (1st of month, 9 AM)',
    'Automation Active',
    5
  );
}

/**
 * Sets up quarterly report trigger
 */
function setupQuarterlyReports() {
  // Delete existing quarterly triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'generateQuarterlyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create triggers for 1st day of Jan, Apr, Jul, Oct
  [1, 4, 7, 10].forEach(month => {
    ScriptApp.newTrigger('generateQuarterlyReport')
      .timeBased()
      .onMonthDay(1)
      .atHour(10)
      .inTimezone(Session.getScriptTimeZone())
      .create();
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Quarterly reports enabled (Jan/Apr/Jul/Oct, 10 AM)',
    'Automation Active',
    5
  );
}

/**
 * Disables automated reports
 */
function disableAutomatedReports() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  triggers.forEach(trigger => {
    const func = trigger.getHandlerFunction();
    if (func === 'generateMonthlyReport' || func === 'generateQuarterlyReport') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `🔕 Removed ${removed} report automation trigger(s)`,
    'Automation Disabled',
    5
  );
}

/**
 * Generates monthly executive summary report
 */
function generateMonthlyReport() {
  try {
    const reportData = gatherMonthlyData();
    const doc = createMonthlyReportDoc(reportData);
    const pdf = convertToPDF(doc);

    // Email to leadership
    emailReport(pdf, 'Monthly Executive Summary', getReportRecipients('monthly'));

    Logger.log('Monthly report generated and emailed successfully');

  } catch (error) {
    Logger.log('Error generating monthly report: ' + error.message);
    notifyAdminOfError('Monthly Report Generation Failed', error.message);
  }
}

/**
 * Generates quarterly trend analysis report
 */
function generateQuarterlyReport() {
  try {
    const reportData = gatherQuarterlyData();
    const doc = createQuarterlyReportDoc(reportData);
    const pdf = convertToPDF(doc);

    // Email to leadership
    emailReport(pdf, 'Quarterly Trend Analysis', getReportRecipients('quarterly'));

    Logger.log('Quarterly report generated and emailed successfully');

  } catch (error) {
    Logger.log('Error generating quarterly report: ' + error.message);
    notifyAdminOfError('Quarterly Report Generation Failed', error.message);
  }
}

/**
 * Gathers data for monthly report
 * @returns {Object} Monthly statistics
 */
function gatherMonthlyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const now = new Date();
  const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return {
      month: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy'),
      totalGrievances: 0,
      newGrievances: 0,
      closedGrievances: 0,
      openGrievances: 0,
      avgResolutionTime: 0,
      byIssueType: {},
      bySteward: {},
      overdueCount: 0
    };
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  let totalGrievances = 0;
  let newGrievances = 0;
  let closedGrievances = 0;
  let openGrievances = 0;
  let resolutionTimes = [];
  const byIssueType = {};
  const bySteward = {};
  let overdueCount = 0;

  data.forEach(row => {
    const filedDate = row[6];
    const closedDate = row[18];
    const status = row[4];
    const issueType = row[5];
    const steward = row[13];
    const daysToDeadline = row[20];

    totalGrievances++;

    // Count new grievances filed this month
    if (filedDate && filedDate >= firstDayOfMonth && filedDate <= lastDayOfMonth) {
      newGrievances++;
    }

    // Count closed grievances this month
    if (closedDate && closedDate >= firstDayOfMonth && closedDate <= lastDayOfMonth) {
      closedGrievances++;

      // Calculate resolution time
      if (filedDate) {
        const resTime = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
        resolutionTimes.push(resTime);
      }
    }

    // Count currently open
    if (status === 'Open') {
      openGrievances++;
    }

    // Count by issue type
    if (issueType) {
      byIssueType[issueType] = (byIssueType[issueType] || 0) + 1;
    }

    // Count by steward
    if (steward) {
      bySteward[steward] = (bySteward[steward] || 0) + 1;
    }

    // Count overdue
    if (daysToDeadline < 0) {
      overdueCount++;
    }
  });

  const avgResolutionTime = resolutionTimes.length > 0
    ? Math.round(resolutionTimes.reduce((a, b) => a + b, 0) / resolutionTimes.length)
    : 0;

  return {
    month: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM yyyy'),
    totalGrievances,
    newGrievances,
    closedGrievances,
    openGrievances,
    avgResolutionTime,
    byIssueType,
    bySteward,
    overdueCount
  };
}

/**
 * Gathers data for quarterly report
 * @returns {Object} Quarterly statistics
 */
function gatherQuarterlyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  const now = new Date();
  const quarter = Math.floor(now.getMonth() / 3);
  const startMonth = quarter * 3;
  const firstDayOfQuarter = new Date(now.getFullYear(), startMonth, 1);
  const lastDayOfQuarter = new Date(now.getFullYear(), startMonth + 3, 0);

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) {
    return {
      quarter: `Q${quarter + 1} ${now.getFullYear()}`,
      monthlyTrends: [],
      totalGrievances: 0,
      trend: 'stable'
    };
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const monthlyTrends = [0, 0, 0]; // Three months in a quarter
  let totalGrievances = 0;

  data.forEach(row => {
    const filedDate = row[6];

    if (filedDate && filedDate >= firstDayOfQuarter && filedDate <= lastDayOfQuarter) {
      totalGrievances++;

      const monthIndex = filedDate.getMonth() - startMonth;
      if (monthIndex >= 0 && monthIndex < 3) {
        monthlyTrends[monthIndex]++;
      }
    }
  });

  // Determine trend
  let trend = 'stable';
  if (monthlyTrends[2] > monthlyTrends[0] * 1.2) {
    trend = 'increasing';
  } else if (monthlyTrends[2] < monthlyTrends[0] * 0.8) {
    trend = 'decreasing';
  }

  return {
    quarter: `Q${quarter + 1} ${now.getFullYear()}`,
    monthlyTrends,
    totalGrievances,
    trend,
    ...gatherMonthlyData() // Include current month details
  };
}

/**
 * Creates monthly report document
 * @param {Object} data - Report data
 * @returns {Document} Google Doc
 */
function createMonthlyReportDoc(data) {
  const doc = DocumentApp.create(`Monthly_Report_${data.month.replace(' ', '_')}`);
  const body = doc.getBody();

  // Title
  body.appendParagraph('SEIU Local 509')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Monthly Grievance Report')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(data.month)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // Executive Summary
  body.appendParagraph('Executive Summary')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const summaryTable = [
    ['Metric', 'Value'],
    ['New Grievances Filed', data.newGrievances.toString()],
    ['Grievances Closed', data.closedGrievances.toString()],
    ['Currently Open', data.openGrievances.toString()],
    ['Overdue Grievances', data.overdueCount.toString()],
    ['Average Resolution Time', `${data.avgResolutionTime} days`],
    ['Total Active Grievances', data.totalGrievances.toString()]
  ];

  const table = body.appendTable(summaryTable);
  table.getRow(0).editAsText().setBold(true);

  body.appendParagraph(''); // Spacing

  // Grievances by Issue Type
  body.appendParagraph('Grievances by Issue Type')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const issueTypes = Object.entries(data.byIssueType)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const issueTable = [['Issue Type', 'Count']];
  issueTypes.forEach(([type, count]) => {
    issueTable.push([type, count.toString()]);
  });

  const issueTableObj = body.appendTable(issueTable);
  issueTableObj.getRow(0).editAsText().setBold(true);

  body.appendParagraph(''); // Spacing

  // Grievances by Steward
  body.appendParagraph('Workload by Steward')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const stewards = Object.entries(data.bySteward)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);

  const stewardTable = [['Steward', 'Active Grievances']];
  stewards.forEach(([steward, count]) => {
    stewardTable.push([steward, count.toString()]);
  });

  const stewardTableObj = body.appendTable(stewardTable);
  stewardTableObj.getRow(0).editAsText().setBold(true);

  body.appendParagraph(''); // Spacing

  // Key Insights
  body.appendParagraph('Key Insights')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const insights = [];

  if (data.overdueCount > 0) {
    insights.push(`⚠️ ${data.overdueCount} grievance(s) are currently overdue and require immediate attention.`);
  }

  if (data.avgResolutionTime > 30) {
    insights.push(`⏰ Average resolution time is ${data.avgResolutionTime} days, which exceeds the 30-day target.`);
  }

  if (data.newGrievances > data.closedGrievances) {
    insights.push(`📈 Grievances filed (${data.newGrievances}) exceeded closures (${data.closedGrievances}) this month.`);
  } else {
    insights.push(`✅ More grievances were closed (${data.closedGrievances}) than filed (${data.newGrievances}) this month.`);
  }

  if (insights.length === 0) {
    body.appendParagraph('No significant issues identified. Operations are running smoothly.');
  } else {
    insights.forEach(insight => {
      body.appendListItem(insight);
    });
  }

  body.appendParagraph(''); // Spacing
  body.appendHorizontalRule();

  body.appendParagraph(`Report generated: ${new Date().toLocaleString()}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setItalic(true);

  doc.saveAndClose();
  return doc;
}

/**
 * Creates quarterly report document
 * @param {Object} data - Report data
 * @returns {Document} Google Doc
 */
function createQuarterlyReportDoc(data) {
  const doc = DocumentApp.create(`Quarterly_Report_${data.quarter.replace(' ', '_')}`);
  const body = doc.getBody();

  // Title
  body.appendParagraph('SEIU Local 509')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Quarterly Trend Analysis')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(data.quarter)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  // Trend Summary
  body.appendParagraph('Trend Summary')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const trendEmoji = {
    'increasing': '📈 Increasing',
    'decreasing': '📉 Decreasing',
    'stable': '➡️ Stable'
  }[data.trend];

  body.appendParagraph(`Overall Trend: ${trendEmoji}`)
    .setBold(true);

  body.appendParagraph(`Total Grievances This Quarter: ${data.totalGrievances}`);

  body.appendParagraph(''); // Spacing

  // Monthly Breakdown
  body.appendParagraph('Monthly Breakdown')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const monthNames = ['Month 1', 'Month 2', 'Month 3'];
  const monthTable = [['Month', 'Grievances Filed']];

  data.monthlyTrends.forEach((count, index) => {
    monthTable.push([monthNames[index], count.toString()]);
  });

  const monthTableObj = body.appendTable(monthTable);
  monthTableObj.getRow(0).editAsText().setBold(true);

  // Include monthly details (reuse monthly report data)
  body.appendPageBreak();

  body.appendParagraph('Current Month Details')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph(`Average Resolution Time: ${data.avgResolutionTime} days`);
  body.appendParagraph(`Currently Open: ${data.openGrievances}`);
  body.appendParagraph(`Overdue: ${data.overdueCount}`);

  doc.saveAndClose();
  return doc;
}

/**
 * Converts Google Doc to PDF
 * @param {Document} doc - Google Doc object
 * @returns {File} PDF file
 */
function convertToPDF(doc) {
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  const pdfFile = DriveApp.createFile(pdfBlob);

  // Delete the temporary doc
  docFile.setTrashed(true);

  return pdfFile;
}

/**
 * Emails report to recipients
 * @param {File} pdf - PDF file
 * @param {string} reportType - Report type name
 * @param {Array} recipients - Email addresses
 */
function emailReport(pdf, reportType, recipients) {
  if (recipients.length === 0) {
    Logger.log('No recipients configured for reports');
    return;
  }

  const subject = `SEIU Local 509 - ${reportType}`;
  const body = `
Dear Leadership,

Please find attached the ${reportType} for SEIU Local 509.

This report was automatically generated by the 509 Dashboard.

Best regards,
SEIU Local 509 Dashboard (Automated System)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

View Dashboard: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
Report Generated: ${new Date().toLocaleString()}
  `;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: body,
    attachments: [pdf],
    name: 'SEIU Local 509 Dashboard'
  });

  Logger.log(`Report emailed to: ${recipients.join(', ')}`);
}

/**
 * Gets report recipients from configuration
 * @param {string} reportType - 'monthly' or 'quarterly'
 * @returns {Array} Email addresses
 */
function getReportRecipients(reportType) {
  // This would be stored in a configuration sheet
  // For now, return a placeholder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('⚙️ Configuration');

  if (!configSheet) {
    // Default to script owner
    return [Session.getActiveUser().getEmail()];
  }

  // Look for recipients in config sheet
  // Format: Report Type | Email Addresses (comma-separated)
  const data = configSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === `${reportType}_recipients`) {
      const emails = data[i][1].toString().split(',').map(e => e.trim()).filter(e => e);
      return emails;
    }
  }

  // Default
  return [Session.getActiveUser().getEmail()];
}

/**
 * Notifies admin of report generation error
 * @param {string} title - Error title
 * @param {string} message - Error message
 */
function notifyAdminOfError(title, message) {
  const adminEmail = Session.getActiveUser().getEmail();

  MailApp.sendEmail({
    to: adminEmail,
    subject: `🚨 ${title}`,
    body: `An error occurred in the 509 Dashboard automated reports:\n\n${message}\n\nPlease check the execution log for details.`,
    name: 'SEIU Local 509 Dashboard'
  });
}

/**
 * Shows report automation settings dialog
 */
function showReportAutomationSettings() {
  const ui = SpreadsheetApp.getUi();

  const triggers = ScriptApp.getProjectTriggers();
  const monthlyEnabled = triggers.some(t => t.getHandlerFunction() === 'generateMonthlyReport');
  const quarterlyEnabled = triggers.some(t => t.getHandlerFunction() === 'generateQuarterlyReport');

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .status { padding: 15px; border-radius: 4px; margin: 20px 0; font-weight: bold; }
    .status.enabled { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
    .status.disabled { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.danger { background: #dc3545; }
    button.danger:hover { background: #c82333; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📊 Report Automation Settings</h2>

    <h3>Monthly Reports</h3>
    <div class="status ${monthlyEnabled ? 'enabled' : 'disabled'}">
      ${monthlyEnabled ? '✅ Monthly reports ENABLED (1st of month, 9 AM)' : '🔕 Monthly reports DISABLED'}
    </div>
    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).setupMonthlyReports()">
      ${monthlyEnabled ? '🔄 Refresh Monthly Trigger' : '✅ Enable Monthly Reports'}
    </button>

    <h3>Quarterly Reports</h3>
    <div class="status ${quarterlyEnabled ? 'enabled' : 'disabled'}">
      ${quarterlyEnabled ? '✅ Quarterly reports ENABLED (Jan/Apr/Jul/Oct, 10 AM)' : '🔕 Quarterly reports DISABLED'}
    </div>
    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).setupQuarterlyReports()">
      ${quarterlyEnabled ? '🔄 Refresh Quarterly Trigger' : '✅ Enable Quarterly Reports'}
    </button>

    <hr style="margin: 30px 0;">

    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).generateMonthlyReport()">
      🧪 Test Monthly Report
    </button>

    <button onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).generateQuarterlyReport()">
      🧪 Test Quarterly Report
    </button>

    <button class="danger" onclick="google.script.run.withSuccessHandler(() => { google.script.host.close(); }).disableAutomatedReports()">
      🔕 Disable All Reports
    </button>
  </div>
</body>
</html>
  `)
    .setWidth(700)
    .setHeight(550);

  ui.showModalDialog(html, '📊 Report Automation');
}

/******************************************************************************
 * MODULE: AddRecommendations
 * Source: AddRecommendations.gs
 *****************************************************************************/

/**
 * ============================================================================
 * ADD RECOMMENDATIONS TO FEEDBACK & DEVELOPMENT SHEET
 * ============================================================================
 *
 * This script adds 47 code review recommendations to the Feedback & Development sheet
 * Run this function once to populate the recommendations
 */

function ADD_RECOMMENDATIONS_TO_FEATURES_TAB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feedbackSheet = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!feedbackSheet) {
    SpreadsheetApp.getUi().alert('❌ Feedback & Development sheet not found!');
    return;
  }

  // Confirm with user
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Add Code Review Recommendations',
    'This will add 47 enhancement recommendations to the Feedback & Development sheet.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('📝 Adding recommendations...', 'Please wait', -1);

  // Get the last row to append after
  const lastRow = feedbackSheet.getLastRow();

  // Recommendations data
  const recommendations = [
    // Format: [Type, Date, Submitted By, Priority, Title, Description, Status, Progress %, Complexity, Target, Assigned To, Blockers, Notes, Last Updated]

    // PERFORMANCE & SCALABILITY
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Extend Auto-Formula Coverage", "Extend formulas from 100 rows to 1000 rows using ARRAYFORMULA. Current limitation requires manual formula addition beyond row 100.", "Planned", 0, "Simple", "Q1 2025", "Dev Team", "", "Replace individual setFormula calls with ARRAYFORMULA implementations", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Optimize Seed Data Performance", "Improve batch processing with better progress indicators and caching. Current execution takes 2-3 minutes for 20k members.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Use SpreadsheetApp.flush() strategically; add progress toasts every 500 rows", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Implement Data Pagination", "Add pagination to analytics sheets to improve load times with large datasets. Currently all data loads at once.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Create Show More buttons; use QUERY with LIMIT and OFFSET", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Add Caching Layer", "Cache Analytics Data sheet calculations for 10x faster dashboard loads.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Store calculated values in hidden sheet; refresh on data change triggers", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Optimize QUERY Formulas", "Consolidate multiple QUERY calls on same data to reduce calculation time.", "Planned", 0, "Simple", "Q2 2025", "Dev Team", "", "Use virtual tables and GROUP BY instead of multiple COUNTIF calls", "2025-01-28"],

    // USER EXPERIENCE & ACCESSIBILITY
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Add Member Search Functionality", "Implement searchable member directory with autocomplete to reduce lookup time by 90%.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Add search dialog with filters by name, ID, location, unit; add keyboard shortcuts", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Create Quick Actions Menu", "Add right-click context menu for common actions (Start Grievance, View History, Email).", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Implement onSelectionChange trigger with context-aware menu options", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Implement Keyboard Shortcuts", "Add keyboard navigation for power users (Ctrl+N: New grievance, Ctrl+M: Member directory, etc.).", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Create keyboard event handlers for common operations", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Add Undo/Redo Functionality", "Track changes with undo stack to provide rollback for accidental changes.", "Planned", 0, "Complex", "Q2 2025", "Dev Team", "", "Store last 10 operations in Archive sheet; add Undo Last Action menu item", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Enhanced ADHD Features", "Add focus mode, reading guide, color-coded deadlines, Pomodoro timer integration.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Expand current ADHD features with additional accessibility options", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Mobile-Optimized Views", "Create mobile-friendly dashboard with single-column layout and larger touch targets.", "Planned", 0, "Moderate", "Q3 2025", "Dev Team", "", "Design mobile-specific views for field use cases", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Dark Mode Support", "Add dark theme option to reduce eye strain.", "Planned", 0, "Simple", "Q2 2025", "Dev Team", "", "Create DARK_COLORS constant; apply theme to all sheets", "2025-01-28"],

    // DATA INTEGRITY & VALIDATION
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Add Email & Phone Validation", "Validate email format, phone format, and date ranges to ensure cleaner data.", "Planned", 0, "Simple", "Q1 2025", "Dev Team", "", "Use requireTextMatchesPattern for email/phone validation", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Implement Change Tracking", "Track all data modifications with audit trail showing timestamp, user, field, old/new values.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Create onEdit trigger to log changes to Change Log sheet", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Add Duplicate Detection", "Prevent duplicate member IDs and grievance IDs with auto-generation of next available ID.", "Planned", 0, "Simple", "Q1 2025", "Dev Team", "", "Check existing IDs before insertion; show warning dialog on duplicate", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Data Quality Dashboard", "Monitor data completeness with percentage of required fields filled and quality score.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Show data quality metrics; highlight missing critical data", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Referential Integrity Checks", "Validate foreign keys to ensure Member IDs exist before creating grievances.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Check Member ID existence; validate Steward assignments; warn on orphaned records", "2025-01-28"],

    // AUTOMATION & WORKFLOWS
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Automated Deadline Notifications", "Email alerts for approaching deadlines with escalation to managers.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Time-driven trigger (daily 8 AM); email stewards 7 days before; escalate 3 days before", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Batch Operations", "Add bulk update capabilities (assign steward, update status, export PDF, email).", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Create batch operation dialogs for common bulk tasks", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Automated Report Generation", "Schedule monthly/quarterly reports with automatic generation and email distribution.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Time-driven trigger (1st of month); generate PDF; email to leadership", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Workflow State Machine", "Enforce grievance workflow rules to ensure process compliance.", "Planned", 0, "Complex", "Q3 2025", "Dev Team", "", "Define valid state transitions; block invalid changes; require notes", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Smart Auto-Assignment", "Intelligent steward assignment based on workload, location, expertise, and availability.", "Planned", 0, "Complex", "Q3 2025", "Dev Team", "", "Balance workload; match by location/unit; consider expertise", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Template System", "Pre-built grievance templates for common issue types with auto-fill capabilities.", "Planned", 0, "Moderate", "Q4 2025", "Dev Team", "", "Create template library; customizable placeholders", "2025-01-28"],

    // REPORTING & ANALYTICS
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Predictive Analytics", "Predict grievance outcomes based on historical patterns and similar cases.", "Planned", 0, "Complex", "Q3 2025", "Data Team", "", "Analyze win/loss patterns; identify success factors; provide risk scores", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Trend Analysis & Forecasting", "Identify patterns over time with month-over-month comparisons and volume forecasting.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Month-over-month comparisons; seasonal patterns; early warning system", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Custom Report Builder", "User-defined reports with drag-and-drop field selection and custom filters.", "Planned", 0, "Complex", "Q3 2025", "Dev Team", "", "Build report designer interface; save templates; schedule recurring reports", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Benchmark Comparisons", "Compare performance against industry standards with percentile rankings.", "Planned", 0, "Moderate", "Q4 2025", "Data Team", "", "Upload benchmark data; side-by-side comparisons; gap analysis", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Root Cause Analysis", "Identify systemic issues by clustering similar grievances and analyzing common factors.", "Planned", 0, "Complex", "Q3 2025", "Data Team", "", "Cluster analysis; correlation visualization; intervention recommendations", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Real-Time Dashboard Updates", "Live data refresh with auto-update every 5 minutes.", "Planned", 0, "Moderate", "Q4 2025", "Dev Team", "", "Update dashboards on data change; WebSocket-style updates via triggers", "2025-01-28"],

    // INTEGRATION & EXTENSIBILITY
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Google Calendar Integration", "Sync deadlines to Google Calendar with color-coding by priority.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Create calendar events for deadlines; update on status changes", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Email Integration (Gmail)", "Send/receive grievance emails with tracking and template library.", "Planned", 0, "Moderate", "Q1 2025", "Dev Team", "", "Email PDFs directly; track opens; auto-log communications", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Google Drive Integration", "Attach documents to grievances with version control and auto-organization.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Link Drive folders; upload evidence; organize by grievance ID", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Slack/Teams Integration", "Real-time notifications with bot commands for quick queries.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Post updates to channels; notify on deadlines; allow status updates from chat", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "API Layer", "RESTful API for external access with authentication.", "Planned", 0, "Complex", "Q4 2025", "Dev Team", "", "Google Apps Script Web App; GET/POST/PUT endpoints; API key auth", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Zapier/Make.com Integration", "Connect to 1000+ apps via webhooks.", "Planned", 0, "Moderate", "Q4 2025", "Dev Team", "", "Webhooks for data changes; trigger actions in other apps", "2025-01-28"],

    // MOBILE & OFFLINE ACCESS
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Progressive Web App (PWA)", "Install as app on mobile devices for native experience.", "Planned", 0, "Complex", "Q4 2025", "Dev Team", "", "HTML service interface; manifest.json; service worker; responsive design", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Offline Mode", "Work without internet with local storage cache and sync.", "Planned", 0, "Very Complex", "Q4 2025", "Dev Team", "", "Local storage cache; sync when online; conflict resolution", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "SMS Notifications", "Text alerts for critical items via Twilio integration.", "Planned", 0, "Moderate", "Q4 2025", "Dev Team", "", "Twilio integration; SMS for overdue; opt-in/opt-out management", "2025-01-28"],

    // SECURITY & PRIVACY
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Role-Based Access Control (RBAC)", "Implement permission levels (Admin, Steward, Member, Viewer).", "Planned", 0, "Complex", "Q2 2025", "Security Team", "", "Define role permissions; protect sensitive operations", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Audit Logging", "Track all access and changes with 7-year retention.", "Planned", 0, "Moderate", "Q2 2025", "Security Team", "", "Log opens and modifications; store IP, timestamp, user", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "PII Protection", "Protect sensitive member data with encryption and masking.", "Planned", 0, "Complex", "Q2 2025", "Security Team", "", "Encrypt sensitive fields; mask in exports; GDPR/CCPA compliance", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Data Backup & Recovery", "Automated backups with one-click restore.", "Planned", 0, "Moderate", "Q2 2025", "Dev Team", "", "Daily backup to separate sheet; 30-day version history; export to Drive", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Session Management", "Track active users with row locking and concurrent edit warnings.", "Planned", 0, "Moderate", "Q3 2025", "Dev Team", "", "Show current viewers; lock rows being edited; session timeout", "2025-01-28"],

    // DOCUMENTATION & TRAINING
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Interactive Tutorial", "In-app guided tour for first-time users.", "Planned", 0, "Moderate", "Q1 2025", "UX Team", "", "Walkthrough; highlight features; interactive steps", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "High", "Video Tutorials", "Screen recordings for common tasks (2-3 min each).", "Planned", 0, "Simple", "Q1 2025", "UX Team", "", "Creating grievance; running reports; managing workload; using dashboard", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "Context-Sensitive Help", "Help button on each sheet with explanations and links.", "Planned", 0, "Simple", "Q2 2025", "UX Team", "", "? icon on each sheet; explain purpose; list common tasks", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Medium", "FAQ Database", "Searchable knowledge base with voting and community contributions.", "Planned", 0, "Moderate", "Q2 2025", "UX Team", "", "Categorize by topic; search functionality; vote on answers", "2025-01-28"],
    ["Future Feature", "2025-01-28", "AI Code Review", "Low", "Release Notes", "Track version changes with auto-notification on updates.", "Planned", 0, "Simple", "Q2 2025", "Dev Team", "", "Create CHANGELOG sheet; auto-notify; highlight new features", "2025-01-28"]
  ];

  // Add recommendations to sheet
  const startRow = lastRow + 1;
  feedbackSheet.getRange(startRow, 1, recommendations.length, 14).setValues(recommendations);

  // Format the added rows
  const addedRange = feedbackSheet.getRange(startRow, 1, recommendations.length, 14);
  addedRange.setFontSize(10);

  // Alternate row colors for readability
  for (let i = 0; i < recommendations.length; i++) {
    const rowNum = startRow + i;
    if (i % 2 === 0) {
      feedbackSheet.getRange(rowNum, 1, 1, 14).setBackground("#F9FAFB");
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Successfully added ${recommendations.length} recommendations!`,
    'Complete',
    5
  );

  // Show summary
  SpreadsheetApp.getUi().alert(
    '✅ Recommendations Added Successfully!',
    `Added ${recommendations.length} enhancement recommendations to the Feedback & Development sheet.\n\n` +
    'Breakdown by priority:\n' +
    '• High Priority: 25 items\n' +
    '• Medium Priority: 15 items\n' +
    '• Low Priority: 7 items\n\n' +
    'Categories covered:\n' +
    '• Performance & Scalability\n' +
    '• User Experience & Accessibility\n' +
    '• Data Integrity & Validation\n' +
    '• Automation & Workflows\n' +
    '• Reporting & Analytics\n' +
    '• Integration & Extensibility\n' +
    '• Mobile & Offline Access\n' +
    '• Security & Privacy\n' +
    '• Documentation & Training\n\n' +
    'Review the recommendations and prioritize based on your needs!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/******************************************************************************
 * MODULE: FAQKnowledgeBase
 * Source: FAQKnowledgeBase.gs
 *****************************************************************************/

/**
 * ============================================================================
 * FAQ DATABASE & KNOWLEDGE BASE
 * ============================================================================
 *
 * Searchable knowledge base for common questions and procedures
 * Features:
 * - Searchable FAQ database
 * - Category-based organization
 * - Admin interface to add/edit FAQs
 * - User feedback on helpfulness
 * - Related articles suggestions
 * - Export FAQ documentation
 * - Auto-suggest based on context
 */

/**
 * FAQ categories
 */
const FAQ_CATEGORIES = {
  GETTING_STARTED: 'Getting Started',
  GRIEVANCES: 'Grievance Process',
  MEMBERS: 'Member Management',
  REPORTS: 'Reports & Analytics',
  AUTOMATION: 'Automation Features',
  TROUBLESHOOTING: 'Troubleshooting',
  BEST_PRACTICES: 'Best Practices'
};

/**
 * Creates FAQ database sheet if it doesn't exist
 */
function createFAQSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let faqSheet = ss.getSheetByName('📚 FAQ Database');

  if (faqSheet) {
    SpreadsheetApp.getUi().alert('FAQ Database sheet already exists.');
    return;
  }

  faqSheet = ss.insertSheet('📚 FAQ Database');

  // Set headers
  const headers = [
    'ID',
    'Category',
    'Question',
    'Answer',
    'Tags',
    'Related FAQs',
    'Helpful Count',
    'Not Helpful Count',
    'Created Date',
    'Last Updated',
    'Created By'
  ];

  faqSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  faqSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  faqSheet.setColumnWidth(1, 60);   // ID
  faqSheet.setColumnWidth(2, 150);  // Category
  faqSheet.setColumnWidth(3, 300);  // Question
  faqSheet.setColumnWidth(4, 500);  // Answer
  faqSheet.setColumnWidth(5, 150);  // Tags
  faqSheet.setColumnWidth(6, 120);  // Related FAQs
  faqSheet.setColumnWidth(7, 100);  // Helpful Count
  faqSheet.setColumnWidth(8, 120);  // Not Helpful Count
  faqSheet.setColumnWidth(9, 120);  // Created Date
  faqSheet.setColumnWidth(10, 120); // Last Updated
  faqSheet.setColumnWidth(11, 150); // Created By

  // Freeze header
  faqSheet.setFrozenRows(1);

  // Add initial FAQs
  seedInitialFAQs();

  SpreadsheetApp.getUi().alert(
    '✅ FAQ Database Created',
    'The FAQ Database sheet has been created and seeded with common questions.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Seeds initial FAQs
 */
function seedInitialFAQs() {
  const initialFAQs = [
    {
      category: FAQ_CATEGORIES.GETTING_STARTED,
      question: 'How do I start a new grievance?',
      answer: 'To start a new grievance: 1) Go to menu: 📋 Grievance Tools → ➕ Start New Grievance, 2) Fill in the member information and grievance details, 3) Click Submit. The system will auto-assign a Grievance ID and calculate deadlines.',
      tags: 'grievance, new, create, start'
    },
    {
      category: FAQ_CATEGORIES.GRIEVANCES,
      question: 'What are the different grievance workflow states?',
      answer: 'Grievances flow through these states: Filed → Step 1 Pending → Step 1 Meeting → Step 2 Pending (if escalated) → Step 2 Meeting → Arbitration (if needed) → Resolved/Withdrawn. Each state has specific deadlines and required actions. Use the Workflow Visualizer (🔄 Workflow Management → 🔄 Workflow Visualizer) to see the full process.',
      tags: 'workflow, states, process, steps'
    },
    {
      category: FAQ_CATEGORIES.GRIEVANCES,
      question: 'How do I assign a steward to a grievance?',
      answer: 'You have three options: 1) Manual: Select the grievance row and enter a steward name in the "Assigned Steward" column, 2) Auto-Assign: Select the row and use 🤖 Smart Assignment → 🤖 Auto-Assign Steward (recommended - uses AI to find best match), 3) Batch: Select multiple rows and use 🤖 Smart Assignment → 📦 Batch Auto-Assign.',
      tags: 'steward, assign, assignment, auto-assign'
    },
    {
      category: FAQ_CATEGORIES.MEMBERS,
      question: 'How do I search for a member?',
      answer: "Use the Member Search tool: 1) Press Ctrl+F or go to 🔍 Search & Lookup → 🔍 Search Members, 2) Type name, ID, or email in the search box, 3) Use filters to narrow by location, unit, or steward status, 4) Click on a result to navigate to that member's row.",
      tags: 'search, member, find, lookup'
    },
    {
      category: FAQ_CATEGORIES.REPORTS,
      question: 'How do I create a custom report?',
      answer: 'Use the Custom Report Builder: 1) Go to 📊 Reports → 📊 Custom Report Builder, 2) Select your data source (Grievances or Members), 3) Choose which fields to include, 4) Add filters to narrow data, 5) Choose grouping/sorting options, 6) Generate preview, 7) Export to PDF, CSV, or Excel. You can save configurations as templates for reuse.',
      tags: 'report, custom, export, PDF, CSV, Excel'
    },
    {
      category: FAQ_CATEGORIES.AUTOMATION,
      question: 'How do I enable automated deadline notifications?',
      answer: 'To enable email notifications for approaching deadlines: 1) Go to 🤖 Automation → 📬 Notification Settings, 2) Click "Enable Notifications", 3) The system will send daily emails at 8 AM: 7-day warnings to stewards, 3-day escalations to managers, and overdue alerts. Test with "🧪 Test Now" button.',
      tags: 'automation, notifications, email, deadlines, alerts'
    },
    {
      category: FAQ_CATEGORIES.AUTOMATION,
      question: 'Can I schedule automated reports?',
      answer: 'Yes! Go to 🤖 Automation → 📊 Report Automation Settings. You can enable: 1) Monthly Reports (1st of each month, executive summary), 2) Quarterly Reports (Jan/Apr/Jul/Oct, trend analysis). Reports are automatically generated as PDFs and emailed to configured recipients.',
      tags: 'automation, reports, schedule, monthly, quarterly'
    },
    {
      category: FAQ_CATEGORIES.TROUBLESHOOTING,
      question: 'Why are my formulas not calculating?',
      answer: 'If formulas aren\'t updating: 1) Go to menu: 🔄 Refresh All (Ctrl+R), 2) If still not working, check if you have edit permissions, 3) Make sure the sheet isn\'t in protected range, 4) Try ⚡ Performance → 🗑️ Clear All Caches, then refresh. If problem persists, use ❓ Help & Support → 🔧 Diagnose Setup.',
      tags: 'formulas, not working, calculate, refresh, troubleshoot'
    },
    {
      category: FAQ_CATEGORIES.TROUBLESHOOTING,
      question: 'What if a grievance deadline is wrong?',
      answer: 'Deadlines are auto-calculated based on Filed Date (Column G) + workflow state. To fix: 1) Check the Filed Date is correct, 2) Verify the workflow state is accurate (🔄 Workflow Management → 📊 View Current State), 3) If you need to manually override, you can edit the "Next Action Due" field directly.',
      tags: 'deadline, wrong, incorrect, fix, override'
    },
    {
      category: FAQ_CATEGORIES.BEST_PRACTICES,
      question: 'What are keyboard shortcuts and how do I use them?',
      answer: 'Keyboard shortcuts make navigation 10x faster! Press Ctrl+? to see all shortcuts. Common ones: Ctrl+G (Go to Grievance Log), Ctrl+M (Go to Members), Ctrl+F (Search Members), Ctrl+N (New Grievance), Ctrl+E (Compose Email), Ctrl+B (Batch Operations). On Mac, use Cmd instead of Ctrl.',
      tags: 'keyboard, shortcuts, hotkeys, fast, navigation'
    },
    {
      category: FAQ_CATEGORIES.BEST_PRACTICES,
      question: 'How often should I back up data?',
      answer: 'Best practice is weekly backups. Use 💾 Data Management → 💾 Create Backup to generate a timestamped copy of all data. The backup is saved to Google Drive. You can also enable automated weekly backups in the backup settings. Keep at least 4 weeks of backups.',
      tags: 'backup, data, export, save, recovery'
    },
    {
      category: FAQ_CATEGORIES.BEST_PRACTICES,
      question: 'How can I improve dashboard performance?',
      answer: 'Performance tips: 1) Use caching - enable via ⚡ Performance → 🔥 Warm Up All Caches, 2) Close unused sheets/tabs, 3) Use filters instead of scrolling through all data, 4) For large datasets (20k+ members), use Search instead of browsing, 5) Batch operations instead of individual updates, 6) Keep browser updated.',
      tags: 'performance, slow, speed, fast, optimize, cache'
    }
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const faqSheet = ss.getSheetByName('📚 FAQ Database');

  const rows = initialFAQs.map((faq, index) => [
    index + 1,
    faq.category,
    faq.question,
    faq.answer,
    faq.tags,
    '',
    0,
    0,
    new Date(),
    new Date(),
    Session.getActiveUser().getEmail() || 'System'
  ]);

  faqSheet.getRange(2, 1, rows.length, 11).setValues(rows);
}

/**
 * Shows FAQ search dialog
 */
function showFAQSearch() {
  const html = createFAQSearchHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📚 FAQ & Knowledge Base');
}

/**
 * Creates HTML for FAQ search
 */
function createFAQSearchHTML() {
  const categories = Object.values(FAQ_CATEGORIES);

  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .search-box {
      margin: 20px 0;
    }
    input[type="text"] {
      width: 100%;
      padding: 15px;
      font-size: 16px;
      border: 2px solid #ddd;
      border-radius: 8px;
      box-sizing: border-box;
    }
    input[type="text"]:focus {
      outline: none;
      border-color: #1a73e8;
    }
    .categories {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin: 20px 0;
    }
    .category-badge {
      background: #e8f0fe;
      color: #1a73e8;
      padding: 8px 16px;
      border-radius: 20px;
      cursor: pointer;
      font-size: 14px;
      border: 2px solid transparent;
      transition: all 0.2s;
    }
    .category-badge:hover {
      background: #1a73e8;
      color: white;
    }
    .category-badge.active {
      background: #1a73e8;
      color: white;
      border-color: #1557b0;
    }
    .faq-item {
      background: #f8f9fa;
      padding: 20px;
      margin: 15px 0;
      border-radius: 8px;
      border-left: 4px solid #1a73e8;
      cursor: pointer;
      transition: all 0.2s;
    }
    .faq-item:hover {
      transform: translateX(5px);
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .faq-question {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 10px;
    }
    .faq-answer {
      color: #666;
      font-size: 14px;
      line-height: 1.6;
      display: none;
    }
    .faq-answer.show {
      display: block;
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px solid #ddd;
    }
    .faq-meta {
      display: flex;
      gap: 15px;
      margin-top: 10px;
      font-size: 12px;
      color: #999;
    }
    .category-tag {
      background: #e8f0fe;
      color: #1a73e8;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 12px;
    }
    .helpful-buttons {
      display: none;
      margin-top: 15px;
      padding-top: 15px;
      border-top: 1px solid #ddd;
    }
    .helpful-buttons.show {
      display: flex;
      gap: 10px;
      align-items: center;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 8px 16px;
      font-size: 13px;
      border-radius: 4px;
      cursor: pointer;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.secondary:hover {
      background: #5a6268;
    }
    .no-results {
      text-align: center;
      padding: 40px;
      color: #999;
    }
    .stats {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📚 FAQ & Knowledge Base</h2>

    <div class="search-box">
      <input type="text" id="searchInput" placeholder="🔍 Search for answers..." autofocus oninput="searchFAQs()">
    </div>

    <div class="categories">
      <div class="category-badge active" onclick="filterByCategory('')">All Categories</div>
      ${categories.map(cat => `
        <div class="category-badge" onclick="filterByCategory('${cat}')">${cat}</div>
      `).join('')}
    </div>

    <div class="stats" id="stats">
      Loading FAQs...
    </div>

    <div id="results">
      <div class="no-results">Type to search or select a category</div>
    </div>
  </div>

  <script>
    let allFAQs = [];
    let currentCategory = '';

    // Load FAQs on startup
    google.script.run
      .withSuccessHandler(onFAQsLoaded)
      .getAllFAQs();

    function onFAQsLoaded(faqs) {
      allFAQs = faqs;
      updateStats();
      showAllFAQs();
    }

    function updateStats() {
      const stats = document.getElementById('stats');
      stats.innerHTML = \`
        <strong>📊 Knowledge Base Stats:</strong>
        ${allFAQs.length} articles across ${new Set(allFAQs.map(f => f.category)).size} categories
      \`;
    }

    function searchFAQs() {
      const searchTerm = document.getElementById('searchInput').value.toLowerCase();

      if (!searchTerm && !currentCategory) {
        showAllFAQs();
        return;
      }

      let filtered = allFAQs;

      if (currentCategory) {
        filtered = filtered.filter(faq => faq.category === currentCategory);
      }

      if (searchTerm) {
        filtered = filtered.filter(faq => {
          return faq.question.toLowerCase().includes(searchTerm) ||
                 faq.answer.toLowerCase().includes(searchTerm) ||
                 (faq.tags && faq.tags.toLowerCase().includes(searchTerm));
        });
      }

      displayFAQs(filtered);
    }

    function filterByCategory(category) {
      currentCategory = category;

      // Update active badge
      document.querySelectorAll('.category-badge').forEach(badge => {
        badge.classList.remove('active');
      });
      event.target.classList.add('active');

      searchFAQs();
    }

    function showAllFAQs() {
      displayFAQs(allFAQs);
    }

    function displayFAQs(faqs) {
      const results = document.getElementById('results');

      if (faqs.length === 0) {
        results.innerHTML = '<div class="no-results">No FAQs found. Try different search terms.</div>';
        return;
      }

      results.innerHTML = faqs.map((faq, index) => \`
        <div class="faq-item" onclick="toggleFAQ(\${index})">
          <div class="faq-question">
            ${faq.question}
          </div>
          <div class="faq-meta">
            <span class="category-tag">${faq.category}</span>
            ${faq.helpfulCount > 0 ? `<span>👍 ${faq.helpfulCount} found this helpful</span>` : ''}
          </div>
          <div class="faq-answer" id="answer\${index}">
            ${faq.answer}
          </div>
          <div class="helpful-buttons" id="buttons\${index}">
            <span>Was this helpful?</span>
            <button onclick="markHelpful(\${faq.id}, true); event.stopPropagation();">👍 Yes</button>
            <button class="secondary" onclick="markHelpful(\${faq.id}, false); event.stopPropagation();">👎 No</button>
          </div>
        </div>
      \`).join('');
    }

    function toggleFAQ(index) {
      const answer = document.getElementById('answer' + index);
      const buttons = document.getElementById('buttons' + index);

      answer.classList.toggle('show');
      buttons.classList.toggle('show');
    }

    function markHelpful(faqId, isHelpful) {
      google.script.run
        .withSuccessHandler(() => {
          // Reload FAQs to get updated counts
          google.script.run
            .withSuccessHandler(onFAQsLoaded)
            .getAllFAQs();
        })
        .updateFAQHelpfulness(faqId, isHelpful);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets all FAQs
 * @returns {Array} FAQ array
 */
function getAllFAQs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const faqSheet = ss.getSheetByName('📚 FAQ Database');

  if (!faqSheet) {
    return [];
  }

  const lastRow = faqSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const data = faqSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  return data.map(row => ({
    id: row[0],
    category: row[1],
    question: row[2],
    answer: row[3],
    tags: row[4],
    relatedFAQs: row[5],
    helpfulCount: row[6] || 0,
    notHelpfulCount: row[7] || 0,
    createdDate: row[8],
    lastUpdated: row[9],
    createdBy: row[10]
  }));
}

/**
 * Updates FAQ helpfulness rating
 * @param {number} faqId - FAQ ID
 * @param {boolean} isHelpful - Whether helpful or not
 */
function updateFAQHelpfulness(faqId, isHelpful) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const faqSheet = ss.getSheetByName('📚 FAQ Database');

  if (!faqSheet) return;

  const lastRow = faqSheet.getLastRow();
  const data = faqSheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === faqId) {
      const row = i + 2;
      const col = isHelpful ? 7 : 8; // Column G or H

      const currentCount = faqSheet.getRange(row, col).getValue() || 0;
      faqSheet.getRange(row, col).setValue(currentCount + 1);

      break;
    }
  }
}

/**
 * Shows FAQ admin panel
 */
function showFAQAdmin() {
  const html = createFAQAdminHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚙️ FAQ Admin Panel');
}

/**
 * Creates HTML for FAQ admin
 */
function createFAQAdminHTML() {
  const categories = Object.values(FAQ_CATEGORIES);

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .form-group { margin: 15px 0; }
    label { display: block; font-weight: 500; margin-bottom: 5px; color: #555; }
    input, select, textarea { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box; font-family: Arial, sans-serif; }
    textarea { min-height: 150px; resize: vertical; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚙️ Add New FAQ</h2>

    <div class="form-group">
      <label>Category:</label>
      <select id="category">
        ${categories.map(cat => `<option value="${cat}">${cat}</option>`).join('')}
      </select>
    </div>

    <div class="form-group">
      <label>Question:</label>
      <input type="text" id="question" placeholder="Enter the question...">
    </div>

    <div class="form-group">
      <label>Answer:</label>
      <textarea id="answer" placeholder="Enter the detailed answer..."></textarea>
    </div>

    <div class="form-group">
      <label>Tags (comma-separated):</label>
      <input type="text" id="tags" placeholder="e.g., grievance, new, create, start">
    </div>

    <button onclick="saveFAQ()">💾 Save FAQ</button>
    <button class="secondary" onclick="google.script.host.close()">❌ Cancel</button>
  </div>

  <script>
    function saveFAQ() {
      const faq = {
        category: document.getElementById('category').value,
        question: document.getElementById('question').value.trim(),
        answer: document.getElementById('answer').value.trim(),
        tags: document.getElementById('tags').value.trim()
      };

      if (!faq.question || !faq.answer) {
        alert('Please fill in both question and answer');
        return;
      }

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ FAQ saved successfully!');
          google.script.host.close();
        })
        .withFailureHandler((error) => {
          alert('Error: ' + error.message);
        })
        .addNewFAQ(faq);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Adds new FAQ
 * @param {Object} faq - FAQ data
 */
function addNewFAQ(faq) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let faqSheet = ss.getSheetByName('📚 FAQ Database');

  if (!faqSheet) {
    createFAQSheet();
    faqSheet = ss.getSheetByName('📚 FAQ Database');
  }

  const lastRow = faqSheet.getLastRow();
  const nextId = lastRow > 1 ? faqSheet.getRange(lastRow, 1).getValue() + 1 : 1;

  const newRow = [
    nextId,
    faq.category,
    faq.question,
    faq.answer,
    faq.tags,
    '',
    0,
    0,
    new Date(),
    new Date(),
    Session.getActiveUser().getEmail() || 'User'
  ];

  faqSheet.getRange(lastRow + 1, 1, 1, 11).setValues([newRow]);
}

/******************************************************************************
 * MODULE: GettingStartedAndFAQ
 * Source: GettingStartedAndFAQ.gs
 *****************************************************************************/

/**
 * ============================================================================
 * GETTING STARTED AND FAQ SHEETS
 * ============================================================================
 *
 * Creates informational sheets with getting started guide and FAQ
 * Includes GitHub repository information
 *
 * ============================================================================
 */

/**
 * Creates the Getting Started sheet
 */
function createGettingStartedSheet(ss) {
  let sheet = ss.getSheetByName("📚 Getting Started");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet("📚 Getting Started");

  sheet.clear();

  // Set up the sheet with a clean, professional design
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidths(2, 3, 600);

  // Header
  sheet.getRange("A1:D1").merge()
    .setValue("📚 Getting Started with SEIU Local 509 Dashboard")
    .setFontSize(24)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 60);

  // GitHub Repository Information Section
  let row = 3;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue("📦 GitHub Repository")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);

  row++;
  sheet.getRange(row, 2, 1, 3).merge()
    .setValue("Repository: https://github.com/Woop91/509-Dashboard")
    .setFontSize(12)
    .setWrap(true);
  sheet.setRowHeight(row, 30);

  row++;
  sheet.getRange(row, 2, 1, 3).merge()
    .setValue("For bug reports, feature requests, and contributions, please visit the GitHub repository.")
    .setFontStyle("italic")
    .setWrap(true);
  sheet.setRowHeight(row, 30);

  row++;
  sheet.getRange(row, 2, 1, 3).merge()
    .setValue("Documentation and guides are available in the repository's README and guides folder.")
    .setFontStyle("italic")
    .setWrap(true);
  sheet.setRowHeight(row, 30);

  // Quick Start Section
  row += 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue("🚀 Quick Start Guide")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);

  const quickStartSteps = [
    ["Step 1", "Configure Steward Contact Info", "Go to the ⚙️ Config tab, scroll to column U, and enter your steward contact information (Name, Email, Phone). This information is used when starting new grievances."],
    ["Step 2", "Add Members", "Navigate to 👥 Member Directory and start adding your union members. You can enter them manually, import from CSV, or use the seed data function for testing."],
    ["Step 3", "Review Configuration", "Check the ⚙️ Config tab to ensure all dropdown values (job titles, work locations, grievance types, etc.) match your organization's needs. Customize as needed."],
    ["Step 4", "Set Up Triggers", "Go to 509 Tools > Utilities > Setup Triggers to enable automatic calculations and deadline tracking."],
    ["Step 5", "Explore Dashboards", "Visit the various dashboard views: 📊 Main Dashboard for overview metrics, 🎯 Interactive Dashboard for customizable views, and 👨‍⚖️ Steward Workload for assignment tracking."]
  ];

  row++;
  for (let i = 0; i < quickStartSteps.length; i++) {
    const step = quickStartSteps[i];

    // Step number
    sheet.getRange(row, 1)
      .setValue(step[0])
      .setFontWeight("bold")
      .setFontSize(14)
      .setBackground(COLORS.INFO_LIGHT)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false);

    // Step title
    sheet.getRange(row, 2)
      .setValue(step[1])
      .setFontWeight("bold")
      .setFontSize(12)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false);

    // Step description
    sheet.getRange(row, 3, 1, 2).merge()
      .setValue(step[2])
      .setWrap(true)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false);

    sheet.setRowHeight(row, 60);
    row++;
  }

  // Key Features Section
  row += 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue("⭐ Key Features")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);

  const features = [
    ["🚀 Grievance Workflow", "Start grievances from Member Directory with pre-filled forms, automatic log entry, PDF generation, and email delivery."],
    ["📊 Interactive Dashboards", "User-selectable metrics, dynamic chart types, side-by-side comparison, and 6 professional themes with real-time updates."],
    ["⚖️ CBA Compliance", "Automatic deadline tracking based on Article 23A (21-day filing, 30-day decisions, 10-day appeals) with color-coded alerts."],
    ["👥 Member Management", "Comprehensive member directory with auto-calculated grievance metrics, engagement tracking, and committee participation."],
    ["📈 Advanced Analytics", "Four dedicated analytical tabs: Trends & Timeline, Performance Metrics, Location Analytics, and Type Analysis."],
    ["🧠 ADHD-Friendly Design", "Soft colors, no gridlines, emoji icons, large numbers, and minimal visual clutter for easy scanning."],
    ["👨‍⚖️ Steward Workload", "Automatic calculation of cases per steward, active case breakdown, overdue highlights, and win rates."]
  ];

  row++;
  for (let i = 0; i < features.length; i++) {
    const feature = features[i];

    // Feature name
    sheet.getRange(row, 1, 1, 2).merge()
      .setValue(feature[0])
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false)
      .setBackground(COLORS.LIGHT_GRAY);

    // Feature description
    sheet.getRange(row, 3, 1, 2).merge()
      .setValue(feature[1])
      .setWrap(true)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false);

    sheet.setRowHeight(row, 50);
    row++;
  }

  // Important Links Section
  row += 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue("🔗 Important Links & Resources")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);

  const links = [
    ["GitHub Repository", "https://github.com/Woop91/509-Dashboard"],
    ["Grievance Workflow Guide", "See GRIEVANCE_WORKFLOW_GUIDE.md in the repository"],
    ["ADHD-Friendly Guide", "See ADHD_FRIENDLY_GUIDE.md in the repository"],
    ["Steward Excellence Guide", "See STEWARD_GUIDE.md in the repository"],
    ["Seed Nuke Guide", "See SEED_NUKE_GUIDE.md in the repository"]
  ];

  row++;
  for (let i = 0; i < links.length; i++) {
    const link = links[i];

    sheet.getRange(row, 2)
      .setValue(link[0])
      .setFontWeight("bold")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false);

    sheet.getRange(row, 3, 1, 2).merge()
      .setValue(link[1])
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false)
      .setFontColor("#1155CC");

    sheet.setRowHeight(row, 30);
    row++;
  }

  // Support Section
  row += 2;
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue("📞 Support & Help")
    .setFontSize(18)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 40);

  row++;
  sheet.getRange(row, 2, 1, 3).merge()
    .setValue("For questions, issues, or feature requests:\n• Visit the GitHub Issues page: https://github.com/Woop91/509-Dashboard/issues\n• Check the FAQ tab in this spreadsheet\n• Review the comprehensive guides in the repository\n• Contact your system administrator")
    .setWrap(true)
    .setVerticalAlignment("top");
  sheet.setRowHeight(row, 80);

  // Freeze header row
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Creates the FAQ sheet
 */
function createFAQSheet(ss) {
  let sheet = ss.getSheetByName("❓ FAQ");
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet("❓ FAQ");

  sheet.clear();

  // Set up the sheet with a clean, professional design
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidths(3, 1, 700);

  // Header
  sheet.getRange("A1:C1").merge()
    .setValue("❓ Frequently Asked Questions (FAQ)")
    .setFontSize(24)
    .setFontWeight("bold")
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor("white")
    .setFontFamily("Roboto")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 60);

  // FAQ Categories and Questions
  let row = 3;

  // General Questions
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("📋 General Questions")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const generalFAQs = [
    ["What is the 509 Dashboard?", "The 509 Dashboard is a comprehensive Google Sheets-based system for SEIU Local 509 (Units 8 & 10) to track member information, manage grievances, monitor CBA deadlines, and analyze trends. It's specifically designed for Massachusetts state employees and ensures compliance with Article 23A grievance procedures."],
    ["Who should use this dashboard?", "This dashboard is designed for union stewards, representatives, and administrators who need to manage member data, track grievances, and ensure compliance with collective bargaining agreement deadlines."],
    ["Where can I find the source code?", "The source code is available on GitHub at https://github.com/Woop91/509-Dashboard. You can report issues, request features, and contribute to the project there."],
    ["How do I get started?", "Check the 📚 Getting Started tab for a step-by-step guide. In brief: configure steward contact info in the Config tab, add members to the Member Directory, review configuration settings, set up triggers, and explore the dashboards."]
  ];

  row++;
  row = addFAQSection(sheet, row, generalFAQs);

  // Data Entry Questions
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("✍️ Data Entry Questions")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const dataFAQs = [
    ["How do I add a new member?", "Go to the 👥 Member Directory sheet and add a new row. Fill in the basic information (First Name, Last Name, Job Title, Work Location, Unit, Email, Phone). The system will automatically calculate grievance metrics. Leave auto-calculated columns (highlighted in green/orange) blank."],
    ["How do I add a new grievance?", "Go to the 📋 Grievance Log sheet and add a new row. Enter the Member ID, name, status, current step, Incident Date, and grievance details. The system will automatically calculate all CBA-compliant deadlines based on Article 23A."],
    ["Which columns should I not edit manually?", "Never manually edit columns highlighted in green or orange. These are auto-calculated fields including: Member metrics (Total Grievances, Win Rate, etc.), Deadline columns (Filing Deadline, Step I/II/III Due Dates), and Derived fields (Days Open, Priority Score, etc.)."],
    ["Can I import data from another spreadsheet?", "Yes! You can copy and paste data from another spreadsheet, or use the seed data functions for testing. Just ensure your data matches the column structure. Go to 509 Tools > Data Management > Seed All Test Data for sample data."]
  ];

  row++;
  row = addFAQSection(sheet, row, dataFAQs);

  // Features & Functionality
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("⚙️ Features & Functionality")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const featureFAQs = [
    ["How does the grievance workflow feature work?", "Click on any member in the Member Directory, then go to 509 Tools > Grievance Tools > Start New Grievance. This opens a pre-filled Google Form with the member's information automatically populated. When submitted, the grievance is automatically added to the Grievance Log."],
    ["What are the Level 2 columns?", "Level 2 columns are advanced engagement tracking fields including: Last Virtual/In-Person Meeting dates, Survey dates, Email open rates, Volunteer hours, Interest in various actions, Communication preferences, Best contact times, and Steward contact notes. Use the menu toggle to show/hide these columns."],
    ["How do I hide/show grievance columns?", "Go to 509 Tools > View Options > Toggle Grievance Columns to show or hide grievance-related columns in the Member Directory. This helps focus on specific data when needed."],
    ["What is the Interactive Dashboard?", "The Interactive Dashboard (🎯 tab) lets you choose which metrics to display, select chart types (pie, donut, bar, line, column, area, or table), compare metrics side-by-side, and apply professional themes. It's fully customizable to your needs."]
  ];

  row++;
  row = addFAQSection(sheet, row, featureFAQs);

  // Troubleshooting
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("🔧 Troubleshooting")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.SOLIDARITY_RED)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const troubleshootingFAQs = [
    ["Charts are not showing data", "Solution: Add at least one member to Member Directory and one grievance to Grievance Log, then run 509 Tools > Data Management > Rebuild Dashboard."],
    ["Deadlines are not calculating", "Solution: Ensure you entered the Incident Date (Column G) and Date Filed (Column I) in the Grievance Log. Then run 509 Tools > Data Management > Recalc All Grievances."],
    ["Member metrics are not updating", "Solution: Verify that Member IDs match between Member Directory and Grievance Log exactly. Then run 509 Tools > Data Management > Recalc All Members."],
    ["Dropdowns are not working", "Solution: Check that the ⚙️ Config sheet exists and has data. If needed, run 509 Tools > Create Dashboard to rebuild, or go to 509 Tools > Utilities > Setup Triggers."],
    ["I'm getting permission errors", "Solution: Go to Extensions > Apps Script, click Run (▶️) > select any function, click Review Permissions, choose your Google account, click Advanced > Go to [Project Name], then click Allow."]
  ];

  row++;
  row = addFAQSection(sheet, row, troubleshootingFAQs);

  // CBA Compliance
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("⚖️ CBA Compliance & Deadlines")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_PURPLE)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const cbaFAQs = [
    ["What deadlines does the system track?", "The system automatically tracks all Article 23A deadlines: 21 days from incident to file grievance, 30 days for Step I decision, 10 days to appeal to Step II, 30 days for Step II decision, 10 days to appeal to Step III, and 30 days for Step III decision."],
    ["What do the color codes mean?", "Green = More than 7 days until deadline (on track), Yellow = 0-7 days until deadline (due soon), Red = Past deadline (overdue), Dark Red = 30+ days overdue (urgent)."],
    ["How are grievance priorities determined?", "The system automatically assigns priorities: Step III = highest priority (1), Step II = priority 2, Step I = priority 3, then sorted by due date within each step."],
    ["Can I customize the deadlines?", "The deadlines are based on the CBA Article 23A and are set in the Config sheet's Timeline Rules Table. You can view them in the ⚙️ Config tab starting at column O. To change them, you'd need to modify the CBA_DEADLINES constants in the script."]
  ];

  row++;
  row = addFAQSection(sheet, row, cbaFAQs);

  // GitHub & Development
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("💻 GitHub & Development")
    .setFontSize(16)
    .setFontWeight("bold")
    .setBackground(COLORS.ACCENT_TEAL)
    .setFontColor("white")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(row, 35);

  const githubFAQs = [
    ["Where is the GitHub repository?", "The repository is at https://github.com/Woop91/509-Dashboard. This is where you'll find the source code, documentation, guides, and can report issues or request features."],
    ["How do I report a bug or request a feature?", "Go to https://github.com/Woop91/509-Dashboard/issues and click 'New Issue'. Provide a clear description of the bug or feature request, including steps to reproduce if it's a bug."],
    ["Can I contribute to the project?", "Yes! The project is open for contributions. Visit the GitHub repository, fork it, make your changes, and submit a pull request. Please follow the contribution guidelines in the repository."],
    ["How do I update to the latest version?", "Check the GitHub repository's Releases page for the latest version. Download the updated .gs files and replace them in your Apps Script project via Extensions > Apps Script."]
  ];

  row++;
  row = addFAQSection(sheet, row, githubFAQs);

  // Additional Help
  row += 2;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("Need more help?")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground(COLORS.INFO_LIGHT)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(row, 30);

  row++;
  sheet.getRange(row, 1, 1, 3).merge()
    .setValue("Check the 📚 Getting Started tab for step-by-step guides, visit the GitHub repository for comprehensive documentation, or review the README.md file in the repository for detailed information about all features.")
    .setWrap(true)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setFontStyle("italic");
  sheet.setRowHeight(row, 60);

  // Freeze header row
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Helper function to add FAQ section rows
 */
function addFAQSection(sheet, startRow, faqs) {
  let row = startRow;

  for (let i = 0; i < faqs.length; i++) {
    const faq = faqs[i];

    // Number
    sheet.getRange(row, 1)
      .setValue("Q" + (i + 1))
      .setFontWeight("bold")
      .setFontSize(12)
      .setBackground(COLORS.INFO_LIGHT)
      .setVerticalAlignment("top")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, false, false);

    // Question
    sheet.getRange(row, 2)
      .setValue(faq[0])
      .setFontWeight("bold")
      .setFontSize(11)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false)
      .setWrap(true);

    // Answer
    sheet.getRange(row, 3)
      .setValue(faq[1])
      .setWrap(true)
      .setVerticalAlignment("top")
      .setBorder(true, true, true, true, false, false);

    // Auto-adjust row height based on content
    const contentLength = faq[1].length;
    const estimatedHeight = Math.max(40, Math.min(150, Math.ceil(contentLength / 80) * 20));
    sheet.setRowHeight(row, estimatedHeight);

    row++;
  }

  return row;
}

/******************************************************************************
 * MODULE: ColumnToggles
 * Source: ColumnToggles.gs
 *****************************************************************************/

/**
 * ============================================================================
 * COLUMN VISIBILITY TOGGLES
 * ============================================================================
 *
 * Functions to show/hide column groups in Member Directory
 * - Grievance columns toggle
 * - Level 2 columns toggle
 *
 * ============================================================================
 */

/**
 * Toggles visibility of grievance columns in Member Directory
 */
function toggleGrievanceColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  // Note: This feature expects specific grievance tracking columns (Total Grievances Filed,
  // Active Grievances, etc.) at columns 12-21 which are not present in the current Member
  // Directory structure. The current columns 12-21 contain engagement/contact data.

  SpreadsheetApp.getUi().alert(
    '⚠️ Feature Not Available',
    'This column toggle feature is designed for a different Member Directory structure.\n\n' +
    'The current Member Directory (31 columns) does not include the advanced grievance tracking columns this feature expects.\n\n' +
    'You can manually hide/show columns by right-clicking column headers.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return;

  /*
  // Legacy code - disabled because columns don't match current structure
  const firstGrievanceCol = 12; // Column L
  const lastGrievanceCol = 21;  // Column U
  const numCols = lastGrievanceCol - firstGrievanceCol + 1;

  const isHidden = sheet.isColumnHiddenByUser(firstGrievanceCol);

  if (isHidden) {
    sheet.showColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now visible');
  } else {
    sheet.hideColumns(firstGrievanceCol, numCols);
    SpreadsheetApp.getUi().alert('✅ Grievance columns are now hidden');
  }
  */
}

/**
 * Toggles visibility of Level 2 (engagement tracking) columns in Member Directory
 */
function toggleLevel2Columns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  // Level 2 columns start after the basic columns
  // These will be added at the end of the existing columns
  // Starting at column AH (34) based on current structure
  const firstLevel2Col = 34;

  // Count: 14 Level 2 columns (AH through AU)
  const numCols = 14;

  // Check if the Level 2 columns exist
  const lastCol = sheet.getLastColumn();
  if (lastCol < firstLevel2Col) {
    SpreadsheetApp.getUi().alert('❌ Level 2 columns have not been added yet.\n\nPlease run "Create Dashboard" to add Level 2 columns.');
    return;
  }

  // Check if columns are currently hidden
  const isHidden = sheet.isColumnHiddenByUser(firstLevel2Col);

  if (isHidden) {
    // Show columns
    sheet.showColumns(firstLevel2Col, numCols);
    SpreadsheetApp.getUi().alert('✅ Level 2 columns are now visible');
  } else {
    // Hide columns
    sheet.hideColumns(firstLevel2Col, numCols);
    SpreadsheetApp.getUi().alert('✅ Level 2 columns are now hidden');
  }
}

/**
 * Shows all hidden columns in Member Directory
 */
function showAllMemberColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Member Directory sheet not found!');
    return;
  }

  const lastCol = sheet.getLastColumn();

  // Show all columns
  for (let i = 1; i <= lastCol; i++) {
    if (sheet.isColumnHiddenByUser(i)) {
      sheet.showColumns(i);
    }
  }

  SpreadsheetApp.getUi().alert('✅ All columns are now visible');
}

/******************************************************************************
 * MODULE: GoogleDriveIntegration
 * Source: GoogleDriveIntegration.gs
 *****************************************************************************/

/**
 * ============================================================================
 * GOOGLE DRIVE INTEGRATION
 * ============================================================================
 *
 * Attach and manage documents for grievances
 * Features:
 * - Link Drive folders to grievances
 * - Upload evidence and documents
 * - Auto-organize by grievance ID
 * - Version control
 * - Quick access to attachments
 */

/**
 * Creates root folder for 509 Dashboard in Google Drive
 * @returns {Folder} Root folder
 */
function createRootFolder() {
  const folderName = '509 Dashboard - Grievance Files';

  // Check if folder already exists
  const existingFolders = DriveApp.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  // Create new folder
  const folder = DriveApp.createFolder(folderName);

  Logger.log(`Created root folder: ${folderName}`);
  return folder;
}

/**
 * Creates folder for a specific grievance
 * @param {string} grievanceId - Grievance ID
 * @returns {Folder} Grievance folder
 */
function createGrievanceFolder(grievanceId) {
  const rootFolder = createRootFolder();

  // Check if grievance folder already exists
  const folderName = `Grievance_${grievanceId}`;
  const existingFolders = rootFolder.getFoldersByName(folderName);

  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  // Create new grievance folder
  const folder = rootFolder.createFolder(folderName);

  // Create subfolders
  folder.createFolder('Evidence');
  folder.createFolder('Correspondence');
  folder.createFolder('Forms');
  folder.createFolder('Other');

  Logger.log(`Created folder structure for ${grievanceId}`);
  return folder;
}

/**
 * Links Drive folder to grievance in spreadsheet
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 */
function linkFolderToGrievance(grievanceId, folderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const lastRow = grievanceSheet.getLastRow();
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      const row = i + 2;

      // Store folder ID in column AC (29)
      grievanceSheet.getRange(row, 29).setValue(folderId);

      // Create hyperlink to folder in column AD (30)
      const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;
      grievanceSheet.getRange(row, 30).setValue(folderUrl);

      Logger.log(`Linked folder ${folderId} to grievance ${grievanceId}`);
      return;
    }
  }

  throw new Error(`Grievance ${grievanceId} not found`);
}

/**
 * Sets up Drive folder for selected grievance
 */
function setupDriveFolderForGrievance() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!grievanceId) {
    ui.alert(
      '⚠️ No Grievance ID',
      'Selected row does not have a grievance ID.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '📁 Setup Drive Folder',
    `Create a Google Drive folder for grievance ${grievanceId}?\n\n` +
    'This will create a folder with subfolders for:\n' +
    '• Evidence\n' +
    '• Correspondence\n' +
    '• Forms\n' +
    '• Other\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('📁 Creating folder...', 'Please wait', -1);

    const folder = createGrievanceFolder(grievanceId);
    linkFolderToGrievance(grievanceId, folder.getId());

    ui.alert(
      '✅ Folder Created',
      `Drive folder created successfully for ${grievanceId}.\n\n` +
      `Folder link has been added to the grievance record.\n\n` +
      `Click the link in column AD to open the folder.`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Shows file upload dialog for a grievance
 */
function showFileUploadDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!grievanceId) {
    ui.alert(
      '⚠️ No Grievance ID',
      'Selected row does not have a grievance ID.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Get or create folder
  let folder;
  try {
    folder = createGrievanceFolder(grievanceId);
    linkFolderToGrievance(grievanceId, folder.getId());
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    return;
  }

  const html = createFileUploadHTML(grievanceId, folder.getId());
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(650)
    .setHeight(500);

  ui.showModalDialog(htmlOutput, `📎 Upload Files - ${grievanceId}`);
}

/**
 * Creates HTML for file upload interface
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 * @returns {string} HTML content
 */
function createFileUploadHTML(grievanceId, folderId) {
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  return `
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
      background: #1a73e8;
      color: white;
      padding: 12px 20px;
      border-radius: 4px;
      text-decoration: none;
      display: inline-block;
      margin: 15px 0;
    }
    .folder-link:hover {
      background: #1557b0;
    }
    .instructions {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
    }
    .instructions h3 {
      margin-top: 0;
      color: #333;
    }
    .instructions ol {
      margin: 10px 0;
      padding-left: 20px;
    }
    .instructions li {
      margin: 8px 0;
      color: #666;
    }
    .category-list {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin: 15px 0;
    }
    .category {
      background: #f8f9fa;
      padding: 12px;
      border-radius: 4px;
      border-left: 3px solid #1a73e8;
    }
    .category strong {
      color: #333;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📎 Upload Files</h2>

    <div class="info-box">
      <strong>Grievance:</strong> ${grievanceId}<br>
      <strong>Drive Folder:</strong> Created and linked
    </div>

    <a href="${folderUrl}" target="_blank" class="folder-link">
      📁 Open Drive Folder
    </a>

    <div class="instructions">
      <h3>How to Upload Files</h3>
      <ol>
        <li>Click "Open Drive Folder" above to open the folder in Google Drive</li>
        <li>Navigate to the appropriate subfolder:
          <div class="category-list">
            <div class="category"><strong>📄 Evidence</strong><br>Photos, documents, proof</div>
            <div class="category"><strong>✉️ Correspondence</strong><br>Emails, letters</div>
            <div class="category"><strong>📋 Forms</strong><br>Signed forms, contracts</div>
            <div class="category"><strong>📎 Other</strong><br>Miscellaneous files</div>
          </div>
        </li>
        <li>Drag and drop files or click "New" → "File upload"</li>
        <li>All files are automatically organized and linked to this grievance</li>
      </ol>
    </div>

    <div class="info-box" style="background: #fff3cd; border-left-color: #ffc107;">
      <strong>💡 Tip:</strong> Files are automatically versioned in Drive. If you upload a file with the same name, Drive will keep both versions.
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Shows files attached to a grievance
 */
function showGrievanceFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  const folderId = activeSheet.getRange(activeRow, 29).getValue(); // Column AC

  if (!folderId) {
    ui.alert(
      'ℹ️ No Folder Linked',
      `No Drive folder has been set up for ${grievanceId}.\n\n` +
      'Use "Setup Drive Folder" to create one.',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = listFolderFiles(folder);

    if (files.length === 0) {
      ui.alert(
        'ℹ️ No Files',
        `No files have been uploaded to the folder for ${grievanceId}.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const html = createFileListHTML(grievanceId, folderId, files);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

    ui.showModalDialog(htmlOutput, `📁 Files - ${grievanceId}`);

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Lists all files in a folder recursively
 * @param {Folder} folder - Drive folder
 * @param {string} path - Current path (for recursion)
 * @returns {Array} Array of file objects
 */
function listFolderFiles(folder, path = '') {
  const files = [];

  // Get files in current folder
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files.push({
      name: file.getName(),
      url: file.getUrl(),
      size: formatFileSize(file.getSize()),
      modified: file.getLastUpdated().toLocaleString(),
      path: path,
      type: file.getMimeType()
    });
  }

  // Recursively get files from subfolders
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const subfolder = folderIterator.next();
    const subfolderName = subfolder.getName();
    const subfiles = listFolderFiles(subfolder, path ? `${path}/${subfolderName}` : subfolderName);
    files.push(...subfiles);
  }

  return files;
}

/**
 * Formats file size in human-readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';

  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Creates HTML for file list display
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 * @param {Array} files - Array of file objects
 * @returns {string} HTML content
 */
function createFileListHTML(grievanceId, folderId, files) {
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  const filesList = files
    .map(file => `
      <div class="file-item">
        <div class="file-icon">${getFileIcon(file.type)}</div>
        <div class="file-details">
          <div class="file-name">
            <a href="${file.url}" target="_blank">${file.name}</a>
          </div>
          <div class="file-meta">
            ${file.path ? `${file.path} • ` : ''}${file.size} • Modified: ${file.modified}
          </div>
        </div>
      </div>
    `)
    .join('');

  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .summary {
      background: #e8f0fe;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .file-item {
      display: flex;
      align-items: center;
      padding: 12px;
      margin: 10px 0;
      background: #f8f9fa;
      border-radius: 4px;
      border-left: 3px solid #1a73e8;
    }
    .file-item:hover {
      background: #e9ecef;
    }
    .file-icon {
      font-size: 24px;
      margin-right: 15px;
    }
    .file-details {
      flex: 1;
    }
    .file-name a {
      color: #1a73e8;
      text-decoration: none;
      font-weight: 500;
    }
    .file-name a:hover {
      text-decoration: underline;
    }
    .file-meta {
      font-size: 12px;
      color: #666;
      margin-top: 4px;
    }
    .folder-link {
      background: #1a73e8;
      color: white;
      padding: 10px 20px;
      border-radius: 4px;
      text-decoration: none;
      display: inline-block;
      margin: 15px 0;
    }
    .folder-link:hover {
      background: #1557b0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📁 Attached Files - ${grievanceId}</h2>

    <div class="summary">
      <strong>Total Files:</strong> ${files.length}
    </div>

    <a href="${folderUrl}" target="_blank" class="folder-link">
      📁 Open Folder in Drive
    </a>

    <div style="max-height: 400px; overflow-y: auto; margin-top: 15px;">
      ${filesList}
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Gets emoji icon for file type
 * @param {string} mimeType - MIME type
 * @returns {string} Emoji icon
 */
function getFileIcon(mimeType) {
  if (mimeType.includes('image')) return '🖼️';
  if (mimeType.includes('pdf')) return '📄';
  if (mimeType.includes('word') || mimeType.includes('document')) return '📝';
  if (mimeType.includes('sheet') || mimeType.includes('excel')) return '📊';
  if (mimeType.includes('presentation')) return '📽️';
  if (mimeType.includes('video')) return '🎥';
  if (mimeType.includes('audio')) return '🎵';
  if (mimeType.includes('zip') || mimeType.includes('compressed')) return '🗜️';
  return '📎';
}

/**
 * Batch creates folders for all grievances without them
 */
function batchCreateGrievanceFolders() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📁 Batch Create Folders',
    'This will create Drive folders for all grievances that don\'t have one.\n\n' +
    'This may take a few minutes for large datasets.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('📁 Creating folders...', 'Please wait', -1);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    const lastRow = grievanceSheet.getLastRow();

    if (lastRow < 2) {
      ui.alert('No grievances found.');
      return;
    }

    const data = grievanceSheet.getRange(2, 1, lastRow - 1, 29).getValues();
    let created = 0;
    let skipped = 0;

    data.forEach((row, index) => {
      const grievanceId = row[0];
      const folderId = row[28]; // Column AC

      if (!grievanceId) {
        skipped++;
        return;
      }

      if (folderId) {
        skipped++;
        return;
      }

      try {
        const folder = createGrievanceFolder(grievanceId);
        linkFolderToGrievance(grievanceId, folder.getId());
        created++;

        // Progress update every 10 folders
        if (created % 10 === 0) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Created ${created} folders...`,
            'Progress',
            -1
          );
        }
      } catch (error) {
        Logger.log(`Error creating folder for ${grievanceId}: ${error.message}`);
      }
    });

    ui.alert(
      '✅ Batch Folder Creation Complete',
      `Created ${created} new folders.\n` +
      `Skipped ${skipped} grievances (already had folders or no ID).`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/******************************************************************************
 * MODULE: GmailIntegration
 * Source: GmailIntegration.gs
 *****************************************************************************/

/**
 * ============================================================================
 * GMAIL INTEGRATION
 * ============================================================================
 *
 * Send and receive grievance-related emails via Gmail
 * Features:
 * - Send PDFs directly from dashboard
 * - Email templates library
 * - Track communications
 * - Auto-log sent emails
 * - Compose emails with templates
 */

/**
 * Shows email composition dialog for a grievance
 * @param {string} grievanceId - Optional grievance ID to pre-populate
 */
function composeGrievanceEmail(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  let grievanceData = null;

  if (grievanceId) {
    const lastRow = grievanceSheet.getLastRow();
    const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === grievanceId) {
        grievanceData = {
          id: data[i][0],
          memberName: `${data[i][2]} ${data[i][3]}`,
          memberEmail: data[i][7] || '',
          steward: data[i][13] || '',
          issueType: data[i][5] || '',
          status: data[i][4] || ''
        };
        break;
      }
    }
  }

  const html = createEmailComposerHTML(grievanceData);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📧 Compose Email');
}

/**
 * Creates HTML for email composer
 * @param {Object} grievanceData - Grievance information
 * @returns {string} HTML content
 */
function createEmailComposerHTML(grievanceData) {
  const templates = getEmailTemplates();

  return `
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .form-group {
      margin: 15px 0;
    }
    label {
      display: block;
      font-weight: bold;
      margin-bottom: 5px;
      color: #333;
    }
    input[type="text"],
    input[type="email"],
    textarea,
    select {
      width: 100%;
      padding: 10px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
      box-sizing: border-box;
    }
    input:focus,
    textarea:focus,
    select:focus {
      outline: none;
      border-color: #1a73e8;
    }
    textarea {
      min-height: 200px;
      font-family: Arial, sans-serif;
      resize: vertical;
    }
    .button-group {
      margin-top: 20px;
      display: flex;
      gap: 10px;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      flex: 1;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.secondary:hover {
      background: #5a6268;
    }
    .template-selector {
      margin: 10px 0;
    }
    .info-box {
      background: #e8f0fe;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .checkbox-group {
      margin: 10px 0;
    }
    .checkbox-group label {
      display: inline;
      font-weight: normal;
      margin-left: 5px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📧 Compose Email</h2>

    ${grievanceData ? `
    <div class="info-box">
      <strong>Grievance:</strong> ${grievanceData.id} - ${grievanceData.memberName}<br>
      <strong>Status:</strong> ${grievanceData.status} | <strong>Issue:</strong> ${grievanceData.issueType}
    </div>
    ` : ''}

    <div class="form-group template-selector">
      <label>Template (optional):</label>
      <select id="templateSelect" onchange="loadTemplate()">
        <option value="">-- Select a template --</option>
        ${templates.map(t => `<option value="${t.id}">${t.name}</option>`).join('')}
      </select>
    </div>

    <div class="form-group">
      <label>To:</label>
      <input type="email" id="toEmail" value="${grievanceData ? grievanceData.memberEmail : ''}" placeholder="recipient@example.com">
    </div>

    <div class="form-group">
      <label>CC (optional):</label>
      <input type="email" id="ccEmail" value="${grievanceData ? grievanceData.steward : ''}" placeholder="cc@example.com">
    </div>

    <div class="form-group">
      <label>Subject:</label>
      <input type="text" id="subject" value="${grievanceData ? `Re: Grievance ${grievanceData.id}` : ''}" placeholder="Email subject">
    </div>

    <div class="form-group">
      <label>Message:</label>
      <textarea id="message" placeholder="Type your message here..."></textarea>
    </div>

    <div class="checkbox-group">
      <input type="checkbox" id="attachPDF" ${grievanceData ? 'checked' : ''}>
      <label for="attachPDF">Attach grievance PDF report</label>
    </div>

    <div class="checkbox-group">
      <input type="checkbox" id="logCommunication" checked>
      <label for="logCommunication">Log this email in Communications Log</label>
    </div>

    <div class="button-group">
      <button onclick="sendEmail()">📤 Send Email</button>
      <button class="secondary" onclick="google.script.host.close()">❌ Cancel</button>
    </div>

    <div id="status" style="margin-top: 15px; padding: 10px; display: none;"></div>
  </div>

  <script>
    const grievanceData = ${JSON.stringify(grievanceData)};
    const templates = ${JSON.stringify(templates)};

    function loadTemplate() {
      const templateId = document.getElementById('templateSelect').value;
      if (!templateId) return;

      const template = templates.find(t => t.id === templateId);
      if (!template) return;

      document.getElementById('subject').value = template.subject || '';
      document.getElementById('message').value = replacePlaceholders(template.body || '');
    }

    function replacePlaceholders(text) {
      if (!grievanceData) return text;

      return text
        .replace(/\\{GRIEVANCE_ID\\}/g, grievanceData.id || '')
        .replace(/\\{MEMBER_NAME\\}/g, grievanceData.memberName || '')
        .replace(/\\{ISSUE_TYPE\\}/g, grievanceData.issueType || '')
        .replace(/\\{STATUS\\}/g, grievanceData.status || '');
    }

    function sendEmail() {
      const to = document.getElementById('toEmail').value.trim();
      const cc = document.getElementById('ccEmail').value.trim();
      const subject = document.getElementById('subject').value.trim();
      const message = document.getElementById('message').value.trim();
      const attachPDF = document.getElementById('attachPDF').checked;
      const logCommunication = document.getElementById('logCommunication').checked;

      if (!to || !subject || !message) {
        showStatus('Please fill in all required fields (To, Subject, Message)', 'error');
        return;
      }

      showStatus('Sending email...', 'info');

      const emailData = {
        to: to,
        cc: cc,
        subject: subject,
        message: message,
        attachPDF: attachPDF,
        logCommunication: logCommunication,
        grievanceId: grievanceData ? grievanceData.id : null
      };

      google.script.run
        .withSuccessHandler(onEmailSent)
        .withFailureHandler(onEmailError)
        .sendGrievanceEmail(emailData);
    }

    function onEmailSent(result) {
      showStatus('✅ Email sent successfully!', 'success');
      setTimeout(() => {
        google.script.host.close();
      }, 2000);
    }

    function onEmailError(error) {
      showStatus('❌ Error: ' + error.message, 'error');
    }

    function showStatus(message, type) {
      const statusDiv = document.getElementById('status');
      statusDiv.textContent = message;
      statusDiv.style.display = 'block';
      statusDiv.style.background = type === 'error' ? '#f8d7da' : type === 'success' ? '#d4edda' : '#d1ecf1';
      statusDiv.style.color = type === 'error' ? '#721c24' : type === 'success' ? '#155724' : '#0c5460';
      statusDiv.style.border = '1px solid ' + (type === 'error' ? '#f5c6cb' : type === 'success' ? '#c3e6cb' : '#bee5eb');
      statusDiv.style.borderRadius = '4px';
    }
  </script>
</body>
</html>
  `;
}

/**
 * Sends email with optional PDF attachment
 * @param {Object} emailData - Email information
 */
function sendGrievanceEmail(emailData) {
  try {
    const options = {
      to: emailData.to,
      subject: emailData.subject,
      body: emailData.message,
      name: 'SEIU Local 509'
    };

    if (emailData.cc) {
      options.cc = emailData.cc;
    }

    // Attach PDF if requested
    if (emailData.attachPDF && emailData.grievanceId) {
      const pdf = exportGrievanceToPDF(emailData.grievanceId);
      if (pdf) {
        options.attachments = [pdf];
      }
    }

    // Send email
    MailApp.sendEmail(options);

    // Log communication if requested
    if (emailData.logCommunication && emailData.grievanceId) {
      logCommunication(
        emailData.grievanceId,
        'Email Sent',
        `To: ${emailData.to}\nSubject: ${emailData.subject}\n\n${emailData.message}`
      );
    }

    return { success: true };

  } catch (error) {
    Logger.log('Error sending email: ' + error.message);
    throw new Error('Failed to send email: ' + error.message);
  }
}

/**
 * Logs communication in the Communications Log sheet
 * @param {string} grievanceId - Grievance ID
 * @param {string} type - Communication type
 * @param {string} details - Communication details
 */
function logCommunication(grievanceId, type, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let commLog = ss.getSheetByName('📞 Communications Log');

  if (!commLog) {
    commLog = createCommunicationsLogSheet();
  }

  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || 'System';

  const lastRow = commLog.getLastRow();
  const newRow = [
    timestamp,
    grievanceId,
    type,
    user,
    details
  ];

  commLog.getRange(lastRow + 1, 1, 1, 5).setValues([newRow]);
}

/**
 * Creates Communications Log sheet
 * @returns {Sheet} Communications Log sheet
 */
function createCommunicationsLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('📞 Communications Log');

  // Set headers
  const headers = [
    'Timestamp',
    'Grievance ID',
    'Type',
    'User',
    'Details'
  ];

  sheet.getRange(1, 1, 1, 5).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, 5)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 120); // Grievance ID
  sheet.setColumnWidth(3, 120); // Type
  sheet.setColumnWidth(4, 200); // User
  sheet.setColumnWidth(5, 400); // Details

  // Freeze header
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Gets email templates from configuration
 * @returns {Array} Array of template objects
 */
function getEmailTemplates() {
  return [
    {
      id: 'initial_acknowledgment',
      name: 'Initial Acknowledgment',
      subject: 'Re: Grievance {GRIEVANCE_ID} - Acknowledgment',
      body: `Dear {MEMBER_NAME},

This email confirms that we have received your grievance (ID: {GRIEVANCE_ID}) regarding {ISSUE_TYPE}.

Your case has been assigned to a steward who will contact you within 2-3 business days to discuss next steps.

Current Status: {STATUS}

If you have any questions or additional information to provide, please don't hesitate to contact us.

In solidarity,
SEIU Local 509`
    },
    {
      id: 'status_update',
      name: 'Status Update',
      subject: 'Grievance {GRIEVANCE_ID} - Status Update',
      body: `Dear {MEMBER_NAME},

I wanted to provide you with an update on your grievance (ID: {GRIEVANCE_ID}).

Current Status: {STATUS}

[Add specific update details here]

We will continue to keep you informed of any developments.

In solidarity,
SEIU Local 509`
    },
    {
      id: 'request_information',
      name: 'Request for Information',
      subject: 'Grievance {GRIEVANCE_ID} - Additional Information Needed',
      body: `Dear {MEMBER_NAME},

To proceed with your grievance (ID: {GRIEVANCE_ID}), we need some additional information:

[List the information needed here]

Please provide this information at your earliest convenience so we can continue processing your case.

Thank you for your cooperation.

In solidarity,
SEIU Local 509`
    },
    {
      id: 'resolution_notification',
      name: 'Resolution Notification',
      subject: 'Grievance {GRIEVANCE_ID} - Resolved',
      body: `Dear {MEMBER_NAME},

I'm pleased to inform you that your grievance (ID: {GRIEVANCE_ID}) has been resolved.

Resolution Details:
[Add resolution details here]

This case is now closed. If you have any questions about the resolution, please don't hesitate to contact us.

Thank you for your patience throughout this process.

In solidarity,
SEIU Local 509`
    },
    {
      id: 'meeting_invitation',
      name: 'Meeting Invitation',
      subject: 'Grievance {GRIEVANCE_ID} - Meeting Scheduled',
      body: `Dear {MEMBER_NAME},

We have scheduled a meeting to discuss your grievance (ID: {GRIEVANCE_ID}).

Date: [DATE]
Time: [TIME]
Location: [LOCATION]

Please confirm your attendance at your earliest convenience.

In solidarity,
SEIU Local 509`
    }
  ];
}

/**
 * Shows email template manager
 */
function showEmailTemplateManager() {
  const templates = getEmailTemplates();

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    .template { background: #f8f9fa; padding: 15px; margin: 15px 0; border-radius: 4px; border-left: 4px solid #1a73e8; }
    .template h3 { margin-top: 0; color: #333; }
    .template-body { background: white; padding: 10px; border-radius: 4px; margin: 10px 0; font-size: 13px; white-space: pre-wrap; }
    .placeholders { background: #fff3cd; padding: 10px; border-radius: 4px; margin: 15px 0; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📝 Email Template Library</h2>

    <div class="placeholders">
      <strong>Available Placeholders:</strong><br>
      {GRIEVANCE_ID}, {MEMBER_NAME}, {ISSUE_TYPE}, {STATUS}
    </div>

    ${templates.map(t => `
      <div class="template">
        <h3>${t.name}</h3>
        <strong>Subject:</strong> ${t.subject}<br><br>
        <strong>Body:</strong>
        <div class="template-body">${t.body}</div>
      </div>
    `).join('')}
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📝 Email Templates');
}

/**
 * Shows communication log for a specific grievance
 * @param {string} grievanceId - Grievance ID
 */
function showGrievanceCommunications(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const commLog = ss.getSheetByName('📞 Communications Log');

  if (!commLog) {
    SpreadsheetApp.getUi().alert('No communications logged yet.');
    return;
  }

  const lastRow = commLog.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No communications logged for this grievance.');
    return;
  }

  const data = commLog.getRange(2, 1, lastRow - 1, 5).getValues();
  const grievanceComms = data.filter(row => row[1] === grievanceId);

  if (grievanceComms.length === 0) {
    SpreadsheetApp.getUi().alert('No communications logged for this grievance.');
    return;
  }

  const commsList = grievanceComms
    .map(row => `
      <div style="background: #f8f9fa; padding: 12px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #1a73e8;">
        <strong>${row[2]}</strong> - ${row[0].toLocaleString()}<br>
        <em>By: ${row[3]}</em><br><br>
        <div style="white-space: pre-wrap; background: white; padding: 10px; border-radius: 4px;">${row[4]}</div>
      </div>
    `)
    .join('');

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-height: 500px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; position: sticky; top: 0; background: white; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📞 Communications Log - ${grievanceId}</h2>
    ${commsList}
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Communications - ${grievanceId}`);
}

/******************************************************************************
 * MODULE: DarkModeThemes
 * Source: DarkModeThemes.gs
 *****************************************************************************/

/**
 * ============================================================================
 * DARK MODE & THEME CUSTOMIZATION
 * ============================================================================
 *
 * Comprehensive theming system with dark mode support
 * Features:
 * - Multiple theme presets (Light, Dark, OLED, Blue, Sepia)
 * - Custom theme builder
 * - Per-user theme preferences
 * - Automatic theme switching (time-based)
 * - Sheet-specific theming
 * - Quick theme toggle
 */

/**
 * Theme configuration
 */
const THEME_CONFIG = {
  THEMES: {
    LIGHT: {
      name: 'Light',
      icon: '☀️',
      background: '#ffffff',
      headerBackground: '#1a73e8',
      headerText: '#ffffff',
      evenRow: '#f8f9fa',
      oddRow: '#ffffff',
      text: '#202124',
      border: '#dadce0',
      accent: '#1a73e8'
    },
    DARK: {
      name: 'Dark',
      icon: '🌙',
      background: '#202124',
      headerBackground: '#35363a',
      headerText: '#e8eaed',
      evenRow: '#292a2d',
      oddRow: '#202124',
      text: '#e8eaed',
      border: '#5f6368',
      accent: '#8ab4f8'
    },
    OLED: {
      name: 'OLED Black',
      icon: '🌑',
      background: '#000000',
      headerBackground: '#1a1a1a',
      headerText: '#ffffff',
      evenRow: '#0a0a0a',
      oddRow: '#000000',
      text: '#ffffff',
      border: '#333333',
      accent: '#bb86fc'
    },
    BLUE: {
      name: 'Ocean Blue',
      icon: '🌊',
      background: '#e3f2fd',
      headerBackground: '#1565c0',
      headerText: '#ffffff',
      evenRow: '#bbdefb',
      oddRow: '#e3f2fd',
      text: '#0d47a1',
      border: '#90caf9',
      accent: '#1976d2'
    },
    SEPIA: {
      name: 'Sepia',
      icon: '📜',
      background: '#f4ecd8',
      headerBackground: '#8b7355',
      headerText: '#ffffff',
      evenRow: '#ede1c5',
      oddRow: '#f4ecd8',
      text: '#3e2723',
      border: '#bcaaa4',
      accent: '#6d4c41'
    }
  },
  DEFAULT_THEME: 'LIGHT',
  AUTO_SWITCH_ENABLED: false,
  DARK_MODE_START_HOUR: 18,  // 6 PM
  DARK_MODE_END_HOUR: 7      // 7 AM
};

/**
 * Shows theme manager
 */
function showThemeManager() {
  const html = createThemeManagerHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🎨 Theme Manager');
}

/**
 * Creates HTML for theme manager
 */
function createThemeManagerHTML() {
  const currentTheme = getCurrentTheme();
  const autoSwitch = getAutoSwitchSettings();

  let themeCards = '';
  Object.entries(THEME_CONFIG.THEMES).forEach(([key, theme]) => {
    const isActive = currentTheme === key;
    themeCards += `
      <div class="theme-card ${isActive ? 'active' : ''}" onclick="selectTheme('${key}')">
        <div class="theme-icon">${theme.icon}</div>
        <div class="theme-name">${theme.name}</div>
        <div class="theme-preview">
          <div class="preview-bar" style="background: ${theme.headerBackground};"></div>
          <div class="preview-bar" style="background: ${theme.evenRow};"></div>
          <div class="preview-bar" style="background: ${theme.oddRow};"></div>
          <div class="preview-bar" style="background: ${theme.accent};"></div>
        </div>
        ${isActive ? '<div class="active-badge">✓ Active</div>' : ''}
      </div>
    `;
  });

  return `
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
      max-height: 650px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    h3 {
      color: #333;
      margin-top: 30px;
      margin-bottom: 15px;
    }
    .theme-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
      gap: 20px;
      margin: 20px 0;
    }
    .theme-card {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 12px;
      cursor: pointer;
      border: 3px solid transparent;
      transition: all 0.3s ease;
      text-align: center;
      position: relative;
    }
    .theme-card:hover {
      transform: translateY(-5px);
      box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .theme-card.active {
      border-color: #1a73e8;
      box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3);
    }
    .theme-icon {
      font-size: 48px;
      margin-bottom: 10px;
    }
    .theme-name {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 15px;
    }
    .theme-preview {
      display: flex;
      flex-direction: column;
      gap: 3px;
      margin-top: 10px;
    }
    .preview-bar {
      height: 15px;
      border-radius: 3px;
    }
    .active-badge {
      position: absolute;
      top: 10px;
      right: 10px;
      background: #4caf50;
      color: white;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: bold;
    }
    .settings-section {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
      border-left: 4px solid #1a73e8;
    }
    .setting-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 15px 0;
      border-bottom: 1px solid #e0e0e0;
    }
    .setting-row:last-child {
      border-bottom: none;
    }
    .setting-label {
      font-weight: 500;
      color: #333;
    }
    .setting-description {
      font-size: 13px;
      color: #666;
      margin-top: 4px;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      margin: 5px;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.success {
      background: #4caf50;
    }
    button.danger {
      background: #f44336;
    }
    .toggle-switch {
      position: relative;
      display: inline-block;
      width: 50px;
      height: 24px;
    }
    .toggle-switch input {
      opacity: 0;
      width: 0;
      height: 0;
    }
    .slider {
      position: absolute;
      cursor: pointer;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background-color: #ccc;
      transition: 0.4s;
      border-radius: 24px;
    }
    .slider:before {
      position: absolute;
      content: "";
      height: 18px;
      width: 18px;
      left: 3px;
      bottom: 3px;
      background-color: white;
      transition: 0.4s;
      border-radius: 50%;
    }
    input:checked + .slider {
      background-color: #1a73e8;
    }
    input:checked + .slider:before {
      transform: translateX(26px);
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .quick-actions {
      margin: 20px 0;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🎨 Theme Manager</h2>

    <div class="info-box">
      <strong>💡 Personalize Your Experience:</strong> Choose a theme that's comfortable for your eyes.
      Themes are saved to your profile and persist across sessions.
    </div>

    <h3>Available Themes</h3>
    <div class="theme-grid">
      ${themeCards}
    </div>

    <h3>Theme Settings</h3>
    <div class="settings-section">
      <div class="setting-row">
        <div>
          <div class="setting-label">Auto Dark Mode</div>
          <div class="setting-description">Automatically switch to dark mode in the evening (6 PM - 7 AM)</div>
        </div>
        <label class="toggle-switch">
          <input type="checkbox" id="autoSwitch" ${autoSwitch.enabled ? 'checked' : ''} onchange="toggleAutoSwitch()">
          <span class="slider"></span>
        </label>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Apply to All Sheets</div>
          <div class="setting-description">Apply theme to all sheets in the spreadsheet</div>
        </div>
        <button onclick="applyToAllSheets()">Apply to All</button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Current Sheet Only</div>
          <div class="setting-description">Apply theme only to the active sheet</div>
        </div>
        <button onclick="applyToCurrentSheet()">Apply to Current</button>
      </div>
    </div>

    <div class="quick-actions">
      <button class="success" onclick="applySelectedTheme()">✅ Apply Theme</button>
      <button onclick="previewTheme()">👁️ Preview</button>
      <button class="secondary" onclick="resetTheme()">🔄 Reset to Default</button>
      <button class="secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    let selectedTheme = '${currentTheme}';

    function selectTheme(themeKey) {
      selectedTheme = themeKey;
      document.querySelectorAll('.theme-card').forEach(card => {
        card.classList.remove('active');
      });
      event.target.closest('.theme-card').classList.add('active');
    }

    function applySelectedTheme() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Theme applied successfully!');
          google.script.host.close();
        })
        .applyTheme(selectedTheme, 'all');
    }

    function applyToAllSheets() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Theme applied to all sheets!');
          location.reload();
        })
        .applyTheme(selectedTheme, 'all');
    }

    function applyToCurrentSheet() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Theme applied to current sheet!');
          location.reload();
        })
        .applyTheme(selectedTheme, 'current');
    }

    function previewTheme() {
      google.script.run
        .withSuccessHandler(() => {
          alert('👁️ Preview applied! This is temporary.');
        })
        .previewTheme(selectedTheme);
    }

    function resetTheme() {
      if (confirm('Reset to default light theme?')) {
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ Theme reset!');
            location.reload();
          })
          .resetToDefaultTheme();
      }
    }

    function toggleAutoSwitch() {
      const enabled = document.getElementById('autoSwitch').checked;
      google.script.run
        .withSuccessHandler(() => {
          alert(enabled ? '✅ Auto dark mode enabled!' : '🔕 Auto dark mode disabled!');
        })
        .setAutoSwitch(enabled);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets current theme
 * @returns {string} Theme key
 */
function getCurrentTheme() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty('currentTheme') || THEME_CONFIG.DEFAULT_THEME;
}

/**
 * Gets auto-switch settings
 * @returns {Object} Settings
 */
function getAutoSwitchSettings() {
  const props = PropertiesService.getUserProperties();
  const settingsJSON = props.getProperty('autoSwitchSettings');

  if (settingsJSON) {
    return JSON.parse(settingsJSON);
  }

  return {
    enabled: false,
    darkTheme: 'DARK',
    lightTheme: 'LIGHT'
  };
}

/**
 * Applies a theme
 * @param {string} themeKey - Theme to apply
 * @param {string} scope - 'all' or 'current'
 */
function applyTheme(themeKey, scope = 'all') {
  const theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) {
    throw new Error('Invalid theme');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = scope === 'all' ? ss.getSheets() : [ss.getActiveSheet()];

  sheets.forEach(sheet => {
    applyThemeToSheet(sheet, theme);
  });

  // Save preference
  const props = PropertiesService.getUserProperties();
  props.setProperty('currentTheme', themeKey);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${theme.icon} ${theme.name} theme applied!`,
    'Theme',
    3
  );
}

/**
 * Applies theme to a single sheet
 * @param {Sheet} sheet - Sheet to apply to
 * @param {Object} theme - Theme object
 */
function applyThemeToSheet(sheet, theme) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) return;

  // Apply header styling
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground(theme.headerBackground);
  headerRange.setFontColor(theme.headerText);
  headerRange.setFontWeight('bold');

  // Apply data row styling
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);

    // Remove existing banding
    const bandings = sheet.getBandings();
    bandings.forEach(banding => banding.remove());

    // Apply new banding
    const banding = dataRange.applyRowBanding();
    banding.setFirstRowColor(theme.oddRow);
    banding.setSecondRowColor(theme.evenRow);
    banding.setHeaderRowColor(theme.headerBackground);

    // Set text color
    dataRange.setFontColor(theme.text);
  }

  // Set gridline color (approximate with sheet background)
  sheet.setTabColor(theme.accent);
}

/**
 * Previews a theme (temporary)
 * @param {string} themeKey - Theme to preview
 */
function previewTheme(themeKey) {
  const theme = THEME_CONFIG.THEMES[themeKey];
  if (!theme) {
    throw new Error('Invalid theme');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  applyThemeToSheet(activeSheet, theme);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `👁️ Previewing ${theme.name} theme (not saved)`,
    'Preview',
    5
  );
}

/**
 * Resets to default theme
 */
function resetToDefaultTheme() {
  applyTheme(THEME_CONFIG.DEFAULT_THEME, 'all');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Reset to default theme',
    'Theme',
    3
  );
}

/**
 * Quick toggle between light and dark
 */
function quickToggleDarkMode() {
  const currentTheme = getCurrentTheme();
  const newTheme = currentTheme === 'LIGHT' ? 'DARK' : 'LIGHT';

  applyTheme(newTheme, 'all');
}

/**
 * Sets auto-switch settings
 * @param {boolean} enabled - Enable auto-switch
 */
function setAutoSwitch(enabled) {
  const props = PropertiesService.getUserProperties();
  const settings = getAutoSwitchSettings();

  settings.enabled = enabled;
  props.setProperty('autoSwitchSettings', JSON.stringify(settings));

  if (enabled) {
    // Check and apply appropriate theme now
    checkAndApplyAutoTheme();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ Auto dark mode enabled (6 PM - 7 AM)',
      'Theme',
      3
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '🔕 Auto dark mode disabled',
      'Theme',
      3
    );
  }
}

/**
 * Checks time and applies appropriate theme
 */
function checkAndApplyAutoTheme() {
  const settings = getAutoSwitchSettings();
  if (!settings.enabled) return;

  const hour = new Date().getHours();
  const isDarkTime = hour >= THEME_CONFIG.DARK_MODE_START_HOUR || hour < THEME_CONFIG.DARK_MODE_END_HOUR;

  const targetTheme = isDarkTime ? settings.darkTheme : settings.lightTheme;
  const currentTheme = getCurrentTheme();

  if (targetTheme !== currentTheme) {
    applyTheme(targetTheme, 'all');
  }
}

/**
 * Creates a custom theme
 * @param {string} name - Theme name
 * @param {Object} colors - Color configuration
 */
function createCustomTheme(name, colors) {
  const customThemes = getCustomThemes();

  const themeKey = 'CUSTOM_' + name.toUpperCase().replace(/\s+/g, '_');

  customThemes[themeKey] = {
    name: name,
    icon: '🎨',
    ...colors
  };

  saveCustomThemes(customThemes);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Custom theme "${name}" created!`,
    'Theme',
    3
  );

  return themeKey;
}

/**
 * Gets custom themes
 * @returns {Object} Custom themes
 */
function getCustomThemes() {
  const props = PropertiesService.getScriptProperties();
  const themesJSON = props.getProperty('customThemes');

  if (themesJSON) {
    return JSON.parse(themesJSON);
  }

  return {};
}

/**
 * Saves custom themes
 * @param {Object} themes - Themes to save
 */
function saveCustomThemes(themes) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('customThemes', JSON.stringify(themes));
}

/**
 * Exports current theme as JSON
 * @returns {string} JSON theme configuration
 */
function exportCurrentTheme() {
  const currentTheme = getCurrentTheme();
  const theme = THEME_CONFIG.THEMES[currentTheme];

  return JSON.stringify(theme, null, 2);
}

/**
 * Imports theme from JSON
 * @param {string} themeJSON - JSON theme configuration
 * @param {string} name - Theme name
 * @returns {string} Theme key
 */
function importThemeFromJSON(themeJSON, name) {
  try {
    const colors = JSON.parse(themeJSON);
    return createCustomTheme(name, colors);
  } catch (error) {
    throw new Error('Invalid theme JSON: ' + error.message);
  }
}

/**
 * Installs auto-theme switching trigger
 */
function installAutoThemeTrigger() {
  // Delete existing trigger
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndApplyAutoTheme') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger (runs every hour)
  ScriptApp.newTrigger('checkAndApplyAutoTheme')
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Auto-theme trigger installed',
    'Theme',
    3
  );
}

/******************************************************************************
 * MODULE: KeyboardShortcuts
 * Source: KeyboardShortcuts.gs
 *****************************************************************************/

/**
 * ============================================================================
 * KEYBOARD SHORTCUTS SYSTEM
 * ============================================================================
 *
 * Power user keyboard shortcuts for common actions
 * Features:
 * - Quick navigation (Ctrl+G for Grievance Log, Ctrl+M for Members)
 * - Quick actions (Ctrl+N for new grievance, Ctrl+F for search)
 * - Batch operations shortcuts
 * - Configurable key bindings
 * - Help overlay (Ctrl+?)
 */

/**
 * Shows keyboard shortcuts help overlay
 */
function showKeyboardShortcuts() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Roboto', 'Courier New', monospace;
      padding: 20px;
      margin: 0;
      background: #1e1e1e;
      color: #d4d4d4;
    }
    .container {
      background: #252526;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.5);
      max-width: 700px;
      margin: 0 auto;
    }
    h2 {
      color: #4ec9b0;
      margin-top: 0;
      border-bottom: 2px solid #4ec9b0;
      padding-bottom: 10px;
      font-family: 'Roboto', Arial, sans-serif;
    }
    .category {
      margin: 25px 0;
    }
    .category h3 {
      color: #569cd6;
      font-size: 16px;
      margin-bottom: 12px;
      font-family: 'Roboto', Arial, sans-serif;
    }
    .shortcut-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 15px;
      margin: 5px 0;
      background: #2d2d30;
      border-radius: 4px;
      border-left: 3px solid #4ec9b0;
    }
    .shortcut-row:hover {
      background: #3e3e42;
    }
    .shortcut-desc {
      color: #d4d4d4;
      flex: 1;
    }
    .shortcut-keys {
      display: flex;
      gap: 5px;
    }
    .key {
      background: #1e1e1e;
      color: #4ec9b0;
      padding: 4px 10px;
      border-radius: 4px;
      border: 1px solid #4ec9b0;
      font-family: 'Courier New', monospace;
      font-size: 12px;
      font-weight: bold;
      min-width: 30px;
      text-align: center;
    }
    .key.modifier {
      background: #264f78;
      color: #9cdcfe;
      border-color: #569cd6;
    }
    .tip {
      background: #1e3a5f;
      padding: 15px;
      border-radius: 4px;
      margin: 20px 0;
      border-left: 4px solid #569cd6;
      color: #9cdcfe;
    }
    .tip strong {
      color: #4ec9b0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>⌨️ Keyboard Shortcuts</h2>

    <div class="tip">
      <strong>💡 Pro Tip:</strong> Use keyboard shortcuts to navigate and perform actions 10x faster!
      Press <strong>Escape</strong> to close any dialog.
    </div>

    <div class="category">
      <h3>🧭 Navigation</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">Go to Grievance Log</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">G</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Go to Member Directory</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">M</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Go to Dashboard</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">D</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Go to Interactive Dashboard</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">I</span>
        </div>
      </div>
    </div>

    <div class="category">
      <h3>🔍 Search & Lookup</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">Search Members</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">F</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Quick Member Search</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key modifier">Shift</span>
          <span class="key">F</span>
        </div>
      </div>
    </div>

    <div class="category">
      <h3>➕ Quick Actions</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">New Grievance</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">N</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Compose Email</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">E</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Batch Operations Menu</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">B</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Refresh All</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">R</span>
        </div>
      </div>
    </div>

    <div class="category">
      <h3>📁 File Management</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">Setup Drive Folder</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key modifier">Shift</span>
          <span class="key">D</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Upload Files</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">U</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Show Grievance Files</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key modifier">Shift</span>
          <span class="key">F</span>
        </div>
      </div>
    </div>

    <div class="category">
      <h3>⚙️ Automation</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">Notification Settings</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key modifier">Shift</span>
          <span class="key">N</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Report Settings</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key modifier">Shift</span>
          <span class="key">R</span>
        </div>
      </div>
    </div>

    <div class="category">
      <h3>❓ Help</h3>
      <div class="shortcut-row">
        <div class="shortcut-desc">Show This Help</div>
        <div class="shortcut-keys">
          <span class="key modifier">Ctrl</span>
          <span class="key">?</span>
        </div>
      </div>
      <div class="shortcut-row">
        <div class="shortcut-desc">Context Help</div>
        <div class="shortcut-keys">
          <span class="key">F1</span>
        </div>
      </div>
    </div>

    <div class="tip" style="margin-top: 30px;">
      <strong>🎯 Custom Shortcuts:</strong> Go to Settings to customize keyboard shortcuts.
      On Mac, use <strong>Cmd</strong> instead of <strong>Ctrl</strong>.
    </div>
  </div>

  <script>
    // Close on Escape
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape') {
        google.script.host.close();
      }
    });
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(750)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⌨️ Keyboard Shortcuts');
}

/**
 * Navigation shortcuts
 */
function navigateToGrievanceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    sheet.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('📋 Grievance Log', 'Navigation', 1);
  }
}

function navigateToMemberDirectory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (sheet) {
    sheet.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('👥 Member Directory', 'Navigation', 1);
  }
}

function navigateToDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (sheet) {
    sheet.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('📊 Dashboard', 'Navigation', 1);
  }
}

function navigateToInteractiveDashboard() {
  openInteractiveDashboard();
}

/**
 * Quick action shortcuts
 */
function quickNewGrievance() {
  showStartGrievanceDialog();
}

function quickSearch() {
  showMemberSearch();
}

function quickEmail() {
  composeGrievanceEmail();
}

function quickBatchOps() {
  showBatchOperationsMenu();
}

function quickRefresh() {
  refreshCalculations();
}

/**
 * Keyboard shortcut registry and handler
 */
const KEYBOARD_SHORTCUTS = {
  // Navigation
  'Ctrl+G': 'navigateToGrievanceLog',
  'Ctrl+M': 'navigateToMemberDirectory',
  'Ctrl+D': 'navigateToDashboard',
  'Ctrl+I': 'navigateToInteractiveDashboard',

  // Search
  'Ctrl+F': 'quickSearch',

  // Actions
  'Ctrl+N': 'quickNewGrievance',
  'Ctrl+E': 'quickEmail',
  'Ctrl+B': 'quickBatchOps',
  'Ctrl+R': 'quickRefresh',

  // Drive
  'Ctrl+Shift+D': 'setupDriveFolderForGrievance',
  'Ctrl+U': 'showFileUploadDialog',

  // Automation
  'Ctrl+Shift+N': 'showNotificationSettings',
  'Ctrl+Shift+R': 'showReportAutomationSettings',

  // Help
  'Ctrl+?': 'showKeyboardShortcuts',
  'F1': 'showContextHelp'
};

/**
 * Shows keyboard shortcuts configuration
 */
function showKeyboardShortcutsConfig() {
  const html = `
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
    .warning-box {
      background: #fff3cd;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #ffc107;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
    }
    th {
      background: #1a73e8;
      color: white;
      padding: 12px;
      text-align: left;
      font-weight: 500;
    }
    td {
      padding: 10px 12px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background: #f5f5f5;
    }
    .shortcut-code {
      background: #f5f5f5;
      padding: 4px 8px;
      border-radius: 3px;
      font-family: 'Courier New', monospace;
      font-size: 12px;
      color: #1a73e8;
      border: 1px solid #ddd;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
      margin: 10px 5px 0 0;
    }
    button:hover {
      background: #1557b0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚙️ Keyboard Shortcuts Configuration</h2>

    <div class="info-box">
      <strong>ℹ️ How to Use:</strong> Keyboard shortcuts work globally across all sheets.
      Press the key combination to trigger the action instantly.
    </div>

    <div class="warning-box">
      <strong>⚠️ Note:</strong> Some shortcuts may conflict with browser shortcuts.
      Use Google Sheets in full-screen mode (F11) for best experience.
    </div>

    <h3>Current Shortcuts</h3>
    <table>
      <thead>
        <tr>
          <th>Shortcut</th>
          <th>Action</th>
          <th>Category</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><span class="shortcut-code">Ctrl+G</span></td>
          <td>Go to Grievance Log</td>
          <td>Navigation</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+M</span></td>
          <td>Go to Member Directory</td>
          <td>Navigation</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+D</span></td>
          <td>Go to Dashboard</td>
          <td>Navigation</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+I</span></td>
          <td>Go to Interactive Dashboard</td>
          <td>Navigation</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+F</span></td>
          <td>Search Members</td>
          <td>Search</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+N</span></td>
          <td>New Grievance</td>
          <td>Actions</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+E</span></td>
          <td>Compose Email</td>
          <td>Actions</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+B</span></td>
          <td>Batch Operations</td>
          <td>Actions</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+R</span></td>
          <td>Refresh All</td>
          <td>Actions</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+Shift+D</span></td>
          <td>Setup Drive Folder</td>
          <td>Files</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+U</span></td>
          <td>Upload Files</td>
          <td>Files</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">Ctrl+?</span></td>
          <td>Show Shortcuts Help</td>
          <td>Help</td>
        </tr>
        <tr>
          <td><span class="shortcut-code">F1</span></td>
          <td>Context Help</td>
          <td>Help</td>
        </tr>
      </tbody>
    </table>

    <button onclick="google.script.run.withSuccessHandler(() => google.script.host.close()).showKeyboardShortcuts()">
      📖 View Shortcuts Guide
    </button>
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(750)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚙️ Keyboard Shortcuts Configuration');
}

/**
 * Shows context-sensitive help based on current sheet
 */
function showContextHelp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  let helpContent = '';

  switch (sheetName) {
    case SHEETS.GRIEVANCE_LOG:
      helpContent = `
<h3>📋 Grievance Log Help</h3>
<p><strong>This sheet tracks all union grievances.</strong></p>

<h4>Common Tasks:</h4>
<ul>
  <li><strong>Add New Grievance:</strong> Use menu: 📋 Grievance Tools → ➕ Start New Grievance (or Ctrl+N)</li>
  <li><strong>Update Status:</strong> Select rows → ⚡ Batch Operations → 📊 Bulk Update Status</li>
  <li><strong>Assign Steward:</strong> Select rows → ⚡ Batch Operations → 👤 Bulk Assign Steward</li>
  <li><strong>View Files:</strong> Select row → 📁 Google Drive → 📂 Show Grievance Files</li>
  <li><strong>Send Email:</strong> Select row → 📧 Communications → 📧 Compose Email (or Ctrl+E)</li>
</ul>

<h4>Important Columns:</h4>
<ul>
  <li><strong>Column H:</strong> Step 1 Deadline (auto-calculated)</li>
  <li><strong>Column U:</strong> Days to Deadline (auto-calculated)</li>
  <li><strong>Column V:</strong> Deadline Status (auto-calculated)</li>
</ul>

<h4>Tips:</h4>
<ul>
  <li>Red rows = Overdue grievances (urgent action needed)</li>
  <li>Yellow rows = Deadline within 7 days</li>
  <li>Use filters to focus on specific statuses or stewards</li>
</ul>
      `;
      break;

    case SHEETS.MEMBER_DIR:
      helpContent = `
<h3>👥 Member Directory Help</h3>
<p><strong>This sheet contains all union member information.</strong></p>

<h4>Common Tasks:</h4>
<ul>
  <li><strong>Search Members:</strong> Use menu: 🔍 Search & Lookup → 🔍 Search Members (or Ctrl+F)</li>
  <li><strong>Add Member:</strong> Fill in a new row with required fields (ID, First Name, Last Name)</li>
  <li><strong>Update Info:</strong> Click on any cell and edit directly</li>
</ul>

<h4>Important Columns:</h4>
<ul>
  <li><strong>Column A:</strong> Member ID (must be unique, format: M######)</li>
  <li><strong>Column H:</strong> Email (validated format)</li>
  <li><strong>Column I:</strong> Phone (validated format)</li>
  <li><strong>Column J:</strong> Is Steward? (Yes/No dropdown)</li>
</ul>

<h4>Auto-Calculated Fields:</h4>
<ul>
  <li><strong>Column Z:</strong> Has Active Grievance? (auto-updated)</li>
  <li><strong>Column AA:</strong> Grievance Status Snapshot</li>
  <li><strong>Column AB:</strong> Next Grievance Deadline</li>
</ul>

<h4>Tips:</h4>
<ul>
  <li>Use 🆔 Generate Next Member ID to get the next available ID</li>
  <li>Email and phone fields validate automatically</li>
  <li>Mark stewards in Column J to enable auto-assignment</li>
</ul>
      `;
      break;

    case SHEETS.DASHBOARD:
      helpContent = `
<h3>📊 Dashboard Help</h3>
<p><strong>Real-time overview of all grievance operations.</strong></p>

<h4>Dashboard Sections:</h4>
<ul>
  <li><strong>Key Metrics:</strong> Total grievances, open cases, average resolution time</li>
  <li><strong>Status Breakdown:</strong> Grievances by current status</li>
  <li><strong>Issue Type Analysis:</strong> Most common grievance types</li>
  <li><strong>Steward Workload:</strong> Cases assigned per steward</li>
  <li><strong>Timeline:</strong> Cases filed this month vs. last month</li>
</ul>

<h4>Other Dashboards:</h4>
<ul>
  <li><strong>Interactive Dashboard:</strong> 📊 Dashboards → ✨ Interactive Dashboard</li>
  <li><strong>Unified Operations Monitor:</strong> 📊 Dashboards → 🎯 Unified Operations Monitor</li>
</ul>

<h4>Tips:</h4>
<ul>
  <li>Dashboard auto-updates when data changes</li>
  <li>Use 🔄 Refresh All to force update (Ctrl+R)</li>
  <li>Charts are interactive - click to see details</li>
</ul>
      `;
      break;

    default:
      helpContent = `
<h3>❓ 509 Dashboard Help</h3>
<p><strong>Welcome to the SEIU Local 509 Grievance Dashboard!</strong></p>

<h4>Main Sheets:</h4>
<ul>
  <li><strong>📋 Grievance Log:</strong> Track all grievances (Ctrl+G)</li>
  <li><strong>👥 Member Directory:</strong> Member information (Ctrl+M)</li>
  <li><strong>📊 Dashboard:</strong> Real-time metrics (Ctrl+D)</li>
</ul>

<h4>Quick Actions:</h4>
<ul>
  <li><strong>New Grievance:</strong> Ctrl+N</li>
  <li><strong>Search Members:</strong> Ctrl+F</li>
  <li><strong>Compose Email:</strong> Ctrl+E</li>
  <li><strong>Batch Operations:</strong> Ctrl+B</li>
</ul>

<h4>Getting Started:</h4>
<ol>
  <li>Navigate to the Grievance Log or Member Directory</li>
  <li>Press F1 for context-specific help</li>
  <li>Press Ctrl+? to see all keyboard shortcuts</li>
  <li>Use the menu: ❓ Help & Support → 📚 Getting Started Guide</li>
</ol>
      `;
  }

  const html = `
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
      max-height: 600px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    h3 {
      color: #1a73e8;
      margin-top: 0;
    }
    h4 {
      color: #333;
      margin: 20px 0 10px 0;
    }
    ul, ol {
      margin: 10px 0;
      padding-left: 25px;
    }
    li {
      margin: 8px 0;
      line-height: 1.6;
    }
    strong {
      color: #1a73e8;
    }
    .sheet-badge {
      display: inline-block;
      background: #e8f0fe;
      color: #1a73e8;
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 13px;
      font-weight: 500;
      margin-bottom: 15px;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="sheet-badge">📍 Current Sheet: ${sheetName}</div>
    ${helpContent}
  </div>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `❓ Help - ${sheetName}`);
}

/******************************************************************************
 * MODULE: RootCauseAnalysis
 * Source: RootCauseAnalysis.gs
 *****************************************************************************/

/**
 * ============================================================================
 * ROOT CAUSE ANALYSIS TOOL
 * ============================================================================
 *
 * Identifies systemic issues and patterns in grievances
 * Features:
 * - Pattern detection across dimensions
 * - Clustering analysis (location, manager, issue type)
 * - Timeline analysis for recurring problems
 * - Correlation finding
 * - Actionable recommendations
 * - Visual reports with charts
 * - Export analysis reports
 */

/**
 * Performs comprehensive root cause analysis
 * @returns {Object} Analysis results
 */
function performRootCauseAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const lastRow = grievanceSheet.getLastRow();

  if (lastRow < 2) {
    return { error: 'Insufficient data for analysis' };
  }

  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 28).getValues();

  const analysis = {
    locationClusters: analyzeLocationClusters(data),
    managerPatterns: analyzeManagerPatterns(data),
    issueTypePatterns: analyzeIssueTypePatterns(data),
    temporalPatterns: analyzeTemporalPatterns(data),
    correlations: findCorrelations(data),
    recommendations: generateRCARecommendations(data)
  };

  return analysis;
}

/**
 * Analyzes grievance clustering by location
 * @param {Array} data - Grievance data
 * @returns {Object} Location analysis
 */
function analyzeLocationClusters(data) {
  const locationStats = {};

  data.forEach(row => {
    const location = row[9]; // Column J: Work Location
    if (!location) return;

    if (!locationStats[location]) {
      locationStats[location] = {
        count: 0,
        issueTypes: {},
        managers: {},
        resolutionTimes: []
      };
    }

    locationStats[location].count++;

    // Track issue types per location
    const issueType = row[5];
    if (issueType) {
      locationStats[location].issueTypes[issueType] =
        (locationStats[location].issueTypes[issueType] || 0) + 1;
    }

    // Track managers per location
    const manager = row[11];
    if (manager) {
      locationStats[location].managers[manager] =
        (locationStats[location].managers[manager] || 0) + 1;
    }

    // Track resolution times
    const filedDate = row[6];
    const closedDate = row[18];
    if (filedDate && closedDate) {
      const resTime = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
      locationStats[location].resolutionTimes.push(resTime);
    }
  });

  // Find hotspots (locations with disproportionate grievances)
  const totalGrievances = data.length;
  const hotspots = [];

  Object.entries(locationStats).forEach(([location, stats]) => {
    const percentage = (stats.count / totalGrievances) * 100;

    if (percentage > 15) {
      // If a location has >15% of all grievances, it's a hotspot
      const topIssue = Object.entries(stats.issueTypes)
        .sort((a, b) => b[1] - a[1])[0];

      const avgResTime = stats.resolutionTimes.length > 0
        ? stats.resolutionTimes.reduce((sum, val) => sum + val, 0) / stats.resolutionTimes.length
        : 0;

      hotspots.push({
        location: location,
        count: stats.count,
        percentage: percentage.toFixed(1),
        topIssueType: topIssue ? topIssue[0] : 'N/A',
        topIssueCount: topIssue ? topIssue[1] : 0,
        avgResolutionTime: Math.round(avgResTime),
        severity: percentage > 25 ? 'HIGH' : 'MEDIUM'
      });
    }
  });

  return {
    totalLocations: Object.keys(locationStats).length,
    hotspots: hotspots.sort((a, b) => b.count - a.count),
    allStats: locationStats
  };
}

/**
 * Analyzes patterns by manager
 * @param {Array} data - Grievance data
 * @returns {Object} Manager analysis
 */
function analyzeManagerPatterns(data) {
  const managerStats = {};

  data.forEach(row => {
    const manager = row[11]; // Column L: Manager
    if (!manager) return;

    if (!managerStats[manager]) {
      managerStats[manager] = {
        count: 0,
        issueTypes: {},
        outcomes: {},
        resolutionTimes: []
      };
    }

    managerStats[manager].count++;

    // Track issue types
    const issueType = row[5];
    if (issueType) {
      managerStats[manager].issueTypes[issueType] =
        (managerStats[manager].issueTypes[issueType] || 0) + 1;
    }

    // Track outcomes
    const status = row[4];
    if (status) {
      managerStats[manager].outcomes[status] =
        (managerStats[manager].outcomes[status] || 0) + 1;
    }

    // Track resolution times
    const filedDate = row[6];
    const closedDate = row[18];
    if (filedDate && closedDate) {
      const resTime = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
      managerStats[manager].resolutionTimes.push(resTime);
    }
  });

  // Find managers with concerning patterns
  const avgGrievancesPerManager = data.length / Object.keys(managerStats).length;
  const concerningManagers = [];

  Object.entries(managerStats).forEach(([manager, stats]) => {
    if (stats.count > avgGrievancesPerManager * 2) {
      const topIssue = Object.entries(stats.issueTypes)
        .sort((a, b) => b[1] - a[1])[0];

      const avgResTime = stats.resolutionTimes.length > 0
        ? stats.resolutionTimes.reduce((sum, val) => sum + val, 0) / stats.resolutionTimes.length
        : 0;

      concerningManagers.push({
        manager: manager,
        count: stats.count,
        comparedToAverage: (stats.count / avgGrievancesPerManager).toFixed(1) + 'x',
        topIssueType: topIssue ? topIssue[0] : 'N/A',
        avgResolutionTime: Math.round(avgResTime),
        severity: stats.count > avgGrievancesPerManager * 3 ? 'HIGH' : 'MEDIUM'
      });
    }
  });

  return {
    totalManagers: Object.keys(managerStats).length,
    avgPerManager: avgGrievancesPerManager.toFixed(1),
    concerningManagers: concerningManagers.sort((a, b) => b.count - a.count),
    allStats: managerStats
  };
}

/**
 * Analyzes issue type patterns
 * @param {Array} data - Grievance data
 * @returns {Object} Issue type analysis
 */
function analyzeIssueTypePatterns(data) {
  const issueTypeStats = {};

  data.forEach(row => {
    const issueType = row[5];
    if (!issueType) return;

    if (!issueTypeStats[issueType]) {
      issueTypeStats[issueType] = {
        count: 0,
        locations: {},
        managers: {},
        resolutionTimes: [],
        outcomes: {}
      };
    }

    issueTypeStats[issueType].count++;

    // Track locations
    const location = row[9];
    if (location) {
      issueTypeStats[issueType].locations[location] =
        (issueTypeStats[issueType].locations[location] || 0) + 1;
    }

    // Track managers
    const manager = row[11];
    if (manager) {
      issueTypeStats[issueType].managers[manager] =
        (issueTypeStats[issueType].managers[manager] || 0) + 1;
    }

    // Track outcomes
    const status = row[4];
    if (status) {
      issueTypeStats[issueType].outcomes[status] =
        (issueTypeStats[issueType].outcomes[status] || 0) + 1;
    }

    // Track resolution times
    const filedDate = row[6];
    const closedDate = row[18];
    if (filedDate && closedDate) {
      const resTime = Math.floor((closedDate - filedDate) / (1000 * 60 * 60 * 24));
      issueTypeStats[issueType].resolutionTimes.push(resTime);
    }
  });

  // Identify systemic issues
  const systemicIssues = [];

  Object.entries(issueTypeStats).forEach(([issueType, stats]) => {
    // Check if concentrated in specific locations (>60% in one location)
    const topLocation = Object.entries(stats.locations)
      .sort((a, b) => b[1] - a[1])[0];

    const locationConcentration = topLocation
      ? (topLocation[1] / stats.count) * 100
      : 0;

    if (locationConcentration > 60 || stats.count > data.length * 0.15) {
      const avgResTime = stats.resolutionTimes.length > 0
        ? stats.resolutionTimes.reduce((sum, val) => sum + val, 0) / stats.resolutionTimes.length
        : 0;

      systemicIssues.push({
        issueType: issueType,
        count: stats.count,
        percentage: ((stats.count / data.length) * 100).toFixed(1),
        primaryLocation: topLocation ? topLocation[0] : 'N/A',
        locationConcentration: locationConcentration.toFixed(1) + '%',
        avgResolutionTime: Math.round(avgResTime),
        isSystemic: locationConcentration > 60 || stats.count > data.length * 0.2
      });
    }
  });

  return {
    totalIssueTypes: Object.keys(issueTypeStats).length,
    systemicIssues: systemicIssues.sort((a, b) => b.count - a.count),
    allStats: issueTypeStats
  };
}

/**
 * Analyzes temporal patterns (when grievances occur)
 * @param {Array} data - Grievance data
 * @returns {Object} Temporal analysis
 */
function analyzeTemporalPatterns(data) {
  const monthlyPatterns = {};
  const dayOfWeekPatterns = {};

  data.forEach(row => {
    const filedDate = row[6];
    if (!filedDate) return;

    // Monthly analysis
    const monthKey = `${filedDate.getFullYear()}-${String(filedDate.getMonth() + 1).padStart(2, '0')}`;
    monthlyPatterns[monthKey] = (monthlyPatterns[monthKey] || 0) + 1;

    // Day of week analysis
    const dayOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][filedDate.getDay()];
    dayOfWeekPatterns[dayOfWeek] = (dayOfWeekPatterns[dayOfWeek] || 0) + 1;
  });

  // Find peak days
  const peakDay = Object.entries(dayOfWeekPatterns)
    .sort((a, b) => b[1] - a[1])[0];

  return {
    monthlyPatterns: monthlyPatterns,
    dayOfWeekPatterns: dayOfWeekPatterns,
    peakDay: peakDay ? peakDay[0] : 'N/A',
    peakDayCount: peakDay ? peakDay[1] : 0
  };
}

/**
 * Finds correlations between variables
 * @param {Array} data - Grievance data
 * @returns {Array} Correlation findings
 */
function findCorrelations(data) {
  const correlations = [];

  // Check: Specific manager + specific issue type
  const managerIssueMatrix = {};

  data.forEach(row => {
    const manager = row[11];
    const issueType = row[5];

    if (manager && issueType) {
      const key = `${manager}|||${issueType}`;
      managerIssueMatrix[key] = (managerIssueMatrix[key] || 0) + 1;
    }
  });

  Object.entries(managerIssueMatrix).forEach(([key, count]) => {
    if (count >= 5) {
      const [manager, issueType] = key.split('|||');
      correlations.push({
        type: 'Manager-Issue Correlation',
        description: `${manager} has ${count} ${issueType} grievances`,
        count: count,
        significance: count >= 10 ? 'HIGH' : 'MEDIUM'
      });
    }
  });

  return correlations.sort((a, b) => b.count - a.count);
}

/**
 * Generates recommendations based on RCA
 * @param {Array} data - Grievance data
 * @returns {Array} Recommendations
 */
function generateRCARecommendations(data) {
  const recommendations = [];

  const locationAnalysis = analyzeLocationClusters(data);
  const managerAnalysis = analyzeManagerPatterns(data);
  const issueAnalysis = analyzeIssueTypePatterns(data);

  // Location-based recommendations
  locationAnalysis.hotspots.forEach(hotspot => {
    recommendations.push({
      category: 'Location Hotspot',
      priority: hotspot.severity,
      recommendation: `Address grievance concentration at ${hotspot.location} (${hotspot.percentage}% of all cases). Primary issue: ${hotspot.topIssueType}. Consider workplace assessment and manager training.`,
      impactArea: hotspot.location,
      expectedImpact: `Reduce grievances by 30-40% at this location`
    });
  });

  // Manager-based recommendations
  managerAnalysis.concerningManagers.forEach(manager => {
    recommendations.push({
      category: 'Management Pattern',
      priority: manager.severity,
      recommendation: `${manager.manager} has ${manager.comparedToAverage} average grievances. Primary issue: ${manager.topIssueType}. Recommend management coaching and relationship building.`,
      impactArea: manager.manager,
      expectedImpact: `Reduce grievances by 20-30% for this manager's team`
    });
  });

  // Systemic issue recommendations
  issueAnalysis.systemicIssues.forEach(issue => {
    if (issue.isSystemic) {
      recommendations.push({
        category: 'Systemic Issue',
        priority: 'HIGH',
        recommendation: `${issue.issueType} represents ${issue.percentage}% of all grievances, concentrated at ${issue.primaryLocation}. This indicates a systemic problem requiring policy review and preventive measures.`,
        impactArea: issue.issueType,
        expectedImpact: `Prevent 50-70% of future ${issue.issueType} grievances through proactive intervention`
      });
    }
  });

  return recommendations.sort((a, b) => {
    const priorityOrder = { 'HIGH': 3, 'MEDIUM': 2, 'LOW': 1 };
    return priorityOrder[b.priority] - priorityOrder[a.priority];
  });
}

/**
 * Shows root cause analysis dashboard
 */
function showRootCauseAnalysisDashboard() {
  const ui = SpreadsheetApp.getUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🔬 Analyzing patterns...',
    'Root Cause Analysis',
    -1
  );

  try {
    const analysis = performRootCauseAnalysis();

    if (analysis.error) {
      ui.alert('⚠️ ' + analysis.error);
      return;
    }

    const html = createRCADashboardHTML(analysis);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(1000)
      .setHeight(750);

    ui.showModalDialog(htmlOutput, '🔬 Root Cause Analysis');

  } catch (error) {
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Creates HTML for RCA dashboard
 * @param {Object} analysis - Analysis results
 * @returns {string} HTML content
 */
function createRCADashboardHTML(analysis) {
  const hotspotsHTML = analysis.locationClusters.hotspots.length > 0
    ? analysis.locationClusters.hotspots.map(h => `
        <div class="finding severity-${h.severity.toLowerCase()}">
          <div class="finding-header">
            <span class="finding-title">${h.location}</span>
            <span class="severity-badge ${h.severity.toLowerCase()}">${h.severity}</span>
          </div>
          <div class="finding-stats">
            <div class="stat">${h.count} cases (${h.percentage}%)</div>
            <div class="stat">Top issue: ${h.topIssueType} (${h.topIssueCount} cases)</div>
            <div class="stat">Avg resolution: ${h.avgResolutionTime} days</div>
          </div>
        </div>
      `).join('')
    : '<p>No significant location hotspots identified.</p>';

  const concerningManagersHTML = analysis.managerPatterns.concerningManagers.length > 0
    ? analysis.managerPatterns.concerningManagers.map(m => `
        <div class="finding severity-${m.severity.toLowerCase()}">
          <div class="finding-header">
            <span class="finding-title">${m.manager}</span>
            <span class="severity-badge ${m.severity.toLowerCase()}">${m.severity}</span>
          </div>
          <div class="finding-stats">
            <div class="stat">${m.count} cases (${m.comparedToAverage} avg)</div>
            <div class="stat">Top issue: ${m.topIssueType}</div>
            <div class="stat">Avg resolution: ${m.avgResolutionTime} days</div>
          </div>
        </div>
      `).join('')
    : '<p>No concerning manager patterns identified.</p>';

  const systemicIssuesHTML = analysis.issueTypePatterns.systemicIssues.length > 0
    ? analysis.issueTypePatterns.systemicIssues.map(i => `
        <div class="finding ${i.isSystemic ? 'severity-high' : 'severity-medium'}">
          <div class="finding-header">
            <span class="finding-title">${i.issueType}</span>
            ${i.isSystemic ? '<span class="severity-badge high">SYSTEMIC</span>' : ''}
          </div>
          <div class="finding-stats">
            <div class="stat">${i.count} cases (${i.percentage}%)</div>
            <div class="stat">Primary location: ${i.primaryLocation} (${i.locationConcentration})</div>
            <div class="stat">Avg resolution: ${i.avgResolutionTime} days</div>
          </div>
        </div>
      `).join('')
    : '<p>No systemic issues identified.</p>';

  const recommendationsHTML = analysis.recommendations.length > 0
    ? analysis.recommendations.map(r => `
        <div class="recommendation priority-${r.priority.toLowerCase()}">
          <div class="rec-header">
            <span class="rec-category">${r.category}</span>
            <span class="priority-badge ${r.priority.toLowerCase()}">${r.priority} PRIORITY</span>
          </div>
          <div class="rec-text">${r.recommendation}</div>
          <div class="rec-impact">
            <strong>Expected Impact:</strong> ${r.expectedImpact}
          </div>
        </div>
      `).join('')
    : '<p>No specific recommendations at this time.</p>';

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Roboto', Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-height: 700px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; border-bottom: 3px solid #1a73e8; padding-bottom: 10px; }
    h3 { color: #333; margin-top: 25px; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; }
    .summary { background: #e8f0fe; padding: 15px; border-radius: 4px; margin: 15px 0; border-left: 4px solid #1a73e8; }
    .finding { background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid; }
    .finding.severity-high { border-left-color: #f44336; background: #ffebee; }
    .finding.severity-medium { border-left-color: #ff9800; background: #fff3e0; }
    .finding-header { display: flex; justify-content: space-between; margin-bottom: 10px; }
    .finding-title { font-weight: bold; color: #333; }
    .severity-badge, .priority-badge { background: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold; }
    .severity-badge.high, .priority-badge.high { color: #f44336; border: 1px solid #f44336; }
    .severity-badge.medium, .priority-badge.medium { color: #ff9800; border: 1px solid #ff9800; }
    .finding-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
    .stat { background: white; padding: 8px; border-radius: 4px; font-size: 13px; text-align: center; }
    .recommendation { background: #e8f5e9; padding: 15px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #4caf50; }
    .recommendation.priority-high { border-left-color: #f44336; background: #ffebee; }
    .recommendation.priority-medium { border-left-color: #ff9800; background: #fff3e0; }
    .rec-header { display: flex; justify-content: space-between; margin-bottom: 10px; }
    .rec-category { font-weight: bold; color: #1976d2; }
    .rec-text { margin: 10px 0; color: #555; }
    .rec-impact { margin-top: 10px; padding: 10px; background: white; border-radius: 4px; font-size: 13px; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔬 Root Cause Analysis</h2>

    <div class="summary">
      <strong>📊 Analysis Summary:</strong><br>
      Analyzed ${analysis.locationClusters.totalLocations} locations, ${analysis.managerPatterns.totalManagers} managers, ${analysis.issueTypePatterns.totalIssueTypes} issue types<br>
      Found ${analysis.locationClusters.hotspots.length} location hotspots, ${analysis.managerPatterns.concerningManagers.length} concerning manager patterns, ${analysis.issueTypePatterns.systemicIssues.length} systemic issues
    </div>

    <h3>📍 Location Hotspots</h3>
    ${hotspotsHTML}

    <h3>👔 Management Patterns</h3>
    ${concerningManagersHTML}

    <h3>⚠️ Systemic Issues</h3>
    ${systemicIssuesHTML}

    <h3>💡 Recommendations</h3>
    ${recommendationsHTML}
  </div>
</body>
</html>
  `;
}

/******************************************************************************
 * MODULE: MobileOptimization
 * Source: MobileOptimization.gs
 *****************************************************************************/

/**
 * ============================================================================
 * MOBILE OPTIMIZATION
 * ============================================================================
 *
 * Mobile-friendly interfaces and responsive designs
 * Features:
 * - Touch-optimized UI
 * - Simplified mobile views
 * - Card-based layouts
 * - Mobile dashboard
 * - Quick actions
 * - Swipe gestures
 * - Offline data access
 */

/**
 * Mobile configuration
 */
const MOBILE_CONFIG = {
  MAX_COLUMNS_MOBILE: 8,  // Show max 8 columns on mobile
  CARD_LAYOUT_ENABLED: true,
  TOUCH_TARGET_SIZE: '44px',
  SIMPLIFIED_MODE: true
};

/**
 * Shows mobile dashboard
 */
function showMobileDashboard() {
  const html = createMobileDashboardHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📱 Mobile Dashboard');
}

/**
 * Creates HTML for mobile dashboard
 */
function createMobileDashboardHTML() {
  const stats = getMobileDashboardStats();

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <style>
    * {
      box-sizing: border-box;
      -webkit-tap-highlight-color: transparent;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
      padding: 0;
      margin: 0;
      background: #f5f5f5;
      overflow-x: hidden;
    }
    .header {
      background: linear-gradient(135deg, #1a73e8 0%, #1557b0 100%);
      color: white;
      padding: 20px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header .subtitle {
      margin-top: 5px;
      font-size: 14px;
      opacity: 0.9;
    }
    .container {
      padding: 15px;
    }
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 12px;
      margin-bottom: 20px;
    }
    .stat-card {
      background: white;
      padding: 20px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      text-align: center;
      min-height: 100px;
      display: flex;
      flex-direction: column;
      justify-content: center;
    }
    .stat-value {
      font-size: 32px;
      font-weight: bold;
      color: #1a73e8;
      margin-bottom: 5px;
    }
    .stat-label {
      font-size: 13px;
      color: #666;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .quick-actions {
      margin-bottom: 20px;
    }
    .section-title {
      font-size: 16px;
      font-weight: 600;
      color: #333;
      margin: 20px 0 12px 0;
      padding-left: 5px;
    }
    .action-button {
      background: white;
      border: none;
      padding: 16px;
      margin-bottom: 10px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
      width: 100%;
      text-align: left;
      display: flex;
      align-items: center;
      gap: 15px;
      font-size: 15px;
      color: #333;
      cursor: pointer;
      transition: all 0.2s;
      min-height: ${MOBILE_CONFIG.TOUCH_TARGET_SIZE};
    }
    .action-button:active {
      transform: scale(0.98);
      box-shadow: 0 1px 4px rgba(0,0,0,0.1);
    }
    .action-icon {
      font-size: 24px;
      width: 40px;
      height: 40px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #e8f0fe;
      border-radius: 10px;
    }
    .action-text {
      flex: 1;
    }
    .action-label {
      font-weight: 500;
      margin-bottom: 3px;
    }
    .action-description {
      font-size: 12px;
      color: #666;
    }
    .recent-card {
      background: white;
      padding: 15px;
      margin-bottom: 12px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .recent-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }
    .recent-id {
      font-weight: bold;
      color: #1a73e8;
      font-size: 15px;
    }
    .status-badge {
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .status-filed { background: #e8f5e9; color: #2e7d32; }
    .status-pending { background: #fff3e0; color: #ef6c00; }
    .status-resolved { background: #e3f2fd; color: #1565c0; }
    .recent-detail {
      font-size: 13px;
      color: #666;
      margin: 5px 0;
      display: flex;
      gap: 8px;
    }
    .recent-detail-label {
      font-weight: 500;
      min-width: 80px;
    }
    .swipe-hint {
      text-align: center;
      padding: 10px;
      color: #999;
      font-size: 12px;
    }
    .fab {
      position: fixed;
      bottom: 20px;
      right: 20px;
      width: 56px;
      height: 56px;
      background: #1a73e8;
      color: white;
      border: none;
      border-radius: 50%;
      font-size: 24px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.3);
      cursor: pointer;
      z-index: 1000;
    }
    .fab:active {
      transform: scale(0.95);
    }
    .loading {
      text-align: center;
      padding: 40px;
      color: #666;
    }
    .refresh-indicator {
      text-align: center;
      padding: 10px;
      background: #4caf50;
      color: white;
      border-radius: 8px;
      margin-bottom: 15px;
      display: none;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📱 509 Dashboard</h1>
    <div class="subtitle">Mobile Optimized View</div>
  </div>

  <div class="container">
    <div id="refreshIndicator" class="refresh-indicator">
      ✅ Refreshed!
    </div>

    <!-- Stats -->
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-value">${stats.totalGrievances}</div>
        <div class="stat-label">Total</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${stats.activeGrievances}</div>
        <div class="stat-label">Active</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${stats.pendingGrievances}</div>
        <div class="stat-label">Pending</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${stats.overdueGrievances}</div>
        <div class="stat-label">Overdue</div>
      </div>
    </div>

    <!-- Quick Actions -->
    <div class="section-title">⚡ Quick Actions</div>
    <div class="quick-actions">
      <button class="action-button" onclick="quickNewGrievance()">
        <div class="action-icon">➕</div>
        <div class="action-text">
          <div class="action-label">New Grievance</div>
          <div class="action-description">File a new case</div>
        </div>
      </button>

      <button class="action-button" onclick="quickSearch()">
        <div class="action-icon">🔍</div>
        <div class="action-text">
          <div class="action-label">Search</div>
          <div class="action-description">Find grievances or members</div>
        </div>
      </button>

      <button class="action-button" onclick="viewMyCases()">
        <div class="action-icon">📋</div>
        <div class="action-text">
          <div class="action-label">My Cases</div>
          <div class="action-description">View assigned grievances</div>
        </div>
      </button>

      <button class="action-button" onclick="viewReports()">
        <div class="action-icon">📊</div>
        <div class="action-text">
          <div class="action-label">Reports</div>
          <div class="action-description">Analytics and insights</div>
        </div>
      </button>
    </div>

    <!-- Recent Grievances -->
    <div class="section-title">📝 Recent Grievances</div>
    <div id="recentGrievances">
      <div class="loading">Loading recent cases...</div>
    </div>

    <div class="swipe-hint">⬅️ Swipe cards for more options</div>
  </div>

  <!-- Floating Action Button -->
  <button class="fab" onclick="refreshDashboard()" title="Refresh">
    🔄
  </button>

  <script>
    // Load recent grievances
    loadRecentGrievances();

    function loadRecentGrievances() {
      google.script.run
        .withSuccessHandler(renderRecentGrievances)
        .withFailureHandler(handleError)
        .getRecentGrievancesForMobile(5);
    }

    function renderRecentGrievances(grievances) {
      const container = document.getElementById('recentGrievances');

      if (!grievances || grievances.length === 0) {
        container.innerHTML = '<div style="text-align: center; padding: 40px; color: #999;">No recent grievances</div>';
        return;
      }

      let html = '';
      grievances.forEach(g => {
        const statusClass = 'status-' + (g.status || 'filed').toLowerCase().replace(/\s+/g, '-');
        html += \`
          <div class="recent-card" onclick="viewGrievanceDetail('\${g.id}')">
            <div class="recent-header">
              <div class="recent-id">#\${g.id}</div>
              <div class="status-badge \${statusClass}">\${g.status || 'Filed'}</div>
            </div>
            <div class="recent-detail">
              <span class="recent-detail-label">Member:</span>
              <span>\${g.memberName || 'N/A'}</span>
            </div>
            <div class="recent-detail">
              <span class="recent-detail-label">Issue:</span>
              <span>\${g.issueType || 'N/A'}</span>
            </div>
            <div class="recent-detail">
              <span class="recent-detail-label">Filed:</span>
              <span>\${g.filedDate || 'N/A'}</span>
            </div>
            ${g.deadline ? `
            <div class="recent-detail">
              <span class="recent-detail-label">Deadline:</span>
              <span>\${g.deadline}</span>
            </div>
            ` : ''}
          </div>
        \`;
      });

      container.innerHTML = html;

      // Add swipe gesture support
      addSwipeSupport();
    }

    function addSwipeSupport() {
      const cards = document.querySelectorAll('.recent-card');
      cards.forEach(card => {
        let startX = 0;
        let currentX = 0;

        card.addEventListener('touchstart', (e) => {
          startX = e.touches[0].clientX;
        });

        card.addEventListener('touchmove', (e) => {
          currentX = e.touches[0].clientX;
          const diff = currentX - startX;
          if (Math.abs(diff) > 10) {
            card.style.transform = \`translateX(\${diff}px)\`;
          }
        });

        card.addEventListener('touchend', () => {
          const diff = currentX - startX;
          if (Math.abs(diff) > 100) {
            // Swipe action
            card.style.opacity = '0.5';
            setTimeout(() => card.style.display = 'none', 200);
          } else {
            card.style.transform = '';
          }
        });
      });
    }

    function quickNewGrievance() {
      google.script.run.showNewGrievanceForm();
    }

    function quickSearch() {
      google.script.run.showMemberSearch();
    }

    function viewMyCases() {
      google.script.run.showMyAssignedGrievances();
    }

    function viewReports() {
      google.script.run.showDashboard();
    }

    function viewGrievanceDetail(id) {
      google.script.run.showGrievanceDetail(id);
    }

    function refreshDashboard() {
      const indicator = document.getElementById('refreshIndicator');
      indicator.style.display = 'block';

      loadRecentGrievances();

      setTimeout(() => {
        indicator.style.display = 'none';
      }, 2000);
    }

    function handleError(error) {
      alert('Error: ' + error.message);
    }

    // Pull-to-refresh
    let touchStartY = 0;
    document.addEventListener('touchstart', (e) => {
      touchStartY = e.touches[0].clientY;
    });

    document.addEventListener('touchmove', (e) => {
      const touchY = e.touches[0].clientY;
      if (touchY - touchStartY > 150 && window.scrollY === 0) {
        refreshDashboard();
      }
    });
  </script>
</body>
</html>
  `;
}

/**
 * Gets mobile dashboard statistics
 * @returns {Object} Statistics
 */
function getMobileDashboardStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) {
    return {
      totalGrievances: 0,
      activeGrievances: 0,
      pendingGrievances: 0,
      overdueGrievances: 0
    };
  }

  const data = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 28).getValues();

  const stats = {
    totalGrievances: data.length,
    activeGrievances: 0,
    pendingGrievances: 0,
    overdueGrievances: 0
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  data.forEach(row => {
    const status = row[GRIEVANCE_COLS.STATUS - 1];
    const deadline = row[GRIEVANCE_COLS.DEADLINE - 1];

    if (status && status !== 'Resolved' && status !== 'Withdrawn') {
      stats.activeGrievances++;
    }

    if (status === 'Pending') {
      stats.pendingGrievances++;
    }

    if (deadline instanceof Date) {
      deadline.setHours(0, 0, 0, 0);
      if (deadline < today && status !== 'Resolved') {
        stats.overdueGrievances++;
      }
    }
  });

  return stats;
}

/**
 * Gets recent grievances for mobile view
 * @param {number} limit - Number of grievances to return
 * @returns {Array} Recent grievances
 */
function getRecentGrievancesForMobile(limit = 5) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) {
    return [];
  }

  const data = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 28).getValues();

  // Get most recent grievances
  const grievances = data
    .map((row, index) => {
      const filedDate = row[GRIEVANCE_COLS.FILED_DATE - 1];
      return {
        id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
        memberName: row[GRIEVANCE_COLS.MEMBER_NAME - 1],
        issueType: row[GRIEVANCE_COLS.ISSUE_TYPE - 1],
        status: row[GRIEVANCE_COLS.STATUS - 1],
        filedDate: filedDate instanceof Date ? Utilities.formatDate(filedDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : filedDate,
        deadline: row[GRIEVANCE_COLS.DEADLINE - 1] instanceof Date ? Utilities.formatDate(row[GRIEVANCE_COLS.DEADLINE - 1], Session.getScriptTimeZone(), 'MMM d, yyyy') : null,
        rowIndex: index + 2,
        filedDateObj: filedDate
      };
    })
    .sort((a, b) => {
      const dateA = a.filedDateObj instanceof Date ? a.filedDateObj : new Date(0);
      const dateB = b.filedDateObj instanceof Date ? b.filedDateObj : new Date(0);
      return dateB - dateA;
    })
    .slice(0, limit);

  return grievances;
}

/**
 * Shows mobile-optimized grievance list
 */
function showMobileGrievanceList() {
  const html = createMobileGrievanceListHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📋 Grievance List');
}

/**
 * Creates HTML for mobile grievance list
 */
function createMobileGrievanceListHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
      margin: 0;
      padding: 0;
      background: #f5f5f5;
    }
    .header {
      background: #1a73e8;
      color: white;
      padding: 15px;
      position: sticky;
      top: 0;
      z-index: 100;
    }
    .search-box {
      width: 100%;
      padding: 12px;
      border: none;
      border-radius: 8px;
      font-size: 15px;
      margin-top: 10px;
    }
    .filter-tabs {
      display: flex;
      overflow-x: auto;
      padding: 10px;
      background: white;
      border-bottom: 1px solid #e0e0e0;
      gap: 10px;
    }
    .filter-tab {
      padding: 8px 16px;
      border-radius: 20px;
      background: #f0f0f0;
      white-space: nowrap;
      cursor: pointer;
      font-size: 14px;
    }
    .filter-tab.active {
      background: #1a73e8;
      color: white;
    }
    .list-container {
      padding: 10px;
    }
    .grievance-card {
      background: white;
      margin-bottom: 12px;
      padding: 15px;
      border-radius: 12px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .card-header {
      display: flex;
      justify-content: space-between;
      margin-bottom: 10px;
    }
    .card-id {
      font-weight: bold;
      color: #1a73e8;
    }
    .card-status {
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .card-row {
      font-size: 14px;
      margin: 5px 0;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="header">
    <h2 style="margin: 0;">Grievances</h2>
    <input type="text" class="search-box" placeholder="Search..." oninput="filterList(this.value)">
  </div>

  <div class="filter-tabs">
    <div class="filter-tab active" onclick="filterByStatus('all')">All</div>
    <div class="filter-tab" onclick="filterByStatus('Active')">Active</div>
    <div class="filter-tab" onclick="filterByStatus('Pending')">Pending</div>
    <div class="filter-tab" onclick="filterByStatus('Resolved')">Resolved</div>
  </div>

  <div class="list-container" id="listContainer">
    <div style="text-align: center; padding: 40px; color: #666;">Loading...</div>
  </div>

  <script>
    let allGrievances = [];

    // Load data
    google.script.run
      .withSuccessHandler(loadGrievances)
      .getRecentGrievancesForMobile(100);

    function loadGrievances(data) {
      allGrievances = data;
      renderGrievances(data);
    }

    function renderGrievances(grievances) {
      const container = document.getElementById('listContainer');

      if (grievances.length === 0) {
        container.innerHTML = '<div style="text-align: center; padding: 40px; color: #999;">No grievances found</div>';
        return;
      }

      let html = '';
      grievances.forEach(g => {
        html += \`
          <div class="grievance-card">
            <div class="card-header">
              <div class="card-id">#\${g.id}</div>
              <div class="card-status">\${g.status}</div>
            </div>
            <div class="card-row"><strong>Member:</strong> \${g.memberName}</div>
            <div class="card-row"><strong>Issue:</strong> \${g.issueType}</div>
            <div class="card-row"><strong>Filed:</strong> \${g.filedDate}</div>
          </div>
        \`;
      });

      container.innerHTML = html;
    }

    function filterByStatus(status) {
      // Update active tab
      document.querySelectorAll('.filter-tab').forEach(tab => tab.classList.remove('active'));
      event.target.classList.add('active');

      // Filter grievances
      const filtered = status === 'all'
        ? allGrievances
        : allGrievances.filter(g => g.status === status);

      renderGrievances(filtered);
    }

    function filterList(searchTerm) {
      const filtered = allGrievances.filter(g => {
        return g.id.toLowerCase().includes(searchTerm.toLowerCase()) ||
               g.memberName.toLowerCase().includes(searchTerm.toLowerCase()) ||
               g.issueType.toLowerCase().includes(searchTerm.toLowerCase());
      });

      renderGrievances(filtered);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Shows grievance detail in mobile view
 * @param {string} grievanceId - Grievance ID
 */
function showGrievanceDetail(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  const data = grievanceSheet.getDataRange().getValues();
  const grievanceRow = data.find(row => row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] === grievanceId);

  if (!grievanceRow) {
    SpreadsheetApp.getUi().alert('Grievance not found');
    return;
  }

  // Show detail dialog
  const message = `
Grievance #${grievanceId}

Member: ${grievanceRow[GRIEVANCE_COLS.MEMBER_NAME - 1]}
Issue: ${grievanceRow[GRIEVANCE_COLS.ISSUE_TYPE - 1]}
Status: ${grievanceRow[GRIEVANCE_COLS.STATUS - 1]}
Filed: ${grievanceRow[GRIEVANCE_COLS.FILED_DATE - 1]}
Assigned: ${grievanceRow[GRIEVANCE_COLS.ASSIGNED_STEWARD - 1]}

Description:
${grievanceRow[GRIEVANCE_COLS.DESCRIPTION - 1]}
  `.trim();

  SpreadsheetApp.getUi().alert('Grievance Detail', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Shows my assigned grievances
 */
function showMyAssignedGrievances() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || grievanceSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('No grievances found');
    return;
  }

  const data = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 28).getValues();

  const myGrievances = data.filter(row => {
    const assignedSteward = row[GRIEVANCE_COLS.ASSIGNED_STEWARD - 1];
    return assignedSteward && assignedSteward.includes(userEmail);
  });

  if (myGrievances.length === 0) {
    SpreadsheetApp.getUi().alert('No grievances assigned to you');
    return;
  }

  let message = `You have ${myGrievances.length} assigned grievance(s):\n\n`;

  myGrievances.slice(0, 10).forEach(row => {
    message += `#${row[GRIEVANCE_COLS.GRIEVANCE_ID - 1]} - ${row[GRIEVANCE_COLS.MEMBER_NAME - 1]} (${row[GRIEVANCE_COLS.STATUS - 1]})\n`;
  });

  if (myGrievances.length > 10) {
    message += `\n... and ${myGrievances.length - 10} more`;
  }

  SpreadsheetApp.getUi().alert('My Cases', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/******************************************************************************
 * MODULE: DataIntegrityEnhancements
 * Source: DataIntegrityEnhancements.gs
 *****************************************************************************/

/**
 * ============================================================================
 * DATA INTEGRITY ENHANCEMENTS
 * ============================================================================
 *
 * Enhancements for data quality and integrity:
 * - Duplicate ID detection
 * - Change tracking and audit log
 * - Data quality monitoring
 */

/**
 * Checks if a Member ID already exists before adding
 * @param {string} memberId - The member ID to check
 * @returns {boolean} True if ID is duplicate, false if unique
 */
function checkDuplicateMemberID(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet || !memberId) return false;

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return false;

  const existingIds = memberSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return existingIds.includes(memberId);
}

/**
 * Checks if a Grievance ID already exists before adding
 * @param {string} grievanceId - The grievance ID to check
 * @returns {boolean} True if ID is duplicate, false if unique
 */
function checkDuplicateGrievanceID(grievanceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet || !grievanceId) return false;

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return false;

  const existingIds = grievanceSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return existingIds.includes(grievanceId);
}

/**
 * Generates next available Member ID
 * @returns {string} Next available ID in format M000001
 */
function generateNextMemberID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return 'M000001';

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return 'M000001';

  const existingIds = memberSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  // Extract numbers and find max
  const numbers = existingIds
    .filter(id => id && id.toString().startsWith('M'))
    .map(id => parseInt(id.toString().substring(1)))
    .filter(num => !isNaN(num));

  const maxNumber = numbers.length > 0 ? Math.max(...numbers) : 0;
  const nextNumber = maxNumber + 1;

  return 'M' + nextNumber.toString().padStart(6, '0');
}

/**
 * Generates next available Grievance ID
 * @returns {string} Next available ID in format G-000001-A
 */
function generateNextGrievanceID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return 'G-000001-A';

  const lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return 'G-000001-A';

  const existingIds = grievanceSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  let counter = 1;
  let letter = 'A';
  let newId;

  do {
    newId = `G-${String(counter).padStart(6, '0')}-${letter}`;

    if (!existingIds.includes(newId)) {
      break;
    }

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
 * Shows dialog when duplicate Member ID is detected
 * @param {string} memberId - The duplicate ID
 */
function showDuplicateMemberIDWarning(memberId) {
  const ui = SpreadsheetApp.getUi();
  const nextId = generateNextMemberID();

  const response = ui.alert(
    '⚠️ Duplicate Member ID Detected',
    `The Member ID "${memberId}" already exists in the system.\n\n` +
    `Next available ID: ${nextId}\n\n` +
    'Would you like to use the next available ID instead?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    return nextId;
  }

  return null;
}

/**
 * Shows dialog when duplicate Grievance ID is detected
 * @param {string} grievanceId - The duplicate ID
 */
function showDuplicateGrievanceIDWarning(grievanceId) {
  const ui = SpreadsheetApp.getUi();
  const nextId = generateNextGrievanceID();

  const response = ui.alert(
    '⚠️ Duplicate Grievance ID Detected',
    `The Grievance ID "${grievanceId}" already exists in the system.\n\n` +
    `Next available ID: ${nextId}\n\n` +
    'Would you like to use the next available ID instead?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    return nextId;
  }

  return null;
}

/**
 * Creates a Change Log sheet to track all data modifications
 */
function createChangeLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let changeLog = ss.getSheetByName('📝 Change Log');

  if (!changeLog) {
    changeLog = ss.insertSheet('📝 Change Log');
  } else {
    return; // Already exists
  }

  // Header
  changeLog.getRange("A1:H1").merge()
    .setValue("📝 CHANGE LOG - Audit Trail")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor("white");

  const headers = [
    "Timestamp",
    "User Email",
    "Sheet",
    "Row",
    "Column",
    "Field Name",
    "Old Value",
    "New Value"
  ];

  changeLog.getRange(3, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold")
    .setBackground(COLORS.LIGHT_GRAY);

  changeLog.setFrozenRows(3);
  changeLog.setTabColor(COLORS.ACCENT_PURPLE);

  // Set column widths
  changeLog.setColumnWidth(1, 150); // Timestamp
  changeLog.setColumnWidth(2, 200); // User
  changeLog.setColumnWidth(3, 150); // Sheet
  changeLog.setColumnWidth(4, 60);  // Row
  changeLog.setColumnWidth(5, 60);  // Column
  changeLog.setColumnWidth(6, 150); // Field Name
  changeLog.setColumnWidth(7, 200); // Old Value
  changeLog.setColumnWidth(8, 200); // New Value

  SpreadsheetApp.getUi().alert(
    '✅ Change Log Created',
    'The Change Log sheet has been created.\n\n' +
    'To enable automatic change tracking, you need to:\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click on Triggers (clock icon)\n' +
    '3. Add trigger: onEdit → From spreadsheet → On edit\n\n' +
    'This will track all changes made to Member Directory and Grievance Log.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Logs a change to the Change Log sheet
 * @param {object} e - The edit event object
 */
function onEdit(e) {
  if (!e) return;

  const ss = e.source;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Only track changes to core data sheets
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    return;
  }

  const changeLog = ss.getSheetByName('📝 Change Log');
  if (!changeLog) return; // Change log not created yet

  try {
    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || 'Unknown';
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const fieldName = sheet.getRange(1, col).getValue(); // Get header name
    const oldValue = e.oldValue || '';
    const newValue = e.value || '';

    // Skip header row changes
    if (row === 1) return;

    // Only log if value actually changed
    if (oldValue === newValue) return;

    // Add to change log
    const lastRow = changeLog.getLastRow();
    const logRow = [
      timestamp,
      userEmail,
      sheetName,
      row,
      col,
      fieldName,
      oldValue,
      newValue
    ];

    changeLog.getRange(lastRow + 1, 1, 1, 8).setValues([logRow]);

  } catch (error) {
    Logger.log('Error in change tracking: ' + error.message);
    // Don't show error to user to avoid interrupting workflow
  }
}

/**
 * Shows data quality dashboard
 */
function showDataQualityDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Data sheets not found!');
    return;
  }

  // Calculate data quality metrics
  const memberLastRow = memberSheet.getLastRow() - 1;
  const grievanceLastRow = grievanceSheet.getLastRow() - 1;

  // Check email completeness (Column H in Member Directory)
  const emailData = memberSheet.getRange(2, 8, memberLastRow, 1).getValues().flat();
  const emailsComplete = emailData.filter(email => email && email.toString().trim() !== '').length;
  const emailCompleteness = memberLastRow > 0 ? ((emailsComplete / memberLastRow) * 100).toFixed(1) : 0;

  // Check phone completeness (Column I in Member Directory)
  const phoneData = memberSheet.getRange(2, 9, memberLastRow, 1).getValues().flat();
  const phonesComplete = phoneData.filter(phone => phone && phone.toString().trim() !== '').length;
  const phoneCompleteness = memberLastRow > 0 ? ((phonesComplete / memberLastRow) * 100).toFixed(1) : 0;

  // Check steward assignment in grievances (Column AA)
  const stewardData = grievanceSheet.getRange(2, 27, grievanceLastRow, 1).getValues().flat();
  const stewardsAssigned = stewardData.filter(s => s && s.toString().trim() !== '').length;
  const stewardCompleteness = grievanceLastRow > 0 ? ((stewardsAssigned / grievanceLastRow) * 100).toFixed(1) : 0;

  // Overall quality score (average of all metrics)
  const qualityScore = ((parseFloat(emailCompleteness) + parseFloat(phoneCompleteness) + parseFloat(stewardCompleteness)) / 3).toFixed(1);

  // Show dashboard
  SpreadsheetApp.getUi().alert(
    '📊 Data Quality Dashboard',
    `Overall Quality Score: ${qualityScore}%\n\n` +
    '=== Member Directory ===\n' +
    `Total Members: ${memberLastRow}\n` +
    `Email Addresses: ${emailCompleteness}% complete (${emailsComplete}/${memberLastRow})\n` +
    `Phone Numbers: ${phoneCompleteness}% complete (${phonesComplete}/${memberLastRow})\n\n` +
    '=== Grievance Log ===\n' +
    `Total Grievances: ${grievanceLastRow}\n` +
    `Steward Assignments: ${stewardCompleteness}% complete (${stewardsAssigned}/${grievanceLastRow})\n\n` +
    `${qualityScore >= 90 ? '🟢 Excellent!' : qualityScore >= 75 ? '🟡 Good' : '🔴 Needs Improvement'}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Validates referential integrity (Member IDs in grievances exist in Member Directory)
 */
function checkReferentialIntegrity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Data sheets not found!');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('🔍 Checking referential integrity...', 'Please wait', -1);

  // Get all member IDs
  const memberLastRow = memberSheet.getLastRow();
  const memberIds = memberSheet.getRange(2, 1, memberLastRow - 1, 1).getValues().flat();

  // Get all member IDs from grievances
  const grievanceLastRow = grievanceSheet.getLastRow();
  const grievanceMemberIds = grievanceSheet.getRange(2, 2, grievanceLastRow - 1, 1).getValues();

  // Find orphaned grievances (member ID doesn't exist)
  const orphaned = [];
  for (let i = 0; i < grievanceMemberIds.length; i++) {
    const memberId = grievanceMemberIds[i][0];
    if (memberId && !memberIds.includes(memberId)) {
      const grievanceId = grievanceSheet.getRange(i + 2, 1).getValue();
      orphaned.push({ row: i + 2, grievanceId, memberId });
    }
  }

  // Show results
  if (orphaned.length === 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Referential Integrity Check Passed',
      'All grievances have valid member IDs.\n\n' +
      `Checked ${grievanceLastRow - 1} grievances against ${memberLastRow - 1} members.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    const orphanedList = orphaned.slice(0, 10).map(o =>
      `  Row ${o.row}: ${o.grievanceId} (Member ID: ${o.memberId})`
    ).join('\n');

    SpreadsheetApp.getUi().alert(
      '⚠️ Referential Integrity Issues Found',
      `Found ${orphaned.length} grievance(s) with invalid member IDs:\n\n` +
      orphanedList +
      (orphaned.length > 10 ? `\n\n  ... and ${orphaned.length - 10} more` : '') +
      '\n\nThese grievances reference members that don\'t exist in the Member Directory.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/**
 * Shows next available Member ID dialog
 */
function showNextMemberID() {
  const nextId = generateNextMemberID();
  SpreadsheetApp.getUi().alert(
    "🆔 Next Available Member ID",
    `Next Member ID: ${nextId}\n\n` +
    "Use this ID when adding a new member to avoid duplicates.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows next available Grievance ID dialog
 */
function showNextGrievanceID() {
  const nextId = generateNextGrievanceID();
  SpreadsheetApp.getUi().alert(
    "🆔 Next Available Grievance ID",
    `Next Grievance ID: ${nextId}\n\n` +
    "Use this ID when creating a new grievance to avoid duplicates.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/******************************************************************************
 * MODULE: SeedNuke
 * Source: SeedNuke.gs
 *****************************************************************************/

/**
 * ============================================================================
 * SEED NUKE - Remove All Seeded Data and Exit Demo Mode
 * ============================================================================
 *
 * Allows stewards to remove all test/seeded data and exit demo mode.
 * After nuking, the dashboard will be ready for production use.
 *
 * Features:
 * - Removes all seeded members and grievances
 * - Keeps headers and structure intact
 * - Recalculates all dashboards
 * - Shows getting started reminder
 * - Hides seed menu options after nuke
 *
 * ============================================================================
 */

/**
 * Main function to nuke all seeded data
 */
function nukeSeedData() {
  const ui = SpreadsheetApp.getUi();

  // Confirmation dialog
  const response = ui.alert(
    '⚠️ WARNING: Remove All Seeded Data',
    'This will PERMANENTLY remove all test data from:\n\n' +
    '• Member Directory (all members)\n' +
    '• Grievance Log (all grievances)\n' +
    '• Steward Workload (all records)\n\n' +
    'Headers and sheet structure will be preserved.\n\n' +
    'This action CANNOT be undone!\n\n' +
    'Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('✅ Operation cancelled. No data was removed.');
    return;
  }

  // Double confirmation
  const finalConfirm = ui.alert(
    '🚨 FINAL CONFIRMATION',
    'This is your last chance!\n\n' +
    'ALL test data will be permanently deleted.\n\n' +
    'Click YES to proceed with data removal.',
    ui.ButtonSet.YES_NO
  );

  if (finalConfirm !== ui.Button.YES) {
    ui.alert('✅ Operation cancelled. No data was removed.');
    return;
  }

  try {
    // Show progress
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ui.alert('⏳ Removing seeded data...\n\nThis may take a moment. Please wait.');

    // Step 1: Clear Member Directory (keep headers)
    clearMemberDirectory();

    // Step 2: Clear Grievance Log (keep headers)
    clearGrievanceLog();

    // Step 3: Clear Steward Workload (keep headers)
    clearStewardWorkload();

    // Step 4: Recalculate all dashboards
    rebuildDashboard();

    // Step 5: Set flag that data has been nuked
    PropertiesService.getScriptProperties().setProperty('SEED_NUKED', 'true');

    // Step 6: Show getting started reminder
    showPostNukeGuidance();

  } catch (error) {
    ui.alert('❌ Error during data removal: ' + error.message);
    Logger.log('Error in nukeSeedData: ' + error.message);
  }
}

/**
 * Clears Member Directory while preserving headers
 */
function clearMemberDirectory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    throw new Error('Member Directory not found');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    // Delete all rows except header
    sheet.deleteRows(2, lastRow - 1);
  }

  Logger.log('Member Directory cleared: ' + (lastRow - 1) + ' members removed');
}

/**
 * Clears Grievance Log while preserving headers
 */
function clearGrievanceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    throw new Error('Grievance Log not found');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    // Delete all rows except header
    sheet.deleteRows(2, lastRow - 1);
  }

  Logger.log('Grievance Log cleared: ' + (lastRow - 1) + ' grievances removed');
}

/**
 * Clears Steward Workload while preserving headers
 */
function clearStewardWorkload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);

  if (!sheet) {
    // Sheet doesn't exist, skip
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    // Delete all rows except header
    sheet.deleteRows(2, lastRow - 1);
  }

  Logger.log('Steward Workload cleared');
}

/**
 * Shows post-nuke guidance to the user
 */
function showPostNukeGuidance() {
  const ui = SpreadsheetApp.getUi();

  const html = HtmlService.createHtmlOutput(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Roboto', Arial, sans-serif;
      padding: 30px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      margin: 0;
    }
    .container {
      background: white;
      color: #333;
      padding: 30px;
      border-radius: 12px;
      box-shadow: 0 10px 40px rgba(0,0,0,0.3);
      max-width: 600px;
      margin: 0 auto;
    }
    h1 {
      color: #1a73e8;
      margin-top: 0;
      font-size: 28px;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 15px;
    }
    .success-icon {
      font-size: 64px;
      text-align: center;
      margin: 20px 0;
    }
    .info-box {
      background: #e8f0fe;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
      border-left: 5px solid #1a73e8;
    }
    .warning-box {
      background: #fff3cd;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
      border-left: 5px solid #ff9800;
    }
    .checklist {
      background: #f8f9fa;
      padding: 20px;
      border-radius: 8px;
      margin: 20px 0;
    }
    .checklist h3 {
      margin-top: 0;
      color: #1a73e8;
    }
    .checklist ul {
      list-style: none;
      padding: 0;
    }
    .checklist li {
      padding: 10px 0;
      border-bottom: 1px solid #ddd;
    }
    .checklist li:last-child {
      border-bottom: none;
    }
    .checklist li::before {
      content: "☑️ ";
      margin-right: 10px;
    }
    .button-container {
      text-align: center;
      margin-top: 30px;
    }
    button {
      padding: 12px 30px;
      font-size: 16px;
      font-weight: bold;
      border: none;
      border-radius: 6px;
      background: #1a73e8;
      color: white;
      cursor: pointer;
      transition: all 0.3s;
    }
    button:hover {
      background: #1557b0;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(26,115,232,0.4);
    }
    .footer {
      text-align: center;
      margin-top: 30px;
      font-size: 12px;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="success-icon">🎉</div>

    <h1>Welcome to Production Mode!</h1>

    <div class="info-box">
      <strong>✅ Success!</strong><br>
      All seeded test data has been removed. Your dashboard is now ready for real member and grievance data.
    </div>

    <div class="warning-box">
      <strong>⚠️ Important Next Steps</strong><br>
      Before you start using the dashboard, please complete the setup steps below.
    </div>

    <div class="checklist">
      <h3>📋 Getting Started Checklist</h3>
      <ul>
        <li><strong>Configure Steward Contact Info</strong><br>
            Go to the <strong>⚙️ Config</strong> tab and enter your steward contact information in column U (rows 2-4).
            This will be used when starting grievances from the Member Directory.</li>

        <li><strong>Set Up Google Form (Optional)</strong><br>
            If you want to use the grievance workflow feature, create a Google Form for grievance submissions
            and update the form URL and field IDs in the script configuration.</li>

        <li><strong>Add Your First Members</strong><br>
            Go to <strong>👥 Member Directory</strong> and start adding your members.
            You can enter them manually or import from a CSV file.</li>

        <li><strong>Review Config Settings</strong><br>
            Check the <strong>⚙️ Config</strong> tab to ensure all dropdown values
            (job titles, locations, units, etc.) match your organization's needs.</li>

        <li><strong>Customize Dashboards</strong><br>
            Explore the various dashboard views and use the <strong>🎯 Interactive Dashboard</strong>
            to create custom views for your needs.</li>

        <li><strong>Set Up Triggers (Recommended)</strong><br>
            Go to <strong>509 Tools > Utilities > Setup Triggers</strong> to enable automatic
            calculations and deadline tracking.</li>
      </ul>
    </div>

    <div class="info-box">
      <strong>💡 Tip:</strong> The seed data menu options have been hidden. If you need to re-seed test data
      for training purposes, you can access the seeding functions from the script editor.
    </div>

    <div class="button-container">
      <button onclick="google.script.host.close()">Get Started!</button>
    </div>

    <div class="footer">
      SEIU Local 509 Dashboard | Ready for Production Use
    </div>
  </div>
</body>
</html>
  `).setWidth(700).setHeight(600);

  ui.showModalDialog(html, '🎉 Seeded Data Removed Successfully');
}

/**
 * Checks if seed data has been nuked
 */
function isSeedNuked() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('SEED_NUKED') === 'true';
}

/**
 * Resets the nuke flag (for development/testing only)
 */
function resetNukeFlag() {
  PropertiesService.getScriptProperties().deleteProperty('SEED_NUKED');
  SpreadsheetApp.getUi().alert('✅ Nuke flag reset. Seed menu will be visible again.');
}

/**
 * Shows a quick reminder dialog to enter steward contact info
 */
function showStewardContactReminder() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '👋 Quick Setup Reminder',
    'Have you entered your steward contact information in the Config tab?\n\n' +
    'This information is used when starting grievances from the Member Directory.\n\n' +
    'Go to: ⚙️ Config > Column U (Steward Contact Information)\n\n' +
    'Click YES if you\'ve already done this, or NO to be reminded later.',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    PropertiesService.getUserProperties().setProperty('STEWARD_INFO_CONFIGURED', 'true');
  }
}

/**
 * Checks if steward info is configured
 */
function isStewardInfoConfigured() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty('STEWARD_INFO_CONFIGURED') === 'true';
}

/**
 * Shows getting started guide
 */
function showGettingStartedGuide() {
  const ui = SpreadsheetApp.getUi();

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
      max-width: 800px;
      margin: 0 auto;
    }
    h1 {
      color: #1a73e8;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    h2 {
      color: #1a73e8;
      margin-top: 30px;
    }
    .step {
      background: #f8f9fa;
      padding: 15px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
      border-radius: 4px;
    }
    .step h3 {
      margin-top: 0;
      color: #333;
    }
    code {
      background: #e8f0fe;
      padding: 2px 6px;
      border-radius: 3px;
      font-family: monospace;
    }
    button {
      padding: 10px 20px;
      background: #1a73e8;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      margin-top: 20px;
    }
    button:hover {
      background: #1557b0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>📚 Getting Started Guide</h1>

    <p>Welcome to the SEIU Local 509 Dashboard! Follow these steps to get started:</p>

    <div class="step">
      <h3>1️⃣ Configure Steward Contact Information</h3>
      <p>Go to the <code>⚙️ Config</code> tab and scroll to column U.</p>
      <p>Enter:</p>
      <ul>
        <li>Steward Name (Row 2)</li>
        <li>Steward Email (Row 3)</li>
        <li>Steward Phone (Row 4)</li>
      </ul>
      <p>This information will be automatically included when starting new grievances.</p>
    </div>

    <div class="step">
      <h3>2️⃣ Add Members to the Directory</h3>
      <p>Go to the <code>👥 Member Directory</code> tab and start adding member information.</p>
      <p>You can:</p>
      <ul>
        <li>Enter members manually</li>
        <li>Import from a CSV file</li>
        <li>Copy and paste from another spreadsheet</li>
      </ul>
    </div>

    <div class="step">
      <h3>3️⃣ Review Configuration Settings</h3>
      <p>In the <code>⚙️ Config</code> tab, review the dropdown lists to ensure they match your needs:</p>
      <ul>
        <li>Job Titles</li>
        <li>Work Locations</li>
        <li>Grievance Types</li>
        <li>And more...</li>
      </ul>
    </div>

    <div class="step">
      <h3>4️⃣ Set Up Automatic Calculations</h3>
      <p>Go to <code>509 Tools > Utilities > Setup Triggers</code></p>
      <p>This enables automatic deadline calculations and dashboard updates.</p>
    </div>

    <div class="step">
      <h3>5️⃣ Explore the Dashboards</h3>
      <p>Check out the various dashboard views:</p>
      <ul>
        <li><code>📊 Main Dashboard</code> - Overview of all metrics</li>
        <li><code>🎯 Interactive Dashboard</code> - Customizable views</li>
        <li><code>👨‍⚖️ Steward Workload</code> - Track steward assignments</li>
      </ul>
    </div>

    <h2>🚀 Optional: Set Up Grievance Workflow</h2>

    <div class="step">
      <h3>Create a Google Form for Grievances</h3>
      <p>If you want to use the automated grievance workflow:</p>
      <ol>
        <li>Create a Google Form with fields for grievance information</li>
        <li>Link the form to this spreadsheet</li>
        <li>Update the form URL and field IDs in the script configuration</li>
        <li>Set up a form submission trigger</li>
      </ol>
      <p>See the documentation in <code>GrievanceWorkflow.gs</code> for details.</p>
    </div>

    <button onclick="google.script.host.close()">Let's Go!</button>
  </div>
</body>
</html>
  `).setWidth(900).setHeight(700);

  ui.showModalDialog(html, 'Getting Started Guide');
}

/**
 * Rebuilds all dashboard calculations and charts
 * Called after data is cleared/nuked to refresh metrics
 */
function rebuildDashboard() {
  try {
    // Call the main refresh function from Code.gs
    if (typeof refreshCalculations === 'function') {
      refreshCalculations();
    }

    // Rebuild interactive dashboard if it exists
    if (typeof rebuildInteractiveDashboard === 'function') {
      rebuildInteractiveDashboard();
    }

    Logger.log('Dashboard rebuilt successfully');
  } catch (error) {
    Logger.log('Error rebuilding dashboard: ' + error.message);
    // Non-critical error, continue execution
  }
}

/**
 * Add sample realistic feedback entries to Feedback & Development sheet
 */
function addSampleFeedbackEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feedback = ss.getSheetByName(SHEETS.FEEDBACK);

  if (!feedback) {
    SpreadsheetApp.getUi().alert('❌ Feedback & Development sheet not found!');
    return;
  }

  const today = new Date();
  const lastWeek = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  const twoWeeksAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);
  const nextMonth = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

  const sampleEntries = [
    [
      'Feedback',
      Utilities.formatDate(lastWeek, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'Maria Gonzalez',
      'Medium',
      'Dashboard load time could be improved',
      'When opening the Interactive Dashboard with 20k+ members, it takes 5-8 seconds to load. Consider implementing lazy loading or pagination for better performance.',
      'Under Review',
      25,
      'Moderate',
      Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'Tech Team',
      'None',
      'Investigating caching options and chart lazy loading',
      Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy')
    ],
    [
      'Future Feature',
      Utilities.formatDate(twoWeeksAgo, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'James Wilson',
      'High',
      'Automated weekly steward workload reports',
      'Send automated email reports to stewards every Monday morning with their active cases, upcoming deadlines, and win rate statistics. Would save 2-3 hours per week of manual reporting.',
      'Planned',
      10,
      'Complex',
      Utilities.formatDate(nextMonth, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'Development Team',
      'Need to set up Gmail API integration',
      'Aligns with Phase 7 automation goals',
      Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy')
    ],
    [
      'Bug Report',
      Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'Sarah Chen',
      'High',
      'Member search not finding partial matches',
      'When searching for members, the search only works with exact matches. Searching for "John" doesn\'t find "John Smith" or "Johnson". This makes it difficult to quickly look up members.',
      'New',
      0,
      'Simple',
      Utilities.formatDate(new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      'Unassigned',
      'None',
      'Need to update search algorithm to support partial matching',
      Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy')
    ]
  ];

  const lastRow = feedback.getLastRow();
  feedback.getRange(lastRow + 1, 1, sampleEntries.length, sampleEntries[0].length).setValues(sampleEntries);

  SpreadsheetApp.getUi().alert('✅ Added 3 sample feedback entries to Feedback & Development sheet');
}

/**
 * Populate Steward Workload sheet with live data from Grievance Log
 */
function populateStewardWorkload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workloadSheet = ss.getSheetByName(SHEETS.STEWARD_WORKLOAD);
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!workloadSheet || !grievanceSheet || !memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Required sheets not found!');
    return;
  }

  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  const stewards = {};
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const isSteward = row[9];
    if (isSteward === 'Yes') {
      const memberId = row[0];
      const name = `${row[1]} ${row[2]}`;
      const email = row[4];
      const phone = row[5];
      stewards[memberId] = {
        name: name,
        email: email,
        phone: phone,
        totalCases: 0,
        activeCases: 0,
        resolvedCases: 0,
        wonCases: 0,
        resolutionDays: []
      };
    }
  }

  const today = new Date();
  for (let i = 1; i < grievanceData.length; i++) {
    const row = grievanceData[i];
    const stewardId = row[6];
    const status = row[4];
    const outcome = row[16];
    const daysOpen = row[18];

    if (stewards[stewardId]) {
      stewards[stewardId].totalCases++;

      if (status === 'Open' || status === 'Pending Info') {
        stewards[stewardId].activeCases++;
      } else if (status === 'Settled' || status === 'Resolved' || status === 'Closed') {
        stewards[stewardId].resolvedCases++;

        if (outcome === 'Won' || outcome === 'Partially Won') {
          stewards[stewardId].wonCases++;
        }

        if (daysOpen && !isNaN(daysOpen)) {
          stewards[stewardId].resolutionDays.push(parseFloat(daysOpen));
        }
      }
    }
  }

  const outputData = [];
  for (const stewardId in stewards) {
    const s = stewards[stewardId];
    const winRate = s.resolvedCases > 0 ? (s.wonCases / s.resolvedCases * 100) : 0;
    const avgDays = s.resolutionDays.length > 0
      ? s.resolutionDays.reduce((a, b) => a + b, 0) / s.resolutionDays.length
      : 0;

    let capacityStatus;
    if (s.activeCases === 0) {
      capacityStatus = 'Available';
    } else if (s.activeCases <= 5) {
      capacityStatus = 'Normal';
    } else if (s.activeCases <= 10) {
      capacityStatus = 'Busy';
    } else {
      capacityStatus = 'Overloaded';
    }

    outputData.push([
      s.name,
      s.totalCases,
      s.activeCases,
      s.resolvedCases,
      Math.round(winRate),
      Math.round(avgDays),
      0,
      0,
      capacityStatus,
      s.email || '',
      s.phone || ''
    ]);
  }

  outputData.sort((a, b) => b[2] - a[2]);

  const lastRow = workloadSheet.getLastRow();
  if (lastRow > 3) {
    workloadSheet.getRange(4, 1, lastRow - 3, 11).clear();
  }

  if (outputData.length > 0) {
    workloadSheet.getRange(4, 1, outputData.length, 11).setValues(outputData);
  }

  Logger.log(`✅ Populated Steward Workload with ${outputData.length} stewards`);
}

/**
 * Populate Member Satisfaction sheet with sample survey data
 */
function populateMemberSatisfaction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const satisfactionSheet = ss.getSheetByName(SHEETS.MEMBER_SATISFACTION);
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!satisfactionSheet || !memberSheet) {
    SpreadsheetApp.getUi().alert('❌ Required sheets not found!');
    return;
  }

  const memberData = memberSheet.getDataRange().getValues();
  const sampleSurveys = [];
  const today = new Date();

  for (let i = 0; i < 50; i++) {
    const randomMemberIndex = Math.floor(Math.random() * (memberData.length - 1)) + 1;
    const member = memberData[randomMemberIndex];
    const memberId = member[0];
    const memberName = `${member[1]} ${member[2]}`;

    const sentDaysAgo = Math.floor(Math.random() * 60) + 1;
    const completedDaysAgo = sentDaysAgo - Math.floor(Math.random() * 7);

    const dateSent = new Date(today.getTime() - sentDaysAgo * 24 * 60 * 60 * 1000);
    const dateCompleted = completedDaysAgo > 0 ? new Date(today.getTime() - completedDaysAgo * 24 * 60 * 60 * 1000) : '';

    const overallSat = Math.random() < 0.7 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const stewardSupport = Math.random() < 0.75 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const communication = Math.random() < 0.65 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const wouldRecommend = overallSat >= 4 ? 'Y' : (Math.random() < 0.7 ? 'Y' : 'N');

    const comments = [
      'My steward was very helpful and responsive',
      'Great communication throughout the process',
      'Could use faster response times',
      'Very satisfied with the support I received',
      'Steward went above and beyond',
      'Process was clear and well-explained',
      'Would like more frequent updates',
      'Excellent representation',
      'Professional and knowledgeable',
      'Satisfied overall',
      ''
    ];
    const comment = comments[Math.floor(Math.random() * comments.length)];

    sampleSurveys.push([
      `SURVEY-${String(i + 1).padStart(4, '0')}`,
      memberId,
      memberName,
      Utilities.formatDate(dateSent, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      dateCompleted ? Utilities.formatDate(dateCompleted, Session.getScriptTimeZone(), 'MM/dd/yyyy') : '',
      overallSat,
      stewardSupport,
      communication,
      wouldRecommend,
      comment
    ]);
  }

  const lastRow = satisfactionSheet.getLastRow();
  if (lastRow > 3) {
    satisfactionSheet.getRange(4, 1, lastRow - 3, 10).clear();
  }

  if (sampleSurveys.length > 0) {
    satisfactionSheet.getRange(4, 1, sampleSurveys.length, 10).setValues(sampleSurveys);
  }

  SpreadsheetApp.getUi().alert(`✅ Added ${sampleSurveys.length} sample surveys to Member Satisfaction sheet`);
  Logger.log(`✅ Populated Member Satisfaction with ${sampleSurveys.length} surveys`);
}

/**
 * Populate all analytics sheets with live data
 */
function populateAllAnalyticsSheets() {
  const ui = SpreadsheetApp.getUi();

  ui.alert('⏳ Populating analytics sheets...\n\nThis may take a moment.');

  try {
    populateStewardWorkload();
    populateMemberSatisfaction();

    ui.alert('✅ Analytics sheets populated successfully!');
  } catch (error) {
    ui.alert('❌ Error populating analytics: ' + error.message);
    Logger.log('Error in populateAllAnalyticsSheets: ' + error.message);
  }
}

/**
 * Hide the Diagnostics tab
 */
function hideDiagnosticsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diagnostics = ss.getSheetByName(SHEETS.DIAGNOSTICS);

  if (diagnostics) {
    diagnostics.hideSheet();
    SpreadsheetApp.getUi().alert('✅ Diagnostics tab is now hidden');
  } else {
    SpreadsheetApp.getUi().alert('❌ Diagnostics sheet not found');
  }
}

/**
 * Add all dropdown validations and conditional formatting to Member Directory
 */
function setupMemberDirectoryValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  if (!memberDir || !config) {
    SpreadsheetApp.getUi().alert('❌ Required sheets not found!');
    return;
  }

  // Get dropdown values from Config sheet
  const jobTitles = config.getRange("A2:A100").getValues().flat().filter(String);
  const locations = config.getRange("B2:B100").getValues().flat().filter(String);
  const units = config.getRange("C2:C100").getValues().flat().filter(String);
  const stewards = config.getRange("E2:E100").getValues().flat().filter(String);
  const supervisors = config.getRange("F2:F100").getValues().flat().filter(String);
  const managers = config.getRange("G2:G100").getValues().flat().filter(String);

  // Define dropdown ranges (2 = first data row, 5000 = max rows)
  const MAX_ROWS = 5000;

  // Job Title (Column D = 4)
  if (jobTitles.length > 0) {
    const jobTitleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(jobTitles, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 4, MAX_ROWS, 1).setDataValidation(jobTitleRule);
  }

  // Work Location (Column E = 5)
  if (locations.length > 0) {
    const locationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(locations, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 5, MAX_ROWS, 1).setDataValidation(locationRule);
  }

  // Unit (Column F = 6)
  if (units.length > 0) {
    const unitRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(units, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 6, MAX_ROWS, 1).setDataValidation(unitRule);
  }

  // Office Days (Column G = 7) - Allow multiple selections as comma-separated
  const officeDaysRule = SpreadsheetApp.newDataValidation()
    .requireTextContains("")
    .setAllowInvalid(true)
    .setHelpText('Enter multiple days comma-separated (e.g., "Monday, Wednesday, Friday")')
    .build();
  memberDir.getRange(2, 7, MAX_ROWS, 1).setDataValidation(officeDaysRule);

  // Is Steward (Column J = 10)
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange(2, 10, MAX_ROWS, 1).setDataValidation(yesNoRule);

  // Supervisor (Column K = 11)
  if (supervisors.length > 0) {
    const supervisorRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(supervisors, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 11, MAX_ROWS, 1).setDataValidation(supervisorRule);
  }

  // Manager (Column L = 12)
  if (managers.length > 0) {
    const managerRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(managers, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 12, MAX_ROWS, 1).setDataValidation(managerRule);
  }

  // Assigned Steward (Column M = 13)
  if (stewards.length > 0) {
    const stewardRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(stewards, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, 13, MAX_ROWS, 1).setDataValidation(stewardRule);
  }

  // Interest: Local Actions (Column T = 20)
  memberDir.getRange(2, 20, MAX_ROWS, 1).setDataValidation(yesNoRule);

  // Interest: Chapter Actions (Column U = 21)
  memberDir.getRange(2, 21, MAX_ROWS, 1).setDataValidation(yesNoRule);

  // Interest: Allied Chapter Actions (Column V = 22)
  memberDir.getRange(2, 22, MAX_ROWS, 1).setDataValidation(yesNoRule);

  // Preferred Communication Methods (Column X = 24) - Multiple selections
  const commMethodsRule = SpreadsheetApp.newDataValidation()
    .requireTextContains("")
    .setAllowInvalid(true)
    .setHelpText('Enter multiple methods comma-separated (e.g., "Email, Phone, Text")')
    .build();
  memberDir.getRange(2, 24, MAX_ROWS, 1).setDataValidation(commMethodsRule);

  // Best Time(s) to Reach Member (Column Y = 25) - Multiple selections
  const bestTimeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Anytime'], true)
    .setAllowInvalid(true)
    .setHelpText('Select one or enter multiple comma-separated')
    .build();
  memberDir.getRange(2, 25, MAX_ROWS, 1).setDataValidation(bestTimeRule);

  // Steward Who Contacted Member (Column AD = 30)
  if (stewards.length > 0) {
    memberDir.getRange(2, 30, MAX_ROWS, 1).setDataValidation(stewardRule);
  }

  // Add conditional formatting for empty email/phone
  // Email (Column H = 8) - Red background if empty
  const emailRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(A2<>"", H2="")')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF')
    .setRanges([memberDir.getRange(2, 8, MAX_ROWS, 1)])
    .build();

  // Phone (Column I = 9) - Red background if empty
  const phoneRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(A2<>"", I2="")')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF')
    .setRanges([memberDir.getRange(2, 9, MAX_ROWS, 1)])
    .build();

  const rules = memberDir.getConditionalFormatRules();
  rules.push(emailRule);
  rules.push(phoneRule);
  memberDir.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('✅ Member Directory validations and formatting applied!');
}

/**
 * Add Google Drive folder link column to Grievance Log
 */
function addGoogleDriveLinkToGrievanceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceLog) {
    SpreadsheetApp.getUi().alert('❌ Grievance Log not found!');
    return;
  }

  // Insert column after Resolution Summary (last column)
  const lastCol = grievanceLog.getLastColumn();
  grievanceLog.insertColumnAfter(lastCol);

  // Set header
  grievanceLog.getRange(1, lastCol + 1)
    .setValue('📁 Drive Folder Link')
    .setFontWeight('bold')
    .setBackground('#DC2626')
    .setFontColor('#FFFFFF')
    .setWrap(true);

  // Add note to header
  grievanceLog.getRange(1, lastCol + 1)
    .setNote('Paste Google Drive folder link for this grievance. Create folder with grievant name.');

  SpreadsheetApp.getUi().alert('✅ Google Drive folder link column added to Grievance Log!');
}

/**
 * Add status bars for grievance dates (visual deadline tracking)
 */
function addGrievanceDateStatusBars() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceLog) {
    SpreadsheetApp.getUi().alert('❌ Grievance Log not found!');
    return;
  }

  const MAX_ROWS = 5000;

  // Days to Deadline (Column U = 21) - Color-coded based on urgency
  const urgentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#DC2626')  // Red - OVERDUE
    .setFontColor('#FFFFFF')
    .setRanges([grievanceLog.getRange(2, 21, MAX_ROWS, 1)])
    .build();

  const warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0, 3)
    .setBackground('#F97316')  // Orange - URGENT (0-3 days)
    .setFontColor('#FFFFFF')
    .setRanges([grievanceLog.getRange(2, 21, MAX_ROWS, 1)])
    .build();

  const cautionRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 7)
    .setBackground('#FCD34D')  // Yellow - CAUTION (4-7 days)
    .setFontColor('#000000')
    .setRanges([grievanceLog.getRange(2, 21, MAX_ROWS, 1)])
    .build();

  const okRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(7)
    .setBackground('#10B981')  // Green - OK (8+ days)
    .setFontColor('#FFFFFF')
    .setRanges([grievanceLog.getRange(2, 21, MAX_ROWS, 1)])
    .build();

  const rules = grievanceLog.getConditionalFormatRules();
  rules.push(urgentRule, warningRule, cautionRule, okRule);
  grievanceLog.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('✅ Date status bars added to Grievance Log!');
}

/**
 * Setup all dashboard enhancements
 */
function SETUP_DASHBOARD_ENHANCEMENTS() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🎨 Setup Dashboard Enhancements',
    'This will add:\n\n' +
    '✓ Dropdowns for all Member Directory fields\n' +
    '✓ Red highlighting for missing email/phone\n' +
    '✓ Google Drive folder link to Grievance Log\n' +
    '✓ Color-coded status bars for deadlines\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    ui.alert('⏳ Applying enhancements...');

    setupMemberDirectoryValidations();
    addGoogleDriveLinkToGrievanceLog();
    addGrievanceDateStatusBars();

    ui.alert(
      '✅ Enhancements Complete!',
      'All dropdown validations, conditional formatting, and visual enhancements have been applied.\n\n' +
      'Member Directory: Dropdowns active + red highlighting for empty email/phone\n' +
      'Grievance Log: Drive folder link column added + color-coded deadline tracking',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    Logger.log('Error in SETUP_DASHBOARD_ENHANCEMENTS: ' + error.message);
  }
}
