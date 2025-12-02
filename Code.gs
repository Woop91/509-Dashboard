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
    createLocationSheet();
    createTypeAnalysisSheet();
    SpreadsheetApp.getActive().toast("✅ Analytics sheets created", "75%", 2);

    createExecutiveDashboard();
    createKPIPerformanceDashboard();
    createMemberEngagementSheet();
    createCostImpactSheet();
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
function SEED_5K_GRIEVANCES() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Seed 5,000 Grievances',
    'Use the 2 toggles instead for better performance:\n\n' +
    '• Seed Grievances - Toggle 1 (2,500)\n' +
    '• Seed Grievances - Toggle 2 (2,500)\n\n' +
    'Would you like to seed all 5,000 at once anyway?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Call both toggles
  seedGrievancesWithCount(2500, "Toggle 1");
  seedGrievancesWithCount(2500, "Toggle 2");
}

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

  // Get all grievance data
  const grievanceData = grievanceSheet.getDataRange().getValues();
  const memberData = memberSheet.getDataRange().getValues();

  // Build steward lookup map (Member ID -> Steward info)
  const stewards = {};
  for (let i = 1; i < memberData.length; i++) {
    const row = memberData[i];
    const isSteward = row[9]; // Column J - Is Steward
    if (isSteward === 'Yes') {
      const memberId = row[0];
      const name = `${row[1]} ${row[2]}`; // First + Last name
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

  // Process grievances
  const today = new Date();
  for (let i = 1; i < grievanceData.length; i++) {
    const row = grievanceData[i];
    const stewardId = row[6]; // Column G - Assigned Steward ID
    const status = row[4]; // Column E - Status
    const outcome = row[16]; // Column Q - Outcome
    const daysOpen = row[18]; // Column S - Days Open

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

  // Build output data
  const outputData = [];
  for (const stewardId in stewards) {
    const s = stewards[stewardId];
    const winRate = s.resolvedCases > 0 ? (s.wonCases / s.resolvedCases * 100) : 0;
    const avgDays = s.resolutionDays.length > 0
      ? s.resolutionDays.reduce((a, b) => a + b, 0) / s.resolutionDays.length
      : 0;

    // Capacity status based on active cases
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
      0, // Overdue cases - would need deadline calculation
      0, // Due this week - would need deadline calculation
      capacityStatus,
      s.email || '',
      s.phone || ''
    ]);
  }

  // Sort by active cases (descending)
  outputData.sort((a, b) => b[2] - a[2]);

  // Clear existing data (keep headers)
  const lastRow = workloadSheet.getLastRow();
  if (lastRow > 3) {
    workloadSheet.getRange(4, 1, lastRow - 3, 11).clear();
  }

  // Write new data
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

  // Get member data
  const memberData = memberSheet.getDataRange().getValues();

  // Generate 50 sample survey responses
  const sampleSurveys = [];
  const today = new Date();

  for (let i = 0; i < 50; i++) {
    // Pick a random member (skip header row)
    const randomMemberIndex = Math.floor(Math.random() * (memberData.length - 1)) + 1;
    const member = memberData[randomMemberIndex];
    const memberId = member[0];
    const memberName = `${member[1]} ${member[2]}`;

    // Generate survey dates
    const sentDaysAgo = Math.floor(Math.random() * 60) + 1; // 1-60 days ago
    const completedDaysAgo = sentDaysAgo - Math.floor(Math.random() * 7); // Within a week of being sent

    const dateSent = new Date(today.getTime() - sentDaysAgo * 24 * 60 * 60 * 1000);
    const dateCompleted = completedDaysAgo > 0 ? new Date(today.getTime() - completedDaysAgo * 24 * 60 * 60 * 1000) : '';

    // Generate ratings (weighted toward positive)
    const overallSat = Math.random() < 0.7 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const stewardSupport = Math.random() < 0.75 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const communication = Math.random() < 0.65 ? (4 + Math.floor(Math.random() * 2)) : (3 + Math.floor(Math.random() * 2));
    const wouldRecommend = overallSat >= 4 ? 'Y' : (Math.random() < 0.7 ? 'Y' : 'N');

    // Generate realistic comments
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

  // Clear existing data (keep headers)
  const lastRow = satisfactionSheet.getLastRow();
  if (lastRow > 3) {
    satisfactionSheet.getRange(4, 1, lastRow - 3, 10).clear();
  }

  // Write new data
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
    // Populate Steward Workload
    populateStewardWorkload();

    // Populate Member Satisfaction
    populateMemberSatisfaction();

    // Note: Other analytics sheets (Trends, Location, etc.) use formulas and auto-populate
    // Performance, Quick Stats, Executive Dashboard, KPI Performance, etc. all use formulas

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
