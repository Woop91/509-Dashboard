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
  LOCATION: "🗺️ Location Analytics",
  TYPE_ANALYSIS: "📊 Type Analysis",
  EXECUTIVE_DASHBOARD: "💼 Executive Dashboard",
  KPI_PERFORMANCE: "📊 KPI Performance Dashboard",
  MEMBER_ENGAGEMENT: "👥 Member Engagement",
  COST_IMPACT: "💰 Cost Impact",
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
     "Supervisors", "Managers", "Stewards", "Grievance Status", "Grievance Step",
     "Issue Category", "Articles Violated", "Communication Methods"],

    ["Coordinator", "Boston HQ", "Unit A - Administrative", "Monday", "Yes",
     "Sarah Johnson", "Michael Chen", "Jane Smith", "Open", "Informal",
     "Discipline", "Art. 1 - Recognition", "Email"],

    ["Analyst", "Worcester Office", "Unit B - Technical", "Tuesday", "No",
     "Mike Wilson", "Lisa Anderson", "John Doe", "Pending Info", "Step I",
     "Workload", "Art. 2 - Union Security", "Phone"],

    ["Case Manager", "Springfield Branch", "Unit C - Support Services", "Wednesday", "",
     "Emily Davis", "Robert Brown", "Mary Johnson", "Settled", "Step II",
     "Scheduling", "Art. 3 - Management Rights", "Text"],

    ["Specialist", "Cambridge Office", "Unit D - Operations", "Thursday", "",
     "Tom Harris", "Jennifer Lee", "Bob Wilson", "Withdrawn", "Step III",
     "Pay", "Art. 4 - No Discrimination", "In Person"],

    ["Senior Analyst", "Lowell Center", "Unit E - Field Services", "Friday", "",
     "Amanda White", "David Martinez", "Alice Brown", "Closed", "Mediation",
     "Discrimination", "Art. 5 - Union Business", ""],

    ["Team Lead", "Quincy Station", "", "Saturday", "",
     "Chris Taylor", "Susan Garcia", "Tom Davis", "Appealed", "Arbitration",
     "Safety", "Art. 23 - Grievance Procedure", ""],

    ["Director", "Remote/Hybrid", "", "Sunday", "",
     "Patricia Moore", "James Wilson", "Sarah Martinez", "", "",
     "Benefits", "Art. 24 - Discipline", ""],

    ["Manager", "Brockton Office", "", "", "",
     "Kevin Anderson", "Nancy Taylor", "Kevin Jones", "", "",
     "Training", "Art. 25 - Hours of Work", ""],

    ["Assistant", "Lynn Location", "", "", "",
     "Michelle Lee", "Richard White", "Linda Garcia", "", "",
     "Other", "Art. 26 - Overtime", ""],

    ["Associate", "Salem Office", "", "", "",
     "Brandon Scott", "Angela Moore", "Daniel Kim", "", "",
     "Harassment", "Art. 27 - Seniority", ""],

    ["Technician", "", "", "", "",
     "Jessica Green", "Christopher Lee", "Rachel Adams", "", "",
     "Equipment", "Art. 28 - Layoff", ""],

    ["Administrator", "", "", "", "",
     "Andrew Clark", "Melissa Wright", "", "", "",
     "Leave", "Art. 29 - Sick Leave", ""],

    ["Support Staff", "", "", "", "",
     "Rachel Brown", "Timothy Davis", "", "", "",
     "Grievance Process", "Art. 30 - Vacation", ""]
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

  const memberMetrics = [
    ["Total Members", "=COUNTA('Member Directory'!A:A)-1", "👥"],
    ["Active Stewards", "=COUNTIF('Member Directory'!J:J,\"Yes\")", "🛡️"],
    ["Avg Open Rate", "=TEXT(AVERAGE('Member Directory'!R:R)/100,\"0.0%\")", "📧"],
    ["YTD Vol. Hours", "=SUM('Member Directory'!S:S)", "🙋"]
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

  const grievanceMetrics = [
    ["Open Grievances", "=COUNTIF('Grievance Log'!E:E,\"Open\")", "🔴"],
    ["Pending Info", "=COUNTIF('Grievance Log'!E:E,\"Pending Info\")", "🟡"],
    ["Settled (This Mo.)", "=COUNTIFS('Grievance Log'!E:E,\"Settled\",'Grievance Log'!R:R,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))", "🟢"],
    ["Avg Days Open", "=ROUND(AVERAGE(FILTER('Grievance Log'!S:S,'Grievance Log'!E:E=\"Open\")),0)", "⏱️"]
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

  const engagementMetrics = [
    ["Virtual Mtgs", "=COUNTIF('Member Directory'!N:N,\">=\"&TODAY()-30)"],
    ["In-Person Mtgs", "=COUNTIF('Member Directory'!O:O,\">=\"&TODAY()-30)"],
    ["Local Interest", "=COUNTIF('Member Directory'!T:T,\"Yes\")"],
    ["Chapter Interest", "=COUNTIF('Member Directory'!U:U,\"Yes\")"]
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

  // Formula to populate upcoming deadlines (open grievances with deadlines in next 14 days)
  dashboard.getRange("A22").setFormula(
    '=IFERROR(QUERY(\'Grievance Log\'!A:U, ' +
    '"SELECT A, C, T, U, E ' +
    'WHERE E = \'Open\' AND T IS NOT NULL AND T <= date \'"&TEXT(TODAY()+14,"yyyy-mm-dd")&"\' ' +
    'ORDER BY T ASC ' +
    'LIMIT 10", 0), "No upcoming deadlines")'
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

  // Grievances by Status
  analytics.getRange("A3").setValue("Grievances by Status");
  analytics.getRange("A4:B4").setValues([["Status", "Count"]]).setFontWeight("bold");
  analytics.getRange("A5").setFormula('=UNIQUE(FILTER(\'Grievance Log\'!E:E, \'Grievance Log\'!E:E<>"", \'Grievance Log\'!E:E<>"Status"))');
  analytics.getRange("B5").setFormula('=ARRAYFORMULA(IF(A5:A<>"", COUNTIF(\'Grievance Log\'!E:E, A5:A), ""))');

  // Grievances by Unit
  analytics.getRange("D3").setValue("Grievances by Unit");
  analytics.getRange("D4:E4").setValues([["Unit", "Count"]]).setFontWeight("bold");
  analytics.getRange("D5").setFormula('=UNIQUE(FILTER(\'Grievance Log\'!Y:Y, \'Grievance Log\'!Y:Y<>"", \'Grievance Log\'!Y:Y<>"Unit"))');
  analytics.getRange("E5").setFormula('=ARRAYFORMULA(IF(D5:D<>"", COUNTIF(\'Grievance Log\'!Y:Y, D5:D), ""))');

  // Members by Location
  analytics.getRange("G3").setValue("Members by Location");
  analytics.getRange("G4:H4").setValues([["Location", "Count"]]).setFontWeight("bold");
  analytics.getRange("G5").setFormula('=UNIQUE(FILTER(\'Member Directory\'!E:E, \'Member Directory\'!E:E<>"", \'Member Directory\'!E:E<>"Work Location (Site)"))');
  analytics.getRange("H5").setFormula('=ARRAYFORMULA(IF(G5:G<>"", COUNTIF(\'Member Directory\'!E:E, G5:G), ""))');

  // Steward Workload
  analytics.getRange("J3").setValue("Steward Workload");
  analytics.getRange("J4:K4").setValues([["Steward", "Open Cases"]]).setFontWeight("bold");
  analytics.getRange("J5").setFormula('=UNIQUE(FILTER(\'Grievance Log\'!AA:AA, \'Grievance Log\'!AA:AA<>"", \'Grievance Log\'!AA:AA<>"Assigned Steward (Name)"))');
  analytics.getRange("K5").setFormula('=ARRAYFORMULA(IF(J5:J<>"", COUNTIFS(\'Grievance Log\'!AA:AA, J5:J, \'Grievance Log\'!E:E, "Open"), ""))');

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

  const quickStats = [
    ["Active Members", "=COUNTA('Member Directory'!A2:A)", "", ""],
    ["Active Grievances", "=COUNTIF('Grievance Log'!E:E,\"Open\")", "", ""],
    ["Win Rate", "=TEXT(IFERROR(COUNTIFS('Grievance Log'!E:E,\"Resolved*\",'Grievance Log'!AB:AB,\"*Won*\")/COUNTIF('Grievance Log'!E:E,\"Resolved*\"),0),\"0%\")", "", ""],
    ["Avg Resolution (Days)", "=ROUND(AVERAGE('Grievance Log'!S:S),1)", "", ""],
    ["Overdue Cases", "=COUNTIF('Grievance Log'!U:U,\"<0\")", "", ""],
    ["Active Stewards", "=COUNTIF('Member Directory'!J:J,\"Yes\")", "", ""]
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
    ["Total Active Members", "=COUNTA('Member Directory'!A2:A)", ""],
    ["Total Active Grievances", "=COUNTIF('Grievance Log'!E:E,\"Open\")", ""],
    ["Overall Win Rate", "=TEXT(IFERROR(COUNTIFS('Grievance Log'!E:E,\"Resolved*\",'Grievance Log'!AB:AB,\"*Won*\")/COUNTIF('Grievance Log'!E:E,\"Resolved*\"),0),\"0.0%\")", ""],
    ["Avg Resolution Time (Days)", "=ROUND(AVERAGE('Grievance Log'!S:S),1)", ""],
    ["Cases Overdue", "=COUNTIF('Grievance Log'!U:U,\"<0\")", ""],
    ["Member Satisfaction Score", "=TEXT(AVERAGE('Member Satisfaction'!C:C),\"0.0\")", ""],
    ["Total Grievances Filed YTD", "=COUNTA('Grievance Log'!A2:A)", ""],
    ["Resolved Grievances", "=COUNTIF('Grievance Log'!E:E,\"Resolved\")", ""]
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

  // Member Directory validations
  const memberValidations = [
    { col: 4, configCol: 1 },   // Job Title
    { col: 5, configCol: 2 },   // Work Location
    { col: 6, configCol: 3 },   // Unit
    { col: 10, configCol: 5 },  // Is Steward
    { col: 11, configCol: 6 },  // Supervisor
    { col: 12, configCol: 7 },  // Manager
    { col: 13, configCol: 8 },  // Assigned Steward
    { col: 20, configCol: 5 },  // Interest: Local
    { col: 21, configCol: 5 },  // Interest: Chapter
    { col: 22, configCol: 5 },  // Interest: Allied
    { col: 24, configCol: 13 }  // Comm Methods
  ];

  memberValidations.forEach(v => {
    const configRange = config.getRange(2, v.configCol, 50, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configRange, true)
      .setAllowInvalid(false)
      .build();
    memberDir.getRange(2, v.col, 5000, 1).setDataValidation(rule);
  });

  // Grievance Log validations
  const grievanceValidations = [
    { col: 5, configCol: 9 },   // Status
    { col: 6, configCol: 10 },  // Current Step
    { col: 22, configCol: 12 }, // Articles Violated
    { col: 23, configCol: 11 }, // Issue Category
    { col: 25, configCol: 3 },  // Unit
    { col: 26, configCol: 2 },  // Work Location
    { col: 27, configCol: 8 }   // Assigned Steward
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

  // Grievance Log formulas
  for (let row = 2; row <= 100; row++) {
    // Filing Deadline (Incident Date + 21 days)
    grievanceLog.getRange(row, 8).setFormula(
      `=IF(G${row}<>"",G${row}+21,"")`
    );

    // Step I Decision Due (Date Filed + 30 days)
    grievanceLog.getRange(row, 10).setFormula(
      `=IF(I${row}<>"",I${row}+30,"")`
    );

    // Step II Appeal Due (Step I Decision Rcvd + 10 days)
    grievanceLog.getRange(row, 12).setFormula(
      `=IF(K${row}<>"",K${row}+10,"")`
    );

    // Step II Decision Due (Step II Appeal Filed + 30 days)
    grievanceLog.getRange(row, 14).setFormula(
      `=IF(M${row}<>"",M${row}+30,"")`
    );

    // Step III Appeal Due (Step II Decision Rcvd + 30 days)
    grievanceLog.getRange(row, 16).setFormula(
      `=IF(O${row}<>"",O${row}+30,"")`
    );

    // Days Open
    grievanceLog.getRange(row, 19).setFormula(
      `=IF(I${row}<>"",IF(R${row}<>"",R${row}-I${row},TODAY()-I${row}),"")`
    );

    // Next Action Due
    grievanceLog.getRange(row, 20).setFormula(
      `=IF(E${row}="Open",IF(F${row}="Step I",J${row},IF(F${row}="Step II",N${row},IF(F${row}="Step III",P${row},H${row}))),"")`
    );

    // Days to Deadline
    grievanceLog.getRange(row, 21).setFormula(
      `=IF(T${row}<>"",T${row}-TODAY(),"")`
    );
  }

  // Member Directory formulas
  for (let row = 2; row <= 100; row++) {
    // Has Open Grievance?
    memberDir.getRange(row, 26).setFormula(
      `=IF(COUNTIFS('Grievance Log'!B:B,A${row},'Grievance Log'!E:E,"Open")>0,"Yes","No")`
    );

    // Grievance Status Snapshot
    memberDir.getRange(row, 27).setFormula(
      `=IFERROR(INDEX('Grievance Log'!E:E,MATCH(A${row},'Grievance Log'!B:B,0)),"")`
    );

    // Next Grievance Deadline
    memberDir.getRange(row, 28).setFormula(
      `=IFERROR(INDEX('Grievance Log'!T:T,MATCH(A${row},'Grievance Log'!B:B,0)),"")`
    );
  }
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
    .addSubMenu(ui.createMenu("📋 Grievance Tools")
      .addItem("➕ Start New Grievance", "showStartGrievanceDialog"))
    .addSeparator()
    .addSubMenu(ui.createMenu("⚙️ Admin")
      .addItem("Seed 20k Members", "SEED_20K_MEMBERS")
      .addItem("Seed 5k Grievances", "SEED_5K_GRIEVANCES")
      .addSeparator()
      .addItem("Clear All Data", "clearAllData")
      .addItem("🗑️ Nuke All Seed Data", "nukeSeedData"))
    .addSeparator()
    .addSubMenu(ui.createMenu("♿ ADHD Features")
      .addItem("Hide Gridlines (Focus Mode)", "hideAllGridlines")
      .addItem("Show Gridlines", "showAllGridlines")
      .addItem("Reorder Sheets Logically", "reorderSheetsLogically")
      .addItem("Setup ADHD Defaults", "setupADHDDefaults"))
    .addSubMenu(ui.createMenu("👁️ Column Toggles")
      .addItem("Toggle Advanced Grievance Columns", "toggleGrievanceColumns")
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
function SEED_20K_MEMBERS() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Seed 20,000 Members',
    'This will add 20,000 member records. This may take 2-3 minutes. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  SpreadsheetApp.getActive().toast("🚀 Seeding 20,000 members...", "Processing", -1);

  const firstNames = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah", "Charles", "Karen", "Christopher", "Nancy", "Daniel", "Lisa", "Matthew", "Betty", "Anthony", "Margaret", "Mark", "Sandra", "Donald", "Ashley", "Steven", "Kimberly", "Paul", "Emily", "Andrew", "Donna", "Joshua", "Michelle"];
  const lastNames = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores"];

  const jobTitles = config.getRange("A2:A14").getValues().flat().filter(String);
  const locations = config.getRange("B2:B14").getValues().flat().filter(String);
  const units = config.getRange("C2:C7").getValues().flat().filter(String);
  const supervisors = config.getRange("F2:F14").getValues().flat().filter(String);
  const managers = config.getRange("G2:G14").getValues().flat().filter(String);
  const stewards = config.getRange("H2:H14").getValues().flat().filter(String);
  const commMethods = ["Email", "Phone", "Text", "In Person"];
  const times = ["Mornings", "Afternoons", "Evenings", "Weekends", "Flexible"];

  // Validate config data
  if (jobTitles.length === 0 || locations.length === 0 || units.length === 0 ||
      supervisors.length === 0 || managers.length === 0 || stewards.length === 0) {
    ui.alert('Error', 'Config data is incomplete. Please ensure all dropdown lists in Config sheet are populated.', ui.ButtonSet.OK);
    return;
  }

  const BATCH_SIZE = 1000;
  let data = [];

  for (let i = 1; i <= 20000; i++) {
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
    const memberID = "M" + String(i).padStart(6, '0');
    const jobTitle = jobTitles[Math.floor(Math.random() * jobTitles.length)];
    const location = locations[Math.floor(Math.random() * locations.length)];
    const unit = units[Math.floor(Math.random() * units.length)];
    const officeDays = ["Mon", "Tue", "Wed", "Thu", "Fri"][Math.floor(Math.random() * 5)];
    const email = `${firstName.toLowerCase()}.${lastName.toLowerCase()}${i}@union.org`;
    const phone = `(555) ${String(Math.floor(Math.random() * 900) + 100)}-${String(Math.floor(Math.random() * 9000) + 1000)}`;
    const isSteward = Math.random() > 0.95 ? "Yes" : "No";
    const supervisor = supervisors[Math.floor(Math.random() * supervisors.length)];
    const manager = managers[Math.floor(Math.random() * managers.length)];
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
      memberID, firstName, lastName, jobTitle, location, unit, officeDays,
      email, phone, isSteward, supervisor, manager, assignedSteward,
      lastVirtual, lastInPerson, lastSurvey, lastEmailOpen, openRate, volHours,
      localInterest, chapterInterest, alliedInterest, timestamp, commMethod, bestTime,
      "No", "", "", "", "", ""
    ];

    data.push(row);

    if (data.length === BATCH_SIZE) {
      try {
        memberDir.getRange(memberDir.getLastRow() + 1, 1, data.length, row.length).setValues(data);
        SpreadsheetApp.getActive().toast(`Added ${i} of 20,000 members...`, "Progress", 1);
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

  SpreadsheetApp.getActive().toast("✅ 20,000 members added!", "Complete", 5);
}

/* ===================== SEED 5,000 GRIEVANCES ===================== */
function SEED_5K_GRIEVANCES() {
  const ss = SpreadsheetApp.getActive();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Seed 5,000 Grievances',
    'This will add 5,000 grievance records. This may take 1-2 minutes. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  SpreadsheetApp.getActive().toast("🚀 Seeding 5,000 grievances...", "Processing", -1);

  // Get member data ONCE before the loop (CRITICAL FIX)
  const memberLastRow = memberDir.getLastRow();
  if (memberLastRow < 2) {
    ui.alert('Error', 'No members found. Please seed members first.', ui.ButtonSet.OK);
    return;
  }

  const allMemberData = memberDir.getRange(2, 1, memberLastRow - 1, 31).getValues();
  const memberIDs = allMemberData.map(row => row[0]).filter(String);

  const statuses = config.getRange("I2:I8").getValues().flat().filter(String);
  const steps = config.getRange("J2:J7").getValues().flat().filter(String);
  const articles = config.getRange("L2:L14").getValues().flat().filter(String);
  const categories = config.getRange("K2:K12").getValues().flat().filter(String);
  const stewards = config.getRange("H2:H14").getValues().flat().filter(String);

  // Validate config data
  if (statuses.length === 0 || steps.length === 0 || articles.length === 0 ||
      categories.length === 0 || stewards.length === 0) {
    ui.alert('Error', 'Config data is incomplete. Please ensure all dropdown lists in Config sheet are populated.', ui.ButtonSet.OK);
    return;
  }

  const BATCH_SIZE = 500;
  let data = [];
  let successCount = 0;

  for (let i = 1; i <= 5000; i++) {
    // Get random member
    const memberIndex = Math.floor(Math.random() * memberIDs.length);
    const memberID = memberIDs[memberIndex];
    const memberData = allMemberData[memberIndex];

    if (!memberData || !memberID) continue;

    const grievanceID = "G-" + String(i).padStart(6, '0');
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
    const resolution = isClosed ? ["Resolved favorably", "Withdrawn by member", "Settled with compromise", "No violation found"][Math.floor(Math.random() * 4)] : "";

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
        SpreadsheetApp.getActive().toast(`Added ${successCount} of 5,000 grievances...`, "Progress", 1);
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

  SpreadsheetApp.getActive().toast(`✅ ${successCount} grievances added! Updating member snapshots...`, "Processing", 2);
  updateMemberDirectorySnapshots();
  SpreadsheetApp.getActive().toast(`✅ ${successCount} grievances added and member snapshots updated!`, "Complete", 5);
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
