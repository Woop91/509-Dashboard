/**
 * 509 DASHBOARD - Member Directory & Grievance Tracking System
 *
 * This Google Sheets script creates and manages a comprehensive dashboard
 * for tracking union members and grievances for AFSCME-SEIU Local 509.
 *
 * Features:
 * - Member Directory with comprehensive member information
 * - Grievance Log with auto-calculated deadlines based on contract rules
 * - Real-time Dashboard with metrics and analytics
 * - Data validation using centralized Config sheet
 * - Batch processing for handling large datasets (20k members, 5k grievances)
 *
 * Reference: Collective Bargaining Agreement (CBA) Articles 8, 9, 23A
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const SHEETS = {
  CONFIG: "Config",
  MEMBER_DIR: "Member Directory",
  GRIEVANCE_LOG: "Grievance Log",
  DASHBOARD: "Dashboard",
  ANALYTICS: "Analytics Data",
  FEEDBACK: "Feedback & Development",
  MEMBER_SATISFACTION: "Member Satisfaction"
};

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main function to create and setup the entire 509 Dashboard
 * Creates all sheets, headers, formulas, and data validation
 */
function CREATE_509_DASHBOARD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clear existing sheets (except the first one to avoid errors)
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }

  // Rename first sheet to Config
  sheets[0].setName(SHEETS.CONFIG);

  // Create all sheets
  createConfigSheet(ss);
  createMemberDirectorySheet(ss);
  createGrievanceLogSheet(ss);
  createDashboardSheet(ss);
  createAnalyticsDataSheet(ss);
  createMemberSatisfactionSheet(ss);
  createFeedbackDevelopmentSheet(ss);

  // Setup data validation (must be done after all sheets are created)
  setupDataValidation(ss);

  SpreadsheetApp.getUi().alert('509 Dashboard created successfully!');
}

// ============================================================================
// CONFIG SHEET
// ============================================================================

/**
 * Creates the Config sheet with all dropdown options
 * This serves as the master source for data validation across all sheets
 */
function createConfigSheet(ss) {
  let config = ss.getSheetByName(SHEETS.CONFIG);
  if (!config) {
    config = ss.insertSheet(SHEETS.CONFIG);
  }

  config.clear();

  // Set up headers
  const headers = [
    "Job Titles", "Work Locations", "Units", "Steward Status",
    "Grievance Status", "Grievance Steps", "Grievance Types",
    "Outcomes", "Satisfaction Levels", "Engagement Levels"
  ];

  config.getRange(1, 1, 1, headers.length).setValues([headers]);
  config.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  // Job Titles (from CBA Appendix C)
  const jobTitles = [
    "Administrative Assistant I",
    "Administrative Assistant II",
    "Administrative Assistant III",
    "Case Manager I",
    "Case Manager II",
    "Case Manager III",
    "Program Coordinator I",
    "Program Coordinator II",
    "Program Coordinator III",
    "Social Worker I",
    "Social Worker II",
    "Social Worker III",
    "Clerk I",
    "Clerk II",
    "Clerk III",
    "Specialist I",
    "Specialist II",
    "Specialist III"
  ];

  // Work Locations (major state facilities)
  const workLocations = [
    "Boston - State House",
    "Boston - McCormack Building",
    "Boston - Saltonstall Building",
    "Springfield - State Office Building",
    "Worcester - State Office Complex",
    "Pittsfield - Regional Office",
    "Lowell - Regional Office",
    "New Bedford - Regional Office",
    "Hyannis - Regional Office",
    "Lawrence - Regional Office",
    "Brockton - Regional Office",
    "Fall River - Regional Office",
    "Framingham - Regional Office",
    "Quincy - Regional Office",
    "Remote/Hybrid"
  ];

  // Units (from CBA)
  const units = ["Unit 8", "Unit 10"];

  // Steward Status
  const stewardStatus = ["Yes", "No"];

  // Grievance Status
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

  // Grievance Steps (from CBA Article 23A)
  const grievanceSteps = [
    "Informal",
    "Step I - Immediate Supervisor",
    "Step II - Agency Head",
    "Step III - Human Resources",
    "Mediation",
    "Arbitration"
  ];

  // Grievance Types
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

  // Outcomes
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

  // Satisfaction Levels
  const satisfactionLevels = [
    "Very Satisfied",
    "Satisfied",
    "Neutral",
    "Dissatisfied",
    "Very Dissatisfied",
    "N/A"
  ];

  // Engagement Levels
  const engagementLevels = [
    "Very Active",
    "Active",
    "Moderately Active",
    "Inactive",
    "New Member"
  ];

  // Write all data to Config sheet
  const data = [
    jobTitles,
    workLocations,
    units,
    stewardStatus,
    grievanceStatus,
    grievanceSteps,
    grievanceTypes,
    outcomes,
    satisfactionLevels,
    engagementLevels
  ];

  for (let i = 0; i < data.length; i++) {
    const column = i + 1;
    const values = data[i].map(item => [item]);
    config.getRange(2, column, values.length, 1).setValues(values);
  }

  // Format the Config sheet
  config.setFrozenRows(1);
  config.autoResizeColumns(1, headers.length);
}

// ============================================================================
// MEMBER DIRECTORY SHEET
// ============================================================================

/**
 * Creates the Member Directory sheet
 * Contains comprehensive information about each union member
 */
function createMemberDirectorySheet(ss) {
  let memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!memberDir) {
    memberDir = ss.insertSheet(SHEETS.MEMBER_DIR);
  }

  memberDir.clear();

  // 31 columns for comprehensive member tracking
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
    "Date Joined Union",
    "Membership Status",
    "Total Grievances Filed",
    "Active Grievances",
    "Resolved Grievances",
    "Grievances Won",
    "Grievances Lost",
    "Last Grievance Date",
    "Engagement Level",
    "Events Attended (Last 12mo)",
    "Training Sessions Attended",
    "Committee Member",
    "Preferred Contact Method",
    "Emergency Contact Name",
    "Emergency Contact Phone",
    "Notes",
    "Date of Birth",
    "Hire Date",
    "Seniority Date",
    "Last Updated",
    "Updated By"
  ];

  memberDir.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = memberDir.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4A86E8");
  headerRange.setFontColor("white");

  // Freeze header row
  memberDir.setFrozenRows(1);

  // Set column widths
  memberDir.setColumnWidth(1, 100); // Member ID
  memberDir.setColumnWidth(2, 120); // First Name
  memberDir.setColumnWidth(3, 120); // Last Name
  memberDir.setColumnWidth(4, 200); // Job Title
  memberDir.setColumnWidth(5, 200); // Work Location
  memberDir.setColumnWidth(8, 200); // Email Address

  // Add formulas for calculated columns
  // These will auto-populate from the Grievance Log
  const formulaRow = 2;

  // Total Grievances Filed (Column M = 13)
  memberDir.getRange(formulaRow, 13).setFormula(
    `=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!B:B,A${formulaRow})`
  );

  // Active Grievances (Column N = 14)
  memberDir.getRange(formulaRow, 14).setFormula(
    `=COUNTIFS('${SHEETS.GRIEVANCE_LOG}'!B:B,A${formulaRow},'${SHEETS.GRIEVANCE_LOG}'!E:E,"Filed*")`
  );

  // Resolved Grievances (Column O = 15)
  memberDir.getRange(formulaRow, 15).setFormula(
    `=COUNTIFS('${SHEETS.GRIEVANCE_LOG}'!B:B,A${formulaRow},'${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved*")`
  );

  // Grievances Won (Column P = 16)
  memberDir.getRange(formulaRow, 16).setFormula(
    `=COUNTIFS('${SHEETS.GRIEVANCE_LOG}'!B:B,A${formulaRow},'${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved - Won")`
  );

  // Grievances Lost (Column Q = 17)
  memberDir.getRange(formulaRow, 17).setFormula(
    `=COUNTIFS('${SHEETS.GRIEVANCE_LOG}'!B:B,A${formulaRow},'${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved - Lost")`
  );

  // Last Grievance Date (Column R = 18)
  memberDir.getRange(formulaRow, 18).setFormula(
    `=IFERROR(MAX(FILTER('${SHEETS.GRIEVANCE_LOG}'!G:G,'${SHEETS.GRIEVANCE_LOG}'!B:B=A${formulaRow})),"")`
  );

  // Last Updated (Column AD = 30)
  memberDir.getRange(formulaRow, 30).setFormula(`=NOW()`);
}

// ============================================================================
// GRIEVANCE LOG SHEET
// ============================================================================

/**
 * Creates the Grievance Log sheet
 * Tracks all grievances with auto-calculated deadlines based on CBA rules
 *
 * CBA Article 23A Deadlines:
 * - Filing Deadline: 21 days from incident date
 * - Step I Decision: 30 days from filing
 * - Step II Appeal: 10 days from Step I decision
 * - Step II Decision: 30 days from appeal
 * - Step III Appeal: 10 days from Step II decision
 */
function createGrievanceLogSheet(ss) {
  let grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!grievanceLog) {
    grievanceLog = ss.insertSheet(SHEETS.GRIEVANCE_LOG);
  }

  grievanceLog.clear();

  // 28 columns for comprehensive grievance tracking
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
    "Step I Decision Date",
    "Step I Outcome",
    "Step II Appeal Deadline (10d)",
    "Step II Filed Date",
    "Step II Decision Due (30d)",
    "Step II Decision Date",
    "Step II Outcome",
    "Step III Appeal Deadline (10d)",
    "Step III Filed Date",
    "Step III Decision Date",
    "Mediation Date",
    "Arbitration Date",
    "Final Outcome",
    "Grievance Type",
    "Description",
    "Representative",
    "Notes",
    "Last Updated"
  ];

  grievanceLog.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = grievanceLog.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#E06666");
  headerRange.setFontColor("white");

  // Freeze header row
  grievanceLog.setFrozenRows(1);

  // Set column widths
  grievanceLog.setColumnWidth(1, 120); // Grievance ID
  grievanceLog.setColumnWidth(2, 100); // Member ID
  grievanceLog.setColumnWidth(25, 300); // Description
  grievanceLog.setColumnWidth(27, 300); // Notes

  // Add formulas for auto-calculated deadlines (starting at row 2)
  const formulaRow = 2;

  // Filing Deadline = Incident Date + 21 days (Column H = 8)
  grievanceLog.getRange(formulaRow, 8).setFormula(
    `=IF(G${formulaRow}<>"",G${formulaRow}+21,"")`
  );

  // Step I Decision Due = Date Filed + 30 days (Column J = 10)
  grievanceLog.getRange(formulaRow, 10).setFormula(
    `=IF(I${formulaRow}<>"",I${formulaRow}+30,"")`
  );

  // Step II Appeal Deadline = Step I Decision Date + 10 days (Column M = 13)
  grievanceLog.getRange(formulaRow, 13).setFormula(
    `=IF(K${formulaRow}<>"",K${formulaRow}+10,"")`
  );

  // Step II Decision Due = Step II Filed + 30 days (Column O = 15)
  grievanceLog.getRange(formulaRow, 15).setFormula(
    `=IF(N${formulaRow}<>"",N${formulaRow}+30,"")`
  );

  // Step III Appeal Deadline = Step II Decision Date + 10 days (Column R = 18)
  grievanceLog.getRange(formulaRow, 18).setFormula(
    `=IF(P${formulaRow}<>"",P${formulaRow}+10,"")`
  );

  // Last Updated (Column AB = 28)
  grievanceLog.getRange(formulaRow, 28).setFormula(`=NOW()`);

  // Format date columns
  const dateColumns = [7, 8, 9, 10, 11, 13, 14, 15, 16, 18, 19, 20, 21, 22, 28];
  dateColumns.forEach(col => {
    grievanceLog.getRange(2, col, 1000, 1).setNumberFormat("MM/dd/yyyy");
  });
}

// ============================================================================
// DASHBOARD SHEET
// ============================================================================

/**
 * Creates the Dashboard sheet with real-time metrics
 * All metrics are calculated from actual data (no simulated/fake data)
 */
function createDashboardSheet(ss) {
  let dashboard = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dashboard) {
    dashboard = ss.insertSheet(SHEETS.DASHBOARD);
  }

  dashboard.clear();

  // Title
  dashboard.getRange("A1").setValue("509 DASHBOARD - REAL-TIME METRICS");
  dashboard.getRange("A1").setFontSize(16).setFontWeight("bold");

  // Member Metrics Section
  dashboard.getRange("A3").setValue("MEMBER METRICS").setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");

  dashboard.getRange("A4").setValue("Total Members:");
  dashboard.getRange("B4").setFormula(`=COUNTA('${SHEETS.MEMBER_DIR}'!A:A)-1`);

  dashboard.getRange("A5").setValue("Active Members:");
  dashboard.getRange("B5").setFormula(`=COUNTIF('${SHEETS.MEMBER_DIR}'!L:L,"Active")`);

  dashboard.getRange("A6").setValue("Total Stewards:");
  dashboard.getRange("B6").setFormula(`=COUNTIF('${SHEETS.MEMBER_DIR}'!J:J,"Yes")`);

  dashboard.getRange("A7").setValue("Unit 8 Members:");
  dashboard.getRange("B7").setFormula(`=COUNTIF('${SHEETS.MEMBER_DIR}'!F:F,"Unit 8")`);

  dashboard.getRange("A8").setValue("Unit 10 Members:");
  dashboard.getRange("B8").setFormula(`=COUNTIF('${SHEETS.MEMBER_DIR}'!F:F,"Unit 10")`);

  // Grievance Metrics Section
  dashboard.getRange("A10").setValue("GRIEVANCE METRICS").setFontWeight("bold").setBackground("#E06666").setFontColor("white");

  dashboard.getRange("A11").setValue("Total Grievances:");
  dashboard.getRange("B11").setFormula(`=COUNTA('${SHEETS.GRIEVANCE_LOG}'!A:A)-1`);

  dashboard.getRange("A12").setValue("Active Grievances:");
  dashboard.getRange("B12").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"Filed*")`);

  dashboard.getRange("A13").setValue("Resolved Grievances:");
  dashboard.getRange("B13").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved*")`);

  dashboard.getRange("A14").setValue("Grievances Won:");
  dashboard.getRange("B14").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved - Won")`);

  dashboard.getRange("A15").setValue("Grievances Lost:");
  dashboard.getRange("B15").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"Resolved - Lost")`);

  dashboard.getRange("A16").setValue("Win Rate:");
  dashboard.getRange("B16").setFormula(`=IF(B13>0,B14/B13,0)`);
  dashboard.getRange("B16").setNumberFormat("0.00%");

  dashboard.getRange("A17").setValue("In Mediation:");
  dashboard.getRange("B17").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"In Mediation")`);

  dashboard.getRange("A18").setValue("In Arbitration:");
  dashboard.getRange("B18").setFormula(`=COUNTIF('${SHEETS.GRIEVANCE_LOG}'!E:E,"In Arbitration")`);

  // Engagement Metrics Section
  dashboard.getRange("A20").setValue("ENGAGEMENT METRICS").setFontWeight("bold").setBackground("#93C47D").setFontColor("white");

  dashboard.getRange("A21").setValue("Avg Events Attended:");
  dashboard.getRange("B21").setFormula(`=AVERAGE('${SHEETS.MEMBER_DIR}'!T:T)`);
  dashboard.getRange("B21").setNumberFormat("0.00");

  dashboard.getRange("A22").setValue("Avg Training Sessions:");
  dashboard.getRange("B22").setFormula(`=AVERAGE('${SHEETS.MEMBER_DIR}'!U:U)`);
  dashboard.getRange("B22").setNumberFormat("0.00");

  // Format the dashboard
  dashboard.setColumnWidth(1, 200);
  dashboard.setColumnWidth(2, 150);
}

// ============================================================================
// ANALYTICS DATA SHEET
// ============================================================================

/**
 * Creates the Analytics Data sheet
 * For advanced reporting and trend analysis
 */
function createAnalyticsDataSheet(ss) {
  let analytics = ss.getSheetByName(SHEETS.ANALYTICS);
  if (!analytics) {
    analytics = ss.insertSheet(SHEETS.ANALYTICS);
  }

  analytics.clear();

  const headers = [
    "Date",
    "Total Members",
    "Active Members",
    "Total Grievances",
    "Active Grievances",
    "Resolved Grievances",
    "Win Rate",
    "Avg Member Satisfaction",
    "Notes"
  ];

  analytics.getRange(1, 1, 1, headers.length).setValues([headers]);
  analytics.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  analytics.setFrozenRows(1);
}

// ============================================================================
// MEMBER SATISFACTION SHEET
// ============================================================================

/**
 * Creates the Member Satisfaction sheet
 * For tracking member feedback and satisfaction metrics
 */
function createMemberSatisfactionSheet(ss) {
  let satisfaction = ss.getSheetByName(SHEETS.MEMBER_SATISFACTION);
  if (!satisfaction) {
    satisfaction = ss.insertSheet(SHEETS.MEMBER_SATISFACTION);
  }

  satisfaction.clear();

  const headers = [
    "Response ID",
    "Member ID",
    "Date",
    "Overall Satisfaction",
    "Union Representation",
    "Communication",
    "Grievance Process",
    "Training & Development",
    "Events & Activities",
    "Comments",
    "Follow-up Needed"
  ];

  satisfaction.getRange(1, 1, 1, headers.length).setValues([headers]);
  satisfaction.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  satisfaction.setFrozenRows(1);
}

// ============================================================================
// FEEDBACK & DEVELOPMENT SHEET
// ============================================================================

/**
 * Creates the Feedback & Development sheet
 * For tracking member development and training needs
 */
function createFeedbackDevelopmentSheet(ss) {
  let feedback = ss.getSheetByName(SHEETS.FEEDBACK);
  if (!feedback) {
    feedback = ss.insertSheet(SHEETS.FEEDBACK);
  }

  feedback.clear();

  const headers = [
    "Feedback ID",
    "Member ID",
    "Date",
    "Type",
    "Topic",
    "Description",
    "Priority",
    "Status",
    "Assigned To",
    "Resolution",
    "Date Resolved"
  ];

  feedback.getRange(1, 1, 1, headers.length).setValues([headers]);
  feedback.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  feedback.setFrozenRows(1);
}

// ============================================================================
// DATA VALIDATION
// ============================================================================

/**
 * Sets up data validation rules using Config sheet as the source
 * Must be called after all sheets are created
 */
function setupDataValidation(ss) {
  const config = ss.getSheetByName(SHEETS.CONFIG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Member Directory validations
  // Job Title (Column D)
  const jobTitleRange = config.getRange("A2:A" + (config.getLastRow()));
  const jobTitleRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(jobTitleRange, true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange("D2:D1000").setDataValidation(jobTitleRule);

  // Work Location (Column E)
  const locationRange = config.getRange("B2:B" + (config.getLastRow()));
  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(locationRange, true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange("E2:E1000").setDataValidation(locationRule);

  // Unit (Column F)
  const unitRange = config.getRange("C2:C" + (config.getLastRow()));
  const unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(unitRange, true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange("F2:F1000").setDataValidation(unitRule);

  // Steward Status (Column J)
  const stewardRange = config.getRange("D2:D" + (config.getLastRow()));
  const stewardRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(stewardRange, true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange("J2:J1000").setDataValidation(stewardRule);

  // Engagement Level (Column S)
  const engagementRange = config.getRange("J2:J" + (config.getLastRow()));
  const engagementRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(engagementRange, true)
    .setAllowInvalid(false)
    .build();
  memberDir.getRange("S2:S1000").setDataValidation(engagementRule);

  // Grievance Log validations
  // Status (Column E)
  const statusRange = config.getRange("E2:E" + (config.getLastRow()));
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(statusRange, true)
    .setAllowInvalid(false)
    .build();
  grievanceLog.getRange("E2:E5000").setDataValidation(statusRule);

  // Current Step (Column F)
  const stepRange = config.getRange("F2:F" + (config.getLastRow()));
  const stepRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(stepRange, true)
    .setAllowInvalid(false)
    .build();
  grievanceLog.getRange("F2:F5000").setDataValidation(stepRule);

  // Grievance Type (Column X = 24)
  const typeRange = config.getRange("G2:G" + (config.getLastRow()));
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(typeRange, true)
    .setAllowInvalid(false)
    .build();
  grievanceLog.getRange("X2:X5000").setDataValidation(typeRule);

  // Outcomes (Columns L, P for Step I and Step II outcomes)
  const outcomeRange = config.getRange("H2:H" + (config.getLastRow()));
  const outcomeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(outcomeRange, true)
    .setAllowInvalid(false)
    .build();
  grievanceLog.getRange("L2:L5000").setDataValidation(outcomeRule);
  grievanceLog.getRange("Q2:Q5000").setDataValidation(outcomeRule);
}

// ============================================================================
// DATA SEEDING FUNCTIONS
// ============================================================================

/**
 * Seed 20,000 members into the Member Directory
 * Uses batch processing for performance (1000 rows at a time)
 */
function SEED_20K_MEMBERS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  // Get dropdown values from Config
  const jobTitles = config.getRange("A2:A19").getValues().flat().filter(String);
  const locations = config.getRange("B2:B16").getValues().flat().filter(String);
  const units = ["Unit 8", "Unit 10"];
  const stewardStatus = ["Yes", "No"];
  const membershipStatus = ["Active", "Inactive", "On Leave"];
  const engagementLevels = config.getRange("J2:J6").getValues().flat().filter(String);
  const contactMethods = ["Email", "Phone", "Text", "Mail"];

  const firstNames = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah", "Charles", "Karen"];
  const lastNames = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin"];

  const TOTAL_MEMBERS = 20000;
  const BATCH_SIZE = 1000;
  const batches = Math.ceil(TOTAL_MEMBERS / BATCH_SIZE);

  SpreadsheetApp.getUi().alert(`Starting to seed ${TOTAL_MEMBERS} members in ${batches} batches...`);

  for (let batch = 0; batch < batches; batch++) {
    const startRow = batch * BATCH_SIZE + 2; // +2 for header row
    const rowsInBatch = Math.min(BATCH_SIZE, TOTAL_MEMBERS - (batch * BATCH_SIZE));
    const data = [];

    for (let i = 0; i < rowsInBatch; i++) {
      const memberNum = (batch * BATCH_SIZE) + i + 1;
      const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
      const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
      const isSteward = stewardStatus[Math.floor(Math.random() * 10) < 1 ? 0 : 1]; // 10% stewards

      const row = [
        `MEM${String(memberNum).padStart(6, '0')}`, // Member ID
        firstName, // First Name
        lastName, // Last Name
        jobTitles[Math.floor(Math.random() * jobTitles.length)], // Job Title
        locations[Math.floor(Math.random() * locations.length)], // Work Location
        units[Math.floor(Math.random() * units.length)], // Unit
        "Mon-Fri", // Office Days
        `${firstName.toLowerCase()}.${lastName.toLowerCase()}@mass.gov`, // Email
        `617-555-${String(Math.floor(Math.random() * 10000)).padStart(4, '0')}`, // Phone
        isSteward, // Steward
        new Date(2020 + Math.floor(Math.random() * 5), Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1), // Date Joined
        membershipStatus[Math.floor(Math.random() * membershipStatus.length)], // Membership Status
        "", // Total Grievances (formula will calculate)
        "", // Active Grievances (formula will calculate)
        "", // Resolved Grievances (formula will calculate)
        "", // Grievances Won (formula will calculate)
        "", // Grievances Lost (formula will calculate)
        "", // Last Grievance Date (formula will calculate)
        engagementLevels[Math.floor(Math.random() * engagementLevels.length)], // Engagement Level
        Math.floor(Math.random() * 25), // Events Attended
        Math.floor(Math.random() * 15), // Training Sessions
        isSteward === "Yes" ? "Executive Board" : "", // Committee
        contactMethods[Math.floor(Math.random() * contactMethods.length)], // Preferred Contact
        "", // Emergency Contact Name
        "", // Emergency Contact Phone
        "", // Notes
        new Date(1960 + Math.floor(Math.random() * 40), Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1), // DOB
        new Date(2015 + Math.floor(Math.random() * 10), Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1), // Hire Date
        new Date(2015 + Math.floor(Math.random() * 10), Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1), // Seniority Date
        new Date(), // Last Updated
        "SEED_SCRIPT" // Updated By
      ];

      data.push(row);
    }

    // Write batch to sheet
    memberDir.getRange(startRow, 1, rowsInBatch, data[0].length).setValues(data);

    // Copy formulas down for this batch
    const formulaSourceRow = 2;
    memberDir.getRange(formulaSourceRow, 13, 1, 6).copyTo(
      memberDir.getRange(startRow, 13, rowsInBatch, 6),
      SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
      false
    );

    SpreadsheetApp.flush();
    Logger.log(`Batch ${batch + 1}/${batches} complete (${rowsInBatch} members)`);
  }

  SpreadsheetApp.getUi().alert(`Successfully seeded ${TOTAL_MEMBERS} members!`);
}

/**
 * Seed 5,000 grievances into the Grievance Log
 * Uses batch processing for performance (500 rows at a time)
 */
function SEED_5K_GRIEVANCES() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  // Get member IDs from Member Directory
  const memberData = memberDir.getRange("A2:D" + memberDir.getLastRow()).getValues();
  const members = memberData.filter(row => row[0]); // Filter out empty rows

  if (members.length === 0) {
    SpreadsheetApp.getUi().alert("Please seed members first using SEED_20K_MEMBERS()");
    return;
  }

  // Get dropdown values from Config
  const grievanceStatus = config.getRange("E2:E12").getValues().flat().filter(String);
  const grievanceSteps = config.getRange("F2:F7").getValues().flat().filter(String);
  const grievanceTypes = config.getRange("G2:G16").getValues().flat().filter(String);
  const outcomes = config.getRange("H2:H10").getValues().flat().filter(String);

  const TOTAL_GRIEVANCES = 5000;
  const BATCH_SIZE = 500;
  const batches = Math.ceil(TOTAL_GRIEVANCES / BATCH_SIZE);

  SpreadsheetApp.getUi().alert(`Starting to seed ${TOTAL_GRIEVANCES} grievances in ${batches} batches...`);

  for (let batch = 0; batch < batches; batch++) {
    const startRow = batch * BATCH_SIZE + 2;
    const rowsInBatch = Math.min(BATCH_SIZE, TOTAL_GRIEVANCES - (batch * BATCH_SIZE));
    const data = [];

    for (let i = 0; i < rowsInBatch; i++) {
      const grievanceNum = (batch * BATCH_SIZE) + i + 1;
      const member = members[Math.floor(Math.random() * members.length)];
      const memberId = member[0];
      const firstName = member[1];
      const lastName = member[2];

      // Random dates within last 2 years
      const incidentDate = new Date(Date.now() - Math.floor(Math.random() * 730) * 24 * 60 * 60 * 1000);
      const status = grievanceStatus[Math.floor(Math.random() * grievanceStatus.length)];
      const step = grievanceSteps[Math.floor(Math.random() * grievanceSteps.length)];
      const type = grievanceTypes[Math.floor(Math.random() * grievanceTypes.length)];

      // Determine if grievance has been filed
      const hasFiled = Math.random() > 0.1; // 90% have been filed
      const filedDate = hasFiled ? new Date(incidentDate.getTime() + Math.floor(Math.random() * 20) * 24 * 60 * 60 * 1000) : null;

      const row = [
        `GRV${String(grievanceNum).padStart(6, '0')}`, // Grievance ID
        memberId, // Member ID
        firstName, // First Name
        lastName, // Last Name
        status, // Status
        step, // Current Step
        incidentDate, // Incident Date
        "", // Filing Deadline (formula)
        filedDate, // Date Filed
        "", // Step I Decision Due (formula)
        null, // Step I Decision Date
        "", // Step I Outcome
        "", // Step II Appeal Deadline (formula)
        null, // Step II Filed Date
        "", // Step II Decision Due (formula)
        null, // Step II Decision Date
        "", // Step II Outcome
        "", // Step III Appeal Deadline (formula)
        null, // Step III Filed Date
        null, // Step III Decision Date
        null, // Mediation Date
        null, // Arbitration Date
        status.startsWith("Resolved") ? outcomes[Math.floor(Math.random() * outcomes.length)] : "Pending", // Final Outcome
        type, // Grievance Type
        `${type} - Auto-generated grievance description`, // Description
        "Union Rep " + (Math.floor(Math.random() * 10) + 1), // Representative
        "", // Notes
        new Date() // Last Updated
      ];

      data.push(row);
    }

    // Write batch to sheet
    grievanceLog.getRange(startRow, 1, rowsInBatch, data[0].length).setValues(data);

    // Copy formulas down for this batch
    const formulaSourceRow = 2;
    // Copy deadline calculation formulas (columns H, J, M, O, R, AB)
    [8, 10, 13, 15, 18, 28].forEach(col => {
      grievanceLog.getRange(formulaSourceRow, col).copyTo(
        grievanceLog.getRange(startRow, col, rowsInBatch, 1),
        SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
        false
      );
    });

    SpreadsheetApp.flush();
    Logger.log(`Batch ${batch + 1}/${batches} complete (${rowsInBatch} grievances)`);
  }

  SpreadsheetApp.getUi().alert(`Successfully seeded ${TOTAL_GRIEVANCES} grievances!`);
}

// ============================================================================
// CUSTOM MENU
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('509 Dashboard')
    .addItem('Create Dashboard', 'CREATE_509_DASHBOARD')
    .addSeparator()
    .addItem('Seed 20K Members', 'SEED_20K_MEMBERS')
    .addItem('Seed 5K Grievances', 'SEED_5K_GRIEVANCES')
    .addToUi();
}
