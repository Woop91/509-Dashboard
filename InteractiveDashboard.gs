// ------------------------------------------------------------------------====
// INTERACTIVE DASHBOARD - USER-SELECTABLE METRICS & CHARTS
// ------------------------------------------------------------------------====
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
// ------------------------------------------------------------------------====

/**
 * Creates the Interactive Dashboard sheet with user-selectable controls
 */
function createInteractiveDashboardSheet(ss) {
  var sheet = ss.getSheetByName(SHEETS.INTERACTIVE_DASHBOARD);
  if (!sheet) sheet = ss.insertSheet(SHEETS.INTERACTIVE_DASHBOARD);

  sheet.clear();

  // ------------------------------------------------------------=========
  // HEADER SECTION
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // CONTROL PANEL - Row 4
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // METRIC CARDS SECTION - Row 10-18 (4 cards)
  // ------------------------------------------------------------=========
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

  cardPositions.forEachfunction((card, idx) {
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

  // ------------------------------------------------------------=========
  // CHART AREA 1 - Row 21-42 (Selected Metric 1)
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // CHART AREA 2 - Row 21-42 (Comparison or Selected Metric 2)
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // PIE CHARTS SECTION - Row 45-65
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // WAREHOUSE-STYLE LOCATION CHART - Row 68-88
  // ------------------------------------------------------------=========
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

  // ------------------------------------------------------------=========
  // DATA TABLE - Top Performers/Issues - Row 91-110
  // ------------------------------------------------------------=========
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
    "Light & Clean"
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

    // Create primary chart - pass data to avoid refetch
    createDynamicChart(sheet, metric1, chartType1, metrics, "A22", 10, 20, grievanceData, memberData);

    // Create comparison chart if enabled
    if (enableComparison === "Yes") {
      createDynamicChart(sheet, metric2, chartType2, metrics, "L22", 9, 20, grievanceData, memberData);
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
  metrics.activeMembers = memberData.slice(1).filter(function(row) { return row[10] === 'Active'; }).length;
  metrics.totalStewards = memberData.slice(1).filter(function(row) { return row[9] === 'Yes'; }).length;
  metrics.unit8Members = memberData.slice(1).filter(function(row) { return row[5] === 'Unit 8'; }).length;
  metrics.unit10Members = memberData.slice(1).filter(function(row) { return row[5] === 'Unit 10'; }).length;

  // Grievance metrics
  metrics.totalGrievances = grievanceData.length - 1;
  metrics.activeGrievances = grievanceData.slice(1).filter(function(row) {
    row[4] && (row[4].startsWith('Filed') || row[4] === 'Pending Decision')).length;
  metrics.resolvedGrievances = grievanceData.slice(1).filter(function(row) {
    row[4] && row[4].startsWith('Resolved')).length;

  const resolvedData = grievanceData.slice(1).filter(function(row) { return row[4] && row[4].startsWith('Resolved'); });
  metrics.grievancesWon = resolvedData.filter(function(row) { return row[24] && row[24].includes('Won'); }).length;
  metrics.grievancesLost = resolvedData.filter(function(row) { return row[24] && row[24].includes('Lost'); }).length;

  metrics.winRate = metrics.resolvedGrievances > 0
    ? ((metrics.grievancesWon / metrics.resolvedGrievances) * 100).toFixed(1)
    : 0;

  metrics.overdueGrievances = grievanceData.slice(1).filter(function(row) { return row[28] === 'YES'; }).length;

  // Additional metrics
  metrics.inMediation = grievanceData.slice(1).filter(function(row) { return row[4] === 'In Mediation'; }).length;
  metrics.inArbitration = grievanceData.slice(1).filter(function(row) { return row[4] === 'In Arbitration'; }).length;

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
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} metricName - Metric to display
 * @param {string} chartType - Type of chart (Donut, Pie, Bar, etc)
 * @param {Object} metrics - Pre-calculated metrics
 * @param {string} startCell - Starting cell for chart
 * @param {number} width - Chart width multiplier
 * @param {number} height - Chart height multiplier
 * @param {Array<Array>} grievanceData - Grievance data (to avoid refetch)
 * @param {Array<Array>} memberData - Member data (to avoid refetch)
 */
function createDynamicChart(sheet, metricName, chartType, metrics, startCell, width, height, grievanceData, memberData) {
  try {
    // Remove existing charts in this area first
    const charts = sheet.getCharts();
    charts.forEach(function(chart) {
      const anchor = chart.getContainerInfo().getAnchorRow();
      // Extract row number correctly (handles cells like "AA22" not just "A22")
      if (anchor >= parseInt(startCell.match(/\d+/)[0])) {
        sheet.removeChart(chart);
      }
    });

    // Get data based on metric selection - pass data to avoid refetch
    const chartData = getChartDataForMetric(metricName, metrics, grievanceData, memberData);

  if (!chartData || chartData.length === 0) return;

  // Create chart based on type
  var chartBuilder;
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
  } catch (error) {
    handleError(error, `createDynamicChart(${metricName})`, false);
  }
}

/**
 * Write chart data to sheet (hidden area)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target sheet
 * @param {string} startCell - Starting cell reference
 * @param {Array<Array>} data - Data to write
 */
function writeChartData(sheet, startCell, data) {
  if (!data || data.length === 0) return;

  try {
    const range = sheet.getRange(startCell).offset(1, 0, data.length, 2);
    range.setValues(data);

    // Hide this data area - OPTIMIZED: batch hide instead of one-by-one
    const startRow = range.getRow();
    sheet.hideRows(startRow, data.length);
  } catch (error) {
    handleError(error, 'writeChartData', false);
  }
}

/**
 * Get chart data for specific metric
 */
/**
 * Gets chart data based on selected metric
 * @param {string} metricName - Name of the metric to chart
 * @param {Object} metrics - Pre-calculated metrics object
 * @param {Array<Array>} grievanceData - Grievance data (passed to avoid refetch)
 * @param {Array<Array>} memberData - Member data (passed to avoid refetch)
 * @returns {Array<Array>} Chart data as [label, value] pairs
 */
function getChartDataForMetric(metricName, metrics, grievanceData, memberData) {
  // Data validation
  if (!grievanceData || !memberData) {
    logWarning('getChartDataForMetric', 'Missing data parameters');
    return [["No Data", 0]];
  }

  switch (metricName) {
    case "Total Members":
      return [
        ["Active", metrics.activeMembers],
        ["Inactive", metrics.totalMembers - metrics.activeMembers]
      ];

    case "Active Grievances":
      // Count actual grievances by step from Grievance Log
      const stepCounts = {};
      grievanceData.slice(1).forEach(function(row) {
        const status = row[4]; // Status column (E)
        const step = row[5];    // Current Step column (F)
        if (status && (status.includes('Filed') || status === 'Pending Decision' || status === 'Open')) {
          stepCounts[step] = (stepCounts[step] || 0) + 1;
        }
      });

      const stepData = Object.entries(stepCounts)
        .filterfunction(([step]) { return step && step !== 'Current Step'; })
        .mapfunction(([step, count]) { return [step, count]; });

      return stepData.length > 0 ? stepData : [["No Active Grievances", 0]];

    case "Win Rate %":
      return [
        ["Won", metrics.grievancesWon],
        ["Lost", metrics.grievancesLost]
      ];

    case "Grievances by Type":
      // Count grievances by type/category
      const typeCounts = {};
      grievanceData.slice(1).forEach(function(row) {
        const type = row[22]; // Issue Category column (W)
        if (type && type !== 'Issue Category') {
          typeCounts[type] = (typeCounts[type] || 0) + 1;
        }
      });

      const typeData = Object.entries(typeCounts)
        .sortfunction((a, b) { return b[1] - a[1]; })
        .slice(0, 10)
        .mapfunction(([type, count]) { return [type, count]; });

      return typeData.length > 0 ? typeData : [["No Data", 0]];

    case "Grievances by Location":
      // Count grievances by location
      const locationCounts = {};
      grievanceData.slice(1).forEach(function(row) {
        const location = row[25]; // Work Location column (Z)
        if (location && location !== 'Work Location (Site)') {
          locationCounts[location] = (locationCounts[location] || 0) + 1;
        }
      });

      const locationData = Object.entries(locationCounts)
        .sortfunction((a, b) { return b[1] - a[1]; })
        .slice(0, 10)
        .mapfunction(([location, count]) { return [location, count]; });

      return locationData.length > 0 ? locationData : [["No Data", 0]];

    case "Grievances by Step":
      // Count all grievances by step
      const allStepCounts = {};
      grievanceData.slice(1).forEach(function(row) {
        const step = row[5]; // Current Step column (F)
        if (step && step !== 'Current Step') {
          allStepCounts[step] = (allStepCounts[step] || 0) + 1;
        }
      });

      const allStepData = Object.entries(allStepCounts)
        .mapfunction(([step, count]) { return [step, count]; });

      return allStepData.length > 0 ? allStepData : [["No Data", 0]];

    case "Unit 8 Members":
      const unit8Count = memberData.slice(1).filter(function(row) { return row[5] === 'Unit 8'; }).length;
      return [["Unit 8", unit8Count]];

    case "Unit 10 Members":
      const unit10Count = memberData.slice(1).filter(function(row) { return row[5] === 'Unit 10'; }).length;
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
  grievanceData.slice(1).forEach(function(row) {
    const status = row[4] || 'Unknown';
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  });

  const data = Object.entries(statusCounts).mapfunction(([status, count]) { return [status, count]; });

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
  grievanceData.slice(1).forEach(function(row) {
    const location = row[4] || 'Unknown';  // Adjust column as needed
    locationCounts[location] = (locationCounts[location] || 0) + 1;
  });

  // Get top 10
  const topLocations = Object.entries(locationCounts)
    .sortfunction((a, b) { return b[1] - a[1]; })
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
  grievanceData.slice(1).forEach(function(row) {
    const location = row[4] || 'Unknown';
    locationCounts[location] = (locationCounts[location] || 0) + 1;
  });

  const topLocations = Object.entries(locationCounts)
    .sortfunction((a, b) { return b[1] - a[1]; })
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

  var tableData = [];

  // Generate table based on selected metric
  switch (metricName) {
    case "Grievances by Type":
    case "Issue Category":
      // Show top grievance types with detailed breakdown
      const typeCounts = {};
      const typeActive = {};
      const typeResolved = {};
      const typeWon = {};

      grievanceData.slice(1).forEach(function(row) {
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
        .sortfunction((a, b) { return b[1] - a[1]; })
        .slice(0, 15)
        .mapfunction(([type, total], index) {
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

      grievanceData.slice(1).forEach(function(row) {
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
        .sortfunction((a, b) { return b[1] - a[1]; })
        .slice(0, 15)
        .mapfunction(([location, total], index) {
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

      grievanceData.slice(1).forEach(function(row) {
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
        .sortfunction((a, b) { return (stewardActive[b.name] || 0) - (stewardActive[a.name] || 0); })
        .slice(0, 15)
        .mapfunction(([steward, total], index) {
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
      grievanceData.slice(1).forEach(function(row) {
        const location = row[25];
        if (location && location !== 'Work Location (Site)') {
          defaultLocationCounts[location] = (defaultLocationCounts[location] || 0) + 1;
        }
      });

      tableData = Object.entries(defaultLocationCounts)
        .sortfunction((a, b) { return b[1] - a[1]; })
        .slice(0, 15)
        .mapfunction(([location, count], index) {
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
  var primaryColor, accentColor;

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
    default:
      primaryColor = COLORS.PRIMARY_BLUE;
      accentColor = COLORS.ACCENT_TEAL;
  }

  // Apply theme colors to headers
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
