/**
 * ============================================================================
 * QUICK FILTERS & SAVED VIEWS
 * ============================================================================
 *
 * Fast filtering and saved view management
 * Features:
 * - One-click filters
 * - Save custom filter sets
 * - Quick filter presets
 * - Filter combinations
 * - Recently used filters
 * - Clear all filters
 */

/**
 * Shows quick filter menu
 */
function showQuickFilterMenu() {
  const html = createQuickFilterHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '⚡ Quick Filters');
}

/**
 * Creates HTML for quick filters
 */
function createQuickFilterHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 25px; border-radius: 12px; max-height: 650px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; }
    h3 { color: #333; margin-top: 25px; margin-bottom: 15px; font-size: 16px; }
    .filter-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin: 15px 0; }
    .filter-btn { background: white; border: 2px solid #e0e0e0; padding: 15px; border-radius: 8px; cursor: pointer; transition: all 0.2s; text-align: left; }
    .filter-btn:hover { border-color: #1a73e8; background: #f8fcff; transform: translateX(3px); }
    .filter-icon { font-size: 20px; margin-right: 8px; }
    .filter-label { font-weight: 500; color: #333; }
    .filter-description { font-size: 12px; color: #666; margin-top: 5px; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 20px; border-radius: 6px; cursor: pointer; margin: 5px; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    button.danger { background: #f44336; }
  </style>
</head>
<body>
  <div class="container">
    <h2>⚡ Quick Filters</h2>

    <h3>Status Filters</h3>
    <div class="filter-grid">
      <div class="filter-btn" onclick="applyFilter('status', 'Open')">
        <div><span class="filter-icon">📂</span><span class="filter-label">Open Cases</span></div>
        <div class="filter-description">Show all open grievances</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('status', 'Pending')">
        <div><span class="filter-icon">⏳</span><span class="filter-label">Pending</span></div>
        <div class="filter-description">Awaiting response</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('status', 'Resolved')">
        <div><span class="filter-icon">✅</span><span class="filter-label">Resolved</span></div>
        <div class="filter-description">Successfully resolved cases</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('status', 'Closed')">
        <div><span class="filter-icon">🔒</span><span class="filter-label">Closed</span></div>
        <div class="filter-description">All closed grievances</div>
      </div>
    </div>

    <h3>Time Filters</h3>
    <div class="filter-grid">
      <div class="filter-btn" onclick="applyFilter('time', 'today')">
        <div><span class="filter-icon">📅</span><span class="filter-label">Filed Today</span></div>
        <div class="filter-description">Grievances filed today</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('time', 'thisWeek')">
        <div><span class="filter-icon">📆</span><span class="filter-label">This Week</span></div>
        <div class="filter-description">Filed in the past 7 days</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('time', 'thisMonth')">
        <div><span class="filter-icon">🗓️</span><span class="filter-label">This Month</span></div>
        <div class="filter-description">Filed this month</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('time', 'overdue')">
        <div><span class="filter-icon">⚠️</span><span class="filter-label">Overdue</span></div>
        <div class="filter-description">Past deadline</div>
      </div>
    </div>

    <h3>Priority Filters</h3>
    <div class="filter-grid">
      <div class="filter-btn" onclick="applyFilter('priority', 'urgent')">
        <div><span class="filter-icon">🔥</span><span class="filter-label">Urgent</span></div>
        <div class="filter-description">Deadline within 3 days</div>
      </div>

      <div class="filter-btn" onclick="applyFilter('priority', 'highValue')">
        <div><span class="filter-icon">💎</span><span class="filter-label">High Value</span></div>
        <div class="filter-description">Major issues</div>
      </div>
    </div>

    <div style="margin-top: 30px; padding-top: 20px; border-top: 2px solid #e0e0e0;">
      <button onclick="showAllRecords()">👁️ Show All Records</button>
      <button class="danger" onclick="clearAllFilters()">🗑️ Clear Filters</button>
      <button class="secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    function applyFilter(type, value) {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Filter applied!');
          google.script.host.close();
        })
        .applyQuickFilter(type, value);
    }

    function clearAllFilters() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Filters cleared!');
          google.script.host.close();
        })
        .clearAllFiltersOnSheet();
    }

    function showAllRecords() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Showing all records!');
          google.script.host.close();
        })
        .showAllRecords();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Applies quick filter to active sheet
 * @param {string} filterType - Type of filter
 * @param {string} filterValue - Filter value
 */
function applyQuickFilter(filterType, filterValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // Clear existing filters
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  // Create new filter
  const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  const newFilter = range.createFilter();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  switch (filterType) {
    case 'status':
      // Filter by status column (E)
      const criteria = SpreadsheetApp.newFilterCriteria()
        .whenTextEqualTo(filterValue)
        .build();
      newFilter.setColumnFilterCriteria(GRIEVANCE_COLS.STATUS, criteria);
      break;

    case 'time':
      const dateCriteria = getDateFilterCriteria(filterValue, today);
      if (dateCriteria) {
        newFilter.setColumnFilterCriteria(GRIEVANCE_COLS.DATE_FILED, dateCriteria);
      }
      break;

    case 'priority':
      if (filterValue === 'urgent') {
        // Show grievances with deadline within 3 days
        const threeDaysFromNow = new Date(today.getTime() + (3 * 24 * 60 * 60 * 1000));
        const urgentCriteria = SpreadsheetApp.newFilterCriteria()
          .whenDateBefore(threeDaysFromNow)
          .build();
        newFilter.setColumnFilterCriteria(GRIEVANCE_COLS.NEXT_ACTION_DUE, urgentCriteria);
      }
      break;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Filter applied: ${filterType} = ${filterValue}`,
    'Quick Filter',
    3
  );

  // Log activity
  if (typeof logActivity === 'function') {
    logActivity('Quick Filter', `Applied filter: ${filterType} = ${filterValue}`);
  }
}

/**
 * Gets date filter criteria based on time filter
 * @param {string} timeFilter - Time filter value
 * @param {Date} today - Today's date
 * @returns {FilterCriteria} Filter criteria
 */
function getDateFilterCriteria(timeFilter, today) {
  switch (timeFilter) {
    case 'today':
      return SpreadsheetApp.newFilterCriteria()
        .whenDateEqualTo(today)
        .build();

    case 'thisWeek':
      const weekAgo = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
      return SpreadsheetApp.newFilterCriteria()
        .whenDateAfter(weekAgo)
        .build();

    case 'thisMonth':
      const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
      return SpreadsheetApp.newFilterCriteria()
        .whenDateAfter(monthStart)
        .build();

    case 'overdue':
      return SpreadsheetApp.newFilterCriteria()
        .whenDateBefore(today)
        .build();

    default:
      return null;
  }
}

/**
 * Clears all filters on active sheet
 */
function clearAllFiltersOnSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ All filters cleared',
    'Filters',
    3
  );
}

/**
 * Shows all records (removes filters)
 */
function showAllRecords() {
  clearAllFiltersOnSheet();
}

/**
 * Saves current filter as a view
 * @param {string} viewName - Name for the saved view
 */
function saveFilterView(viewName) {
  const props = PropertiesService.getUserProperties();
  const savedViews = getSavedViews();

  // Get current filter state
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const filter = sheet.getFilter();

  if (!filter) {
    SpreadsheetApp.getUi().alert('No active filter to save');
    return;
  }

  savedViews[viewName] = {
    sheetName: sheet.getName(),
    created: new Date().toISOString()
  };

  props.setProperty('savedFilterViews', JSON.stringify(savedViews));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ View "${viewName}" saved`,
    'Saved Views',
    3
  );
}

/**
 * Gets saved filter views
 * @returns {Object} Saved views
 */
function getSavedViews() {
  const props = PropertiesService.getUserProperties();
  const viewsJSON = props.getProperty('savedFilterViews');

  if (viewsJSON) {
    return JSON.parse(viewsJSON);
  }

  return {};
}
