/**
 * ============================================================================
 * AUTO-REFRESH & REAL-TIME UPDATES
 * ============================================================================
 *
 * Automatic dashboard refresh and real-time data updates
 * Features:
 * - Auto-refresh dashboards
 * - Configurable refresh intervals
 * - Manual refresh controls
 * - Last updated timestamps
 * - Change detection
 * - Refresh notifications
 * - Pause/resume refresh
 */

/**
 * Auto-refresh configuration
 */
const AUTO_REFRESH_CONFIG = {
  DEFAULT_INTERVAL_SECONDS: 300,  // 5 minutes
  INTERVALS: [60, 120, 300, 600, 1800],  // 1m, 2m, 5m, 10m, 30m
  ENABLED_BY_DEFAULT: false
};

/**
 * Shows auto-refresh settings
 */
function showAutoRefreshSettings() {
  const html = createAutoRefreshSettingsHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔄 Auto-Refresh Settings');
}

/**
 * Creates HTML for auto-refresh settings
 */
function createAutoRefreshSettingsHTML() {
  const settings = getAutoRefreshSettings();

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 30px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 30px; border-radius: 12px; max-width: 500px; margin: 0 auto; }
    h2 { color: #1a73e8; margin-top: 0; }
    .setting-row { margin: 25px 0; }
    label { display: block; font-weight: 500; margin-bottom: 10px; color: #333; }
    select { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px; }
    .toggle-container { display: flex; align-items: center; gap: 15px; }
    .toggle-switch { position: relative; display: inline-block; width: 60px; height: 30px; }
    .toggle-switch input { opacity: 0; width: 0; height: 0; }
    .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: 0.4s; border-radius: 30px; }
    .slider:before { position: absolute; content: ""; height: 22px; width: 22px; left: 4px; bottom: 4px; background-color: white; transition: 0.4s; border-radius: 50%; }
    input:checked + .slider { background-color: #1a73e8; }
    input:checked + .slider:before { transform: translateX(30px); }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 6px; cursor: pointer; width: 100%; margin: 5px 0; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    .info-box { background: #e8f0fe; padding: 15px; border-radius: 6px; margin: 20px 0; font-size: 13px; color: #1557b0; }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔄 Auto-Refresh Settings</h2>

    <div class="info-box">
      ℹ️ Auto-refresh will automatically update dashboard data at the specified interval.
      This keeps your data current without manual refreshing.
    </div>

    <div class="setting-row">
      <div class="toggle-container">
        <label style="margin: 0;">Enable Auto-Refresh</label>
        <label class="toggle-switch">
          <input type="checkbox" id="enabled" ${settings.enabled ? 'checked' : ''}>
          <span class="slider"></span>
        </label>
      </div>
    </div>

    <div class="setting-row">
      <label>Refresh Interval</label>
      <select id="interval">
        <option value="60" ${settings.interval === 60 ? 'selected' : ''}>Every 1 minute</option>
        <option value="120" ${settings.interval === 120 ? 'selected' : ''}>Every 2 minutes</option>
        <option value="300" ${settings.interval === 300 ? 'selected' : ''}>Every 5 minutes</option>
        <option value="600" ${settings.interval === 600 ? 'selected' : ''}>Every 10 minutes</option>
        <option value="1800" ${settings.interval === 1800 ? 'selected' : ''}>Every 30 minutes</option>
      </select>
    </div>

    <div class="setting-row">
      <label>Sheets to Auto-Refresh</label>
      <select id="sheets" multiple size="5">
        <option value="Dashboard" ${settings.sheets.includes('Dashboard') ? 'selected' : ''}>Dashboard</option>
        <option value="Analytics Data" ${settings.sheets.includes('Analytics Data') ? 'selected' : ''}>Analytics Data</option>
        <option value="Interactive" ${settings.sheets.includes('Interactive') ? 'selected' : ''}>Interactive Dashboard</option>
        <option value="Steward Workload" ${settings.sheets.includes('Steward Workload') ? 'selected' : ''}>Steward Workload</option>
        <option value="Trends" ${settings.sheets.includes('Trends') ? 'selected' : ''}>Trends & Timeline</option>
      </select>
      <small style="color: #666;">Hold Ctrl/Cmd to select multiple</small>
    </div>

    <button onclick="saveSettings()">💾 Save Settings</button>
    <button class="secondary" onclick="testRefresh()">🧪 Test Refresh Now</button>
    <button class="secondary" onclick="google.script.host.close()">Cancel</button>
  </div>

  <script>
    function saveSettings() {
      const enabled = document.getElementById('enabled').checked;
      const interval = parseInt(document.getElementById('interval').value);
      const sheetsSelect = document.getElementById('sheets');
      const sheets = Array.from(sheetsSelect.selectedOptions).map(o => o.value);

      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Auto-refresh settings saved!');
          google.script.host.close();
        })
        .saveAutoRefreshSettings({
          enabled: enabled,
          interval: interval,
          sheets: sheets
        });
    }

    function testRefresh() {
      google.script.run
        .withSuccessHandler(() => {
          alert('✅ Manual refresh complete!');
        })
        .performManualRefresh();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets auto-refresh settings
 * @returns {Object} Settings
 */
function getAutoRefreshSettings() {
  const props = PropertiesService.getDocumentProperties();
  const settingsJSON = props.getProperty('autoRefreshSettings');

  if (settingsJSON) {
    return JSON.parse(settingsJSON);
  }

  return {
    enabled: AUTO_REFRESH_CONFIG.ENABLED_BY_DEFAULT,
    interval: AUTO_REFRESH_CONFIG.DEFAULT_INTERVAL_SECONDS,
    sheets: ['Dashboard']
  };
}

/**
 * Saves auto-refresh settings
 * @param {Object} settings - Settings to save
 */
function saveAutoRefreshSettings(settings) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('autoRefreshSettings', JSON.stringify(settings));

  // Update triggers
  if (settings.enabled) {
    setupAutoRefreshTrigger(settings.interval);
  } else {
    removeAutoRefreshTrigger();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    settings.enabled ? '✅ Auto-refresh enabled' : '🔕 Auto-refresh disabled',
    'Auto-Refresh',
    3
  );
}

/**
 * Sets up auto-refresh trigger
 * @param {number} intervalSeconds - Refresh interval in seconds
 */
function setupAutoRefreshTrigger(intervalSeconds) {
  // Remove existing triggers
  removeAutoRefreshTrigger();

  // Create new trigger
  const intervalMinutes = Math.max(1, Math.floor(intervalSeconds / 60));

  ScriptApp.newTrigger('performAutoRefresh')
    .timeBased()
    .everyMinutes(intervalMinutes)
    .create();
}

/**
 * Removes auto-refresh trigger
 */
function removeAutoRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'performAutoRefresh') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Performs auto-refresh (triggered automatically)
 */
function performAutoRefresh() {
  const settings = getAutoRefreshSettings();

  if (!settings.enabled) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  settings.sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      refreshSheet(sheet);
    }
  });

  // Update last refresh timestamp
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('lastAutoRefresh', new Date().toISOString());

  // Log activity
  if (typeof logActivity === 'function') {
    logActivity('Auto-Refresh', 'Dashboards automatically refreshed', { sheets: settings.sheets });
  }
}

/**
 * Performs manual refresh
 */
function performManualRefresh() {
  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  refreshSheet(activeSheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ ' + activeSheet.getName() + ' refreshed',
    'Refresh Complete',
    3
  );

  // Log activity
  if (typeof logActivity === 'function') {
    logActivity('Manual Refresh', 'Sheet manually refreshed: ' + activeSheet.getName());
  }
}

/**
 * Refreshes a specific sheet
 * @param {Sheet} sheet - Sheet to refresh
 */
function refreshSheet(sheet) {
  const sheetName = sheet.getName();

  // Add timestamp
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow > 0 && lastCol > 0) {
    // Update "Last Updated" cell if it exists
    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < Math.min(5, values.length); i++) {
      for (let j = 0; j < values[i].length; j++) {
        if (String(values[i][j]).toLowerCase().includes('last updated')) {
          sheet.getRange(i + 1, j + 2).setValue(new Date());
          break;
        }
      }
    }
  }

  SpreadsheetApp.flush();
}

/**
 * Gets last refresh timestamp
 * @returns {string} Timestamp
 */
function getLastRefreshTimestamp() {
  const props = PropertiesService.getDocumentProperties();
  const timestamp = props.getProperty('lastAutoRefresh');

  if (timestamp) {
    return new Date(timestamp).toLocaleString();
  }

  return 'Never';
}

/**
 * Shows refresh status
 */
function showRefreshStatus() {
  const settings = getAutoRefreshSettings();
  const lastRefresh = getLastRefreshTimestamp();

  const message = `
AUTO-REFRESH STATUS

Status: ${settings.enabled ? '✅ Enabled' : '🔕 Disabled'}
Interval: ${settings.interval / 60} minutes
Sheets: ${settings.sheets.join(', ')}
Last Refresh: ${lastRefresh}

${settings.enabled ? 'Dashboards will automatically refresh every ' + (settings.interval / 60) + ' minutes.' : 'Enable auto-refresh in settings to keep dashboards current.'}
  `;

  SpreadsheetApp.getUi().alert('Refresh Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Quick refresh all dashboards
 */
function quickRefreshAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dashboardSheets = [
    SHEETS.DASHBOARD,
    SHEETS.ANALYTICS,
    SHEETS.STEWARD_WORKLOAD,
    SHEETS.TRENDS,
    SHEETS.INTERACTIVE_DASHBOARD
  ];

  let refreshed = 0;

  dashboardSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      refreshSheet(sheet);
      refreshed++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ ${refreshed} dashboard${refreshed !== 1 ? 's' : ''} refreshed`,
    'Refresh Complete',
    3
  );
}
