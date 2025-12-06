/**
 * ============================================================================
 * RELEASE NOTES & VERSION TRACKING
 * ============================================================================
 *
 * Version management and release notes system
 * Features:
 * - Version tracking
 * - Release notes viewer
 * - Change log management
 * - Feature announcements
 * - What's New notifications
 * - Version comparison
 * - Rollback support
 * - Update notifications
 */

/**
 * Version configuration
 */
const VERSION_CONFIG = {
  CURRENT_VERSION: '2.5.0',
  RELEASE_DATE: '2025-01-15',
  VERSION_NAME: 'Enhanced Edition',
  SHOW_WHATS_NEW_ON_LOAD: false
};

/**
 * Release history with all versions
 */
const RELEASE_HISTORY = [
  {
    version: '2.5.0',
    name: 'Enhanced Edition',
    date: '2025-01-15',
    type: 'Major',
    highlights: [
      'Added Session Management & User Preferences',
      'Advanced Data Visualization with interactive charts',
      'Release Notes & Version Tracking system',
      'Context-sensitive help enhancements',
      'Notification Center for system alerts',
      'Real-time dashboard updates'
    ],
    features: [
      'Session tracking with activity logging',
      'User preference management',
      'Interactive chart builder with 8 chart types',
      'Google Charts integration',
      'Release notes viewer',
      'Version comparison tools',
      'Enhanced help system with search',
      'Notification center with priorities',
      'Auto-refresh dashboards'
    ],
    bugFixes: [],
    breaking: []
  },
  {
    version: '2.0.0',
    name: 'Accessibility Update',
    date: '2025-01-10',
    type: 'Major',
    highlights: [
      'Enhanced ADHD Features with control panel',
      'Undo/Redo System with 50-action history',
      'Dark Mode & Theme Customization',
      'Data Pagination for large datasets',
      'Mobile-Optimized responsive views',
      'Enhanced Error Handling & Recovery'
    ],
    features: [
      'ADHD Control Panel with Pomodoro timer',
      'Quick capture notepad',
      'Break reminders',
      '5 theme presets (Light/Dark/OLED/Blue/Sepia)',
      'Auto dark mode (time-based)',
      'Undo/Redo with keyboard shortcuts',
      'Paginated viewer (25-500 items)',
      'Mobile dashboard with touch support',
      'Error Dashboard with health monitoring',
      'Comprehensive error logging'
    ],
    bugFixes: [
      'Fixed theme persistence across sessions',
      'Resolved mobile touch event conflicts',
      'Corrected undo/redo state management'
    ],
    breaking: []
  },
  {
    version: '1.5.0',
    name: 'Advanced Features',
    date: '2025-01-05',
    type: 'Major',
    highlights: [
      'Custom Report Builder',
      'FAQ Knowledge Base',
      'Root Cause Analysis',
      'Data Backup & Recovery',
      'Predictive Analytics'
    ],
    features: [
      'Interactive custom report builder',
      'Searchable FAQ system with 12 articles',
      'Pattern detection for systemic issues',
      'Automated weekly backups',
      'Trend forecasting',
      'Anomaly detection'
    ],
    bugFixes: [],
    breaking: []
  },
  {
    version: '1.0.0',
    name: 'Foundation Release',
    date: '2024-12-20',
    type: 'Major',
    highlights: [
      'Core grievance management system',
      'Member directory with 20k capacity',
      'Interactive dashboards',
      'Batch operations',
      'Email integration'
    ],
    features: [
      'Grievance Log with auto-calculations',
      'Member Directory (20k members)',
      'Dashboard with real-time metrics',
      'Batch assign, update, export',
      'Gmail integration with templates',
      'Google Drive file management',
      'Calendar deadline sync'
    ],
    bugFixes: [],
    breaking: []
  }
];

/**
 * Shows release notes viewer
 */
function showReleaseNotes() {
  const html = createReleaseNotesHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📋 Release Notes');

  // Mark as viewed
  markReleaseNotesViewed();
}

/**
 * Creates HTML for release notes
 */
function createReleaseNotesHTML() {
  let versionSections = '';

  RELEASE_HISTORY.forEach((release, index) => {
    const isLatest = index === 0;

    versionSections += `
      <div class="version-card ${isLatest ? 'latest' : ''}">
        <div class="version-header">
          <div>
            <div class="version-number">
              ${release.version}
              ${isLatest ? '<span class="badge-latest">LATEST</span>' : ''}
            </div>
            <div class="version-name">${release.name}</div>
          </div>
          <div class="version-meta">
            <div class="version-date">${release.date}</div>
            <div class="version-type type-${release.type.toLowerCase()}">${release.type} Release</div>
          </div>
        </div>

        <div class="version-body">
          ${release.highlights.length > 0 ? `
            <h4>✨ Highlights</h4>
            <ul class="highlights">
              ${release.highlights.map(h => `<li>${h}</li>`).join('')}
            </ul>
          ` : ''}

          ${release.features.length > 0 ? `
            <h4>🎁 New Features</h4>
            <ul>
              ${release.features.map(f => `<li>${f}</li>`).join('')}
            </ul>
          ` : ''}

          ${release.bugFixes && release.bugFixes.length > 0 ? `
            <h4>🐛 Bug Fixes</h4>
            <ul>
              ${release.bugFixes.map(b => `<li>${b}</li>`).join('')}
            </ul>
          ` : ''}

          ${release.breaking && release.breaking.length > 0 ? `
            <h4>⚠️ Breaking Changes</h4>
            <ul class="breaking">
              ${release.breaking.map(b => `<li>${b}</li>`).join('')}
            </ul>
          ` : ''}
        </div>
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
      max-width: 850px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    .header {
      text-align: center;
      margin-bottom: 40px;
      padding-bottom: 20px;
      border-bottom: 3px solid #1a73e8;
    }
    .title {
      font-size: 36px;
      font-weight: bold;
      color: #1a73e8;
      margin: 0;
    }
    .subtitle {
      font-size: 18px;
      color: #666;
      margin-top: 10px;
    }
    .current-version {
      background: linear-gradient(135deg, #1a73e8 0%, #1557b0 100%);
      color: white;
      padding: 20px;
      border-radius: 12px;
      margin-bottom: 30px;
      text-align: center;
    }
    .current-version-number {
      font-size: 48px;
      font-weight: bold;
      margin-bottom: 10px;
    }
    .current-version-name {
      font-size: 20px;
      opacity: 0.9;
    }
    .version-card {
      background: #fafafa;
      border: 2px solid #e0e0e0;
      border-radius: 12px;
      padding: 25px;
      margin-bottom: 25px;
      transition: all 0.2s;
    }
    .version-card.latest {
      border-color: #1a73e8;
      background: #f8fcff;
    }
    .version-card:hover {
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .version-header {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 20px;
      padding-bottom: 15px;
      border-bottom: 2px solid #e0e0e0;
    }
    .version-number {
      font-size: 28px;
      font-weight: bold;
      color: #1a73e8;
    }
    .version-name {
      font-size: 16px;
      color: #666;
      margin-top: 5px;
    }
    .version-meta {
      text-align: right;
    }
    .version-date {
      font-size: 14px;
      color: #888;
      margin-bottom: 5px;
    }
    .version-type {
      display: inline-block;
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .type-major { background: #ffebee; color: #c62828; }
    .type-minor { background: #fff3e0; color: #ef6c00; }
    .type-patch { background: #e8f5e9; color: #2e7d32; }
    .badge-latest {
      display: inline-block;
      background: #4caf50;
      color: white;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
      margin-left: 10px;
    }
    .version-body h4 {
      color: #333;
      font-size: 16px;
      margin-top: 20px;
      margin-bottom: 10px;
    }
    .version-body ul {
      margin: 10px 0;
      padding-left: 25px;
    }
    .version-body li {
      margin: 8px 0;
      color: #555;
      line-height: 1.6;
    }
    .highlights li {
      font-weight: 500;
      color: #333;
    }
    .breaking li {
      color: #d32f2f;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 12px 24px;
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
    .actions {
      text-align: center;
      margin-top: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="title">📋 Release Notes</div>
      <div class="subtitle">509 Dashboard Version History</div>
    </div>

    <div class="current-version">
      <div class="current-version-number">v${VERSION_CONFIG.CURRENT_VERSION}</div>
      <div class="current-version-name">${VERSION_CONFIG.VERSION_NAME}</div>
      <div style="margin-top: 10px; font-size: 14px;">Released ${VERSION_CONFIG.RELEASE_DATE}</div>
    </div>

    ${versionSections}

    <div class="actions">
      <button onclick="exportChangelog()">📥 Export Changelog</button>
      <button onclick="viewVersionHistory()">📊 Version History</button>
      <button class="secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    function exportChangelog() {
      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ Changelog exported!');
          window.open(url, '_blank');
        })
        .exportChangelogToSheet();
    }

    function viewVersionHistory() {
      google.script.run.showVersionHistory();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Shows what's new dialog
 */
function showWhatsNew() {
  const latestRelease = RELEASE_HISTORY[0];

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 30px; margin: 0; background: linear-gradient(135deg, #1a73e8 0%, #1557b0 100%); color: white; text-align: center; }
    .container { max-width: 600px; margin: 0 auto; }
    h1 { font-size: 48px; margin: 0 0 10px 0; }
    .version { font-size: 24px; opacity: 0.9; margin-bottom: 30px; }
    .highlights { background: white; color: #333; padding: 30px; border-radius: 12px; text-align: left; margin: 20px 0; }
    .highlights h2 { color: #1a73e8; margin-top: 0; }
    .highlights ul { padding-left: 20px; }
    .highlights li { margin: 12px 0; font-size: 16px; line-height: 1.6; }
    button { background: white; color: #1a73e8; border: none; padding: 15px 30px; font-size: 16px; border-radius: 8px; cursor: pointer; margin: 10px; font-weight: bold; }
    button:hover { background: #f0f0f0; }
  </style>
</head>
<body>
  <div class="container">
    <h1>🎉 What's New</h1>
    <div class="version">Version ${latestRelease.version} - ${latestRelease.name}</div>

    <div class="highlights">
      <h2>✨ Highlights</h2>
      <ul>
        ${latestRelease.highlights.map(h => `<li>${h}</li>`).join('')}
      </ul>

      ${latestRelease.features.length > 0 ? `
        <h2>🎁 New Features</h2>
        <ul>
          ${latestRelease.features.slice(0, 5).map(f => `<li>${f}</li>`).join('')}
          ${latestRelease.features.length > 5 ? `<li><em>...and ${latestRelease.features.length - 5} more!</em></li>` : ''}
        </ul>
      ` : ''}
    </div>

    <button onclick="viewFullNotes()">📋 View Full Release Notes</button>
    <button onclick="google.script.host.close()">Got it!</button>
  </div>

  <script>
    function viewFullNotes() {
      google.script.run.showReleaseNotes();
      google.script.host.close();
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "What's New");

  markReleaseNotesViewed();
}

/**
 * Marks release notes as viewed
 */
function markReleaseNotesViewed() {
  const props = PropertiesService.getUserProperties();
  props.setProperty('lastViewedVersion', VERSION_CONFIG.CURRENT_VERSION);
  props.setProperty('lastViewedDate', new Date().toISOString());
}

/**
 * Checks if user has seen latest release notes
 * @returns {boolean}
 */
function hasSeenLatestRelease() {
  const props = PropertiesService.getUserProperties();
  const lastViewed = props.getProperty('lastViewedVersion');
  return lastViewed === VERSION_CONFIG.CURRENT_VERSION;
}

/**
 * Shows version history
 */
function showVersionHistory() {
  let message = '📊 VERSION HISTORY\n\n';

  RELEASE_HISTORY.forEach(release => {
    message += `v${release.version} - ${release.name}\n`;
    message += `  ${release.date} (${release.type} Release)\n`;
    message += `  Features: ${release.features.length}\n\n`;
  });

  message += `\nCurrent Version: v${VERSION_CONFIG.CURRENT_VERSION}`;

  SpreadsheetApp.getUi().alert('Version History', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Exports changelog to sheet
 * @returns {string} Sheet URL
 */
function exportChangelogToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let changelogSheet = ss.getSheetByName('Changelog_Export');
  if (changelogSheet) {
    changelogSheet.clear();
  } else {
    changelogSheet = ss.insertSheet('Changelog_Export');
  }

  // Headers
  const headers = ['Version', 'Name', 'Date', 'Type', 'Category', 'Description'];
  changelogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  changelogSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  let row = 2;

  RELEASE_HISTORY.forEach(release => {
    // Highlights
    release.highlights.forEach(highlight => {
      changelogSheet.getRange(row, 1, 1, headers.length).setValues([[
        release.version,
        release.name,
        release.date,
        release.type,
        'Highlight',
        highlight
      ]]);
      row++;
    });

    // Features
    release.features.forEach(feature => {
      changelogSheet.getRange(row, 1, 1, headers.length).setValues([[
        release.version,
        release.name,
        release.date,
        release.type,
        'Feature',
        feature
      ]]);
      row++;
    });

    // Bug Fixes
    if (release.bugFixes) {
      release.bugFixes.forEach(fix => {
        changelogSheet.getRange(row, 1, 1, headers.length).setValues([[
          release.version,
          release.name,
          release.date,
          release.type,
          'Bug Fix',
          fix
        ]]);
        row++;
      });
    }
  });

  // Auto-resize
  for (let col = 1; col <= headers.length; col++) {
    changelogSheet.autoResizeColumn(col);
  }

  return ss.getUrl() + '#gid=' + changelogSheet.getSheetId();
}

/**
 * Gets current version info
 * @returns {Object} Version information
 */
function getCurrentVersionInfo() {
  return {
    version: VERSION_CONFIG.CURRENT_VERSION,
    name: VERSION_CONFIG.VERSION_NAME,
    releaseDate: VERSION_CONFIG.RELEASE_DATE,
    hasSeenRelease: hasSeenLatestRelease()
  };
}

/**
 * Checks for updates (placeholder for future API integration)
 * @returns {Object} Update information
 */
function checkForUpdates() {
  // Placeholder - in a real system, this would check an API
  return {
    updateAvailable: false,
    latestVersion: VERSION_CONFIG.CURRENT_VERSION,
    currentVersion: VERSION_CONFIG.CURRENT_VERSION,
    message: 'You are using the latest version!'
  };
}
