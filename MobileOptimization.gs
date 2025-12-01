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
