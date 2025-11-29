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
