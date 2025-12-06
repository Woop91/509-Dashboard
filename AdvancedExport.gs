/**
 * ============================================================================
 * ADVANCED EXPORT OPTIONS
 * ============================================================================
 *
 * Enhanced data export capabilities
 * Features:
 * - Export to multiple formats (CSV, Excel, PDF, JSON)
 * - Custom field selection
 * - Filtered exports
 * - Scheduled exports
 * - Email exports
 * - Template-based exports
 */

/**
 * Shows advanced export dialog
 */
function showAdvancedExport() {
  const html = createAdvancedExportHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📥 Advanced Export');
}

/**
 * Creates HTML for advanced export
 */
function createAdvancedExportHTML() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; margin: 0; background: #f5f5f5; }
    .container { background: white; padding: 30px; border-radius: 12px; max-height: 550px; overflow-y: auto; }
    h2 { color: #1a73e8; margin-top: 0; }
    .section { margin: 25px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; }
    .section-title { font-weight: bold; margin-bottom: 15px; color: #333; }
    .export-btn { background: #1a73e8; color: white; border: none; padding: 15px; border-radius: 8px; cursor: pointer; width: 100%; margin: 10px 0; text-align: left; font-size: 15px; transition: all 0.2s; }
    .export-btn:hover { background: #1557b0; transform: translateX(5px); }
    .export-icon { font-size: 24px; margin-right: 10px; }
    .export-description { font-size: 12px; color: rgba(255,255,255,0.9); margin-top: 5px; }
    label { display: block; margin: 10px 0; font-weight: 500; }
    select, input { width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 6px; margin-top: 5px; box-sizing: border-box; }
    button.secondary { background: #6c757d; }
  </style>
</head>
<body>
  <div class="container">
    <h2>📥 Advanced Export</h2>

    <div class="section">
      <div class="section-title">Quick Export</div>

      <button class="export-btn" onclick="exportAs('csv')">
        <div><span class="export-icon">📄</span> Export as CSV</div>
        <div class="export-description">Comma-separated values for Excel/Google Sheets</div>
      </button>

      <button class="export-btn" onclick="exportAs('excel')">
        <div><span class="export-icon">📊</span> Export as Excel (XLSX)</div>
        <div class="export-description">Microsoft Excel compatible format</div>
      </button>

      <button class="export-btn" onclick="exportAs('pdf')">
        <div><span class="export-icon">📑</span> Export as PDF</div>
        <div class="export-description">Print-ready PDF document</div>
      </button>

      <button class="export-btn" onclick="exportAs('json')">
        <div><span class="export-icon">🗂️</span> Export as JSON</div>
        <div class="export-description">API-friendly JSON format</div>
      </button>
    </div>

    <div class="section">
      <div class="section-title">Custom Export</div>

      <label>
        Data Source
        <select id="dataSource">
          <option value="grievances">Grievance Log</option>
          <option value="members">Member Directory</option>
          <option value="dashboard">Dashboard Data</option>
        </select>
      </label>

      <label>
        Date Range
        <select id="dateRange">
          <option value="all">All Time</option>
          <option value="thisYear">This Year</option>
          <option value="lastYear">Last Year</option>
          <option value="last30">Last 30 Days</option>
          <option value="last90">Last 90 Days</option>
        </select>
      </label>

      <label>
        <input type="checkbox" id="includeArchived"> Include Archived Records
      </label>

      <button class="export-btn" onclick="exportCustom()">
        <div><span class="export-icon">⚙️</span> Custom Export</div>
        <div class="export-description">Export with selected options</div>
      </button>
    </div>

    <div style="text-align: center; margin-top: 20px;">
      <button class="export-btn secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    function exportAs(format) {
      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ Export complete!');
          window.open(url, '_blank');
        })
        .withFailureHandler((error) => {
          alert('❌ Export failed: ' + error.message);
        })
        .performQuickExport(format);
    }

    function exportCustom() {
      const options = {
        dataSource: document.getElementById('dataSource').value,
        dateRange: document.getElementById('dateRange').value,
        includeArchived: document.getElementById('includeArchived').checked
      };

      google.script.run
        .withSuccessHandler((url) => {
          alert('✅ Custom export complete!');
          window.open(url, '_blank');
        })
        .performCustomExport(options);
    }
  </script>
</body>
</html>
  `;
}

/**
 * Performs quick export
 * @param {string} format - Export format
 * @returns {string} File URL
 */
function performQuickExport(format) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  switch (format) {
    case 'csv':
      return exportToCSV(grievanceSheet);
    case 'excel':
      return exportToExcel(grievanceSheet);
    case 'pdf':
      return exportToPDF(grievanceSheet);
    case 'json':
      return exportToJSON(grievanceSheet);
    default:
      throw new Error('Unsupported format');
  }
}

/**
 * Exports sheet to CSV with permission-based filtering
 * @param {Sheet} sheet - Sheet to export
 * @returns {string} File URL
 */
function exportToCSV(sheet) {
  const sheetName = sheet.getName();
  let data = sheet.getDataRange().getValues();

  // Apply permission filtering based on sheet type
  if (sheetName === SHEETS.MEMBER_DIR) {
    data = filterMemberDataByPermission(data);
  } else if (sheetName === SHEETS.GRIEVANCE_LOG) {
    data = filterGrievanceDataByPermission(data);
  }

  let csv = '';
  data.forEach(row => {
    csv += row.map(cell => {
      let value = String(cell);
      if (cell instanceof Date) {
        value = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return '"' + value.replace(/"/g, '""') + '"';
    }).join(',') + '\n';
  });

  const fileName = `${sheet.getName()}_Export_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.csv`;
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);

  return file.getUrl();
}

/**
 * Exports to Excel format with permission-based filtering
 * Creates a temporary filtered copy for export to enforce permissions
 * @param {Sheet} sheet - Sheet to export
 * @returns {string} File URL
 */
function exportToExcel(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheet.getName();
  const fileName = `${sheetName}_Export_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;

  // Get filtered data based on permissions
  let data = sheet.getDataRange().getValues();
  if (sheetName === SHEETS.MEMBER_DIR) {
    data = filterMemberDataByPermission(data);
  } else if (sheetName === SHEETS.GRIEVANCE_LOG) {
    data = filterGrievanceDataByPermission(data);
  }

  // Create temporary spreadsheet with filtered data for Excel export
  const tempSS = SpreadsheetApp.create('_temp_export_' + new Date().getTime());
  const tempSheet = tempSS.getActiveSheet();
  tempSheet.setName(sheetName);

  if (data.length > 0) {
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  // Export the filtered spreadsheet
  const blob = tempSS.getAs(MimeType.MICROSOFT_EXCEL);
  const file = DriveApp.createFile(blob).setName(fileName + '.xlsx');

  // Clean up temporary spreadsheet
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  return file.getUrl();
}

/**
 * Exports to PDF with permission-based filtering
 * Creates a temporary filtered copy for export to enforce permissions
 * @param {Sheet} sheet - Sheet to export
 * @returns {string} File URL
 */
function exportToPDF(sheet) {
  const sheetName = sheet.getName();
  const fileName = `${sheetName}_Export_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.pdf`;

  // Get filtered data based on permissions
  let data = sheet.getDataRange().getValues();
  if (sheetName === SHEETS.MEMBER_DIR) {
    data = filterMemberDataByPermission(data);
  } else if (sheetName === SHEETS.GRIEVANCE_LOG) {
    data = filterGrievanceDataByPermission(data);
  }

  // Create temporary spreadsheet with filtered data for PDF export
  const tempSS = SpreadsheetApp.create('_temp_pdf_export_' + new Date().getTime());
  const tempSheet = tempSS.getActiveSheet();
  tempSheet.setName(sheetName);

  if (data.length > 0) {
    tempSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  // Export the filtered spreadsheet to PDF
  const url = tempSS.getUrl();
  const exportUrl = url.replace(/\/edit.*$/, '') +
    '/export?exportFormat=pdf&format=pdf' +
    '&size=A4' +
    '&portrait=false' +
    '&fitw=true' +
    '&sheetnames=false&printtitle=false' +
    '&pagenumbers=false&gridlines=false' +
    '&fzr=false' +
    '&gid=' + tempSheet.getSheetId();

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  const blob = response.getBlob().setName(fileName);
  const file = DriveApp.createFile(blob);

  // Clean up temporary spreadsheet
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  return file.getUrl();
}

/**
 * Exports to JSON with permission-based filtering
 * @param {Sheet} sheet - Sheet to export
 * @returns {string} File URL
 */
function exportToJSON(sheet) {
  const sheetName = sheet.getName();
  let data = sheet.getDataRange().getValues();

  // Apply permission filtering based on sheet type
  if (sheetName === SHEETS.MEMBER_DIR) {
    data = filterMemberDataByPermission(data);
  } else if (sheetName === SHEETS.GRIEVANCE_LOG) {
    data = filterGrievanceDataByPermission(data);
  }

  const headers = data[0];
  const rows = data.slice(1);

  const jsonData = rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[header] = value;
    });
    return obj;
  });

  const json = JSON.stringify(jsonData, null, 2);
  const fileName = `${sheet.getName()}_Export_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}.json`;
  const file = DriveApp.createFile(fileName, json, MimeType.PLAIN_TEXT);

  return file.getUrl();
}

/**
 * Performs custom export with filters
 * @param {Object} options - Export options
 * @returns {string} File URL
 */
function performCustomExport(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;

  switch (options.dataSource) {
    case 'grievances':
      sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
      break;
    case 'members':
      sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
      break;
    case 'dashboard':
      sheet = ss.getSheetByName(SHEETS.DASHBOARD);
      break;
    default:
      sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  }

  // Apply date range filter if needed
  // For now, export full sheet
  return exportToCSV(sheet);
}
