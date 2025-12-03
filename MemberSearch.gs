/**
 * ------------------------------------------------------------------------====
 * MEMBER SEARCH & LOOKUP FUNCTIONALITY
 * ------------------------------------------------------------------------====
 *
 * Fast member search with autocomplete and filtering
 * Features:
 * - Search by name, ID, location, unit
 * - Fuzzy matching
 * - Quick navigation to member row
 * - Show member details
 */

/**
 * Shows member search dialog
 */
function showMemberSearch() {
  const html = HtmlService.createHtmlOutput(createMemberSearchHTML())
    .setWidth(700)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search Members');
}

/**
 * Creates HTML for member search dialog
 */
function createMemberSearchHTML() {
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
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .search-box {
      margin: 20px 0;
    }
    input[type="text"] {
      width: 100%;
      padding: 12px;
      font-size: 16px;
      border: 2px solid #ddd;
      border-radius: 4px;
      box-sizing: border-box;
    }
    input[type="text"]:focus {
      outline: none;
      border-color: #1a73e8;
    }
    .filters {
      display: flex;
      gap: 10px;
      margin: 15px 0;
    }
    select {
      flex: 1;
      padding: 8px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    .results {
      margin-top: 20px;
      max-height: 350px;
      overflow-y: auto;
      border: 1px solid #ddd;
      border-radius: 4px;
    }
    .result-item {
      padding: 15px;
      border-bottom: 1px solid #eee;
      cursor: pointer;
      transition: background 0.2s;
    }
    .result-item:hover {
      background: #f0f7ff;
    }
    .result-item:last-child {
      border-bottom: none;
    }
    .member-name {
      font-size: 16px;
      font-weight: bold;
      color: #1a73e8;
    }
    .member-details {
      font-size: 13px;
      color: #666;
      margin-top: 5px;
    }
    .member-id {
      display: inline-block;
      background: #e8f0fe;
      padding: 2px 8px;
      border-radius: 3px;
      font-size: 12px;
      margin-right: 8px;
    }
    .no-results {
      text-align: center;
      padding: 40px;
      color: #999;
      font-style: italic;
    }
    .loading {
      text-align: center;
      padding: 20px;
      color: #1a73e8;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>🔍 Search Members</h2>

    <div class="search-box">
      <input type="text" id="searchInput" placeholder="Search by name, ID, email..." autofocus>
    </div>

    <div class="filters">
      <select id="locationFilter">
        <option value="">All Locations</option>
      </select>
      <select id="unitFilter">
        <option value="">All Units</option>
      </select>
      <select id="stewardFilter">
        <option value="">All Members</option>
        <option value="Yes">Stewards Only</option>
        <option value="No">Non-Stewards Only</option>
      </select>
    </div>

    <div id="results" class="results">
      <div class="loading">Loading members...</div>
    </div>
  </div>

  <script>
    var allMembers = [];
    var filteredMembers = [];

    // Load members on page load
    window.onload = function() {
      google.script.run
        .withSuccessHandler(onMembersLoaded)
        .withFailureHandler(onError)
        .getAllMembers();
    };

    function onMembersLoaded(members) {
      allMembers = members;
      filteredMembers = members;

      // Populate filter dropdowns
      populateFilters();

      // Display all members initially
      displayResults(members);

      // Set up event listeners
      document.getElementById('searchInput').addEventListener('input', performSearch);
      document.getElementById('locationFilter').addEventListener('change', performSearch);
      document.getElementById('unitFilter').addEventListener('change', performSearch);
      document.getElementById('stewardFilter').addEventListener('change', performSearch);
    }

    function populateFilters() {
      // Get unique locations
      const locations = [...new Set(allMembers.map(function(m) { return m.location).filter(function(l) { return l; }); })];
      const locationFilter = document.getElementById('locationFilter');
      locations.forEach(function(loc) {
        const option = document.createElement('option');
        option.value = loc;
        option.textContent = loc;
        locationFilter.appendChild(option);
      });

      // Get unique units
      const units = [...new Set(allMembers.map(function(m) { return m.unit).filter(function(u) { return u; }); })];
      const unitFilter = document.getElementById('unitFilter');
      units.forEach(function(unit) {
        const option = document.createElement('option');
        option.value = unit;
        option.textContent = unit;
        unitFilter.appendChild(option);
      });
    }

    function performSearch() {
      const searchTerm = document.getElementById('searchInput').value.toLowerCase();
      const locationFilter = document.getElementById('locationFilter').value;
      const unitFilter = document.getElementById('unitFilter').value;
      const stewardFilter = document.getElementById('stewardFilter').value;

      filteredMembers = allMembers.filter(function(member) {
        // Text search
        const matchesSearch = !searchTerm ||
          member.name.toLowerCase().includes(searchTerm) ||
          member.id.toLowerCase().includes(searchTerm) ||
          (member.email && member.email.toLowerCase().includes(searchTerm));

        // Location filter
        const matchesLocation = !locationFilter || member.location === locationFilter;

        // Unit filter
        const matchesUnit = !unitFilter || member.unit === unitFilter;

        // Steward filter
        const matchesSteward = !stewardFilter || member.isSteward === stewardFilter;

        return matchesSearch && matchesLocation && matchesUnit && matchesSteward;
      });

      displayResults(filteredMembers);
    }

    function displayResults(members) {
      const resultsDiv = document.getElementById('results');

      if (members.length === 0) {
        resultsDiv.innerHTML = '<div class="no-results">No members found</div>';
        return;
      }

      resultsDiv.innerHTML = members.slice(0, 50).map(member => \`
        <div class="result-item" onclick="selectMember('\${member.row}', '\${member.id}')">
          <div class="member-name">\${member.name}</div>
          <div class="member-details">
            <span class="member-id">\${member.id}</span>
            <span>\${member.location || 'No location'} • \${member.unit || 'No unit'}</span>
            \${member.isSteward === 'Yes' ? ' • 🛡️ Steward' : ''}
          </div>
          \${member.email ? \`<div class="member-details">📧 \${member.email}</div>\` : ''}
        </div>
      \`).join('');

      if (members.length > 50) {
        resultsDiv.innerHTML += \`<div class="no-results">Showing first 50 of \${members.length} results. Refine your search.</div>\`;
      }
    }

    function selectMember(row, memberId) {
      google.script.run
        .withSuccessHandler(function(() {
          google.script.host.close();
        })
        .navigateToMember(parseInt(row));
    }

    function onError(error) {
      document.getElementById('results').innerHTML =
        '<div class="no-results">❌ Error: ' + error.message + '</div>';
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets all members for search
 * @returns {Array} Array of member objects
 */
function getAllMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return [];

  const lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = memberSheet.getRange(2, 1, lastRow - 1, 13).getValues();

  return data.map(function((row, index) { return ({
    row: index + 2,
    id: row[0] || '',
    name: `${row[1]} ${row[2]}`.trim(),
    firstName: row[1] || '',
    lastName: row[2] || '',
    jobTitle: row[3] || '',
    location: row[4] || '',
    unit: row[5] || '',
    email: row[7] || '',
    phone: row[8] || '',
    isSteward: row[9] || '',
    supervisor: row[10] || '',
    manager: row[11] || '',
    assignedSteward: row[12] || ''
  })).filter(function(m) { return m.id; }); // Filter out empty rows
}

/**
 * Navigates to a member row in the Member Directory
 * @param {number} row - The row number to navigate to
 */
function navigateToMember(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  // Activate the Member Directory sheet
  ss.setActiveSheet(memberSheet);

  // Navigate to the row
  memberSheet.setActiveRange(memberSheet.getRange(row, 1, 1, 10));

  // Show toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Navigated to row ${row}`,
    'Member Found',
    3
  );
}

/**
 * Quick search function for keyboard shortcut
 */
function quickMemberSearch() {
  showMemberSearch();
}
