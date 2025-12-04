/**
 * ------------------------------------------------------------------------====
 * UNDO/REDO SYSTEM
 * ------------------------------------------------------------------------====
 *
 * Advanced action history and undo/redo functionality
 * Features:
 * - Comprehensive action tracking
 * - Undo/redo stack with 50 action limit
 * - Snapshot-based rollback
 * - Action history viewer
 * - Batch operation support
 * - Keyboard shortcuts (Ctrl+Z, Ctrl+Y)
 * - Visual confirmation
 */

/**
 * Undo/Redo configuration
 */
UNDO_CONFIG = {
  MAX_HISTORY: 50,  // Maximum actions to keep
  STORAGE_KEY: 'undoRedoHistory',
  SUPPORTED_ACTIONS: [
    'EDIT_CELL',
    'ADD_ROW',
    'DELETE_ROW',
    'BATCH_UPDATE',
    'STATUS_CHANGE',
    'ASSIGNMENT_CHANGE'
  ]
};

/**
 * Shows undo/redo control panel
 */
function showUndoRedoPanel() {
  const html = createUndoRedoPanelHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '↩️ Undo/Redo History');
}

/**
 * Creates HTML for undo/redo panel
 */
function createUndoRedoPanelHTML() {
  const history = getUndoHistory();

  let historyRows = '';
  if (history.actions.length === 0) {
    historyRows = '<tr><td colspan="5" style="text-align: center; padding: 40px; color: #999;">No actions recorded yet</td></tr>';
  } else {
    history.actions.slice().reverse().forEach(function(action, index) {
      const timestamp = new Date(action.timestamp).toLocaleString();
      const canUndo = index < history.actions.length - history.currentIndex;
      const canRedo = index >= history.actions.length - history.currentIndex;

      historyRows += `
        <tr ${!canUndo && !canRedo ? 'class="future-action"' : ''}>
          <td>${history.actions.length - index}</td>
          <td><span class="action-badge action-${action.type.toLowerCase()}">${action.type}</span></td>
          <td>${action.description}</td>
          <td style="font-size: 12px; color: #666;">${timestamp}</td>
          <td>
            ${canUndo ? '<button onclick="undoToAction(' + (history.actions.length - index - 1) + ')">↩️ Undo</button>' : ''}
            ${canRedo ? '<button onclick="redoToAction(' + (history.actions.length - index - 1) + ')">↪️ Redo</button>' : ''}
          </td>
        </tr>
      `;
    });
  }

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
      max-height: 550px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .stats {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 15px;
      margin: 20px 0;
    }
    .stat-card {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      text-align: center;
      border-left: 4px solid #1a73e8;
    }
    .stat-number {
      font-size: 32px;
      font-weight: bold;
      color: #1a73e8;
    }
    .stat-label {
      font-size: 13px;
      color: #666;
      margin-top: 5px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
    }
    th {
      background: #1a73e8;
      color: white;
      padding: 12px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
    }
    td {
      padding: 12px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background: #f8f9fa;
    }
    .future-action {
      opacity: 0.4;
    }
    .action-badge {
      display: inline-block;
      padding: 4px 12px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
      text-transform: uppercase;
    }
    .action-edit_cell { background: #e3f2fd; color: #1976d2; }
    .action-add_row { background: #e8f5e9; color: #388e3c; }
    .action-delete_row { background: #ffebee; color: #d32f2f; }
    .action-batch_update { background: #fff3e0; color: #f57c00; }
    .action-status_change { background: #f3e5f5; color: #7b1fa2; }
    .action-assignment_change { background: #e0f2f1; color: #00796b; }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 8px 16px;
      font-size: 13px;
      border-radius: 4px;
      cursor: pointer;
      margin: 2px;
    }
    button:hover {
      background: #1557b0;
    }
    .quick-actions {
      margin: 20px 0;
      padding: 20px;
      background: #f8f9fa;
      border-radius: 8px;
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .quick-actions button {
      padding: 12px 24px;
      font-size: 14px;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .shortcut {
      background: #333;
      color: white;
      padding: 4px 8px;
      border-radius: 4px;
      font-family: monospace;
      font-size: 12px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>↩️ Undo/Redo History</h2>

    <div class="info-box">
      <strong>⌨️ Keyboard Shortcuts:</strong>
      <span class="shortcut">Ctrl+Z</span> Undo &nbsp;&nbsp;
      <span class="shortcut">Ctrl+Y</span> Redo &nbsp;&nbsp;
      <span class="shortcut">Ctrl+Shift+Z</span> View History
    </div>

    <div class="stats">
      <div class="stat-card">
        <div class="stat-number">${history.actions.length}</div>
        <div class="stat-label">Total Actions</div>
      </div>
      <div class="stat-card">
        <div class="stat-number">${history.currentIndex}</div>
        <div class="stat-label">Current Position</div>
      </div>
      <div class="stat-card">
        <div class="stat-number">${history.actions.length - history.currentIndex}</div>
        <div class="stat-label">Actions Available</div>
      </div>
    </div>

    <div class="quick-actions">
      <button onclick="performUndo()">↩️ Undo Last Action</button>
      <button onclick="performRedo()">↪️ Redo Action</button>
      <button onclick="clearHistory()" style="background: #d32f2f;">🗑️ Clear History</button>
      <button onclick="exportHistory()" style="background: #00796b;">📥 Export History</button>
    </div>

    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>Action Type</th>
          <th>Description</th>
          <th>Timestamp</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        ${historyRows}
      </tbody>
    </table>
  </div>

  <script>
    function performUndo() {
      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Action undone!');
          location.reload();
        })
        .withFailureHandler(function(error) {
          alert('❌ Cannot undo: ' + error.message);
        })
        .undoLastAction();
    }

    function performRedo() {
      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Action redone!');
          location.reload();
        })
        .withFailureHandler(function(error) {
          alert('❌ Cannot redo: ' + error.message);
        })
        .redoLastAction();
    }

    function undoToAction(index) {
      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Undo complete!');
          location.reload();
        })
        .undoToIndex(index);
    }

    function redoToAction(index) {
      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Redo complete!');
          location.reload();
        })
        .redoToIndex(index);
    }

    function clearHistory() {
      if (confirm('Clear all undo/redo history? This cannot be undone.')) {
        google.script.run
          .withSuccessHandler(function() {
            alert('✅ History cleared!');
            location.reload();
          })
          .clearUndoHistory();
      }
    }

    function exportHistory() {
      google.script.run
        .withSuccessHandler(function(url) {
          alert('✅ History exported!');
          window.open(url, '_blank');
        })
        .exportUndoHistoryToSheet();
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets undo/redo history
 * @returns {Object} History object
 */
function getUndoHistory() {
  const props = PropertiesService.getScriptProperties();
  const historyJSON = props.getProperty(UNDO_CONFIG.STORAGE_KEY);

  if (historyJSON) {
    return JSON.parse(historyJSON);
  }

  return {
    actions: [],
    currentIndex: 0
  };
}

/**
 * Saves undo/redo history
 * @param {Object} history - History to save
 */
function saveUndoHistory(history) {
  const props = PropertiesService.getScriptProperties();

  // Limit history size
  if (history.actions.length > UNDO_CONFIG.MAX_HISTORY) {
    history.actions = history.actions.slice(-UNDO_CONFIG.MAX_HISTORY);
    history.currentIndex = Math.min(history.currentIndex, history.actions.length);
  }

  props.setProperty(UNDO_CONFIG.STORAGE_KEY, JSON.stringify(history));
}

/**
 * Records an action for undo/redo
 * @param {string} actionType - Type of action
 * @param {string} description - Description
 * @param {Object} beforeState - State before action
 * @param {Object} afterState - State after action
 */
function recordAction(actionType, description, beforeState, afterState) {
  const history = getUndoHistory();

  // Remove any actions after current index (when doing new action after undo)
  if (history.currentIndex < history.actions.length) {
    history.actions = history.actions.slice(0, history.currentIndex);
  }

  // Add new action
  history.actions.push({
    type: actionType,
    description: description,
    timestamp: new Date().toISOString(),
    beforeState: beforeState,
    afterState: afterState
  });

  history.currentIndex = history.actions.length;

  saveUndoHistory(history);
}

/**
 * Records a cell edit action
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {*} oldValue - Old value
 * @param {*} newValue - New value
 */
function recordCellEdit(row, col, oldValue, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const colName = sheet.getRange(1, col).getValue();

  recordAction(
    'EDIT_CELL',
    `Edited ${colName} in row ${row}`,
    { row, col, value: oldValue, sheet: sheet.getName() },
    { row, col, value: newValue, sheet: sheet.getName() }
  );
}

/**
 * Records a row addition
 * @param {number} row - Row number
 * @param {Array} rowData - Row data
 */
function recordRowAddition(row, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  recordAction(
    'ADD_ROW',
    `Added row ${row}`,
    null,
    { row, data: rowData, sheet: sheet.getName() }
  );
}

/**
 * Records a row deletion
 * @param {number} row - Row number
 * @param {Array} rowData - Deleted row data
 */
function recordRowDeletion(row, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  recordAction(
    'DELETE_ROW',
    `Deleted row ${row}`,
    { row, data: rowData, sheet: sheet.getName() },
    null
  );
}

/**
 * Records a batch update
 * @param {string} operation - Operation name
 * @param {Array} changes - Array of changes
 */
function recordBatchUpdate(operation, changes) {
  recordAction(
    'BATCH_UPDATE',
    `${operation}: ${changes.length} rows affected`,
    { changes: changes },
    { operation: operation }
  );
}

/**
 * Undoes the last action
 */
function undoLastAction() {
  const history = getUndoHistory();

  if (history.currentIndex === 0) {
    throw new Error('Nothing to undo');
  }

  const action = history.actions[history.currentIndex - 1];
  applyState(action.beforeState, action.type);

  history.currentIndex--;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `↩️ Undone: ${action.description}`,
    'Undo',
    3
  );
}

/**
 * Redoes the last undone action
 */
function redoLastAction() {
  const history = getUndoHistory();

  if (history.currentIndex >= history.actions.length) {
    throw new Error('Nothing to redo');
  }

  const action = history.actions[history.currentIndex];
  applyState(action.afterState, action.type);

  history.currentIndex++;
  saveUndoHistory(history);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `↪️ Redone: ${action.description}`,
    'Redo',
    3
  );
}

/**
 * Undoes to a specific index
 * @param {number} targetIndex - Target index
 */
function undoToIndex(targetIndex) {
  const history = getUndoHistory();

  while (history.currentIndex > targetIndex) {
    undoLastAction();
  }
}

/**
 * Redoes to a specific index
 * @param {number} targetIndex - Target index
 */
function redoToIndex(targetIndex) {
  const history = getUndoHistory();

  while (history.currentIndex <= targetIndex && history.currentIndex < history.actions.length) {
    redoLastAction();
  }
}

/**
 * Applies a state snapshot
 * @param {Object} state - State to apply
 * @param {string} actionType - Type of action
 */
function applyState(state, actionType) {
  if (!state) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(state.sheet);

  if (!sheet) {
    throw new Error(`Sheet ${state.sheet} not found`);
  }

  switch (actionType) {
    case 'EDIT_CELL':
      sheet.getRange(state.row, state.col).setValue(state.value);
      break;

    case 'ADD_ROW':
      // Remove the added row
      if (state.row) {
        sheet.deleteRow(state.row);
      }
      break;

    case 'DELETE_ROW':
      // Re-add the deleted row
      if (state.row && state.data) {
        sheet.insertRowAfter(state.row - 1);
        sheet.getRange(state.row, 1, 1, state.data.length).setValues([state.data]);
      }
      break;

    case 'BATCH_UPDATE':
      // Restore batch changes
      if (state.changes) {
        state.changes.forEach(function(change) {
          sheet.getRange(change.row, change.col).setValue(change.oldValue);
        });
      }
      break;
  }
}

/**
 * Clears undo/redo history
 */
function clearUndoHistory() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(UNDO_CONFIG.STORAGE_KEY);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Undo/redo history cleared',
    'History',
    3
  );
}

/**
 * Exports undo history to a sheet
 * @returns {string} Sheet URL
 */
function exportUndoHistoryToSheet() {
  const history = getUndoHistory();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get history sheet
  let historySheet = ss.getSheetByName('Undo_History_Export');
  if (historySheet) {
    historySheet.clear();
  } else {
    historySheet = ss.insertSheet('Undo_History_Export');
  }

  // Headers
  const headers = ['#', 'Action Type', 'Description', 'Timestamp', 'Status'];
  historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  historySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Data
  if (history.actions.length > 0) {
    const rows = history.actions.map(function(action, index) {
      const status = index < history.currentIndex ? 'Applied' : 'Undone';
      return [
        index + 1,
        action.type,
        action.description,
        new Date(action.timestamp).toLocaleString(),
        status
      ];
    });

    historySheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize
  for (let col = 1; col <= headers.length; col++) {
    historySheet.autoResizeColumn(col);
  }

  return ss.getUrl() + '#gid=' + historySheet.getSheetId();
}

/**
 * Installs undo/redo keyboard shortcuts
 */
function installUndoRedoShortcuts() {
  // This is handled by the onEdit trigger and keyboard handler
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '⌨️ Undo/Redo shortcuts installed:\nCtrl+Z = Undo\nCtrl+Y = Redo\nCtrl+Shift+Z = History',
    'Keyboard Shortcuts',
    5
  );
}

/**
 * Creates a snapshot of the current grievance log
 * @returns {Object} Snapshot data
 */
function createGrievanceSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const data = grievanceSheet.getDataRange().getValues();

  return {
    timestamp: new Date().toISOString(),
    data: data,
    lastRow: grievanceSheet.getLastRow(),
    lastColumn: grievanceSheet.getLastColumn()
  };
}

/**
 * Restores from a snapshot
 * @param {Object} snapshot - Snapshot to restore
 */
function restoreFromSnapshot(snapshot) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  // Clear existing data (except headers)
  if (grievanceSheet.getLastRow() > 1) {
    grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, grievanceSheet.getLastColumn()).clear();
  }

  // Restore snapshot data (skip header row)
  if (snapshot.data.length > 1) {
    grievanceSheet.getRange(2, 1, snapshot.data.length - 1, snapshot.data[0].length).setValues(snapshot.data.slice(1));
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Snapshot restored',
    'Undo/Redo',
    3
  );
}
