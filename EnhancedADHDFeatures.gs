/**
 * ------------------------------------------------------------------------====
 * ENHANCED ADHD FEATURES
 * ------------------------------------------------------------------------====
 *
 * Advanced accessibility and focus features for neurodivergent users
 * Features:
 * - Focus mode with distraction reduction
 * - Custom color themes for readability
 * - Font size adjustments
 * - Row highlighting and zebra striping
 * - Reduced motion mode
 * - Audio cues (optional)
 * - Task timers and reminders
 * - Quick capture notepad
 * - Customizable dashboard layouts
 * - Attention span breaks
 */

/**
 * ADHD configuration
 */
ADHD_CONFIG = {
  FOCUS_MODE_COLORS: {
    background: '#f5f5f5',
    header: '#4a4a4a',
    accent: '#6b9bd1'
  },
  HIGH_CONTRAST_COLORS: {
    background: '#ffffff',
    header: '#000000',
    accent: '#0066cc',
    alternateRow: '#f0f0f0'
  },
  PASTEL_COLORS: {
    background: '#fef9e7',
    header: '#85929e',
    accent: '#7fb3d5',
    alternateRow: '#ebf5fb'
  }
};

/**
 * Shows ADHD features control panel
 */
function showADHDControlPanel() {
  const html = createADHDControlPanelHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '♿ ADHD Control Panel');
}

/**
 * Creates HTML for ADHD control panel
 */
function createADHDControlPanelHTML() {
  const currentSettings = getADHDSettings();

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
      max-height: 650px;
      overflow-y: auto;
    }
    h2 {
      color: #1a73e8;
      margin-top: 0;
      border-bottom: 3px solid #1a73e8;
      padding-bottom: 10px;
    }
    .section {
      background: #f8f9fa;
      padding: 20px;
      margin: 15px 0;
      border-radius: 8px;
      border-left: 4px solid #1a73e8;
    }
    .section-title {
      font-weight: bold;
      font-size: 16px;
      color: #333;
      margin-bottom: 15px;
    }
    .setting-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 0;
      border-bottom: 1px solid #e0e0e0;
    }
    .setting-row:last-child {
      border-bottom: none;
    }
    .setting-label {
      font-weight: 500;
      color: #555;
    }
    .setting-description {
      font-size: 13px;
      color: #888;
      margin-top: 4px;
    }
    button {
      background: #1a73e8;
      color: white;
      border: none;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 4px;
      cursor: pointer;
    }
    button:hover {
      background: #1557b0;
    }
    button.secondary {
      background: #6c757d;
    }
    button.active {
      background: #4caf50;
    }
    select, input[type="range"] {
      padding: 8px;
      border: 2px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    .slider-container {
      display: flex;
      align-items: center;
      gap: 15px;
    }
    .slider-value {
      min-width: 50px;
      text-align: center;
      font-weight: bold;
      color: #1a73e8;
    }
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .theme-preview {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 15px;
      margin: 15px 0;
    }
    .theme-card {
      padding: 15px;
      border-radius: 8px;
      cursor: pointer;
      border: 3px solid transparent;
      transition: all 0.2s;
    }
    .theme-card:hover {
      transform: scale(1.05);
    }
    .theme-card.selected {
      border-color: #1a73e8;
      box-shadow: 0 4px 12px rgba(26, 115, 232, 0.3);
    }
    .theme-name {
      font-weight: bold;
      margin-bottom: 8px;
      text-align: center;
    }
    .color-bar {
      height: 40px;
      border-radius: 4px;
      margin: 5px 0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>♿ ADHD Control Panel</h2>

    <div class="info-box">
      <strong>💡 Neurodiversity Support:</strong> Customize the dashboard for optimal focus and comfort.
      All settings are saved to your profile.
    </div>

    <!-- Visual Settings -->
    <div class="section">
      <div class="section-title">🎨 Visual Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Color Theme</div>
          <div class="setting-description">Choose a color scheme that works best for you</div>
        </div>
      </div>

      <div class="theme-preview">
        <div class="theme-card ${currentSettings.theme === 'default' ? 'selected' : ''}" onclick="selectTheme('default')">
          <div class="theme-name">Default</div>
          <div class="color-bar" style="background: #1a73e8;"></div>
          <div class="color-bar" style="background: #f8f9fa;"></div>
        </div>

        <div class="theme-card ${currentSettings.theme === 'high-contrast' ? 'selected' : ''}" onclick="selectTheme('high-contrast')">
          <div class="theme-name">High Contrast</div>
          <div class="color-bar" style="background: #000000;"></div>
          <div class="color-bar" style="background: #ffffff;"></div>
        </div>

        <div class="theme-card ${currentSettings.theme === 'pastel' ? 'selected' : ''}" onclick="selectTheme('pastel')">
          <div class="theme-name">Soft Pastel</div>
          <div class="color-bar" style="background: #7fb3d5;"></div>
          <div class="color-bar" style="background: #fef9e7;"></div>
        </div>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Font Size</div>
          <div class="setting-description">Adjust text size for better readability</div>
        </div>
        <div class="slider-container">
          <input type="range" id="fontSize" min="8" max="14" value="${currentSettings.fontSize || 10}" oninput="updateSlider('fontSize')">
          <span class="slider-value" id="fontSizeValue">${currentSettings.fontSize || 10}pt</span>
        </div>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Row Highlighting</div>
          <div class="setting-description">Zebra striping for easier row tracking</div>
        </div>
        <button class="${currentSettings.zebraStripes ? 'active' : 'secondary'}" onclick="toggleZebraStripes()">
          ${currentSettings.zebraStripes ? '✅ Enabled' : 'Enable'}
        </button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Gridlines</div>
          <div class="setting-description">Show or hide cell borders</div>
        </div>
        <button class="${currentSettings.gridlines ? 'active' : 'secondary'}" onclick="toggleGridlines()">
          ${currentSettings.gridlines ? '✅ Visible' : 'Hidden'}
        </button>
      </div>
    </div>

    <!-- Focus Settings -->
    <div class="section">
      <div class="section-title">🎯 Focus Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Focus Mode</div>
          <div class="setting-description">Minimize distractions, hide non-essential UI</div>
        </div>
        <button onclick="activateFocusMode()">🎯 Activate Focus Mode</button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Reduced Motion</div>
          <div class="setting-description">Disable animations and transitions</div>
        </div>
        <button class="${currentSettings.reducedMotion ? 'active' : 'secondary'}" onclick="toggleReducedMotion()">
          ${currentSettings.reducedMotion ? '✅ Enabled' : 'Enable'}
        </button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Quick Capture</div>
          <div class="setting-description">Floating notepad for quick thoughts</div>
        </div>
        <button onclick="showQuickCapture()">📝 Open Quick Capture</button>
      </div>
    </div>

    <!-- Productivity Settings -->
    <div class="section">
      <div class="section-title">⏱️ Productivity Settings</div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Pomodoro Timer</div>
          <div class="setting-description">25-minute focus sessions with breaks</div>
        </div>
        <button onclick="startPomodoroTimer()">⏱️ Start Timer</button>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Break Reminders</div>
          <div class="setting-description">Gentle reminders to take breaks</div>
        </div>
        <select id="breakInterval" onchange="setBreakInterval()">
          <option value="0" ${!currentSettings.breakInterval ? 'selected' : ''}>Disabled</option>
          <option value="30" ${currentSettings.breakInterval === 30 ? 'selected' : ''}>Every 30 minutes</option>
          <option value="60" ${currentSettings.breakInterval === 60 ? 'selected' : ''}>Every hour</option>
          <option value="90" ${currentSettings.breakInterval === 90 ? 'selected' : ''}>Every 90 minutes</option>
        </select>
      </div>

      <div class="setting-row">
        <div>
          <div class="setting-label">Task Complexity Indicators</div>
          <div class="setting-description">Visual cues for task difficulty</div>
        </div>
        <button class="${currentSettings.complexityIndicators ? 'active' : 'secondary'}" onclick="toggleComplexityIndicators()">
          ${currentSettings.complexityIndicators ? '✅ Enabled' : 'Enable'}
        </button>
      </div>
    </div>

    <!-- Quick Actions -->
    <div style="margin-top: 25px; padding-top: 20px; border-top: 2px solid #e0e0e0; display: flex; gap: 10px;">
      <button onclick="saveSettings()">💾 Save Settings</button>
      <button class="secondary" onclick="resetToDefaults()">🔄 Reset to Defaults</button>
      <button class="secondary" onclick="google.script.host.close()">Close</button>
    </div>
  </div>

  <script>
    var selectedTheme = '${currentSettings.theme || 'default'}';

    function selectTheme(theme) {
      selectedTheme = theme;
      document.querySelectorAll('.theme-card').forEach(function(card) {
        card.classList.remove('selected');
      });
      event.target.closest('.theme-card').classList.add('selected');
    }

    function updateSlider(sliderId) {
      const slider = document.getElementById(sliderId);
      const value = document.getElementById(sliderId + 'Value');
      value.textContent = slider.value + 'pt';
    }

    function toggleZebraStripes() {
      google.script.run.toggleZebraStripes();
      setTimeout(function() { return location.reload(), 1000; });
    }

    function toggleGridlines() {
      google.script.run.toggleGridlinesADHD();
      setTimeout(function() { return location.reload(), 1000; });
    }

    function toggleReducedMotion() {
      google.script.run.toggleReducedMotion();
      setTimeout(function() { return location.reload(), 1000; });
    }

    function toggleComplexityIndicators() {
      google.script.run.toggleComplexityIndicators();
      setTimeout(function() { return location.reload(), 1000; });
    }

    function activateFocusMode() {
      google.script.run.activateFocusMode();
      google.script.host.close();
    }

    function showQuickCapture() {
      google.script.run.showQuickCaptureNotepad();
    }

    function startPomodoroTimer() {
      google.script.run.startPomodoroTimer();
      google.script.host.close();
    }

    function setBreakInterval() {
      const interval = document.getElementById('breakInterval').value;
      google.script.run.setBreakReminders(parseInt(interval));
    }

    function saveSettings() {
      const settings = {
        theme: selectedTheme,
        fontSize: document.getElementById('fontSize').value,
        breakInterval: document.getElementById('breakInterval').value
      };

      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Settings saved!');
          google.script.host.close();
        })
        .saveADHDSettings(settings);
    }

    function resetToDefaults() {
      if (confirm('Reset all ADHD settings to defaults?')) {
        google.script.run
          .withSuccessHandler(function() {
            alert('✅ Settings reset!');
            location.reload();
          })
          .resetADHDSettings();
      }
    }
  </script>
</body>
</html>
  `;
}

/**
 * Gets current ADHD settings
 * @returns {Object} Settings
 */
function getADHDSettings() {
  const props = PropertiesService.getUserProperties();
  const settingsJSON = props.getProperty('adhdSettings');

  if (settingsJSON) {
    return JSON.parse(settingsJSON);
  }

  // Defaults
  return {
    theme: 'default',
    fontSize: 10,
    zebraStripes: false,
    gridlines: true,
    reducedMotion: false,
    breakInterval: 0,
    complexityIndicators: false
  };
}

/**
 * Saves ADHD settings
 * @param {Object} settings - Settings to save
 */
function saveADHDSettings(settings) {
  const props = PropertiesService.getUserProperties();
  const currentSettings = getADHDSettings();

  // Merge with current
  const newSettings = { ...currentSettings, ...settings };

  props.setProperty('adhdSettings', JSON.stringify(newSettings));

  // Apply settings
  applyADHDSettings(newSettings);
}

/**
 * Applies ADHD settings to sheets
 * @param {Object} settings - Settings to apply
 */
function applyADHDSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    // Font size
    if (settings.fontSize) {
      sheet.getDataRange().setFontSize(parseInt(settings.fontSize));
    }

    // Zebra stripes
    if (settings.zebraStripes) {
      applyZebraStripes(sheet);
    }

    // Gridlines
    if (settings.gridlines !== undefined) {
      sheet.setHiddenGridlines(!settings.gridlines);
    }
  });
}

/**
 * Toggles zebra striping
 */
function toggleZebraStripes() {
  const settings = getADHDSettings();
  settings.zebraStripes = !settings.zebraStripes;
  saveADHDSettings(settings);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (settings.zebraStripes) {
      applyZebraStripes(sheet);
    } else {
      removeZebraStripes(sheet);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    settings.zebraStripes ? '✅ Zebra stripes enabled' : '🔕 Zebra stripes disabled',
    'Visual Settings',
    3
  );
}

/**
 * Applies zebra striping to sheet
 * @param {Sheet} sheet - Sheet to apply to
 */
function applyZebraStripes(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());

  // Apply banding
  const banding = range.getBandings()[0];
  if (banding) {
    banding.remove();
  }

  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
}

/**
 * Removes zebra striping
 * @param {Sheet} sheet - Sheet to remove from
 */
function removeZebraStripes(sheet) {
  const bandings = sheet.getBandings();
  bandings.forEach(function(banding) { return banding.remove(); });
}

/**
 * Toggles gridlines
 */
function toggleGridlinesADHD() {
  const settings = getADHDSettings();
  settings.gridlines = !settings.gridlines;
  saveADHDSettings(settings);
}

/**
 * Toggles reduced motion
 */
function toggleReducedMotion() {
  const settings = getADHDSettings();
  settings.reducedMotion = !settings.reducedMotion;
  saveADHDSettings(settings);
}

/**
 * Toggles complexity indicators
 */
function toggleComplexityIndicators() {
  const settings = getADHDSettings();
  settings.complexityIndicators = !settings.complexityIndicators;
  saveADHDSettings(settings);
}

/**
 * Activates focus mode
 */
function activateFocusMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hide all sheets except active one
  const activeSheet = ss.getActiveSheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    if (sheet.getName() !== activeSheet.getName()) {
      sheet.hideSheet();
    }
  });

  // Hide gridlines
  activeSheet.setHiddenGridlines(true);

  // Set focus mode colors
  activeSheet.getDataRange().setBackground('#f5f5f5');

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '🎯 Focus mode activated. Press Ctrl+Shift+F to exit.',
    'Focus Mode',
    5
  );
}

/**
 * Deactivates focus mode
 */
function deactivateFocusMode() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Show all sheets
  sheets.forEach(function(sheet) {
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  });

  const settings = getADHDSettings();
  const activeSheet = ss.getActiveSheet();

  // Restore gridlines based on settings
  activeSheet.setHiddenGridlines(!settings.gridlines);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ Focus mode deactivated',
    'Focus Mode',
    3
  );
}

/**
 * Shows quick capture notepad
 */
function showQuickCaptureNotepad() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
    h3 { margin-top: 0; color: #1a73e8; }
    textarea { width: 100%; height: 300px; padding: 10px; border: 2px solid #ddd; border-radius: 4px; font-family: Arial, sans-serif; font-size: 14px; resize: vertical; box-sizing: border-box; }
    button { background: #1a73e8; color: white; border: none; padding: 12px 24px; font-size: 14px; border-radius: 4px; cursor: pointer; margin: 10px 5px 0 0; }
    button:hover { background: #1557b0; }
    button.secondary { background: #6c757d; }
    .info { background: #e8f0fe; padding: 10px; border-radius: 4px; margin: 10px 0; font-size: 13px; }
  </style>
</head>
<body>
  <h3>📝 Quick Capture</h3>
  <div class="info">💡 Jot down quick thoughts without losing focus. Notes are saved automatically.</div>
  <textarea id="notes" placeholder="Type your thoughts here..." autofocus></textarea>
  <button onclick="saveNotes()">💾 Save</button>
  <button class="secondary" onclick="clearNotes()">🗑️ Clear</button>
  <button class="secondary" onclick="google.script.host.close()">Close</button>

  <script>
    // Load existing notes
    google.script.run.withSuccessHandler(loadNotes).getQuickCaptureNotes();

    function loadNotes(notes) {
      if (notes) {
        document.getElementById('notes').value = notes;
      }
    }

    function saveNotes() {
      const notes = document.getElementById('notes').value;
      google.script.run
        .withSuccessHandler(function() {
          alert('✅ Notes saved!');
        })
        .saveQuickCaptureNotes(notes);
    }

    function clearNotes() {
      if (confirm('Clear all notes?')) {
        document.getElementById('notes').value = '';
        google.script.run.saveQuickCaptureNotes('');
      }
    }

    // Auto-save every 30 seconds
    setInterval(function() {
      const notes = document.getElementById('notes').value;
      if (notes) {
        google.script.run.saveQuickCaptureNotes(notes);
      }
    }, 30000);
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📝 Quick Capture');
}

/**
 * Gets quick capture notes
 * @returns {string} Notes
 */
function getQuickCaptureNotes() {
  const props = PropertiesService.getUserProperties();
  return props.getProperty('quickCaptureNotes') || '';
}

/**
 * Saves quick capture notes
 * @param {string} notes - Notes to save
 */
function saveQuickCaptureNotes(notes) {
  const props = PropertiesService.getUserProperties();
  props.setProperty('quickCaptureNotes', notes);
}

/**
 * Starts Pomodoro timer
 */
function startPomodoroTimer() {
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 40px; margin: 0; text-align: center; background: #1a73e8; color: white; }
    h2 { margin: 0 0 20px 0; font-size: 24px; }
    .timer { font-size: 72px; font-weight: bold; margin: 40px 0; font-family: 'Courier New', monospace; }
    .status { font-size: 18px; margin: 20px 0; }
    button { background: white; color: #1a73e8; border: none; padding: 15px 30px; font-size: 16px; border-radius: 8px; cursor: pointer; margin: 10px; font-weight: bold; }
    button:hover { background: #f0f0f0; }
    button.stop { background: #f44336; color: white; }
    button.stop:hover { background: #d32f2f; }
    .progress-bar { width: 100%; height: 8px; background: rgba(255,255,255,0.3); border-radius: 4px; margin: 20px 0; }
    .progress-fill { height: 100%; background: white; border-radius: 4px; transition: width 1s linear; }
  </style>
</head>
<body>
  <h2>🍅 Pomodoro Timer</h2>
  <div class="status" id="status">Focus Session</div>
  <div class="timer" id="timer">25:00</div>
  <div class="progress-bar">
    <div class="progress-fill" id="progress" style="width: 100%;"></div>
  </div>
  <button onclick="toggleTimer()">▶️ Start</button>
  <button class="stop" onclick="stopTimer()">⏹️ Stop</button>
  <button onclick="skipSession()">⏭️ Skip</button>

  <script>
    var timeLeft = 25 * 60; // 25 minutes in seconds
    var totalTime = 25 * 60;
    var isRunning = false;
    var interval;
    var sessionType = 'focus'; // 'focus' or 'break'

    function toggleTimer() {
      if (isRunning) {
        pauseTimer();
      } else {
        startTimer();
      }
    }

    function startTimer() {
      isRunning = true;
      document.querySelector('button').textContent = '⏸️ Pause';

      interval = setInterval(function() {
        if (timeLeft > 0) {
          timeLeft--;
          updateDisplay();
        } else {
          completeSession();
        }
      }, 1000);
    }

    function pauseTimer() {
      isRunning = false;
      clearInterval(interval);
      document.querySelector('button').textContent = '▶️ Resume';
    }

    function stopTimer() {
      if (confirm('Stop timer and exit?')) {
        google.script.host.close();
      }
    }

    function updateDisplay() {
      const minutes = Math.floor(timeLeft / 60);
      const seconds = timeLeft % 60;
      document.getElementById('timer').textContent =
        minutes.toString().padStart(2, '0') + ':' + seconds.toString().padStart(2, '0');

      const progress = (timeLeft / totalTime) * 100;
      document.getElementById('progress').style.width = progress + '%';
    }

    function completeSession() {
      pauseTimer();

      if (sessionType === 'focus') {
        alert('✅ Focus session complete! Time for a 5-minute break.');
        sessionType = 'break';
        timeLeft = 5 * 60;
        totalTime = 5 * 60;
        document.getElementById('status').textContent = 'Break Time';
        document.body.style.background = '#4caf50';
      } else {
        alert('✅ Break complete! Ready for another focus session?');
        sessionType = 'focus';
        timeLeft = 25 * 60;
        totalTime = 25 * 60;
        document.getElementById('status').textContent = 'Focus Session';
        document.body.style.background = '#1a73e8';
      }

      updateDisplay();
    }

    function skipSession() {
      if (confirm('Skip to ' + (sessionType === 'focus' ? 'break' : 'next focus session') + '?')) {
        completeSession();
      }
    }
  </script>
</body>
</html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);

  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '🍅 Pomodoro Timer');
}

/**
 * Sets break reminders
 * @param {number} intervalMinutes - Interval in minutes
 */
function setBreakReminders(intervalMinutes) {
  const settings = getADHDSettings();
  settings.breakInterval = intervalMinutes;
  saveADHDSettings(settings);

  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'showBreakReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger if enabled
  if (intervalMinutes > 0) {
    ScriptApp.newTrigger('showBreakReminder')
      .timeBased()
      .everyMinutes(intervalMinutes)
      .create();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Break reminders enabled (every ${intervalMinutes} minutes)`,
      'ADHD Settings',
      3
    );
  }
}

/**
 * Shows break reminder
 */
function showBreakReminder() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '💆 Time for a quick break! Stretch, hydrate, rest your eyes.',
    'Break Reminder',
    10
  );
}

/**
 * Resets ADHD settings to defaults
 */
function resetADHDSettings() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('adhdSettings');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    // Reset font size for data range (setFontSize only works on ranges, not sheets)
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (maxRows > 0 && maxCols > 0) {
      sheet.getRange(1, 1, maxRows, maxCols).setFontSize(10);
    }
    sheet.setHiddenGridlines(false);
    removeZebraStripes(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '✅ ADHD settings reset to defaults',
    'Settings',
    3
  );
}
