/**
 * ------------------------------------------------------------------------====
 * GOOGLE DRIVE INTEGRATION
 * ------------------------------------------------------------------------====
 *
 * Attach and manage documents for grievances
 * Features:
 * - Link Drive folders to grievances
 * - Upload evidence and documents
 * - Auto-organize by grievance ID
 * - Version control
 * - Quick access to attachments
 */

/**
 * Creates root folder for 509 Dashboard in Google Drive
 * @returns {Folder} Root folder
 */
function createRootFolder() {
  const folderName = '509 Dashboard - Grievance Files';

  // Check if folder already exists
  const existingFolders = DriveApp.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  // Create new folder
  const folder = DriveApp.createFolder(folderName);

  Logger.log(`Created root folder: ${folderName}`);
  return folder;
}

/**
 * Creates folder for a specific grievance
 * @param {string} grievanceId - Grievance ID
 * @param {string} grievantName - Optional grievant name for folder naming
 * @returns {Folder} Grievance folder
 */
function createGrievanceFolder(grievanceId, grievantName) {
  const rootFolder = createRootFolder();

  // Create folder name with grievant name if provided
  let folderName = `Grievance_${grievanceId}`;
  if (grievantName) {
    folderName = `Grievance_${grievanceId}_${grievantName}`;
  }

  // Check if grievance folder already exists
  const existingFolders = rootFolder.getFoldersByName(folderName);

  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }

  // Create new grievance folder
  const folder = rootFolder.createFolder(folderName);

  // Create subfolders
  folder.createFolder('Evidence');
  folder.createFolder('Correspondence');
  folder.createFolder('Forms');
  folder.createFolder('Other');

  Logger.log(`Created folder structure for ${grievanceId}`);
  return folder;
}

/**
 * Links Drive folder to grievance in spreadsheet
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 */
function linkFolderToGrievance(grievanceId, folderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) {
    throw new Error('Grievance Log sheet not found');
  }

  const lastRow = grievanceSheet.getLastRow();
  const data = grievanceSheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === grievanceId) {
      const row = i + 2;

      // Store folder ID in column AC (29) - Extension column not in GRIEVANCE_COLS
      // TODO: Add FOLDER_ID constant to GRIEVANCE_COLS
      grievanceSheet.getRange(row, 29).setValue(folderId);

      // Create hyperlink to folder in column AD (30) - Extension column not in GRIEVANCE_COLS
      // TODO: Add FOLDER_URL constant to GRIEVANCE_COLS
      const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;
      grievanceSheet.getRange(row, 30).setValue(folderUrl);

      Logger.log(`Linked folder ${folderId} to grievance ${grievanceId}`);
      return;
    }
  }

  throw new Error(`Grievance ${grievanceId} not found`);
}

/**
 * Sets up Drive folder for selected grievance
 */
function setupDriveFolderForGrievance() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!grievanceId) {
    ui.alert(
      '⚠️ No Grievance ID',
      'Selected row does not have a grievance ID.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '📁 Setup Drive Folder',
    `Create a Google Drive folder for grievance ${grievanceId}?\n\n` +
    'This will create a folder with subfolders for:\n' +
    '• Evidence\n' +
    '• Correspondence\n' +
    '• Forms\n' +
    '• Other\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('📁 Creating folder...', 'Please wait', -1);

    const folder = createGrievanceFolder(grievanceId);
    linkFolderToGrievance(grievanceId, folder.getId());

    ui.alert(
      '✅ Folder Created',
      `Drive folder created successfully for ${grievanceId}.\n\n` +
      `Folder link has been added to the grievance record.\n\n` +
      `Click the link in column AD to open the folder.`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Shows file upload dialog for a grievance
 */
function showFileUploadDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!grievanceId) {
    ui.alert(
      '⚠️ No Grievance ID',
      'Selected row does not have a grievance ID.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Get or create folder
  let folder;
  try {
    folder = createGrievanceFolder(grievanceId);
    linkFolderToGrievance(grievanceId, folder.getId());
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    return;
  }

  const html = createFileUploadHTML(grievanceId, folder.getId());
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(650)
    .setHeight(500);

  ui.showModalDialog(htmlOutput, `📎 Upload Files - ${grievanceId}`);
}

/**
 * Creates HTML for file upload interface
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 * @returns {string} HTML content
 */
function createFileUploadHTML(grievanceId, folderId) {
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

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
    .info-box {
      background: #e8f0fe;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .folder-link {
      background: #1a73e8;
      color: white;
      padding: 12px 20px;
      border-radius: 4px;
      text-decoration: none;
      display: inline-block;
      margin: 15px 0;
    }
    .folder-link:hover {
      background: #1557b0;
    }
    .instructions {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 4px;
      margin: 15px 0;
    }
    .instructions h3 {
      margin-top: 0;
      color: #333;
    }
    .instructions ol {
      margin: 10px 0;
      padding-left: 20px;
    }
    .instructions li {
      margin: 8px 0;
      color: #666;
    }
    .category-list {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin: 15px 0;
    }
    .category {
      background: #f8f9fa;
      padding: 12px;
      border-radius: 4px;
      border-left: 3px solid #1a73e8;
    }
    .category strong {
      color: #333;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📎 Upload Files</h2>

    <div class="info-box">
      <strong>Grievance:</strong> ${grievanceId}<br>
      <strong>Drive Folder:</strong> Created and linked
    </div>

    <a href="${folderUrl}" target="_blank" class="folder-link">
      📁 Open Drive Folder
    </a>

    <div class="instructions">
      <h3>How to Upload Files</h3>
      <ol>
        <li>Click "Open Drive Folder" above to open the folder in Google Drive</li>
        <li>Navigate to the appropriate subfolder:
          <div class="category-list">
            <div class="category"><strong>📄 Evidence</strong><br>Photos, documents, proof</div>
            <div class="category"><strong>✉️ Correspondence</strong><br>Emails, letters</div>
            <div class="category"><strong>📋 Forms</strong><br>Signed forms, contracts</div>
            <div class="category"><strong>📎 Other</strong><br>Miscellaneous files</div>
          </div>
        </li>
        <li>Drag and drop files or click "New" → "File upload"</li>
        <li>All files are automatically organized and linked to this grievance</li>
      </ol>
    </div>

    <div class="info-box" style="background: #fff3cd; border-left-color: #ffc107;">
      <strong>💡 Tip:</strong> Files are automatically versioned in Drive. If you upload a file with the same name, Drive will keep both versions.
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Shows files attached to a grievance
 */
function showGrievanceFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert(
      '⚠️ Wrong Sheet',
      'Please select a grievance row in the Grievance Log sheet.',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRow = activeSheet.getActiveCell().getRow();

  if (activeRow < 2) {
    ui.alert(
      '⚠️ Invalid Selection',
      'Please select a grievance row (not the header).',
      ui.ButtonSet.OK
    );
    return;
  }

  const grievanceId = activeSheet.getRange(activeRow, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  // Column AC (29) - Extension column not in GRIEVANCE_COLS
  // TODO: Add FOLDER_ID constant to GRIEVANCE_COLS
  const folderId = activeSheet.getRange(activeRow, 29).getValue();

  if (!folderId) {
    ui.alert(
      'ℹ️ No Folder Linked',
      `No Drive folder has been set up for ${grievanceId}.\n\n` +
      'Use "Setup Drive Folder" to create one.',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = listFolderFiles(folder);

    if (files.length === 0) {
      ui.alert(
        'ℹ️ No Files',
        `No files have been uploaded to the folder for ${grievanceId}.`,
        ui.ButtonSet.OK
      );
      return;
    }

    const html = createFileListHTML(grievanceId, folderId, files);
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

    ui.showModalDialog(htmlOutput, `📁 Files - ${grievanceId}`);

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

/**
 * Lists all files in a folder recursively
 * @param {Folder} folder - Drive folder
 * @param {string} path - Current path (for recursion)
 * @returns {Array} Array of file objects
 */
function listFolderFiles(folder, path = '') {
  const files = [];

  // Get files in current folder
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files.push({
      name: file.getName(),
      url: file.getUrl(),
      size: formatFileSize(file.getSize()),
      modified: file.getLastUpdated().toLocaleString(),
      path: path,
      type: file.getMimeType()
    });
  }

  // Recursively get files from subfolders
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const subfolder = folderIterator.next();
    const subfolderName = subfolder.getName();
    const subfiles = listFolderFiles(subfolder, path ? `${path}/${subfolderName}` : subfolderName);
    files.push(...subfiles);
  }

  return files;
}

/**
 * Formats file size in human-readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';

  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Creates HTML for file list display
 * @param {string} grievanceId - Grievance ID
 * @param {string} folderId - Drive folder ID
 * @param {Array} files - Array of file objects
 * @returns {string} HTML content
 */
function createFileListHTML(grievanceId, folderId, files) {
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  const filesList = files
    .map(function(file) { return `
      <div class="file-item">
        <div class="file-icon">${getFileIcon(file.type)}</div>
        <div class="file-details">
          <div class="file-name">
            <a href="${file.url}" target="_blank">${file.name}</a>
          </div>
          <div class="file-meta">
            ${file.path ? `${file.path} • ` : ''}${file.size} • Modified: ${file.modified}
          </div>
        </div>
      </div>
    `; })
    .join('');

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
    .summary {
      background: #e8f0fe;
      padding: 12px;
      border-radius: 4px;
      margin: 15px 0;
      border-left: 4px solid #1a73e8;
    }
    .file-item {
      display: flex;
      align-items: center;
      padding: 12px;
      margin: 10px 0;
      background: #f8f9fa;
      border-radius: 4px;
      border-left: 3px solid #1a73e8;
    }
    .file-item:hover {
      background: #e9ecef;
    }
    .file-icon {
      font-size: 24px;
      margin-right: 15px;
    }
    .file-details {
      flex: 1;
    }
    .file-name a {
      color: #1a73e8;
      text-decoration: none;
      font-weight: 500;
    }
    .file-name a:hover {
      text-decoration: underline;
    }
    .file-meta {
      font-size: 12px;
      color: #666;
      margin-top: 4px;
    }
    .folder-link {
      background: #1a73e8;
      color: white;
      padding: 10px 20px;
      border-radius: 4px;
      text-decoration: none;
      display: inline-block;
      margin: 15px 0;
    }
    .folder-link:hover {
      background: #1557b0;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>📁 Attached Files - ${grievanceId}</h2>

    <div class="summary">
      <strong>Total Files:</strong> ${files.length}
    </div>

    <a href="${folderUrl}" target="_blank" class="folder-link">
      📁 Open Folder in Drive
    </a>

    <div style="max-height: 400px; overflow-y: auto; margin-top: 15px;">
      ${filesList}
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * Gets emoji icon for file type
 * @param {string} mimeType - MIME type
 * @returns {string} Emoji icon
 */
function getFileIcon(mimeType) {
  if (mimeType.includes('image')) return '🖼️';
  if (mimeType.includes('pdf')) return '📄';
  if (mimeType.includes('word') || mimeType.includes('document')) return '📝';
  if (mimeType.includes('sheet') || mimeType.includes('excel')) return '📊';
  if (mimeType.includes('presentation')) return '📽️';
  if (mimeType.includes('video')) return '🎥';
  if (mimeType.includes('audio')) return '🎵';
  if (mimeType.includes('zip') || mimeType.includes('compressed')) return '🗜️';
  return '📎';
}

/**
 * Batch creates folders for all grievances without them
 */
function batchCreateGrievanceFolders() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📁 Batch Create Folders',
    'This will create Drive folders for all grievances that don\'t have one.\n\n' +
    'This may take a few minutes for large datasets.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('📁 Creating folders...', 'Please wait', -1);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    const lastRow = grievanceSheet.getLastRow();

    if (lastRow < 2) {
      ui.alert('No grievances found.');
      return;
    }

    const data = grievanceSheet.getRange(2, 1, lastRow - 1, 29).getValues();
    let created = 0;
    let skipped = 0;

    data.forEach(function(row, index) {
      const grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
      // Array index 28 = Column AC (29) - Extension column not in GRIEVANCE_COLS
      // TODO: Add FOLDER_ID constant to GRIEVANCE_COLS
      const folderId = row[28];

      if (!grievanceId) {
        skipped++;
        return;
      }

      if (folderId) {
        skipped++;
        return;
      }

      try {
        const folder = createGrievanceFolder(grievanceId);
        linkFolderToGrievance(grievanceId, folder.getId());
        created++;

        // Progress update every 10 folders
        if (created % 10 === 0) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Created ${created} folders...`,
            'Progress',
            -1
          );
        }
      } catch (error) {
        Logger.log(`Error creating folder for ${grievanceId}: ${error.message}`);
      }
    });

    ui.alert(
      '✅ Batch Folder Creation Complete',
      `Created ${created} new folders.\n` +
      `Skipped ${skipped} grievances (already had folders or no ID).`,
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}
