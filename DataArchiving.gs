/**
 * ============================================================================
 * DATA ARCHIVING - Automatic Archive of Old Grievances
 * ============================================================================
 *
 * Moves closed grievances older than configured threshold to Archive sheet.
 * Maintains performance as data grows over time.
 *
 * @module DataArchiving
 * @version 2.1.0
 * ============================================================================
 */

/**
 * Archives old closed grievances to Archive sheet
 * Moves grievances closed for more than NUMERIC_CONSTANTS.ARCHIVE_AFTER_YEARS years
 *
 * @param {boolean} dryRun - If true, only count what would be archived
 * @returns {Object} Result with counts and details
 */
function archiveOldGrievances(dryRun = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const archive = ss.getSheetByName(SHEETS.ARCHIVE);

  if (!grievanceLog) {
    throw new Error(ERROR_MESSAGES.SHEET_NOT_FOUND(SHEETS.GRIEVANCE_LOG));
  }

  if (!archive) {
    throw new Error(ERROR_MESSAGES.SHEET_NOT_FOUND(SHEETS.ARCHIVE) +
                    ' Please create Archive sheet first.');
  }

  // Calculate cutoff date
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - NUMERIC_CONSTANTS.ARCHIVE_AFTER_YEARS);

  const data = grievanceLog.getDataRange().getValues();
  const headers = data[0];
  const toArchive = [];
  const toKeep = [headers]; // Keep header

  let archivedCount = 0;
  let keptCount = 0;

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[GRIEVANCE_COLS.STATUS - 1];
    const dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];

    // Archive if closed/settled and older than cutoff
    const shouldArchive =
      (status === 'Closed' || status === 'Settled' || status === 'Withdrawn') &&
      dateClosed &&
      dateClosed instanceof Date &&
      dateClosed < cutoffDate;

    if (shouldArchive) {
      toArchive.push(row);
      archivedCount++;
    } else {
      toKeep.push(row);
      keptCount++;
    }
  }

  const result = {
    totalProcessed: data.length - 1, // Exclude header
    archivedCount: archivedCount,
    keptCount: keptCount,
    cutoffDate: cutoffDate,
    dryRun: dryRun
  };

  // If dry run, just return counts
  if (dryRun) {
    Logger.log(`Dry run: Would archive ${archivedCount} grievances`);
    return result;
  }

  // Move to archive in batches
  if (toArchive.length > 0) {
    SpreadsheetApp.getActive().toast(
      ERROR_MESSAGES.IN_PROGRESS(`Archiving ${toArchive.length} grievances`),
      'Archiving',
      -1
    );

    const batchSize = NUMERIC_CONSTANTS.ARCHIVE_BATCH_SIZE;
    for (let i = 0; i < toArchive.length; i += batchSize) {
      const batch = toArchive.slice(i, i + batchSize);
      const lastRow = archive.getLastRow();

      archive.getRange(lastRow + 1, 1, batch.length, batch[0].length)
            .setValues(batch);

      SpreadsheetApp.flush();

      // Update progress
      const progress = Math.min(100, Math.round(((i + batch.length) / toArchive.length) * 100));
      SpreadsheetApp.getActive().toast(
        `Archiving... ${i + batch.length}/${toArchive.length} (${progress}%)`,
        'Progress',
        NUMERIC_CONSTANTS.TOAST_DURATION_SHORT
      );
    }

    // Update main grievance log
    grievanceLog.clear();
    grievanceLog.getRange(1, 1, toKeep.length, toKeep[0].length)
                .setValues(toKeep);

    SpreadsheetApp.flush();

    // Log the archiving event
    logAuditEvent('DATA_ARCHIVED', {
      grievances_archived: archivedCount,
      cutoff_date: cutoffDate.toISOString()
    }, 'INFO');

    SpreadsheetApp.getActive().toast(
      ERROR_MESSAGES.SUCCESS(`Archived ${archivedCount} old grievances`),
      'Complete',
      NUMERIC_CONSTANTS.TOAST_DURATION_MEDIUM
    );
  } else {
    SpreadsheetApp.getActive().toast(
      'No grievances to archive',
      'Info',
      NUMERIC_CONSTANTS.TOAST_DURATION_SHORT
    );
  }

  return result;
}

/**
 * Shows archive dialog with preview
 */
function showArchiveDialog() {
  // Run dry run to get counts
  const dryRunResult = archiveOldGrievances(true);

  const html = createArchiveDialogHTML(dryRunResult);

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(600).setHeight(400),
    '📦 Archive Old Grievances'
  );
}

/**
 * Creates HTML for archive dialog
 *
 * @param {Object} dryRunResult - Result from dry run
 * @returns {string} HTML content
 */
function createArchiveDialogHTML(dryRunResult) {
  const styles = getCommonDialogStyles();

  const body = `
    <div class="container">
      <h2>📦 Archive Old Grievances</h2>

      ${createInfoBox(
        'Archive Policy',
        `Grievances closed for more than ${NUMERIC_CONSTANTS.ARCHIVE_AFTER_YEARS} years will be moved to the Archive sheet.`,
        'info'
      )}

      <div style="background: #f8f9fa; padding: 15px; border-radius: 4px; margin: 15px 0;">
        <h3 style="margin-top: 0;">Preview:</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
          <div>
            <strong>Total Grievances:</strong><br>
            ${dryRunResult.totalProcessed}
          </div>
          <div>
            <strong>To Archive:</strong><br>
            <span style="color: #f97316; font-size: 20px; font-weight: bold;">
              ${dryRunResult.archivedCount}
            </span>
          </div>
          <div>
            <strong>To Keep:</strong><br>
            ${dryRunResult.keptCount}
          </div>
          <div>
            <strong>Cutoff Date:</strong><br>
            ${dryRunResult.cutoffDate.toLocaleDateString()}
          </div>
        </div>
      </div>

      ${dryRunResult.archivedCount > 0 ? createInfoBox(
        'Warning',
        'This action will move ' + dryRunResult.archivedCount + ' grievances to the Archive sheet. ' +
        'The Archive sheet will still be accessible, but archived grievances will not appear in reports.',
        'warning'
      ) : ''}

      <div class="button-group">
        <button class="btn-secondary" onclick="google.script.host.close()">
          Cancel
        </button>
        <button
          class="btn-primary"
          onclick="archiveNow()"
          ${dryRunResult.archivedCount === 0 ? 'disabled' : ''}
        >
          Archive ${dryRunResult.archivedCount} Grievances
        </button>
      </div>

      ${createLoadingSpinner('Archiving...', 'archiving')}
    </div>
  `;

  const scripts = `
    function archiveNow() {
      document.getElementById('archiving').style.display = 'block';
      google.script.run
        .withSuccessHandler(onArchiveComplete)
        .withFailureHandler(onArchiveError)
        .archiveOldGrievances(false);
    }

    function onArchiveComplete(result) {
      alert('✅ Successfully archived ' + result.archivedCount + ' grievances');
      google.script.host.close();
    }

    function onArchiveError(error) {
      document.getElementById('archiving').style.display = 'none';
      alert('❌ Error: ' + error.message);
    }
  `;

  return createHTMLPage('Archive Old Grievances', styles, body, scripts);
}

/**
 * Restores grievances from archive back to main log
 *
 * @param {Array<string>} grievanceIds - Array of grievance IDs to restore
 * @returns {Object} Result with counts
 */
function restoreFromArchive(grievanceIds) {
  requireRole('ADMIN', 'Restore from Archive');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  const archive = ss.getSheetByName(SHEETS.ARCHIVE);

  if (!grievanceLog || !archive) {
    throw new Error('Required sheets not found');
  }

  const archiveData = archive.getDataRange().getValues();
  const headers = archiveData[0];
  const toRestore = [];
  const remainingInArchive = [headers];

  let restoredCount = 0;

  // Find grievances to restore
  for (let i = 1; i < archiveData.length; i++) {
    const row = archiveData[i];
    const grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];

    if (grievanceIds.includes(grievanceId)) {
      toRestore.push(row);
      restoredCount++;
    } else {
      remainingInArchive.push(row);
    }
  }

  if (toRestore.length > 0) {
    // Add back to grievance log
    const lastRow = grievanceLog.getLastRow();
    grievanceLog.getRange(lastRow + 1, 1, toRestore.length, toRestore[0].length)
                .setValues(toRestore);

    // Update archive (remove restored items)
    archive.clear();
    archive.getRange(1, 1, remainingInArchive.length, remainingInArchive[0].length)
           .setValues(remainingInArchive);

    SpreadsheetApp.flush();

    logAuditEvent('DATA_RESTORED', {
      grievances_restored: restoredCount,
      grievance_ids: grievanceIds
    }, 'INFO');
  }

  return {
    restoredCount: restoredCount,
    notFound: grievanceIds.length - restoredCount
  };
}

/**
 * Gets statistics about archived data
 *
 * @returns {Object} Archive statistics
 */
function getArchiveStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archive = ss.getSheetByName(SHEETS.ARCHIVE);

  if (!archive) {
    return {
      totalArchived: 0,
      oldestDate: null,
      newestDate: null,
      byStatus: {}
    };
  }

  const data = archive.getDataRange().getValues();
  const stats = {
    totalArchived: data.length - 1, // Exclude header
    oldestDate: null,
    newestDate: null,
    byStatus: {}
  };

  for (let i = 1; i < data.length; i++) {
    const dateClosed = data[i][GRIEVANCE_COLS.DATE_CLOSED - 1];
    const status = data[i][GRIEVANCE_COLS.STATUS - 1];

    // Track dates
    if (dateClosed instanceof Date) {
      if (!stats.oldestDate || dateClosed < stats.oldestDate) {
        stats.oldestDate = dateClosed;
      }
      if (!stats.newestDate || dateClosed > stats.newestDate) {
        stats.newestDate = dateClosed;
      }
    }

    // Count by status
    stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
  }

  return stats;
}
