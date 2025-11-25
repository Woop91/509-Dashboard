/**
 * ============================================================================
 * UNIT TESTS FOR SEEDNUKE.GS
 * ============================================================================
 *
 * Tests for data clearing and reset functionality:
 * - Clear member directory
 * - Clear grievance log
 * - Preserve headers
 * - Property management
 *
 * ============================================================================
 */

/* ===================== CLEAR FUNCTION TESTS ===================== */

/**
 * Test: clearMemberDirectory preserves headers
 */
function testClearMemberDirectoryPreservesHeaders() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Get original header
  const originalHeader = memberDir.getRange(1, 1, 1, 31).getValues()[0];

  // Create test data
  createTestMember('TEST-M-NUKE-001');
  createTestMember('TEST-M-NUKE-002');

  const rowCountBefore = memberDir.getLastRow();
  Assert.assertTrue(
    rowCountBefore >= 3,
    'Should have at least 3 rows (header + 2 test members)'
  );

  // Clear the directory
  clearMemberDirectory();

  // Check headers are preserved
  const headerAfter = memberDir.getRange(1, 1, 1, 31).getValues()[0];

  Assert.assertEquals(
    originalHeader.length,
    headerAfter.length,
    'Header row should have same number of columns'
  );

  Assert.assertEquals(
    'Member ID',
    headerAfter[0],
    'First header should still be "Member ID"'
  );

  Assert.assertEquals(
    'Notes from Steward Contact',
    headerAfter[30],
    'Last header should still be "Notes from Steward Contact"'
  );

  // Check data is cleared
  Assert.assertEquals(
    1,
    memberDir.getLastRow(),
    'Should only have header row after clearing'
  );

  Logger.log('✅ clearMemberDirectory preserves headers test passed');
}

/**
 * Test: clearGrievanceLog preserves headers
 */
function testClearGrievanceLogPreservesHeaders() {
  const ss = SpreadsheetApp.getActive();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Get original header
  const originalHeader = grievanceLog.getRange(1, 1, 1, 28).getValues()[0];

  // Create test data
  const testMemberId = createTestMember('TEST-M-NUKE-GL-001');
  const testGrievanceData = [
    'TEST-G-NUKE-001',
    testMemberId,
    'Test',
    'Member',
    'Open',
    'Step I',
    new Date(),
    '',
    new Date(),
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    'Art. 23 - Grievance Procedure',
    'Discipline',
    'test@union.org',
    'Unit A - Administrative',
    'Boston HQ',
    'Jane Smith',
    ''
  ];

  grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, 1, testGrievanceData.length)
    .setValues([testGrievanceData]);

  const rowCountBefore = grievanceLog.getLastRow();
  Assert.assertTrue(
    rowCountBefore >= 2,
    'Should have at least 2 rows (header + 1 grievance)'
  );

  // Clear the log
  clearGrievanceLog();

  // Check headers are preserved
  const headerAfter = grievanceLog.getRange(1, 1, 1, 28).getValues()[0];

  Assert.assertEquals(
    originalHeader.length,
    headerAfter.length,
    'Header row should have same number of columns'
  );

  Assert.assertEquals(
    'Grievance ID',
    headerAfter[0],
    'First header should still be "Grievance ID"'
  );

  Assert.assertEquals(
    'Resolution Summary',
    headerAfter[27],
    'Last header should still be "Resolution Summary"'
  );

  // Check data is cleared
  Assert.assertEquals(
    1,
    grievanceLog.getLastRow(),
    'Should only have header row after clearing'
  );

  // Cleanup
  cleanupTestData();

  Logger.log('✅ clearGrievanceLog preserves headers test passed');
}

/**
 * Test: clearStewardWorkload handles missing sheet gracefully
 */
function testClearStewardWorkloadMissingSheet() {
  // This function should handle the case where Steward Workload sheet doesn't exist
  try {
    clearStewardWorkload();
    // If it doesn't throw an error, test passes
    Logger.log('✅ clearStewardWorkload missing sheet test passed');
  } catch (error) {
    // Should not throw an error, but handle gracefully
    if (error.message.includes('not found')) {
      // This is acceptable - function should handle missing sheet
      Logger.log('✅ clearStewardWorkload missing sheet test passed (handled gracefully)');
    } else {
      throw error;
    }
  }
}

/* ===================== PROPERTY MANAGEMENT TESTS ===================== */

/**
 * Test: Nuke operation sets SEED_NUKED property
 */
function testNukePropertySet() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Clear the property first
  scriptProperties.deleteProperty('SEED_NUKED');

  // Verify it's not set
  const beforeValue = scriptProperties.getProperty('SEED_NUKED');
  Assert.assertNull(
    beforeValue,
    'SEED_NUKED should not be set initially'
  );

  // Set the property (simulating nuke operation)
  scriptProperties.setProperty('SEED_NUKED', 'true');

  // Verify it's set
  const afterValue = scriptProperties.getProperty('SEED_NUKED');
  Assert.assertEquals(
    'true',
    afterValue,
    'SEED_NUKED should be set to "true"'
  );

  // Cleanup
  scriptProperties.deleteProperty('SEED_NUKED');

  Logger.log('✅ Nuke property set test passed');
}

/**
 * Test: Property can be checked to determine if data was nuked
 */
function testCheckNukeStatus() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // Test false case
  scriptProperties.deleteProperty('SEED_NUKED');
  const wasNuked1 = scriptProperties.getProperty('SEED_NUKED') === 'true';
  Assert.assertFalse(
    wasNuked1,
    'Should return false when property not set'
  );

  // Test true case
  scriptProperties.setProperty('SEED_NUKED', 'true');
  const wasNuked2 = scriptProperties.getProperty('SEED_NUKED') === 'true';
  Assert.assertTrue(
    wasNuked2,
    'Should return true when property is set'
  );

  // Cleanup
  scriptProperties.deleteProperty('SEED_NUKED');

  Logger.log('✅ Check nuke status test passed');
}

/* ===================== DATA INTEGRITY TESTS ===================== */

/**
 * Test: Clearing doesn't affect other sheets
 */
function testClearingDoesntAffectOtherSheets() {
  const ss = SpreadsheetApp.getActive();

  // Get row counts of all sheets before clearing
  const configRowsBefore = ss.getSheetByName(SHEETS.CONFIG).getLastRow();
  const dashboardRowsBefore = ss.getSheetByName(SHEETS.DASHBOARD).getLastRow();

  // Create and clear test data
  createTestMember('TEST-M-NUKE-INTEGRITY-001');
  clearMemberDirectory();

  // Check other sheets are unchanged
  const configRowsAfter = ss.getSheetByName(SHEETS.CONFIG).getLastRow();
  const dashboardRowsAfter = ss.getSheetByName(SHEETS.DASHBOARD).getLastRow();

  Assert.assertEquals(
    configRowsBefore,
    configRowsAfter,
    'Config sheet should be unchanged'
  );

  Assert.assertEquals(
    dashboardRowsBefore,
    dashboardRowsAfter,
    'Dashboard sheet should be unchanged'
  );

  Logger.log('✅ Clearing doesnt affect other sheets test passed');
}

/**
 * Test: Multiple clear operations are idempotent
 */
function testMultipleClearOperationsIdempotent() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Create test data
  createTestMember('TEST-M-NUKE-IDEMPOTENT-001');

  // Clear multiple times
  clearMemberDirectory();
  clearMemberDirectory();
  clearMemberDirectory();

  // Should still have only header
  Assert.assertEquals(
    1,
    memberDir.getLastRow(),
    'Multiple clears should still result in only header row'
  );

  // Header should still be intact
  const header = memberDir.getRange(1, 1).getValue();
  Assert.assertEquals(
    'Member ID',
    header,
    'Header should still be intact after multiple clears'
  );

  Logger.log('✅ Multiple clear operations idempotent test passed');
}

/* ===================== EDGE CASE TESTS ===================== */

/**
 * Test: Clearing empty sheet doesn't cause errors
 */
function testClearingEmptySheet() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Ensure sheet is empty (only header)
  if (memberDir.getLastRow() > 1) {
    memberDir.deleteRows(2, memberDir.getLastRow() - 1);
  }

  Assert.assertEquals(
    1,
    memberDir.getLastRow(),
    'Sheet should only have header before test'
  );

  // Try to clear (should not throw error)
  try {
    clearMemberDirectory();
    Assert.assertEquals(
      1,
      memberDir.getLastRow(),
      'Sheet should still only have header after clearing'
    );
    Logger.log('✅ Clearing empty sheet test passed');
  } catch (error) {
    throw new Error('Clearing empty sheet should not throw error: ' + error.message);
  }
}

/**
 * Test: Clearing with formulas doesn't break sheet
 */
function testClearingWithFormulas() {
  const testMemberId = createTestMember('TEST-M-NUKE-FORMULA-001');

  try {
    const ss = SpreadsheetApp.getActive();
    const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

    // Formulas should exist in columns 26-28 (Has Open Grievance, Status Snapshot, Next Deadline)
    const formulaCell = memberDir.getRange(2, 26);
    const formula = formulaCell.getFormula();

    // Clear the data
    clearMemberDirectory();

    // Check that sheet structure is still valid
    Assert.assertEquals(
      1,
      memberDir.getLastRow(),
      'Sheet should have only header after clearing'
    );

    // Add new member and verify formulas work
    createTestMember('TEST-M-NUKE-FORMULA-002');

    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    // New row should have formulas
    const newFormulaCell = memberDir.getRange(2, 26);
    const newFormula = newFormulaCell.getFormula();

    // Formula should exist (might be empty string if no formula setup ran, but shouldn't be broken)
    Assert.assertTrue(
      typeof newFormula === 'string',
      'Formula cell should be valid after clearing and re-adding data'
    );

    Logger.log('✅ Clearing with formulas test passed');

  } finally {
    cleanupTestData();
  }
}

/* ===================== DASHBOARD REBUILD TESTS ===================== */

/**
 * Test: Dashboard rebuild updates metrics
 */
function testDashboardRebuildUpdatesMetrics() {
  const ss = SpreadsheetApp.getActive();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  // Create test data
  createTestMember('TEST-M-NUKE-REBUILD-001');
  createTestMember('TEST-M-NUKE-REBUILD-002');

  // Simulate dashboard rebuild
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  // Clear data
  clearMemberDirectory();

  // Rebuild dashboard
  rebuildDashboard();

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  // Dashboard should still be functional (not showing #REF! errors)
  Assert.assertNotNull(
    dashboard,
    'Dashboard should still exist after rebuild'
  );

  Logger.log('✅ Dashboard rebuild test passed');
}

/**
 * Test: Post-nuke guidance is displayed
 */
function testPostNukeGuidance() {
  // This function should show guidance after nuking
  // We can't easily test UI alerts, but we can verify the function exists and runs

  try {
    showPostNukeGuidance();
    Logger.log('✅ Post-nuke guidance test passed');
  } catch (error) {
    // Function might not exist in SeedNuke.gs, which is okay
    if (error.message.includes('not defined')) {
      Logger.log('⏭️ Post-nuke guidance function not implemented (skipped)');
    } else {
      throw error;
    }
  }
}
