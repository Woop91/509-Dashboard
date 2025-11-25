/**
 * ============================================================================
 * UNIT TESTS FOR GRIEVANCEWORKFLOW.GS
 * ============================================================================
 *
 * Tests for grievance workflow functionality:
 * - Member list retrieval
 * - Member selection dialog
 * - Form pre-filling
 * - Error handling
 *
 * ============================================================================
 */

/* ===================== MEMBER LIST TESTS ===================== */

/**
 * Test: getMemberList returns all members
 */
function testGetMemberList() {
  // Create test members
  createTestMember('TEST-M-WF-001');
  createTestMember('TEST-M-WF-002');
  createTestMember('TEST-M-WF-003');

  try {
    const members = getMemberList();

    Assert.assertTrue(
      members.length >= 3,
      'getMemberList should return at least 3 members (test members)'
    );

    // Check that our test members are in the list
    const testMember = members.find(m => m.memberId === 'TEST-M-WF-001');

    Assert.assertNotNull(
      testMember,
      'Test member should be in member list'
    );

    // Verify member structure
    Assert.assertNotNull(
      testMember.firstName,
      'Member should have firstName'
    );

    Assert.assertNotNull(
      testMember.lastName,
      'Member should have lastName'
    );

    Assert.assertNotNull(
      testMember.email,
      'Member should have email'
    );

    Assert.assertNotNull(
      testMember.location,
      'Member should have location'
    );

    Logger.log('✅ getMemberList test passed');

  } finally {
    cleanupTestData();
  }
}

/**
 * Test: getMemberList returns empty array when no members exist
 */
function testGetMemberListEmpty() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Backup current data
  const backup = memberDir.getRange(2, 1, memberDir.getLastRow() - 1, 31).getValues();

  try {
    // Clear all members (keep header)
    if (memberDir.getLastRow() > 1) {
      memberDir.deleteRows(2, memberDir.getLastRow() - 1);
    }

    const members = getMemberList();

    Assert.assertArrayLength(
      members,
      0,
      'getMemberList should return empty array when no members exist'
    );

    Logger.log('✅ getMemberList empty test passed');

  } finally {
    // Restore backup
    if (backup.length > 0) {
      memberDir.getRange(2, 1, backup.length, backup[0].length).setValues(backup);
    }
  }
}

/**
 * Test: getMemberList filters out empty rows
 */
function testGetMemberListFiltersEmptyRows() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Create test member
  createTestMember('TEST-M-WF-FILTER-001');

  // Add an empty row (all empty cells)
  const emptyRow = new Array(31).fill('');
  memberDir.getRange(memberDir.getLastRow() + 1, 1, 1, 31).setValues([emptyRow]);

  try {
    const members = getMemberList();

    // Check that empty rows are filtered out
    const emptyMembers = members.filter(m => !m.memberId || m.memberId === '');

    Assert.assertArrayLength(
      emptyMembers,
      0,
      'getMemberList should filter out rows with no member ID'
    );

    Logger.log('✅ getMemberList filters empty rows test passed');

  } finally {
    cleanupTestData();
    // Remove the empty row
    const data = memberDir.getRange(2, 1, memberDir.getLastRow() - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (!data[i][0] || data[i][0] === '') {
        memberDir.deleteRow(i + 2);
      }
    }
  }
}

/**
 * Test: getMemberList includes all required fields
 */
function testGetMemberListFieldCompleteness() {
  createTestMember('TEST-M-WF-FIELDS-001');

  try {
    const members = getMemberList();
    const testMember = members.find(m => m.memberId === 'TEST-M-WF-FIELDS-001');

    Assert.assertNotNull(testMember, 'Test member should exist');

    // Check all required fields are present
    const requiredFields = [
      'rowIndex',
      'memberId',
      'firstName',
      'lastName',
      'jobTitle',
      'location',
      'unit',
      'email',
      'phone'
    ];

    requiredFields.forEach(field => {
      Assert.assertTrue(
        testMember.hasOwnProperty(field),
        `Member should have ${field} property`
      );
    });

    Logger.log('✅ getMemberList field completeness test passed');

  } finally {
    cleanupTestData();
  }
}

/* ===================== DIALOG GENERATION TESTS ===================== */

/**
 * Test: Member selection dialog generates valid HTML
 */
function testMemberSelectionDialog() {
  createTestMember('TEST-M-WF-DIALOG-001');

  try {
    const members = getMemberList();
    const html = createMemberSelectionDialog(members);

    // Check that HTML contains required elements
    Assert.assertTrue(
      html.includes('<!DOCTYPE html>'),
      'Dialog should contain valid HTML doctype'
    );

    Assert.assertTrue(
      html.includes('<select'),
      'Dialog should contain a select element'
    );

    Assert.assertTrue(
      html.includes('TEST-M-WF-DIALOG-001'),
      'Dialog should include test member in options'
    );

    Assert.assertTrue(
      html.includes('<button'),
      'Dialog should contain buttons'
    );

    Logger.log('✅ Member selection dialog test passed');

  } finally {
    cleanupTestData();
  }
}

/**
 * Test: Dialog handles multiple members correctly
 */
function testMemberSelectionDialogMultipleMembers() {
  createTestMember('TEST-M-WF-MULTI-001');
  createTestMember('TEST-M-WF-MULTI-002');
  createTestMember('TEST-M-WF-MULTI-003');

  try {
    const members = getMemberList();
    const testMembers = members.filter(m => m.memberId.startsWith('TEST-M-WF-MULTI'));

    Assert.assertTrue(
      testMembers.length >= 3,
      'Should have at least 3 test members'
    );

    const html = createMemberSelectionDialog(members);

    // Check that all test members appear in the HTML
    testMembers.forEach(member => {
      Assert.assertTrue(
        html.includes(member.memberId),
        `Dialog should include member ${member.memberId}`
      );
    });

    Logger.log('✅ Dialog multiple members test passed');

  } finally {
    cleanupTestData();
  }
}

/**
 * Test: Dialog handles special characters in member names
 */
function testMemberSelectionDialogSpecialCharacters() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Create member with special characters
  const specialMemberData = [
    'TEST-M-SPECIAL-001',
    "O'Brien",
    'García',
    'Coordinator',
    'Boston HQ',
    'Unit A - Administrative',
    'Monday',
    'obrien.garcia@union.org',
    '(555) 123-4567',
    'No',
    'Sarah Johnson',
    'Michael Chen',
    'Jane Smith',
    new Date(),
    new Date(),
    new Date(),
    new Date(),
    85,
    10,
    'Yes',
    'Yes',
    'No',
    new Date(),
    'Email',
    'Mornings',
    'No',
    '',
    '',
    '',
    '',
    ''
  ];

  memberDir.getRange(memberDir.getLastRow() + 1, 1, 1, specialMemberData.length)
    .setValues([specialMemberData]);

  try {
    const members = getMemberList();
    const html = createMemberSelectionDialog(members);

    // Should handle apostrophes and accents
    Assert.assertTrue(
      html.includes('TEST-M-SPECIAL-001'),
      'Dialog should include member with special characters'
    );

    Logger.log('✅ Dialog special characters test passed');

  } finally {
    cleanupTestData();
  }
}

/* ===================== ERROR HANDLING TESTS ===================== */

/**
 * Test: showStartGrievanceDialog handles missing Member Directory
 */
function testStartGrievanceDialogMissingSheet() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Temporarily rename the sheet
  const originalName = memberDir.getName();
  memberDir.setName('TEMP_HIDDEN_SHEET');

  try {
    // This should handle the error gracefully
    // We can't easily test UI alerts, but we can verify the function doesn't crash
    const result = getMemberList();

    Assert.assertArrayLength(
      result,
      0,
      'getMemberList should return empty array when sheet is missing'
    );

    Logger.log('✅ Missing sheet error handling test passed');

  } finally {
    // Restore sheet name
    const tempSheet = ss.getSheetByName('TEMP_HIDDEN_SHEET');
    if (tempSheet) {
      tempSheet.setName(originalName);
    }
  }
}

/* ===================== FORM CONFIGURATION TESTS ===================== */

/**
 * Test: Grievance form configuration is defined
 */
function testGrievanceFormConfiguration() {
  Assert.assertNotNull(
    GRIEVANCE_FORM_CONFIG,
    'Grievance form config should be defined'
  );

  Assert.assertNotNull(
    GRIEVANCE_FORM_CONFIG.FORM_URL,
    'Form URL should be defined'
  );

  Assert.assertNotNull(
    GRIEVANCE_FORM_CONFIG.FIELD_IDS,
    'Field IDs should be defined'
  );

  // Check required field IDs exist
  const requiredFields = [
    'MEMBER_ID',
    'MEMBER_FIRST_NAME',
    'MEMBER_LAST_NAME',
    'MEMBER_EMAIL',
    'STEWARD_NAME'
  ];

  requiredFields.forEach(field => {
    Assert.assertNotNull(
      GRIEVANCE_FORM_CONFIG.FIELD_IDS[field],
      `Field ID for ${field} should be defined`
    );
  });

  Logger.log('✅ Grievance form configuration test passed');
}

/* ===================== INTEGRATION TESTS ===================== */

/**
 * Test: Complete workflow from member selection to grievance creation
 */
function testCompleteWorkflowMemberToGrievance() {
  const testMemberId = createTestMember('TEST-M-WF-COMPLETE-001');

  try {
    const ss = SpreadsheetApp.getActive();
    const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
    const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    // Step 1: Verify member exists
    const members = getMemberList();
    const testMember = members.find(m => m.memberId === testMemberId);
    Assert.assertNotNull(testMember, 'Member should exist');

    // Step 2: Simulate grievance creation
    const testGrievanceData = [
      'TEST-G-WF-001',
      testMemberId,
      testMember.firstName,
      testMember.lastName,
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
      testMember.email,
      testMember.unit,
      testMember.location,
      'Jane Smith',
      ''
    ];

    grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, 1, testGrievanceData.length)
      .setValues([testGrievanceData]);

    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    // Step 3: Verify grievance was created
    const grievanceData = grievanceLog.getRange(2, 1, grievanceLog.getLastRow() - 1, 2).getValues();
    const testGrievance = grievanceData.find(row => row[0] === 'TEST-G-WF-001');

    Assert.assertNotNull(testGrievance, 'Grievance should be created');
    Assert.assertEquals(
      testMemberId,
      testGrievance[1],
      'Grievance should link to correct member'
    );

    // Step 4: Verify member directory updates
    Utilities.sleep(2000); // Wait for formulas
    const memberData = memberDir.getRange(2, 1, memberDir.getLastRow() - 1, 31).getValues();
    const updatedMember = memberData.find(row => row[0] === testMemberId);

    Assert.assertNotNull(updatedMember, 'Member should still exist');

    Logger.log('✅ Complete workflow test passed');

  } finally {
    cleanupTestData();
  }
}
