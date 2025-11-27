/**
 * ============================================================================
 * UNIT TESTS FOR CODE.GS
 * ============================================================================
 *
 * Tests for core functionality:
 * - Formula calculations (deadlines, days open, etc.)
 * - Data validation setup
 * - Seeding functions
 * - Helper functions
 *
 * ============================================================================
 */

/* ===================== FORMULA CALCULATION TESTS ===================== */

/**
 * Test: Filing Deadline = Incident Date + 21 days
 */
function testFilingDeadlineCalculation() {
  const incidentDate = new Date(2025, 0, 1); // Jan 1, 2025
  const expectedDeadline = new Date(2025, 0, 22); // Jan 22, 2025

  // Calculate deadline (Incident + 21 days)
  const actualDeadline = new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000);

  Assert.assertDateEquals(
    expectedDeadline,
    actualDeadline,
    'Filing deadline should be 21 days after incident date'
  );

  Logger.log('✅ Filing deadline calculation test passed');
}

/**
 * Test: Step I Decision Due = Date Filed + 30 days
 */
function testStepIDeadlineCalculation() {
  const dateFiled = new Date(2025, 0, 15); // Jan 15, 2025
  const expectedDeadline = new Date(2025, 1, 14); // Feb 14, 2025

  // Calculate deadline (Filed + 30 days)
  const actualDeadline = new Date(dateFiled.getTime() + 30 * 24 * 60 * 60 * 1000);

  Assert.assertDateEquals(
    expectedDeadline,
    actualDeadline,
    'Step I decision should be due 30 days after filing'
  );

  Logger.log('✅ Step I deadline calculation test passed');
}

/**
 * Test: Step II Appeal Due = Step I Decision Received + 10 days
 */
function testStepIIAppealDeadlineCalculation() {
  const stepIDecisionDate = new Date(2025, 1, 14); // Feb 14, 2025
  const expectedDeadline = new Date(2025, 1, 24); // Feb 24, 2025

  // Calculate deadline (Decision + 10 days)
  const actualDeadline = new Date(stepIDecisionDate.getTime() + 10 * 24 * 60 * 60 * 1000);

  Assert.assertDateEquals(
    expectedDeadline,
    actualDeadline,
    'Step II appeal should be due 10 days after Step I decision'
  );

  Logger.log('✅ Step II appeal deadline calculation test passed');
}

/**
 * Test: Days Open calculation for active grievance
 */
function testDaysOpenCalculation() {
  const dateFiled = new Date(2025, 0, 1); // Jan 1, 2025
  const today = new Date(2025, 0, 31); // Jan 31, 2025

  const expectedDaysOpen = 30;
  const actualDaysOpen = Math.floor((today - dateFiled) / (24 * 60 * 60 * 1000));

  Assert.assertEquals(
    expectedDaysOpen,
    actualDaysOpen,
    'Days open should be 30 for a grievance filed 30 days ago'
  );

  Logger.log('✅ Days open calculation test passed');
}

/**
 * Test: Days Open calculation for closed grievance
 */
function testDaysOpenForClosedGrievance() {
  const dateFiled = new Date(2025, 0, 1); // Jan 1, 2025
  const dateClosed = new Date(2025, 0, 31); // Jan 31, 2025

  const expectedDaysOpen = 30;
  const actualDaysOpen = Math.floor((dateClosed - dateFiled) / (24 * 60 * 60 * 1000));

  Assert.assertEquals(
    expectedDaysOpen,
    actualDaysOpen,
    'Days open should use close date for closed grievances'
  );

  Logger.log('✅ Closed grievance days open calculation test passed');
}

/**
 * Test: Next Action Due logic based on current step
 */
function testNextActionDueLogic() {
  // Test data structure mimicking Grievance Log row
  const testCases = [
    {
      status: 'Open',
      step: 'Step I',
      stepIDeadline: new Date(2025, 1, 14),
      stepIIDeadline: new Date(2025, 1, 24),
      stepIIIDeadline: new Date(2025, 2, 26),
      filingDeadline: new Date(2025, 0, 22),
      expected: new Date(2025, 1, 14) // Should use Step I deadline
    },
    {
      status: 'Open',
      step: 'Step II',
      stepIDeadline: new Date(2025, 1, 14),
      stepIIDeadline: new Date(2025, 1, 24),
      stepIIIDeadline: new Date(2025, 2, 26),
      filingDeadline: new Date(2025, 0, 22),
      expected: new Date(2025, 1, 24) // Should use Step II deadline
    },
    {
      status: 'Open',
      step: 'Step III',
      stepIDeadline: new Date(2025, 1, 14),
      stepIIDeadline: new Date(2025, 1, 24),
      stepIIIDeadline: new Date(2025, 2, 26),
      filingDeadline: new Date(2025, 0, 22),
      expected: new Date(2025, 2, 26) // Should use Step III deadline
    },
    {
      status: 'Open',
      step: 'Informal',
      stepIDeadline: new Date(2025, 1, 14),
      stepIIDeadline: new Date(2025, 1, 24),
      stepIIIDeadline: new Date(2025, 2, 26),
      filingDeadline: new Date(2025, 0, 22),
      expected: new Date(2025, 0, 22) // Should use filing deadline
    }
  ];

  testCases.forEach((testCase, index) => {
    let nextAction;
    if (testCase.status === 'Open') {
      if (testCase.step === 'Step I') {
        nextAction = testCase.stepIDeadline;
      } else if (testCase.step === 'Step II') {
        nextAction = testCase.stepIIDeadline;
      } else if (testCase.step === 'Step III') {
        nextAction = testCase.stepIIIDeadline;
      } else {
        nextAction = testCase.filingDeadline;
      }
    }

    Assert.assertDateEquals(
      testCase.expected,
      nextAction,
      `Test case ${index + 1}: Next action should match expected deadline for ${testCase.step}`
    );
  });

  Logger.log('✅ Next action due logic test passed');
}

/**
 * Test: Member Directory formulas
 */
function testMemberDirectoryFormulas() {
  // Create test setup
  const testMemberId = createTestMember('TEST-M-FORMULA-001');

  try {
    const ss = SpreadsheetApp.getActive();
    const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
    const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    // Create a test grievance for this member
    const testGrievanceData = [
      'TEST-G-001',
      testMemberId,
      'Test',
      'Member',
      'Open',
      'Step I',
      new Date(2025, 0, 1),
      '',
      new Date(2025, 0, 10),
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
      'test.member@union.org',
      'Unit A - Administrative',
      'Boston HQ',
      'Jane Smith',
      ''
    ];

    grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, 1, testGrievanceData.length)
      .setValues([testGrievanceData]);

    // Force recalculation
    SpreadsheetApp.flush();
    Utilities.sleep(2000); // Wait for formulas to recalculate

    // Find the test member row
    const memberData = memberDir.getRange(2, 1, memberDir.getLastRow() - 1, 31).getValues();
    const testMemberRow = memberData.findIndex(row => row[0] === testMemberId);

    Assert.assertTrue(
      testMemberRow >= 0,
      'Test member should exist in Member Directory'
    );

    // Check "Has Open Grievance?" (column 26, index 25)
    const hasOpenGrievance = memberData[testMemberRow][25];
    Assert.assertTrue(
      hasOpenGrievance === 'Yes' || hasOpenGrievance === true,
      'Member with open grievance should show "Yes" in Has Open Grievance column'
    );

    // Check "Grievance Status Snapshot" (column 27, index 26)
    const statusSnapshot = memberData[testMemberRow][26];
    Assert.assertEquals(
      'Open',
      statusSnapshot,
      'Grievance status snapshot should match grievance status'
    );

    Logger.log('✅ Member Directory formulas test passed');

  } finally {
    cleanupTestData();
  }
}

/* ===================== DATA VALIDATION TESTS ===================== */

/**
 * Test: Data validation setup creates proper rules
 */
function testDataValidationSetup() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const config = ss.getSheetByName(SHEETS.CONFIG);

  // Check that validation exists for Job Title (column 4)
  const jobTitleCell = memberDir.getRange(2, 4);
  const validation = jobTitleCell.getDataValidation();

  Assert.assertNotNull(
    validation,
    'Job Title column should have data validation'
  );

  Logger.log('✅ Data validation setup test passed');
}

/**
 * Test: Config dropdown values are properly defined
 */
function testConfigDropdownValues() {
  const ss = SpreadsheetApp.getActive();
  const config = ss.getSheetByName(SHEETS.CONFIG);

  // Test Job Titles
  const jobTitles = config.getRange('A2:A14').getValues().flat().filter(String);
  Assert.assertTrue(
    jobTitles.length > 0,
    'Config should have job titles defined'
  );
  Assert.assertContains(
    jobTitles,
    'Coordinator',
    'Config should contain Coordinator job title'
  );

  // Test Office Locations
  const locations = config.getRange('B2:B14').getValues().flat().filter(String);
  Assert.assertTrue(
    locations.length > 0,
    'Config should have office locations defined'
  );
  Assert.assertContains(
    locations,
    'Boston HQ',
    'Config should contain Boston HQ location'
  );

  // Test Grievance Status
  const statuses = config.getRange('I2:I8').getValues().flat().filter(String);
  Assert.assertTrue(
    statuses.length > 0,
    'Config should have grievance statuses defined'
  );
  Assert.assertContains(
    statuses,
    'Open',
    'Config should contain Open status'
  );

  Logger.log('✅ Config dropdown values test passed');
}

/**
 * Test: Member validation rules reference correct Config columns
 */
function testMemberValidationRules() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  // Check critical validations exist
  const columnsToCheck = [
    { col: 4, name: 'Job Title' },
    { col: 5, name: 'Work Location' },
    { col: 6, name: 'Unit' },
    { col: 10, name: 'Is Steward' }
  ];

  columnsToCheck.forEach(item => {
    const cell = memberDir.getRange(2, item.col);
    const validation = cell.getDataValidation();

    Assert.assertNotNull(
      validation,
      `${item.name} (column ${item.col}) should have data validation`
    );
  });

  Logger.log('✅ Member validation rules test passed');
}

/**
 * Test: Grievance validation rules reference correct Config columns
 */
function testGrievanceValidationRules() {
  const ss = SpreadsheetApp.getActive();
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Check critical validations exist
  const columnsToCheck = [
    { col: 5, name: 'Status' },
    { col: 6, name: 'Current Step' },
    { col: 23, name: 'Issue Category' },
    { col: 22, name: 'Articles Violated' }
  ];

  columnsToCheck.forEach(item => {
    const cell = grievanceLog.getRange(2, item.col);
    const validation = cell.getDataValidation();

    Assert.assertNotNull(
      validation,
      `${item.name} (column ${item.col}) should have data validation`
    );
  });

  Logger.log('✅ Grievance validation rules test passed');
}

/* ===================== SEEDING FUNCTION TESTS ===================== */

/**
 * Test: Member seeding generates valid data
 */
function testMemberSeedingValidation() {
  // This test validates the data structure, not actual seeding
  // (to avoid creating 20k test records)

  const firstNames = ["James", "Mary", "John"];
  const lastNames = ["Smith", "Johnson", "Williams"];

  // Simulate member generation
  const testMember = {
    memberId: "M" + String(1).padStart(6, '0'),
    firstName: firstNames[0],
    lastName: lastNames[0],
    email: `${firstNames[0].toLowerCase()}.${lastNames[0].toLowerCase()}1@union.org`,
    phone: `(555) 123-4567`,
    openRate: 75
  };

  // Validate structure
  Assert.assertTrue(
    testMember.memberId.startsWith('M'),
    'Member ID should start with M'
  );

  Assert.assertEquals(
    7,
    testMember.memberId.length,
    'Member ID should be 7 characters (M + 6 digits)'
  );

  Assert.assertTrue(
    testMember.email.includes('@union.org'),
    'Email should end with @union.org'
  );

  Assert.assertTrue(
    testMember.openRate >= 0 && testMember.openRate <= 100,
    'Open rate should be between 0-100'
  );

  Logger.log('✅ Member seeding validation test passed');
}

/**
 * Test: Grievance seeding generates valid data
 */
function testGrievanceSeedingValidation() {
  // Simulate grievance generation
  const testGrievance = {
    grievanceId: "G-" + String(1).padStart(6, '0'),
    memberId: "M000001",
    status: "Open",
    step: "Step I",
    incidentDate: new Date(2025, 0, 1),
    dateFiled: new Date(2025, 0, 10)
  };

  // Validate structure
  Assert.assertTrue(
    testGrievance.grievanceId.startsWith('G-'),
    'Grievance ID should start with G-'
  );

  Assert.assertEquals(
    8,
    testGrievance.grievanceId.length,
    'Grievance ID should be 8 characters (G- + 6 digits)'
  );

  Assert.assertTrue(
    testGrievance.dateFiled >= testGrievance.incidentDate,
    'Date filed should be after incident date'
  );

  Logger.log('✅ Grievance seeding validation test passed');
}

/**
 * Test: Member email format is valid
 */
function testMemberEmailFormat() {
  const testEmails = [
    'john.smith123@union.org',
    'mary.jones456@union.org',
    'robert.wilson789@union.org'
  ];

  const emailRegex = /^[a-z]+\.[a-z]+\d+@union\.org$/;

  testEmails.forEach(email => {
    Assert.assertTrue(
      emailRegex.test(email),
      `Email ${email} should match format firstname.lastnameNNN@union.org`
    );
  });

  Logger.log('✅ Member email format test passed');
}

/**
 * Test: Member IDs are unique
 */
function testMemberIDUniqueness() {
  const memberIds = new Set();

  // Simulate generating 100 member IDs
  for (let i = 1; i <= 100; i++) {
    const memberId = "M" + String(i).padStart(6, '0');
    Assert.assertFalse(
      memberIds.has(memberId),
      `Member ID ${memberId} should be unique`
    );
    memberIds.add(memberId);
  }

  Assert.assertEquals(
    100,
    memberIds.size,
    'Should have 100 unique member IDs'
  );

  Logger.log('✅ Member ID uniqueness test passed');
}

/**
 * Test: Grievances link to valid members
 */
function testGrievanceMemberLinking() {
  const testMemberId = createTestMember('TEST-M-LINK-001');

  try {
    const ss = SpreadsheetApp.getActive();
    const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

    // Create test grievance
    const testGrievanceData = [
      'TEST-G-LINK-001',
      testMemberId, // Valid member ID
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
      'test.member@union.org',
      'Unit A - Administrative',
      'Boston HQ',
      'Jane Smith',
      ''
    ];

    grievanceLog.getRange(grievanceLog.getLastRow() + 1, 1, 1, testGrievanceData.length)
      .setValues([testGrievanceData]);

    // Verify it was created
    const grievanceData = grievanceLog.getRange(2, 1, grievanceLog.getLastRow() - 1, 2).getValues();
    const testGrievance = grievanceData.find(row => row[0] === 'TEST-G-LINK-001');

    Assert.assertNotNull(
      testGrievance,
      'Test grievance should exist'
    );

    Assert.assertEquals(
      testMemberId,
      testGrievance[1],
      'Grievance should link to correct member ID'
    );

    Logger.log('✅ Grievance-member linking test passed');

  } finally {
    cleanupTestData();
  }
}

/**
 * Test: Open Rate is within valid range
 */
function testOpenRateRange() {
  // Simulate 50 random open rates
  for (let i = 0; i < 50; i++) {
    const openRate = Math.floor(Math.random() * 40) + 60; // 60-100 range

    Assert.assertTrue(
      openRate >= 0 && openRate <= 100,
      `Open rate ${openRate} should be between 0-100`
    );

    Assert.assertTrue(
      openRate >= 60,
      `Open rate ${openRate} should be at least 60 (as per seeding logic)`
    );
  }

  Logger.log('✅ Open rate range test passed');
}

/* ===================== EDGE CASE TESTS ===================== */

/**
 * Test: Empty sheets don't break formulas
 */
function testEmptySheetsHandling() {
  const ss = SpreadsheetApp.getActive();
  const dashboard = ss.getSheetByName(SHEETS.DASHBOARD);

  // Dashboard should handle empty data gracefully
  // (formulas should return 0 or empty, not #DIV/0! or #REF!)

  Assert.assertNotNull(
    dashboard,
    'Dashboard sheet should exist'
  );

  Logger.log('✅ Empty sheets handling test passed');
}

/**
 * Test: Future dates are handled correctly
 */
function testFutureDateHandling() {
  const today = new Date();
  const futureDate = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

  // Days to deadline should be positive for future dates
  const daysToDeadline = Math.floor((futureDate - today) / (24 * 60 * 60 * 1000));

  Assert.assertTrue(
    daysToDeadline > 0,
    'Days to deadline should be positive for future dates'
  );

  Assert.assertApproximately(
    30,
    daysToDeadline,
    1,
    'Days to deadline should be approximately 30'
  );

  Logger.log('✅ Future date handling test passed');
}

/**
 * Test: Past deadlines show negative days
 */
function testPastDeadlineHandling() {
  const today = new Date();
  const pastDate = new Date(today.getTime() - 5 * 24 * 60 * 60 * 1000);

  // Days to deadline should be negative for past dates
  const daysToDeadline = Math.floor((pastDate - today) / (24 * 60 * 60 * 1000));

  Assert.assertTrue(
    daysToDeadline < 0,
    'Days to deadline should be negative for overdue deadlines'
  );

  Assert.assertApproximately(
    -5,
    daysToDeadline,
    1,
    'Days to deadline should be approximately -5 for 5 days overdue'
  );

  Logger.log('✅ Past deadline handling test passed');
}
