/**
 * ------------------------------------------------------------------------====
 * TEST FRAMEWORK - Simple Testing Library for Google Apps Script
 * ------------------------------------------------------------------------====
 *
 * A lightweight testing framework that runs within the Apps Script environment.
 * Provides assertion methods, test runners, and reporting.
 *
 * Usage:
 *   1. Write test functions (see tests/*.test.gs)
 *   2. Run via menu: 🧪 Tests > Run All Tests
 *   3. View results in test report sheet
 *
 * ------------------------------------------------------------------------====
 */

// Test results storage
var TEST_RESULTS = {
  passed: [],
  failed: [],
  skipped: []
};

/**
 * Assertion library
 */
var Assert = {
  /**
   * Assert that two values are equal
   */
  assertEquals: function(expected, actual, message) {
    if (expected !== actual) {
      throw new Error(
        (message || 'Assertion failed') +
        `\nExpected: ${JSON.stringify(expected)}` +
        `\nActual: ${JSON.stringify(actual)}`
      );
    }
  },

  /**
   * Assert that value is true
   */
  assertTrue: function(value, message) {
    if (value !== true) {
      throw new Error(
        (message || 'Expected true') +
        `\nActual: ${JSON.stringify(value)}`
      );
    }
  },

  /**
   * Assert that value is false
   */
  assertFalse: function(value, message) {
    if (value !== false) {
      throw new Error(
        (message || 'Expected false') +
        `\nActual: ${JSON.stringify(value)}`
      );
    }
  },

  /**
   * Assert that value is not null or undefined
   */
  assertNotNull: function(value, message) {
    if (value === null || value === undefined) {
      throw new Error(message || 'Value should not be null or undefined');
    }
  },

  /**
   * Assert that value is null
   */
  assertNull: function(value, message) {
    if (value !== null) {
      throw new Error(
        (message || 'Expected null') +
        `\nActual: ${JSON.stringify(value)}`
      );
    }
  },

  /**
   * Assert that array contains value
   */
  assertContains: function(array, value, message) {
    if (!Array.isArray(array)) {
      throw new Error('First argument must be an array');
    }
    if (array.indexOf(value) === -1) {
      throw new Error(
        (message || 'Array does not contain value') +
        `\nArray: ${JSON.stringify(array)}` +
        `\nValue: ${JSON.stringify(value)}`
      );
    }
  },

  /**
   * Assert that array has specific length
   */
  assertArrayLength: function(array, expectedLength, message) {
    if (!Array.isArray(array)) {
      throw new Error('First argument must be an array');
    }
    if (array.length !== expectedLength) {
      throw new Error(
        (message || 'Array length mismatch') +
        `\nExpected length: ${expectedLength}` +
        `\nActual length: ${array.length}`
      );
    }
  },

  /**
   * Assert that function throws an error
   */
  assertThrows: function(fn, message) {
    let threw = false;
    try {
      fn();
    } catch (e) {
      threw = true;
    }
    if (!threw) {
      throw new Error(message || 'Expected function to throw an error');
    }
  },

  /**
   * Assert that two values are approximately equal (for floating point)
   */
  assertApproximately: function(expected, actual, tolerance, message) {
    tolerance = tolerance || 0.001;
    if (Math.abs(expected - actual) > tolerance) {
      throw new Error(
        (message || 'Values not approximately equal') +
        `\nExpected: ${expected}` +
        `\nActual: ${actual}` +
        `\nTolerance: ${tolerance}`
      );
    }
  },

  /**
   * Assert that date is within range
   */
  assertDateEquals: function(expected, actual, message) {
    const expectedTime = expected instanceof Date ? expected.getTime() : new Date(expected).getTime();
    const actualTime = actual instanceof Date ? actual.getTime() : new Date(actual).getTime();

    if (expectedTime !== actualTime) {
      throw new Error(
        (message || 'Dates not equal') +
        `\nExpected: ${new Date(expectedTime).toISOString()}` +
        `\nActual: ${new Date(actualTime).toISOString()}`
      );
    }
  }
};

/**
 * Test runner - discovers and runs all test functions
 */
function runAllTests() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '🧪 Running All Tests',
    'This will run the complete test suite. This may take 2-3 minutes.\n\n' +
    'Results will be displayed in a new "Test Results" sheet.',
    ui.ButtonSet.OK
  );

  SpreadsheetApp.getActive().toast('🧪 Running test suite...', 'Testing', -1);

  // Clear previous results
  TEST_RESULTS.passed = [];
  TEST_RESULTS.failed = [];
  TEST_RESULTS.skipped = [];

  const startTime = new Date();

  // Discover and run all test functions
  const testFunctions = [
    // Code.gs tests
    'testFilingDeadlineCalculation',
    'testStepIDeadlineCalculation',
    'testStepIIAppealDeadlineCalculation',
    'testDaysOpenCalculation',
    'testNextActionDueLogic',
    'testMemberDirectoryFormulas',
    'testDataValidationSetup',
    'testConfigDropdownValues',
    'testMemberValidationRules',
    'testGrievanceValidationRules',

    // Seeding tests
    'testMemberSeedingValidation',
    'testGrievanceSeedingValidation',
    'testMemberEmailFormat',
    'testMemberIDUniqueness',
    'testGrievanceMemberLinking',
    'testOpenRateRange',

    // GrievanceWorkflow tests
    'testGetMemberList',
    'testGetMemberListEmpty',
    'testGetMemberListFiltersEmptyRows',
    'testMemberSelectionDialog',

    // SeedNuke tests
    'testClearMemberDirectoryPreservesHeaders',
    'testClearGrievanceLogPreservesHeaders',
    'testNukePropertySet',

    // Integration tests
    'testCompleteGrievanceWorkflow',
    'testDashboardMetricsUpdate',
    'testMemberGrievanceSnapshot'
  ];

  // Run each test
  testFunctions.forEach(testName => {
    try {
      const testFn = this[testName];
      if (typeof testFn === 'function') {
        testFn();
        TEST_RESULTS.passed.push({
          name: testName,
          time: new Date() - startTime
        });
      } else {
        TEST_RESULTS.skipped.push({
          name: testName,
          reason: 'Function not found'
        });
      }
    } catch (error) {
      TEST_RESULTS.failed.push({
        name: testName,
        error: error.message,
        stack: error.stack
      });
    }
  });

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  // Generate test report
  generateTestReport(duration);

  // Show summary
  const total = TEST_RESULTS.passed.length + TEST_RESULTS.failed.length + TEST_RESULTS.skipped.length;
  const passRate = ((TEST_RESULTS.passed.length / total) * 100).toFixed(1);

  SpreadsheetApp.getActive().toast(
    `✅ ${TEST_RESULTS.passed.length} passed | ❌ ${TEST_RESULTS.failed.length} failed | ⏭️ ${TEST_RESULTS.skipped.length} skipped`,
    `Tests Complete (${passRate}% pass rate)`,
    10
  );

  // Show detailed results dialog
  ui.alert(
    '🧪 Test Suite Complete',
    `Results:\n\n` +
    `✅ Passed: ${TEST_RESULTS.passed.length}\n` +
    `❌ Failed: ${TEST_RESULTS.failed.length}\n` +
    `⏭️ Skipped: ${TEST_RESULTS.skipped.length}\n\n` +
    `Total: ${total} tests\n` +
    `Pass Rate: ${passRate}%\n` +
    `Duration: ${duration.toFixed(2)}s\n\n` +
    `View detailed results in the "Test Results" sheet.`,
    ui.ButtonSet.OK
  );
}

/**
 * Generates a detailed test report in a new sheet
 */
function generateTestReport(duration) {
  const ss = SpreadsheetApp.getActive();

  // Create or clear Test Results sheet
  let reportSheet = ss.getSheetByName('Test Results');
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Test Results');
  }
  reportSheet.clear();

  // Header
  reportSheet.getRange('A1:F1').merge()
    .setValue('🧪 TEST RESULTS')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4A5568')
    .setFontColor('#FFFFFF');

  // Summary
  const total = TEST_RESULTS.passed.length + TEST_RESULTS.failed.length + TEST_RESULTS.skipped.length;
  const passRate = ((TEST_RESULTS.passed.length / total) * 100).toFixed(1);

  const summary = [
    ['Total Tests', total],
    ['✅ Passed', TEST_RESULTS.passed.length],
    ['❌ Failed', TEST_RESULTS.failed.length],
    ['⏭️ Skipped', TEST_RESULTS.skipped.length],
    ['Pass Rate', `${passRate}%`],
    ['Duration', `${duration.toFixed(2)}s`],
    ['Timestamp', new Date().toLocaleString()]
  ];

  reportSheet.getRange(3, 1, summary.length, 2).setValues(summary);
  reportSheet.getRange(3, 1, summary.length, 1).setFontWeight('bold');

  let currentRow = 3 + summary.length + 2;

  // Passed tests
  if (TEST_RESULTS.passed.length > 0) {
    reportSheet.getRange(currentRow, 1, 1, 3).merge()
      .setValue('✅ PASSED TESTS')
      .setFontWeight('bold')
      .setBackground('#D1FAE5')
      .setFontColor('#065F46');

    currentRow++;
    reportSheet.getRange(currentRow, 1, 1, 3).setValues([['Test Name', 'Status', 'Duration (ms)']])
      .setFontWeight('bold')
      .setBackground('#F3F4F6');

    currentRow++;
    TEST_RESULTS.passed.forEach(test => {
      reportSheet.getRange(currentRow, 1, 1, 3).setValues([[test.name, '✅ PASS', test.time]]);
      currentRow++;
    });
    currentRow += 2;
  }

  // Failed tests
  if (TEST_RESULTS.failed.length > 0) {
    reportSheet.getRange(currentRow, 1, 1, 4).merge()
      .setValue('❌ FAILED TESTS')
      .setFontWeight('bold')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B');

    currentRow++;
    reportSheet.getRange(currentRow, 1, 1, 4).setValues([['Test Name', 'Status', 'Error', 'Stack Trace']])
      .setFontWeight('bold')
      .setBackground('#F3F4F6');

    currentRow++;
    TEST_RESULTS.failed.forEach(test => {
      reportSheet.getRange(currentRow, 1, 1, 4).setValues([[
        test.name,
        '❌ FAIL',
        test.error,
        test.stack || 'N/A'
      ]]);
      reportSheet.getRange(currentRow, 3).setWrap(true);
      currentRow++;
    });
    currentRow += 2;
  }

  // Skipped tests
  if (TEST_RESULTS.skipped.length > 0) {
    reportSheet.getRange(currentRow, 1, 1, 3).merge()
      .setValue('⏭️ SKIPPED TESTS')
      .setFontWeight('bold')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E');

    currentRow++;
    reportSheet.getRange(currentRow, 1, 1, 3).setValues([['Test Name', 'Status', 'Reason']])
      .setFontWeight('bold')
      .setBackground('#F3F4F6');

    currentRow++;
    TEST_RESULTS.skipped.forEach(test => {
      reportSheet.getRange(currentRow, 1, 1, 3).setValues([[test.name, '⏭️ SKIP', test.reason]]);
      currentRow++;
    });
  }

  // Auto-resize columns
  reportSheet.autoResizeColumns(1, 4);
  reportSheet.setColumnWidth(3, 400);
  reportSheet.setColumnWidth(4, 300);

  reportSheet.setTabColor('#7C3AED');
  reportSheet.activate();
}

/**
 * Run a single test by name
 */
function runSingleTest(testName) {
  try {
    const testFn = this[testName];
    if (typeof testFn !== 'function') {
      throw new Error(`Test function '${testName}' not found`);
    }

    testFn();
    Logger.log(`✅ ${testName} PASSED`);
    return true;
  } catch (error) {
    Logger.log(`❌ ${testName} FAILED: ${error.message}`);
    Logger.log(error.stack);
    return false;
  }
}

/**
 * Test helper: Create a test member in Member Directory
 */
function createTestMember(memberId) {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  const testMemberData = [
    memberId || 'TEST-M001',
    'Test',
    'Member',
    'Coordinator',
    'Boston HQ',
    'Unit A - Administrative',
    'Monday',
    'test.member@union.org',
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

  memberDir.getRange(memberDir.getLastRow() + 1, 1, 1, testMemberData.length).setValues([testMemberData]);
  return memberId || 'TEST-M001';
}

/**
 * Test helper: Clean up test data
 */
function cleanupTestData() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);
  const grievanceLog = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  // Remove all rows starting with "TEST-"
  const memberData = memberDir.getRange(2, 1, memberDir.getLastRow() - 1, 1).getValues();
  for (let i = memberData.length - 1; i >= 0; i--) {
    if (String(memberData[i][0]).startsWith('TEST-')) {
      memberDir.deleteRow(i + 2);
    }
  }

  const grievanceData = grievanceLog.getRange(2, 1, grievanceLog.getLastRow() - 1, 1).getValues();
  for (let i = grievanceData.length - 1; i >= 0; i--) {
    if (String(grievanceData[i][0]).startsWith('TEST-')) {
      grievanceLog.deleteRow(i + 2);
    }
  }
}
