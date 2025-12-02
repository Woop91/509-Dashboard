#!/usr/bin/env node

/**
 * BUILD SCRIPT FOR 509 DASHBOARD
 * ================================
 *
 * This script automatically builds the consolidated dashboard file
 * from individual modules, eliminating manual sync issues.
 *
 * Usage:
 *   node build.js
 *
 * Output:
 *   ConsolidatedDashboard.gs - Auto-generated consolidated file
 *
 * DO NOT edit ConsolidatedDashboard.gs directly!
 * Edit individual modules and run this script to rebuild.
 */

const fs = require('fs');
const path = require('path');

/**
 * MODULE DEPENDENCY GRAPH
 * =======================
 *
 * Module loading order is critical. Dependencies must be loaded before dependents.
 *
 * Legend:
 * [CORE] = No dependencies, can be loaded first
 * [DEPENDS: X, Y] = Requires modules X and Y to be loaded first
 *
 * Dependency Hierarchy:
 *
 * Level 0 (No Dependencies):
 *   - Constants.gs
 *
 * Level 1 (Depends on Constants only):
 *   - SecurityUtils.gs
 *   - TestConfig.gs
 *   - HTMLTemplates.gs
 *   - I18n.gs
 *
 * Level 2 (Depends on Level 0-1):
 *   - Code.gs (Depends: Constants, SecurityUtils)
 *   - DataArchiving.gs (Depends: Constants, SecurityUtils, HTMLTemplates)
 *
 * Level 3+ (Depends on Level 0-2):
 *   - All feature modules (Depends: Constants, SecurityUtils, Code.gs)
 *   - Test modules (Depends: Constants, TestConfig)
 */

// Module order matters - dependencies must come first
const MODULES = [
  // ===== LEVEL 0: CORE INFRASTRUCTURE (NO DEPENDENCIES) =====
  'Constants.gs',  // [CORE] Provides: SHEETS, MEMBER_COLS, COLORS, ERROR_MESSAGES, etc.

  // ===== LEVEL 1: DEPENDS ON CONSTANTS ONLY =====
  'SecurityUtils.gs',   // [DEPENDS: Constants] Provides: sanitizeHTML, requireRole, isAdmin
  'HTMLTemplates.gs',   // [DEPENDS: Constants] Provides: createHTMLPage, createButton, etc.
  'I18n.gs',           // [DEPENDS: Constants] Provides: t(), getUserLanguage, translations
  'TestConfig.gs',     // [DEPENDS: Constants] Provides: TEST_CONFIG, getTestSpreadsheet

  // ===== LEVEL 2: MAIN SETUP & UTILITIES =====
  'Code.gs',           // [DEPENDS: Constants, SecurityUtils] Provides: CREATE_509_DASHBOARD, onOpen
  'DataArchiving.gs',  // [DEPENDS: Constants, SecurityUtils, HTMLTemplates] Provides: archiveOldGrievances

  // ===== LEVEL 3: FEATURE MODULES (ALPHABETICAL) =====
  // All feature modules depend on: Constants, SecurityUtils, Code.gs
  'ADHDEnhancements.gs',
  'EnhancedADHDFeatures.gs',
  'AddRecommendations.gs',
  'AutomatedNotifications.gs',
  'AutomatedReports.gs',
  'BatchGrievanceRecalc.gs',
  'BatchOperations.gs',
  'CalendarIntegration.gs',
  'ColumnToggles.gs',
  'CustomReportBuilder.gs',
  'DarkModeThemes.gs',
  'DataBackupRecovery.gs',
  'DataCachingLayer.gs',
  'DataIntegrityEnhancements.gs',
  'DataPagination.gs',
  'DistributedLocks.gs',
  'EnhancedErrorHandling.gs',
  'FAQKnowledgeBase.gs',
  'GettingStartedAndFAQ.gs',
  'GmailIntegration.gs',
  'GoogleDriveIntegration.gs',
  'GracefulDegradation.gs',
  'GrievanceFloatToggle.gs',
  'GrievanceWorkflow.gs',
  'IdempotentOperations.gs',
  'IncrementalBackupSystem.gs',
  'InteractiveDashboard.gs',
  'KeyboardShortcuts.gs',
  'LazyLoadCharts.gs',
  'MemberDirectoryDropdowns.gs',
  'MemberDirectoryGoogleFormLink.gs',
  'MemberSearch.gs',
  'MobileOptimization.gs',
  'OptimizedDashboardRebuild.gs',
  'PerformanceMonitoring.gs',
  'Phase6Integration.gs',
  'PredictiveAnalytics.gs',
  'ReorganizedMenu.gs',
  'RootCauseAnalysis.gs',
  'SeedNuke.gs',
  'SmartAutoAssignment.gs',
  'TransactionRollback.gs',
  'UndoRedoSystem.gs',
  'UnifiedOperationsMonitor.gs',
  'WorkflowStateMachine.gs',

  // ===== LEVEL 4: TESTING MODULES (LAST) =====
  // Test modules should be loaded last
  'TestFramework.gs',     // [DEPENDS: Constants, TestConfig] Provides: runAllTests, Assert
  'Code.test.gs',         // [DEPENDS: TestFramework, Code.gs] Unit tests
  'Integration.test.gs'   // [DEPENDS: TestFramework, All modules] Integration tests
];

/**
 * Module dependency validation
 * Checks if all module dependencies are satisfied
 */
function validateModuleDependencies() {
  console.log('🔍 Validating module dependencies...\n');

  const moduleSet = new Set(MODULES);
  const warnings = [];

  // Define known dependencies
  const dependencies = {
    'SecurityUtils.gs': ['Constants.gs'],
    'HTMLTemplates.gs': ['Constants.gs'],
    'I18n.gs': ['Constants.gs'],
    'TestConfig.gs': ['Constants.gs'],
    'Code.gs': ['Constants.gs', 'SecurityUtils.gs'],
    'DataArchiving.gs': ['Constants.gs', 'SecurityUtils.gs', 'HTMLTemplates.gs'],
    'TestFramework.gs': ['Constants.gs', 'TestConfig.gs'],
    'Code.test.gs': ['TestFramework.gs', 'Code.gs'],
    'Integration.test.gs': ['TestFramework.gs']
  };

  // Check each module's dependencies
  MODULES.forEach((module, index) => {
    if (dependencies[module]) {
      dependencies[module].forEach(dep => {
        const depIndex = MODULES.indexOf(dep);

        if (depIndex === -1) {
          warnings.push(`⚠️  ${module} depends on ${dep}, but ${dep} is not in build list`);
        } else if (depIndex > index) {
          warnings.push(`⚠️  ${module} (index ${index}) depends on ${dep} (index ${depIndex}), but ${dep} is loaded later`);
        }
      });
    }
  });

  if (warnings.length > 0) {
    console.log('⚠️  DEPENDENCY WARNINGS:\n');
    warnings.forEach(w => console.log(`   ${w}`));
    console.log('');
  } else {
    console.log('✅ All module dependencies are correctly ordered\n');
  }

  return warnings.length === 0;
}

// Files to exclude from build
const EXCLUDED_FILES = [
  'ConsolidatedDashboard.gs',  // Output file
  'Complete509Dashboard.gs',   // Old version
  'build.js'                   // This script
];

/**
 * Main build function
 */
function build() {
  console.log('🔨 Building 509 Dashboard Consolidated File...\n');

  const buildDate = new Date().toISOString();
  const buildVersion = '2.0.0';

  let consolidated = `/**
 * ============================================================================
 * 509 DASHBOARD - CONSOLIDATED BUILD
 * ============================================================================
 *
 * This file is AUTO-GENERATED by build.js
 * DO NOT EDIT THIS FILE DIRECTLY
 *
 * To make changes:
 * 1. Edit individual module files (e.g., SecurityUtils.gs, Code.gs)
 * 2. Run: node build.js
 * 3. This file will be regenerated automatically
 *
 * Build Info:
 * - Version: ${buildVersion}
 * - Build Date: ${buildDate}
 * - Modules: ${MODULES.length} files
 *
 * ============================================================================
 */

`;

  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  // Process each module
  MODULES.forEach((moduleName, index) => {
    const modulePath = path.join(__dirname, moduleName);

    try {
      if (!fs.existsSync(modulePath)) {
        console.log(`⚠️  [${index + 1}/${MODULES.length}] SKIPPED: ${moduleName} (file not found)`);
        return; // Skip missing files (they may not be created yet)
      }

      const content = fs.readFileSync(modulePath, 'utf8');

      // Add module header
      consolidated += `\n${'='.repeat(80)}\n`;
      consolidated += `// MODULE: ${moduleName}\n`;
      consolidated += `// Source: ${moduleName}\n`;
      consolidated += `${'='.repeat(80)}\n\n`;

      // Add module content
      consolidated += content;

      // Add spacing between modules
      consolidated += '\n\n';

      console.log(`✅ [${index + 1}/${MODULES.length}] ${moduleName} (${content.length} bytes)`);
      successCount++;

    } catch (error) {
      console.error(`❌ [${index + 1}/${MODULES.length}] ERROR: ${moduleName}`);
      console.error(`   ${error.message}`);
      errors.push({ module: moduleName, error: error.message });
      errorCount++;
    }
  });

  // Write consolidated file
  const outputPath = path.join(__dirname, 'ConsolidatedDashboard.gs');

  try {
    fs.writeFileSync(outputPath, consolidated, 'utf8');
    console.log(`\n✅ BUILD SUCCESSFUL!`);
    console.log(`   Output: ${outputPath}`);
    console.log(`   Size: ${(consolidated.length / 1024).toFixed(2)} KB`);
    console.log(`   Modules: ${successCount} included, ${errorCount} errors`);

    // Show file size comparison
    const stats = fs.statSync(outputPath);
    console.log(`   File size: ${(stats.size / 1024).toFixed(2)} KB\n`);

    // Show errors if any
    if (errors.length > 0) {
      console.log(`\n⚠️  ERRORS ENCOUNTERED:\n`);
      errors.forEach(err => {
        console.log(`   - ${err.module}: ${err.error}`);
      });
    }

    // Build report
    console.log('\n📊 BUILD REPORT:');
    console.log(`   ✅ Successfully built: ${successCount} modules`);
    console.log(`   ⚠️  Errors: ${errorCount}`);
    console.log(`   📦 Total size: ${(consolidated.length / 1024).toFixed(2)} KB`);
    console.log(`   📅 Build date: ${buildDate}\n`);

    console.log('🎉 Consolidated file ready for deployment!\n');
    console.log('Next steps:');
    console.log('   1. Review ConsolidatedDashboard.gs');
    console.log('   2. Deploy to Google Apps Script');
    console.log('   3. Test in your Google Sheet\n');

  } catch (error) {
    console.error(`\n❌ BUILD FAILED!`);
    console.error(`   Error writing output file: ${error.message}\n`);
    process.exit(1);
  }
}

/**
 * Verify all modules exist
 */
function verifyModules() {
  console.log('🔍 Verifying modules...\n');

  let missing = [];
  let found = 0;

  MODULES.forEach(moduleName => {
    const modulePath = path.join(__dirname, moduleName);
    if (fs.existsSync(modulePath)) {
      found++;
    } else {
      missing.push(moduleName);
    }
  });

  if (missing.length > 0) {
    console.log(`⚠️  Missing modules (${missing.length}):`);
    missing.forEach(name => console.log(`   - ${name}`));
    console.log(`\n   These modules will be skipped in the build.\n`);
  }

  console.log(`✅ Found ${found}/${MODULES.length} modules\n`);
}

/**
 * Clean old build artifacts
 */
function clean() {
  const outputPath = path.join(__dirname, 'ConsolidatedDashboard.gs');

  if (fs.existsSync(outputPath)) {
    fs.unlinkSync(outputPath);
    console.log('🧹 Cleaned old build artifacts\n');
  }
}

// Parse command line arguments
const args = process.argv.slice(2);

if (args.includes('--help') || args.includes('-h')) {
  console.log(`
509 Dashboard Build Script
==========================

Usage:
  node build.js [options]

Options:
  --help, -h     Show this help message
  --verify, -v   Verify modules exist without building
  --clean, -c    Remove old build artifacts
  --quiet, -q    Suppress verbose output

Examples:
  node build.js              # Build consolidated file
  node build.js --verify     # Check which modules exist
  node build.js --clean      # Clean then build
`);
  process.exit(0);
}

if (args.includes('--verify') || args.includes('-v')) {
  verifyModules();
  process.exit(0);
}

if (args.includes('--clean') || args.includes('-c')) {
  clean();
}

// Run the build
try {
  verifyModules();
  validateModuleDependencies();
  build();
} catch (error) {
  console.error('\n💥 FATAL ERROR:', error.message);
  console.error(error.stack);
  process.exit(1);
}
