# 509 Dashboard - Complete Feature Reference (Consolidated Architecture)

**Version:** 2.3 (Consolidated Branch)
**Last Updated:** 2025-11-27
**Purpose:** Union grievance tracking and member engagement system for SEIU Local 509
**Architecture:** Single-file consolidated (Code.gs only) with runtime dynamic column mapping

---

## 🔴 CRITICAL: Always Reference This Document

**Before making ANY changes to the codebase:**

1. **READ AI_REFERENCE.md first** - This document is the single source of truth for the entire system
2. **Check the Changelog** - Understand recent changes and current version
3. **Review Code Quality section** - Avoid repeating fixed issues
4. **Verify dynamic column usage** - ALL column references MUST use runtime column mapping
5. **Follow established patterns** - Don't introduce inconsistencies

**Why This Matters:**
- Prevents re-introducing bugs that were already fixed
- Ensures consistency across all 21 sheets
- Maintains 100% dynamic column coverage (critical for system stability)
- Documents all design decisions and architectural choices

**⚠️ DO NOT:**
- Make changes without consulting this document
- Use hardcoded column references (A:A, AB:AB, etc.)
- Add features without updating this documentation
- Skip the verification commands in the Code Quality section

**✅ ALWAYS:**
- Use `getColumnMapping()` to get runtime column indices
- Use `initializeColumnIndices()` for dynamic header reading
- Update changelog when making significant changes
- Run verification commands after modifications

---

## 🏗️ Consolidated Architecture

**THIS BRANCH uses a different architecture from the modular branch:**

### Key Differences:

**File Structure:**
- **Single File:** Everything consolidated into `Code.gs` (~17,000+ lines)
- **No Modules:** No separate ADHDEnhancements.gs, GrievanceWorkflow.gs, etc.
- **No Complete509Dashboard.gs:** Not needed since this IS the single-file version

**Column Mapping System:**
- **Runtime Dynamic Mapping:** Reads column headers at runtime
- **Functions:** `initializeColumnIndices(sheetName)`, `getColumnMapping(sheetName)`
- **No Constants:** No MEMBER_COLS or GRIEVANCE_COLS compile-time constants
- **Caching:** Column mappings cached globally for performance
- **Flexibility:** Can handle column reordering, additions, deletions automatically

**Advantages:**
- ✅ Handles column structure changes without code modification
- ✅ More resilient to user customizations
- ✅ Single file easier to deploy
- ✅ No module synchronization issues

**Trade-offs:**
- ⚠️ ~50-100ms overhead on first access (then cached)
- ⚠️ Requires column headers to match exactly
- ⚠️ Harder to navigate (one large file)

### Using Runtime Column Mapping

**Instead of:**
```javascript
// Modular branch style (compile-time constants)
const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
formula = `=COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Open")`;
```

**Use:**
```javascript
// Consolidated branch style (runtime mapping)
const cols = getColumnMapping(SHEETS.GRIEVANCE_LOG);
// Access directly by index: cols.STATUS, cols.MEMBER_ID, etc.
// Or build formula with column letters dynamically
```

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Sheet Structure (21 Sheets)](#sheet-structure-21-sheets)
3. [Core Data Sheets](#core-data-sheets)
4. [Dashboard Sheets](#dashboard-sheets)
5. [Analytics Sheets](#analytics-sheets)
6. [Utility Sheets](#utility-sheets)
7. [Menu System](#menu-system)
8. [Data Validation Rules](#data-validation-rules)
9. [Formula System](#formula-system)
10. [Seed Data Functions](#seed-data-functions)
11. [Color Scheme](#color-scheme)
12. [Runtime Column Mapping System](#runtime-column-mapping-system)
13. [Testing Framework](#testing-framework)
14. [Changelog](#changelog)

---

## System Overview

The 509 Dashboard is a comprehensive Google Apps Script-based union management system that tracks:
- **Members:** 20,000+ member records with engagement data
- **Grievances:** 5,000+ grievance cases with deadline tracking
- **Analytics:** Real-time KPIs, trends, and performance metrics
- **Dashboards:** Executive summaries and interactive visualizations

**Key Design Principles:**
- No fake data (CPU, memory, etc.) - all metrics track real union activity
- **⚠️ CRITICAL: ALL column references MUST use runtime column mapping**
- Runtime dynamic column mapping (reads headers at execution time)
- Single-file consolidated architecture for easy deployment
- Real-time formula-based calculations
- Comprehensive data validation

**🔴 MANDATORY RULE: Everything Must Be Dynamic**
- **NEVER** use hardcoded column references like `'Member Directory'!A:A` or `'Grievance Log'!AB:AB`
- **ALWAYS** use `getColumnMapping(sheetName)` to get column indices
- **Example:** `const cols = getColumnMapping(SHEETS.GRIEVANCE_LOG); // cols.STATUS gives column index`
- This allows columns to be reordered without breaking formulas
- Verification: `grep "'Member Directory'![A-Z]:[A-Z]" Code.gs` should return 0 matches

---

## Sheet Structure (21 Sheets)

### Complete Sheet List

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 1 | ⚙️ Config | Core | Master dropdown lists for validation |
| 2 | 👥 Member Directory | Core | All member data (35 columns A-AI) |
| 3 | 📋 Grievance Log | Core | All grievance cases (32 columns A-AF) |
| 4 | 📊 Main Dashboard | Dashboard | Main real-time metrics dashboard |
| 5 | 📈 Analytics Data | Hidden | Computed aggregations for dashboards |
| 6 | Member Satisfaction | Core | Survey tracking and satisfaction scores |
| 7 | 🔮 Future Features | Utility | Feature requests and roadmap |
| 8 | 📝 Pending Features | Utility | In-progress features |
| 9 | 🎯 Interactive (Your Custom View) | Dashboard | User-customizable visualization |
| 10 | 📚 Getting Started | Help | Onboarding guide |
| 11 | ❓ FAQ | Help | Frequently asked questions |
| 12 | ⚙️ User Settings | Utility | Per-user preferences |
| 13 | 👨‍⚖️ Steward Workload | Analytics | Steward capacity and caseload |
| 14 | 📉 Test 1: Trends | Analytics | Monthly trend analysis |
| 15 | 🎯 Test 2: Performance | Analytics | Win rates and efficiency metrics |
| 16 | 📍 Test 3: Locations | Analytics | Geographic grievance distribution |
| 17 | 📁 Test 4: Types | Analytics | Grievance categorization analysis |
| 18 | 👔 Test 5: Executive View | Dashboard | Executive summary dashboard |
| 19 | 💼 Test 6: KPI View | Dashboard | KPI performance dashboard |
| 20 | 🤝 Test 7: Engagement | Analytics | Member participation metrics |
| 21 | 💰 Test 8: Cost Impact | Analytics | Financial recovery tracking |

**Recent Changes:** Consolidated from 25 to 21 sheets

---

## Runtime Column Mapping System

### How It Works

The consolidated branch uses **runtime dynamic column mapping** instead of compile-time constants.

### Core Functions

**1. initializeColumnIndices(sheetName, headerMap)**
```javascript
/**
 * Reads column headers from a sheet and creates index mappings
 * @param {string} sheetName - Name of the sheet
 * @param {Object} headerMap - Optional: Map of header text to property names
 * @returns {Object} Column index mapping object
 */
function initializeColumnIndices(sheetName, headerMap = {}) {
  // Reads first row headers
  // Creates mapping like: { MEMBER_ID: 1, FIRST_NAME: 2, ... }
  // Cached globally for performance
}
```

**2. getColumnMapping(sheetName)**
```javascript
/**
 * Get cached column mapping for a sheet (lazy initialization)
 * @param {string} sheetName - Name of the sheet
 * @returns {Object} Column index mapping
 */
function getColumnMapping(sheetName) {
  // Returns cached mapping or initializes if not cached
  // ~50-100ms first call, then instant from cache
}
```

### Usage Examples

**Member Directory Columns:**
```javascript
const memberCols = getColumnMapping(SHEETS.MEMBER_DIR);
// Access: memberCols.MEMBER_ID, memberCols.FIRST_NAME, etc.
// Indices are 1-based (A=1, B=2, etc.)
```

**Grievance Log Columns:**
```javascript
const grievanceCols = getColumnMapping(SHEETS.GRIEVANCE_LOG);
// Access: grievanceCols.STATUS, grievanceCols.RESOLUTION, etc.
```

### Column Structure

**Member Directory (35 columns A-AI):**
- A: Member ID
- B: First Name
- C: Last Name
- D: Job Title
- E: Work Location (Site)
- F: Unit
- G: Office Days/Week
- H: Email
- I: Phone
- J: Is Steward? (Yes/No)
- K: Supervisor Name
- L: Manager Name
- M: Assigned Steward (Name)
- N: Last Virtual Meeting
- O: Last In-Person Meeting
- P: Last Survey Date
- Q: Last Email Open
- R: Email Open Rate %
- S: Volunteer Hours (YTD)
- T: Interest in Local Roles?
- U: Interest in Chapter Roles?
- V: Interest in Allied Campaigns?
- W: Timestamp
- X: Preferred Communication Method
- Y: Best Time to Contact
- Z: Has Open Grievance? (Formula)
- AA: # Open Grievances (Formula)
- AB: Next Grievance Deadline (Formula)
- AC: Most Recent Steward Contact Date
- AD: Steward Who Contacted Member
- AE: Notes from Steward Contact
- AF: Status (Active/Inactive)
- AG: Date Added
- AH: Last Updated
- AI: Tags (comma-separated)

**Grievance Log (32 columns A-AF):**
- A: Grievance ID
- B: Member ID
- C: First Name
- D: Last Name
- E: Status
- F: Current Step
- G: Incident Date
- H: Filing Deadline
- I: Date Filed
- J: Step I Due Date
- K: Step I Decision Received
- L: Step II Appeal Due
- M: Step II Appeal Filed
- N: Step II Due Date
- O: Step II Decision Received
- P: Step III Appeal Due
- Q: Step III Appeal Filed
- R: Date Closed
- S: Days Open (Formula)
- T: Next Action Due (Formula)
- U: Days to Deadline (Formula)
- V: Article(s) Violated
- W: Issue Category
- X: Member Email
- Y: Unit
- Z: Location
- AA: Assigned Steward (Name)
- AB: Resolution Summary
- AC: Notes
- AD: Priority (Auto-calculated)
- AE: Overdue? (Formula)
- AF: Last Updated

---

## Testing Framework

The consolidated branch now includes a comprehensive test suite:

### Test Files

**1. TestFramework.gs (487 lines)**
- Core testing utilities
- Assertion functions
- Test runner infrastructure
- Result formatting

**2. Code.test.gs (671 lines)**
- Unit tests for core functions
- Sheet creation tests
- Data validation tests
- Formula verification

**3. Integration.test.gs (620 lines)**
- End-to-end workflow tests
- Multi-sheet integration tests
- Performance benchmarks
- Data integrity checks

### Running Tests

**From Menu:**
```
Test & Debug → Run All Tests
```

**Or programmatically:**
```javascript
RUN_ALL_TESTS();
```

### Test Documentation

See `TESTING.md` for comprehensive testing guide and `TEST_QUICK_START.md` for quick start.

---

## Changelog

### Version 2.3 (Consolidated Branch - Current)

**Major Changes:**
- ✅ Merged test suite from testing branch (1,778 lines of tests)
- ✅ Added AI_REFERENCE.md (this document)
- ✅ Updated for consolidated architecture documentation
- ✅ Added TESTING.md, TEST_QUICK_START.md, QUICK_DEPLOY.md
- ✅ Maintained single-file consolidated architecture
- ✅ Preserved runtime dynamic column mapping system

**Files Structure:**
- Code.gs: Single consolidated file (~17,000+ lines)
- TestFramework.gs: Testing infrastructure (487 lines)
- Code.test.gs: Unit tests (671 lines)
- Integration.test.gs: Integration tests (620 lines)
- 13 .md documentation files

**Architecture:**
- Runtime column mapping (initializeColumnIndices, getColumnMapping)
- Global caching for performance (~50-100ms first access, then instant)
- No compile-time MEMBER_COLS/GRIEVANCE_COLS constants
- Handles column structure changes automatically

### Version 2.0-2.2 (Modular Branch - Parallel Development)

The modular branch pursued a different architecture with:
- Separate .gs files for each feature module
- Compile-time MEMBER_COLS/GRIEVANCE_COLS constants
- getColumnLetter() helper function
- Complete509Dashboard.gs for single-file deployment

**Note:** The consolidated and modular branches have diverged architecturally. This consolidated branch prioritizes flexibility and single-file deployment, while the modular branch prioritizes code organization and maintainability.

---

## Verification Commands

**Check for hardcoded column references:**
```bash
grep "'Member Directory'![A-Z]:[A-Z]" Code.gs
grep "'Grievance Log'![A-Z]:[A-Z]" Code.gs
# Should return 0 matches
```

**Verify test files exist:**
```bash
ls -la Code.test.gs Integration.test.gs TestFramework.gs
```

**Check for runtime column mapping usage:**
```bash
grep -c "getColumnMapping\|initializeColumnIndices" Code.gs
# Should return multiple matches
```

---

## Code Quality Standards

**MANDATORY RULES:**
1. ✅ ALL column references must use runtime column mapping
2. ✅ NO hardcoded column letters (A:A, AB:AB, etc.)
3. ✅ Cache column mappings globally for performance
4. ✅ Update AI_REFERENCE.md when making structural changes
5. ✅ Run tests after significant modifications
6. ✅ Maintain single-file consolidated architecture

**Bug Fixes Included:**
- ✅ Win Rate calculation (resolution text includes "Won"/"Lost"/"Settled")
- ✅ Column group errors (delete/recreate pattern)
- ✅ All column references dynamic (runtime mapping)
- ✅ setDate() bug fixed in recalcAllGrievances()
- ✅ Member array column indices corrected after restructuring

---

**Document Version:** 2.3
**Last Updated:** 2025-11-27
**Maintained By:** Claude (AI Assistant)
**Branch:** claude/consolidated-gs-01SWZWcYa9AMkCxVRQk8hkuV
**Repository:** Woop91/509-dashboard

---

## Support & Maintenance

### How to Debug Issues

**1. Run Diagnostics:**
```
Menu → Help & Support → 🔧 Diagnose Setup
```

**2. Check Execution Log:**
```
Apps Script Editor → View → Logs
```

**3. Run Tests:**
```
Menu → Test & Debug → Run All Tests
```

### Getting Help

1. Check this AI_REFERENCE.md document first
2. Review TESTING.md for test suite documentation
3. Run DIAGNOSE_SETUP() to identify issues
4. Check Apps Script logs for errors
5. Review recent commits for changes

---
