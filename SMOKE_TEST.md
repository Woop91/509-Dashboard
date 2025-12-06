# Smoke Test Checklist

Runtime validation tests for the 509 Dashboard. These tests verify functionality that cannot be checked through static analysis.

## Prerequisites

1. Import `ConsolidatedDashboard.gs` as a single script file in the target Google Spreadsheet
2. Or use `clasp push` to deploy all individual `.gs` files

## Step 1: Initial Setup

1. **Run CREATE_509_DASHBOARD**
   - In Script Editor: Run > `CREATE_509_DASHBOARD`
   - Fix any authorization prompts (Drive, Gmail, Calendar, etc.)
   - Wait for completion

2. **Verify Core Sheets Created**
   - [ ] Config
   - [ ] Member Directory
   - [ ] Grievance Log
   - [ ] Dashboard
   - [ ] Analytics Data
   - [ ] Feedback & Development
   - [ ] Getting Started
   - [ ] FAQ & Help

## Step 2: Reload & Menu Test

1. **Close and reopen the spreadsheet**
   - Verify no errors on load
   - Verify custom menus appear:
     - [ ] Dashboard (or 509 Tools)
     - [ ] Sheet Manager
     - [ ] Setup
     - [ ] Administrator

## Step 3: Member Directory Tests

1. **Add a member manually**
   - [ ] Auto-ID generates (M000001 format)
   - [ ] All dropdowns work (Job Title, Location, Unit, etc.)
   - [ ] No broken formula errors

2. **Verify dropdown sources**
   - [ ] Job Titles from Config
   - [ ] Locations from Config
   - [ ] Units from Config
   - [ ] Steward dropdown shows actual stewards

## Step 4: Grievance Workflow Tests

1. **Start a Grievance from Member**
   - [ ] Toggle "Start Grievance" checkbox for a member
   - [ ] Verify grievance row appears in Grievance Log
   - [ ] Member info synced correctly (Name, Email, Location)
   - [ ] Auto-ID generates (G-000001 format)

2. **Verify Computed Deadlines**
   - [ ] Filing Deadline auto-calculates (21 days from incident)
   - [ ] Step I Decision Due auto-calculates (30 days)
   - [ ] Days Open formula works
   - [ ] Next Action Due formula works

3. **Test Status Changes**
   - [ ] Change status to "Pending Info" - verify tracking
   - [ ] Change status to "Settled" - verify Date Closed prompts

## Step 5: Security & Audit Tests

1. **Check Audit Log**
   - [ ] Navigate to Audit Log sheet (may be hidden)
   - [ ] Verify entries are being logged
   - [ ] Confirm no blocking protection issues

2. **Test Role-Based Access**
   - [ ] Run Security Audit from Administrator menu
   - [ ] Verify role detection works

## Step 6: Dashboard & Charts

1. **Main Dashboard**
   - [ ] Navigate to Dashboard sheet
   - [ ] Verify KPIs display correctly
   - [ ] No #REF! or #ERROR! cells

2. **Interactive Dashboard**
   - [ ] Open Interactive Dashboard
   - [ ] Change metric dropdowns
   - [ ] Refresh charts - verify they render
   - [ ] Toggle comparison mode

## Step 7: Integration Tests

1. **Google Drive Integration** (if configured)
   - [ ] Setup Drive Folder for Grievance
   - [ ] Verify folder created

2. **Email Integration** (if configured)
   - [ ] Compose test email
   - [ ] Verify Communications Log entry

3. **Calendar Integration** (if configured)
   - [ ] Sync deadlines to Calendar
   - [ ] Verify calendar events created

## Common Issues & Fixes

| Issue | Likely Cause | Fix |
|-------|--------------|-----|
| Dropdowns empty | Config sheet missing data | Populate Config columns |
| Auto-ID not working | Last row detection issue | Check for blank rows |
| Charts not loading | Data range empty | Add sample data first |
| Permission errors | Triggers not authorized | Re-run authorization |
| #REF! errors | Missing sheet references | Run CREATE_509_DASHBOARD |

## Performance Benchmarks

For a spreadsheet with 20K members and 5K grievances:

| Operation | Expected Time |
|-----------|---------------|
| Dashboard Refresh | < 5 seconds |
| Grievance Search | < 2 seconds |
| Report Generation | < 30 seconds |
| Full Backup | < 60 seconds |

## Notes

- First run may be slower due to cache warming
- Some features require additional API authorizations
- Organization-specific text (SEIU Local 509, Article references) can be customized in Config sheet or FAQ entries
