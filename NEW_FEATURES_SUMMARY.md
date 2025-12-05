# New Features Summary

This document summarizes the new features added to the 509 Dashboard.

## 1. Grievance Float/Sort Toggle Feature

**File:** `GrievanceFloatToggle.gs`

**Description:**
A toggle feature that reorganizes the Grievance Log based on priority:
- Sends closed/settled/inactive grievances to the bottom
- Prioritizes by step: Step 3 > Step 2 > Step 1
- Within each step, sorts by soonest due date

**Functions:**
- `toggleGrievanceFloat()` - Toggle the feature on/off
- `applyGrievanceFloat()` - Apply the sorting immediately
- `showGrievanceFloatPanel()` - Show control panel dialog
- `getGrievanceFloatState()` / `setGrievanceFloatState()` - State management

**Menu Location:**
- Average User > Grievance Tools > Grievance Float Toggle
- Average User > Grievance Tools > Float Control Panel

## 2. Google Drive Folder Auto-Creation

**Files Modified:**
- `GoogleDriveIntegration.gs` - Updated `createGrievanceFolder()` to accept grievant name
- `GrievanceWorkflow.gs` - Updated `addGrievanceToLog()` to auto-create folders

**Description:**
Automatically creates a Google Drive folder with subfolders when a new grievance is created.
Folder is named with format: `Grievance_{ID}_{FirstName_LastName}`

**Subfolders Created:**
- Evidence
- Correspondence
- Forms
- Other

**Menu Location:**
- Existing: Sheet Manager > Google Drive Integration > Batch Create All Folders

## 3. Member Directory Dropdowns

**File:** `MemberDirectoryDropdowns.gs`

**Description:**
Adds data validation dropdowns to Member Directory for consistent data entry across all specified fields.

**Dropdowns Added:**
- Job Title / Position (Column D)
- Department / Unit
- Worksite / Office Location (Column E)
- Work Schedule / Office Days (Column G) - Multiple selections
- Unit (8 or 10) (Column F)
- Is Steward (Y/N) (Column J)
- Assigned Steward (Column M) - Populated from stewards list
- Immediate Supervisor (Column K)
- Manager / Program Director (Column L)
- Engagement Level
- Committee Member - Multiple selections
- Interest: Local Actions (Column T/20) - Y/N
- Interest: Chapter Actions (Column U/21) - Y/N
- Interest: Allied Chapter Actions (Column V/22) - Y/N
- Preferred Communication Methods (Column X/24) - Multiple selections
- Best Time(s) to Reach Member (Column Y/25) - Multiple selections
- Steward Who Contacted Member (Column AD/30) - Populated from stewards list

**Functions:**
- `setupMemberDirectoryDropdowns()` - Sets up all dropdowns
- `refreshStewardDropdowns()` - Refreshes steward lists
- `removeEmergencyContactColumns()` - Removes emergency contact columns

**Menu Location:**
- Administrator > Setup & Configuration > Setup Member Directory Dropdowns
- Administrator > Setup & Configuration > Refresh Steward Dropdowns
- Administrator > Setup & Configuration > Remove Emergency Contact Columns

## 4. Member Directory Google Form Link

**File:** `MemberDirectoryGoogleFormLink.gs`

**Description:**
Adds a feature to open a Google Form with auto-populated member information. This feature is configurable and can be wired to various use cases.

**Functions:**
- `openMemberGoogleForm()` - Opens the form for selected member
- `generatePreFilledMemberForm()` - Generates pre-filled form URL
- `addMemberFormLinkToPendingFeatures()` - Adds to pending features list

**Configuration Required:**
Update `MEMBER_FORM_CONFIG` object with your Google Form URL and field IDs.

**Menu Location:**
- Administrator > Setup & Configuration > Open Member Google Form

## 5. Reorganized Menu System

**File:** `ReorganizedMenu.gs`

**Description:**
Reorganizes the dashboard menus into three categories for better organization and user experience.

**Three Menu Categories:**

### Average User Menu ("👤 Dashboard")
- Dashboards
- Search & Lookup
- Grievance Tools (includes new Float Toggle)
- Google Drive
- Communications
- Reports
- Accessibility
- Help & Support

### Sheet Manager Menu ("📊 Sheet Manager")
- Data Management
- Performance
- Data Integrity
- Automations
- Google Drive Integration
- Calendar Integration
- Analysis & Insights
- Batch Operations
- Smart Assignment
- Knowledge Base

### Administrator Menu ("⚙️ Administrator")
- Seed Functions
- System Health
- Root Cause Analysis
- Workflow Management
- Setup & Configuration
- Column Toggles & View
- History & Undo
- Mobile & Viewing
- Testing

**To Activate:**
Rename `onOpen_Reorganized()` to `onOpen()` in ReorganizedMenu.gs, and rename the existing `onOpen()` to `onOpen_OLD()`.

## 6. Interest Columns

**Note:** The following columns already exist in the Member Directory:
- Interest: Local Actions (Column T/20)
- Interest: Chapter Actions (Column U/21)
- Interest: Allied Chapter Actions (Column V/22)

Dropdowns have been added for these columns via `setupMemberDirectoryDropdowns()`.

## Applying Changes to ConsolidatedDashboard.gs

⚠️ **CRITICAL**: The following changes MUST be applied to both ConsolidatedDashboard.gs and Complete509Dashboard.gs:

### 1. Update GRIEVANCE_COLS Constant
Add two new columns to GRIEVANCE_COLS (around line 96):
```javascript
  STEWARD: 27,          // AA
  RESOLUTION: 28,       // AB
  DRIVE_FOLDER_ID: 29,  // AC - Google Drive folder ID
  DRIVE_FOLDER_URL: 30  // AD - Google Drive folder URL
};
```

### 2. Update createGrievanceLog() Headers
Add two new headers to the headers array (around line 403):
```javascript
    "Resolution Summary",
    "Drive Folder ID",
    "Drive Folder Link"
  ];
```

### 3. Append New Feature Files
Append these files' contents to the consolidated files:
1. `GrievanceFloatToggle.gs` - Append entire file
2. `MemberDirectoryDropdowns.gs` - Append entire file
3. `MemberDirectoryGoogleFormLink.gs` - Append entire file
4. `ReorganizedMenu.gs` - Replace onOpen() function or add as alternative

### 4. Update Existing Functions
The following changes have already been made to modular files and need to be synced:
- `GoogleDriveIntegration.gs`: Updated to use GRIEVANCE_COLS.DRIVE_FOLDER_ID and GRIEVANCE_COLS.DRIVE_FOLDER_URL
- `GrievanceWorkflow.gs`: Added auto-folder creation in `addGrievanceToLog()`

### 5. Fix Remaining Hardcoded Column References (CRITICAL)
The consolidated files still have hardcoded column references that need to be made dynamic:
- Complete509Dashboard.gs: Lines 1288, 1292, 1345, 1350, 1358
- ConsolidatedDashboard.gs: Lines 1058, 1062, 1115, 1120, 1128

These references to 'Grievance Log'!A:A and 'Member Directory'!J:J should use dynamic column constants per AI_REFERENCE requirements.

## Testing Checklist

- [ ] Test Grievance Float Toggle (enable/disable/sort)
- [ ] Test Google Drive folder auto-creation on new grievance
- [ ] Test Member Directory dropdowns (all fields)
- [ ] Test steward dropdown auto-population
- [ ] Test emergency contact column removal
- [ ] Test Member Google Form link (requires form configuration)
- [ ] Test reorganized menu navigation
- [ ] Verify all functions work in ConsolidatedDashboard.gs

## Pending Features Added

The following features have been added to the Feedback & Development sheet:

1. **Grievance Float/Sort Toggle** - Status: Completed
2. **Member Directory Google Form Link** - Status: Planned

Use the following functions to add these to your Feedback sheet:
- `addGrievanceFloatToPendingFeatures()`
- `addMemberFormLinkToPendingFeatures()`
