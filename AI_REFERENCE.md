# 509 Dashboard - Complete Feature Reference

**Version:** 2.4
**Last Updated:** 2025-12-05
**Purpose:** Union grievance tracking and member engagement system for SEIU Local 509

---

## 🔴 CRITICAL: Always Reference This Document

**Before making ANY changes to the codebase:**

1. **READ AI_REFERENCE.md first** - This document is the single source of truth for the entire system
2. **Check the Changelog** - Understand recent changes and current version
3. **Review Code Quality section** - Avoid repeating fixed issues
4. **Verify dynamic column usage** - ALL column references MUST use MEMBER_COLS and GRIEVANCE_COLS
5. **Follow established patterns** - Don't introduce inconsistencies

**Why This Matters:**
- Prevents re-introducing bugs that were already fixed
- Ensures consistency across all 21 sheets and 15+ code files
- Maintains 100% dynamic column coverage (critical for system stability)
- Documents all design decisions and architectural choices

**⚠️ DO NOT:**
- Make changes without consulting this document
- Use hardcoded column references (A:A, AB:AB, etc.)
- Add features without updating this documentation
- Skip the verification commands in the Code Quality section

**✅ ALWAYS:**
- Reference MEMBER_COLS and GRIEVANCE_COLS constants
- Update changelog when making significant changes
- Run verification commands after modifications
- Maintain parity between Code.gs and Complete509Dashboard.gs

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
12. [Column Mapping System](#column-mapping-system)
13. [File Architecture](#file-architecture)

---

## System Overview

The 509 Dashboard is a comprehensive Google Apps Script-based union management system that tracks:
- **Members:** 20,000+ member records with engagement data
- **Grievances:** 5,000+ grievance cases with deadline tracking
- **Analytics:** Real-time KPIs, trends, and performance metrics
- **Dashboards:** Executive summaries and interactive visualizations

**Key Design Principles:**
- No fake data (CPU, memory, etc.) - all metrics track real union activity
- **⚠️ CRITICAL: ALL column references MUST be dynamic (no hardcoded column letters)**
- Dynamic column references (no hardcoded column letters)
- Consolidated sheets to reduce clutter
- Real-time formula-based calculations
- Comprehensive data validation

**🔴 MANDATORY RULE: Everything Must Be Dynamic**
- **NEVER** use hardcoded column references like `'Member Directory'!A:A` or `'Grievance Log'!AB:AB`
- **ALWAYS** use `MEMBER_COLS` and `GRIEVANCE_COLS` constants with `getColumnLetter()`
- **Example:** Use `${getColumnLetter(MEMBER_COLS.IS_STEWARD)}` instead of `J:J`
- This allows columns to be reordered without breaking formulas
- Verification: `grep "'Member Directory'![A-Z]:[A-Z]" *.gs` should return 0 matches

---

## Sheet Structure (21 Sheets)

### Complete Sheet List

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 1 | Config | Core | Master dropdown lists for validation |
| 2 | Member Directory | Core | All member data (31 columns) |
| 3 | Grievance Log | Core | All grievance cases (28 columns) |
| 4 | Dashboard | Dashboard | Main real-time metrics dashboard |
| 5 | Analytics Data | Hidden | Computed aggregations for dashboards |
| 6 | Member Satisfaction | Core | Survey tracking and satisfaction scores |
| 7 | Feedback & Development | Utility | Bug reports, features, roadmap |
| 8 | 🎯 Interactive (Your Custom View) | Dashboard | User-customizable visualization |
| 9 | 📚 Getting Started | Help | Onboarding guide |
| 10 | ❓ FAQ | Help | Frequently asked questions |
| 11 | ⚙️ User Settings | Utility | Per-user preferences |
| 12 | 👨‍⚖️ Steward Workload | Analytics | Steward capacity and caseload |
| 13 | 📈 Trends & Timeline | Analytics | Monthly trend analysis |
| 14 | 🗺️ Location Analytics | Analytics | Geographic breakdown |
| 15 | 📊 Type Analysis | Analytics | Issue category analysis |
| 16 | 💼 Executive Dashboard | Dashboard | Merged: Executive Summary + Quick Stats |
| 17 | 📊 KPI Performance Dashboard | Dashboard | Merged: Performance Metrics + KPI Board |
| 18 | 👥 Member Engagement | Analytics | Engagement scoring and tracking |
| 19 | 💰 Cost Impact | Analytics | Financial impact analysis |
| 20 | 📦 Archive | Utility | Archived records |
| 21 | 🔧 Diagnostics | Utility | System health checks |

**Recent Consolidations (25 → 21 sheets):**
- Merged: Feedback & Development + Future Features + Pending Features → "Feedback & Development"
- Merged: Executive Summary + Quick Stats → "💼 Executive Dashboard"
- Merged: Performance Metrics + KPI Board → "📊 KPI Performance Dashboard"

---

## Core Data Sheets

### 1. Config Sheet

**Purpose:** Master source for all dropdown validations and system configuration

**Columns (32 total) - See CONFIG_COLS constant in Constants.gs:**
```
Employment Info (1-5):
  A: Job Titles (Coordinator, Analyst, Case Manager, etc.)
  B: Office Locations (Boston HQ, Worcester Office, etc.)
  C: Units (Unit A - Administrative, Unit B - Technical, etc.)
  D: Office Days (Monday-Sunday)
  E: Yes/No (generic Y/N validation)

Supervision (6-7):
  F: Supervisors (names - full name format)
  G: Managers (names - full name format)

Steward Info (8-9):
  H: Stewards (names)
  I: Steward Committees (committees stewards can serve on)

Grievance Settings (10-14):
  J: Grievance Status (Open, Pending Info, Settled, Withdrawn, Closed, Appealed, In Arbitration, Denied)
  K: Grievance Step (Informal, Step I, Step II, Step III, Mediation, Arbitration)
  L: Issue Category (Discipline, Workload, Scheduling, Pay, etc.)
  M: Articles Violated (Art. 1 - Recognition, Art. 23 - Grievance Procedure, etc.)
  N: Communication Methods (Email, Phone, Text, In Person)

Links & Coordinators (15-17):
  O: Grievance Coordinators (comma-separated list)
  P: Grievance Form URL
  Q: Contact Form URL

Notifications (18-20):
  R: Admin Emails
  S: Alert Days (days before deadline to alert, e.g., "3, 7, 14")
  T: Notification Recipients (default CC)

Organization (21-24):
  U: Organization Name
  V: Local Number
  W: Main Address
  X: Main Phone

Integration (25-26):
  Y: Drive Folder ID (Google Drive root folder)
  Z: Calendar ID (Google Calendar)

Deadlines (27-30):
  AA: Filing Deadline Days (default: 21)
  AB: Step 1 Response Days (default: 30)
  AC: Step 2 Appeal Days (default: 10)
  AD: Step 2 Response Days (default: 30)

Multi-select Options (31-32):
  AE: Best Times (best times to contact members)
  AF: Home Towns (list of home towns in area)
```

**Styling:**
- Header row: Bold, dark gray background (#4A5568), white text
- Tab color: Blue (#2563EB)
- Auto-resized columns
- Frozen first row

---

### 2. Member Directory

**Purpose:** Complete member database with engagement tracking

**Columns (31 total) - See MEMBER_COLS constant:**
```
A: Member ID (M000001, M000002, etc.)
B: First Name
C: Last Name
D: Job Title (validated from Config!A)
E: Work Location (Site) (validated from Config!B)
F: Unit (validated from Config!C)
G: Office Days
H: Email Address
I: Phone Number
J: Is Steward (Y/N) (validated from Config!E)
K: Committees (multi-select, comma-separated - references Config!I)
L: Supervisor (Name)
M: Manager (Name)
N: Assigned Steward (Name) (validated from Config!H)
O: Preferred Communication Methods (multi-select - references Config!N)
P: Best Time(s) to Reach Member (multi-select - references Config!AE)
Q: Last Virtual Mtg (Date)
R: Last In-Person Mtg (Date)
S: Open Rate (%)
T: Volunteer Hours (YTD)
U: Interest: Local Actions (Y/N) (validated from Config!E)
V: Home Town (validated from Config!AF)
W: Interest: Chapter Actions (Y/N) (validated from Config!E)
X: Interest: Allied Chapter Actions (Y/N) (validated from Config!E)
Y: Has Open Grievance? (Formula: SUMPRODUCT with active statuses)
Z: Grievance Status Snapshot (Formula)
AA: Next Grievance Deadline (Formula)
AB: Most Recent Steward Contact Date
AC: Steward Who Contacted Member
AD: Notes from Steward Contact
AE: Start Grievance (Checkbox to start grievance with prepopulated member info)
```

**Key Formulas (Column Y - Has Open Grievance):**
Uses SUMPRODUCT to check ALL active grievance statuses:
- Open, Pending Info, Appealed, In Arbitration
```
=ARRAYFORMULA(IF(A2:A1000<>"",IF(SUMPRODUCT((('Grievance Log'!B:B=A2:A1000)*(('Grievance Log'!E:E="Open")+('Grievance Log'!E:E="Pending Info")+('Grievance Log'!E:E="Appealed")+('Grievance Log'!E:E="In Arbitration"))))>0,"Yes","No"),""))
```

**Data Validations:**
- Column D (Job Title): Config!A (CONFIG_COLS.JOB_TITLES)
- Column E (Work Location): Config!B (CONFIG_COLS.OFFICE_LOCATIONS)
- Column F (Unit): Config!C (CONFIG_COLS.UNITS)
- Column J (Is Steward): Config!E (CONFIG_COLS.YES_NO)
- Column N (Assigned Steward): Config!H (CONFIG_COLS.STEWARDS)
- Column U, W, X (Interests): Config!E (CONFIG_COLS.YES_NO)
- Column V (Home Town): Config!AF (CONFIG_COLS.HOME_TOWNS)

**Styling:**
- Header: Bold, green background (#059669), white text, wrapped
- Tab color: Green (#059669)
- Column widths: A=90px, H=180px, AD=250px
- Frozen first row, header height 50px

**Implementation Notes:**
- Sheet is deleted and recreated on setup (prevents column group errors)
- Formulas in columns Y, Z, AA are set for first 1000 rows via `setupFormulasAndCalculations()`
- ARRAYFORMULA used for "Has Open Grievance" to handle all rows at once

---

### 3. Grievance Log

**Purpose:** Complete grievance case tracking with automatic deadline calculations

**Columns (28 total) - See GRIEVANCE_COLS constant:**
```
A: Grievance ID (G-000001, G-000002, etc.)
B: Member ID (links to Member Directory)
C: First Name
D: Last Name
E: Status (validated: Open, Pending Info, Settled, Withdrawn, Closed, Appealed)
F: Current Step (validated: Informal, Step I, Step II, Step III, Mediation, Arbitration)
G: Incident Date
H: Filing Deadline (21d) (Formula: =IF(G2<>"",G2+21,""))
I: Date Filed (Step I)
J: Step I Decision Due (30d) (Formula: =IF(I2<>"",I2+30,""))
K: Step I Decision Rcvd
L: Step II Appeal Due (10d) (Formula: =IF(K2<>"",K2+10,""))
M: Step II Appeal Filed
N: Step II Decision Due (30d) (Formula: =IF(M2<>"",M2+30,""))
O: Step II Decision Rcvd
P: Step III Appeal Due (30d) (Formula: =IF(O2<>"",O2+30,""))
Q: Step III Appeal Filed
R: Date Closed
S: Days Open (Formula: =IF(I2<>"",IF(R2<>"",R2-I2,TODAY()-I2),""))
T: Next Action Due (Formula: =IF(E2="Open",IF(F2="Step I",J2,IF(F2="Step II",N2,IF(F2="Step III",P2,H2))),""))
U: Days to Deadline (Formula: =IF(T2<>"",T2-TODAY(),""))
V: Articles Violated (validated from Config)
W: Issue Category (validated from Config)
X: Member Email
Y: Unit (validated from Config)
Z: Work Location (Site) (validated from Config)
AA: Assigned Steward (Name) (validated from Config)
AB: Resolution Summary (e.g., "Won - Resolved favorably", "Lost - No violation found")
```

**Data Validations:**
- Column E (Status): Config!I2:I14
- Column F (Current Step): Config!J2:J14
- Column V (Articles): Config!L2:L14
- Column W (Issue Category): Config!K2:K14
- Column Y (Unit): Config!C2:C7
- Column Z (Location): Config!B2:B14
- Column AA (Steward): Config!H2:H14

**Styling:**
- Header: Bold, red background (#DC2626), white text, wrapped
- Tab color: Red (#DC2626)
- Column widths: A=110px, V=180px, AB=250px
- Frozen first row, header height 50px

**Auto-Calculated Deadlines:**
All deadline formulas (columns H, J, L, N, P, S, T, U) are set for first 100 rows in `setupFormulasAndCalculations()`

**Contract Rules Built Into Formulas:**
- Filing deadline: Incident date + 21 days
- Step I decision due: Date filed + 30 days
- Step II appeal due: Step I decision received + 10 days
- Step II decision due: Step II appeal filed + 30 days
- Step III appeal due: Step II decision received + 30 days

---

## Dashboard Sheets

### 4. Main Dashboard

**Purpose:** Real-time overview of all union metrics

**Layout:**

**Title Section:**
- A1:L2 merged: "📊 LOCAL 509 DASHBOARD" (18pt, purple #7C3AED)
- A3:L3 merged: Last updated timestamp formula

**Member Metrics (A5:L7):**
- 4 cards (3 columns each):
  - Total Members: `=COUNTA('Member Directory'!A:A)-1`
  - Active Stewards: `=COUNTIF('Member Directory'!J:J,"Yes")`
  - Avg Open Rate: `=TEXT(AVERAGE('Member Directory'!R:R)/100,"0.0%")`
  - YTD Vol. Hours: `=SUM('Member Directory'!S:S)`

**Grievance Metrics (A10:L12):**
- 4 cards (3 columns each):
  - Open Grievances: `=COUNTIF('Grievance Log'!E:E,"Open")`
  - Pending Info: `=COUNTIF('Grievance Log'!E:E,"Pending Info")`
  - Settled (This Month): `=COUNTIFS('Grievance Log'!E:E,"Settled",'Grievance Log'!R:R,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))`
  - Avg Days Open: `=ROUND(AVERAGE(FILTER('Grievance Log'!S:S,'Grievance Log'!E:E="Open")),0)`

**Engagement Metrics (A15:L17):**
- Last 30 days:
  - Virtual Mtgs: `=COUNTIF('Member Directory'!N:N,">="&TODAY()-30)`
  - In-Person Mtgs: `=COUNTIF('Member Directory'!O:O,">="&TODAY()-30)`
  - Local Interest: `=COUNTIF('Member Directory'!T:T,"Yes")`
  - Chapter Interest: `=COUNTIF('Member Directory'!U:U,"Yes")`

**Upcoming Deadlines (A20:L30):**
- Table with columns: Grievance ID, Member, Next Action, Days Until, Status
- Formula: QUERY of Grievance Log for open cases with deadlines in next 14 days

**Styling:**
- Tab color: Purple (#7C3AED)
- All metrics: 20pt bold font, centered
- Card headers: Gray background (#F3F4F6)

---

### 16. 💼 Executive Dashboard

**Purpose:** Consolidated executive summary (merged Executive Summary + Quick Stats)

**Section 1: Quick Stats (A4:D11)**
- Header: "⚡ QUICK STATS" (orange #F97316)
- 6 metrics with columns: Metric | Value | Comparison | Trend
  - Active Members (dynamic)
  - Active Grievances (dynamic: STATUS column)
  - Win Rate (dynamic: STATUS + RESOLUTION columns)
  - Avg Resolution Days (dynamic: DAYS_OPEN column)
  - Overdue Cases (dynamic: DAYS_TO_DEADLINE column)
  - Active Stewards

**Section 2: Detailed KPIs (A13:C22)**
- Header: "📊 DETAILED KEY PERFORMANCE INDICATORS" (purple #7C3AED)
- 8 metrics with columns: Metric | Value | Status
  - Total Active Members
  - Total Active Grievances (dynamic)
  - Overall Win Rate (dynamic: STATUS + RESOLUTION)
  - Avg Resolution Time (dynamic: DAYS_OPEN)
  - Cases Overdue (dynamic: DAYS_TO_DEADLINE)
  - Member Satisfaction Score
  - Total Grievances Filed YTD
  - Resolved Grievances (dynamic: STATUS)

**Dynamic Column Implementation:**
```javascript
const resolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);  // AB
const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);          // E
const daysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);    // S
const daysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE); // U

// Win Rate formula (dynamic):
`=TEXT(IFERROR(COUNTIFS('Grievance Log'!${statusCol}:${statusCol},"Resolved*",
  'Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")/
  COUNTIF('Grievance Log'!${statusCol}:${statusCol},"Resolved*"),0),"0%")`
```

**Styling:**
- Tab color: Purple (#7C3AED)
- Column widths: A=250px, B=150px, C=100px, D=100px
- Frozen rows: 4

---

### 17. 📊 KPI Performance Dashboard

**Purpose:** Comprehensive KPI tracking (merged Performance Metrics + KPI Board)

**Columns (12 total):**
```
A: KPI Name
B: Current Value
C: Target
D: Variance (Current - Target)
E: % Change (vs. previous period)
F: Status (dropdown: On Track, At Risk, Off Track, Exceeding)
G: Last Month (previous month value)
H: YTD Average
I: Best (best value this year)
J: Worst (worst value this year)
K: Owner (responsible person)
L: Last Updated (date)
```

**Data Validation:**
- Column F: Dropdown list ['On Track', 'At Risk', 'Off Track', 'Exceeding'] (F4:F1000)

**Styling:**
- Tab color: Green (#059669)
- Header: Bold, light gray background
- Column widths optimized (A=200px, B-L=80-120px)
- Frozen rows: 3

**Purpose:** Track all KPIs against targets with variance analysis and trend tracking

---

### 8. 🎯 Interactive (Your Custom View)

**Purpose:** User-customizable dashboard with live controls

**Features:**
- Dropdown selectors for:
  - Metric to display (20 options)
  - Chart type (Donut, Pie, Bar, Column, Line, Area, Table)
  - Theme (Union Blue, Solidarity Red, Success Green, etc.)
  - Comparison mode (Yes/No)
- Real-time chart rendering based on selections
- Saved per-user with User Settings sheet

**Control Cells:**
- A7: Metric selector
- B7: Chart type selector
- C7: Second metric (for comparison)
- D7: Second chart type
- E7: Theme selector
- G7: Comparison toggle

**Data Validation Setup:**
Done via `setupInteractiveDashboardControls()` function

---

## Analytics Sheets

### 5. Analytics Data (Hidden)

**Purpose:** Pre-computed aggregations to speed up dashboard queries

**Sections:**

**A. Grievances by Status (A3:B30)**
- Formula: `=UNIQUE(FILTER('Grievance Log'!E:E, 'Grievance Log'!E:E<>"", 'Grievance Log'!E:E<>"Status"))`
- Count: `=ARRAYFORMULA(IF(A5:A<>"", COUNTIF('Grievance Log'!E:E, A5:A), ""))`

**B. Grievances by Unit (D3:E30) - DYNAMIC**
```javascript
const unitCol = getColumnLetter(GRIEVANCE_COLS.UNIT);  // Y
analytics.getRange("D5").setFormula(
  `=UNIQUE(FILTER('Grievance Log'!${unitCol}:${unitCol},
   'Grievance Log'!${unitCol}:${unitCol}<>"",
   'Grievance Log'!${unitCol}:${unitCol}<>"Unit"))`
);
```

**C. Members by Location (G3:H30)**
- Formula: `=UNIQUE(FILTER('Member Directory'!E:E, 'Member Directory'!E:E<>"", 'Member Directory'!E:E<>"Work Location (Site)"))`

**D. Steward Workload (J3:K30) - DYNAMIC**
```javascript
const stewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);    // AA
const statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);      // E
analytics.getRange("J5").setFormula(
  `=UNIQUE(FILTER('Grievance Log'!${stewardCol}:${stewardCol},
   'Grievance Log'!${stewardCol}:${stewardCol}<>"",
   'Grievance Log'!${stewardCol}:${stewardCol}<>"Assigned Steward (Name)"))`
);
analytics.getRange("K5").setFormula(
  `=ARRAYFORMULA(IF(J5:J<>"",
   COUNTIFS('Grievance Log'!${stewardCol}:${stewardCol}, J5:J,
   'Grievance Log'!${statusCol}:${statusCol}, "Open"), ""))`
);
```

**Sheet State:** Hidden (analytics.hideSheet())

---

### 12. 👨‍⚖️ Steward Workload

**Purpose:** Track steward capacity and caseload

**Columns (11 total):**
```
A: Steward Name
B: Total Cases (all time)
C: Active Cases (currently open)
D: Resolved Cases
E: Win Rate %
F: Avg Days to Resolution
G: Overdue Cases
H: Due This Week
I: Capacity Status (e.g., "Overloaded", "Normal", "Available")
J: Email
K: Phone
```

**Styling:**
- Header: Purple background (#7C3AED), white text
- Tab color: Purple (#7C3AED)

---

### 13. 📈 Trends & Timeline

**Purpose:** Monthly trend analysis over time

**Columns (12 total):**
```
A: Month (YYYY-MM)
B: New Grievances (filed this month)
C: Resolved (closed this month)
D: Win Rate % (for cases resolved this month)
E: Avg Resolution Days
F: Active at Month End
G: Overdue (at month end)
H: New Members (joined this month)
I: Active Members (at month end)
J: Stewards Active
K: Satisfaction Score (avg for month)
L: Trend (↑ Improving, → Stable, ↓ Declining)
```

**Styling:**
- Header: Green background (#059669), white text
- Tab color: Green (#059669)

---

### 14. 🗺️ Location Analytics

**Purpose:** Geographic breakdown of union activity

**Columns (11 total):**
```
A: Location
B: Total Members
C: Active Members
D: Total Grievances
E: Active Grievances
F: Win Rate %
G: Avg Resolution Days
H: Member Satisfaction
I: Stewards Assigned
J: Risk Score (calculated)
K: Priority (High/Medium/Low)
```

**Styling:**
- Header: Teal background (#14B8A6), white text
- Tab color: Teal (#14B8A6)

---

### 15. 📊 Type Analysis

**Purpose:** Grievance breakdown by issue category

**Columns (11 total):**
```
A: Issue Type
B: Total Cases
C: Active
D: Resolved
E: Win Rate %
F: Avg Days to Resolve
G: Most Common Location
H: Top Article Violated
I: Trend (↑↓→)
J: Priority Level
K: Notes
```

**Styling:**
- Header: Blue background (#7EC8E3), white text
- Tab color: Blue (#7EC8E3)

---

### 18. 👥 Member Engagement

**Purpose:** Engagement scoring system

**Columns (12 total):**
```
A: Member ID
B: Name
C: Engagement Score (0-100)
D: Last Contact
E: Meetings Attended
F: Surveys Completed
G: Volunteer Hours
H: Committee Participation
I: Event Attendance
J: Email Open Rate
K: Status (Active, At Risk, Inactive)
L: Notes
```

**Styling:**
- Header: Purple background (#7C3AED), white text
- Tab color: Purple (#7C3AED)

---

### 19. 💰 Cost Impact

**Purpose:** Financial impact tracking

**Columns (10 total):**
```
A: Category
B: Estimated Cost
C: Actual Cost
D: Variance
E: ROI (Return on Investment)
F: Cases Affected
G: Members Benefited
H: Status
I: Quarter
J: Notes
```

**Styling:**
- Header: Red background (#DC2626), white text
- Tab color: Red (#DC2626)

---

## Utility Sheets

### 6. Member Satisfaction

**Purpose:** Survey tracking and satisfaction analytics

**Columns (10 total):**
```
A: Survey ID
B: Member ID
C: Member Name
D: Date Sent
E: Date Completed
F: Overall Satisfaction (1-5)
G: Steward Support (1-5)
H: Communication (1-5)
I: Would Recommend Union (Y/N)
J: Comments
```

**Metrics Section (A10:E14):**
- Average Overall Satisfaction: `=IFERROR(AVERAGE(F4:F1000),"-")`
- Average Steward Support: `=IFERROR(AVERAGE(G4:G1000),"-")`
- Average Communication: `=IFERROR(AVERAGE(H4:H1000),"-")`
- % Would Recommend: `=IFERROR(TEXT(COUNTIF(I4:I1000,"Y")/COUNTA(I4:I1000),"0.0%"),"-")`

**Styling:**
- Header: Green background (#10B981), white text
- Tab color: Green (#10B981)

---

### 7. Feedback, Features & Development Roadmap

**Purpose:** Consolidated bug tracking, feature requests, and development roadmap

**Columns (14 total):**
```
A: Type (dropdown: Bug Report, Feedback, Future Feature, In Progress, Completed, Archived)
B: Submitted/Started (date)
C: Submitted By
D: Priority (dropdown: Critical, High, Medium, Low)
E: Title
F: Description
G: Status (dropdown: New, Under Review, Planned, In Progress, Testing, Completed, Deferred, Cancelled)
H: Progress % (0-100)
I: Complexity (dropdown: Simple, Moderate, Complex, Very Complex)
J: Target Completion
K: Assigned To
L: Blockers
M: Resolution/Notes
N: Last Updated
```

**Data Validations:**
- Column A (Type): ['Bug Report', 'Feedback', 'Future Feature', 'In Progress', 'Completed', 'Archived'] (A4:A1000)
- Column D (Priority): ['Critical', 'High', 'Medium', 'Low'] (D4:D1000)
- Column G (Status): ['New', 'Under Review', 'Planned', 'In Progress', 'Testing', 'Completed', 'Deferred', 'Cancelled'] (G4:G1000)
- Column I (Complexity): ['Simple', 'Moderate', 'Complex', 'Very Complex'] (I4:I1000)

**Column Widths:**
- A=120px, B=110px, C=120px, D=80px, E=200px, F=300px, G=100px, H=90px, I=100px, J=110px, K=120px, L=200px, M=250px, N=110px

**Styling:**
- Header: Purple background (#7C3AED), white text
- Tab color: Purple (#7C3AED)

**Purpose:** Track everything from bug reports to future features in one unified sheet

---

### 11. ⚙️ User Settings

**Purpose:** Store per-user preferences

**Example Settings:**
- Preferred dashboard theme
- Default filters
- Notification preferences
- Last viewed sheet

---

### 20. 📦 Archive

**Purpose:** Store archived/deleted records

**Columns (6 total):**
```
A: Item Type (Member, Grievance, etc.)
B: Item ID
C: Archive Date
D: Archived By
E: Reason
F: Original Data (JSON or text blob)
```

**Styling:**
- Header: Gray background (#6B7280), white text
- Tab color: Gray (#6B7280)

---

### 21. 🔧 Diagnostics

**Purpose:** System health monitoring and error logging

**Columns (7 total):**
```
A: Timestamp
B: Check Type (Setup, Formula, Performance, etc.)
C: Component (Sheet name, function name, etc.)
D: Status (OK, Warning, Error)
E: Details
F: Severity (Low, Medium, High, Critical)
G: Action Needed
```

**Styling:**
- Header: Red background (#DC2626), white text
- Tab color: Red (#DC2626)

---

## Menu System

### Main Menu: "📊 509 Dashboard"

```
📊 509 Dashboard
├── 🔄 Refresh All                    → refreshCalculations()
├── ──────────────────
├── 📊 Dashboards
│   ├── 🎯 Unified Operations Monitor  → showUnifiedOperationsMonitor()
│   ├── 📊 Main Dashboard              → goToDashboard()
│   ├── ✨ Interactive Dashboard       → openInteractiveDashboard()
│   └── 🔄 Refresh Interactive Dashboard → rebuildInteractiveDashboard()
├── ──────────────────
├── 📋 Grievance Tools
│   └── ➕ Start New Grievance         → showStartGrievanceDialog()
├── ──────────────────
├── ⚙️ Admin
│   ├── Seed 20k Members               → SEED_20K_MEMBERS()
│   ├── Seed 5k Grievances             → SEED_5K_GRIEVANCES()
│   ├── ──────────────────
│   ├── Clear All Data                 → clearAllData()
│   └── 🗑️ Nuke All Seed Data         → nukeSeedData()
├── ──────────────────
├── ♿ ADHD Features
│   ├── Hide Gridlines (Focus Mode)   → hideAllGridlines()
│   ├── Show Gridlines                 → showAllGridlines()
│   ├── Reorder Sheets Logically       → reorderSheetsLogically()
│   └── Setup ADHD Defaults            → setupADHDDefaults()
├── 👁️ Column Toggles
│   ├── Toggle Advanced Grievance Columns → toggleGrievanceColumns() [DISABLED - shows info alert]
│   ├── Toggle Level 2 Member Columns     → toggleLevel2Columns()
│   └── Show All Member Columns           → showAllMemberColumns()
├── ──────────────────
└── ❓ Help & Support
    ├── 📚 Getting Started Guide       → showGettingStartedGuide()
    ├── ❓ Help                         → showHelp()
    └── 🔧 Diagnose Setup               → DIAGNOSE_SETUP()
```

**Menu Creation:** `onOpen()` function builds entire menu structure

**Note on toggleGrievanceColumns():**
This function is currently disabled because it expects grievance tracking columns at positions 12-21 that don't exist in the current Member Directory structure. It now shows an informative alert explaining the feature is unavailable and suggests manual column hiding.

---

## Data Validation Rules

### Member Directory Validations

Applied via `setupDataValidations()` using MEMBER_COLS and CONFIG_COLS constants:

```javascript
const memberValidations = [
  { col: MEMBER_COLS.JOB_TITLE, configCol: CONFIG_COLS.JOB_TITLES },       // Job Title (4) → Config A
  { col: MEMBER_COLS.WORK_LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS }, // Work Location (5) → Config B
  { col: MEMBER_COLS.UNIT, configCol: CONFIG_COLS.UNITS },                 // Unit (6) → Config C
  { col: MEMBER_COLS.IS_STEWARD, configCol: CONFIG_COLS.YES_NO },          // Is Steward (10) → Config E
  { col: MEMBER_COLS.ASSIGNED_STEWARD, configCol: CONFIG_COLS.STEWARDS },  // Assigned Steward (14) → Config H
  { col: MEMBER_COLS.INTEREST_LOCAL, configCol: CONFIG_COLS.YES_NO },      // Interest: Local (21) → Config E
  { col: MEMBER_COLS.HOME_TOWN, configCol: CONFIG_COLS.HOME_TOWNS },       // Home Town (22) → Config AF
  { col: MEMBER_COLS.INTEREST_CHAPTER, configCol: CONFIG_COLS.YES_NO },    // Interest: Chapter (23) → Config E
  { col: MEMBER_COLS.INTEREST_ALLIED, configCol: CONFIG_COLS.YES_NO }      // Interest: Allied (24) → Config E
];
```

Multi-select columns use text validation with help text:
- COMMITTEES (11) → Config!I (STEWARD_COMMITTEES)
- PREFERRED_COMM (15) → Config!N (COMM_METHODS)
- BEST_TIME (16) → Config!AE (BEST_TIMES)

Applied to rows 2-5000 for each column.

### Grievance Log Validations

```javascript
const grievanceValidations = [
  { col: GRIEVANCE_COLS.STATUS, configCol: CONFIG_COLS.GRIEVANCE_STATUS },         // Status (5) → Config J
  { col: GRIEVANCE_COLS.CURRENT_STEP, configCol: CONFIG_COLS.GRIEVANCE_STEP },     // Current Step (6) → Config K
  { col: GRIEVANCE_COLS.ARTICLES, configCol: CONFIG_COLS.ARTICLES_VIOLATED },      // Articles Violated (22) → Config M
  { col: GRIEVANCE_COLS.ISSUE_CATEGORY, configCol: CONFIG_COLS.ISSUE_CATEGORY },   // Issue Category (23) → Config L
  { col: GRIEVANCE_COLS.UNIT, configCol: CONFIG_COLS.UNITS },                      // Unit (25) → Config C
  { col: GRIEVANCE_COLS.LOCATION, configCol: CONFIG_COLS.OFFICE_LOCATIONS },       // Work Location (26) → Config B
  { col: GRIEVANCE_COLS.STEWARD, configCol: CONFIG_COLS.STEWARDS }                 // Assigned Steward (27) → Config H
];
```

Applied to rows 2-5000 for each column.

### Other Sheet Validations

**Feedback & Development:**
- Type (A4:A1000): 6 options
- Priority (D4:D1000): 4 options
- Status (G4:G1000): 8 options
- Complexity (I4:I1000): 4 options

**KPI Performance Dashboard:**
- Status (F4:F1000): ['On Track', 'At Risk', 'Off Track', 'Exceeding']

**Interactive Dashboard:**
- Metric selectors: 20 metric options
- Chart type: 7 chart type options
- Theme: 6 theme options
- Comparison: Yes/No

---

## Formula System

### Auto-Calculated Formulas

Set via `setupFormulasAndCalculations()` for rows 2-100:

**Grievance Log Formulas:**

```javascript
// Filing Deadline (Column H)
grievanceLog.getRange(row, 8).setFormula(`=IF(G${row}<>"",G${row}+21,"")`);

// Step I Decision Due (Column J)
grievanceLog.getRange(row, 10).setFormula(`=IF(I${row}<>"",I${row}+30,"")`);

// Step II Appeal Due (Column L)
grievanceLog.getRange(row, 12).setFormula(`=IF(K${row}<>"",K${row}+10,"")`);

// Step II Decision Due (Column N)
grievanceLog.getRange(row, 14).setFormula(`=IF(M${row}<>"",M${row}+30,"")`);

// Step III Appeal Due (Column P)
grievanceLog.getRange(row, 16).setFormula(`=IF(O${row}<>"",O${row}+30,"")`);

// Days Open (Column S)
grievanceLog.getRange(row, 19).setFormula(
  `=IF(I${row}<>"",IF(R${row}<>"",R${row}-I${row},TODAY()-I${row}),"")`
);

// Next Action Due (Column T)
grievanceLog.getRange(row, 20).setFormula(
  `=IF(E${row}="Open",IF(F${row}="Step I",J${row},IF(F${row}="Step II",N${row},IF(F${row}="Step III",P${row},H${row}))),"")`
);

// Days to Deadline (Column U)
grievanceLog.getRange(row, 21).setFormula(`=IF(T${row}<>"",T${row}-TODAY(),"")`);
```

**Member Directory Formulas:**

```javascript
// Has Open Grievance? (Column Z / 26)
memberDir.getRange(row, 26).setFormula(
  `=IF(COUNTIFS('Grievance Log'!B:B,A${row},'Grievance Log'!E:E,"Open")>0,"Yes","No")`
);

// Grievance Status Snapshot (Column AA / 27)
memberDir.getRange(row, 27).setFormula(
  `=IFERROR(INDEX('Grievance Log'!E:E,MATCH(A${row},'Grievance Log'!B:B,0)),"")`
);

// Next Grievance Deadline (Column AB / 28)
memberDir.getRange(row, 28).setFormula(
  `=IFERROR(INDEX('Grievance Log'!T:T,MATCH(A${row},'Grievance Log'!B:B,0)),"")`
);
```

---

## Seed Data Functions

### SEED_20K_MEMBERS()

**Purpose:** Generate 20,000 realistic member records

**Process:**
1. Prompts user for confirmation (destructive operation)
2. Loads data from Config sheet (job titles, locations, units, supervisors, managers, stewards)
3. Validates config data is complete
4. Generates members in batches of 1,000
5. For each member:
   - Generates Member ID (M000001 - M020000)
   - Randomly selects from name lists (40 first names, 40 last names)
   - Assigns job title, location, unit (from Config)
   - Generates email (firstname.lastname{number}@union.org)
   - Generates phone number (555 area code)
   - Random steward status (5% are stewards)
   - Random engagement dates (last 90 days)
   - Random open rates (60-100%)
   - Random volunteer hours (0-50)
   - Random interest flags
   - Communication preferences and best contact times
6. Progress toasts every 1,000 records
7. Error handling with retry logic

**Batch Size:** 1,000 rows per write (prevents timeout)

**Error Handling:** Retries once with 1-second delay if batch write fails

---

### SEED_5K_GRIEVANCES()

**Purpose:** Generate 5,000 realistic grievance records

**Process:**
1. Prompts user for confirmation
2. Loads member data ONCE (critical for performance)
3. Extracts all member IDs into array
4. Loads config data (statuses, steps, articles, categories, stewards)
5. Validates config data
6. Generates grievances in batches of 500
7. For each grievance:
   - Selects random member from loaded member data
   - Generates Grievance ID (G-000001 - G-005000)
   - Random status, step, dates (last 365 days)
   - Calculates all deadlines based on contract rules:
     * Filing deadline = incident + 21 days
     * Step I decision = filed + 30 days
     * Step II appeal due = Step I decision + 10 days
     * Step II decision = Step II appeal + 30 days
     * Step III appeal = Step II decision + 30 days
   - Calculates days open, next action, days to deadline
   - Random article, category, location, steward
   - For closed cases: resolution with Win/Lost/Settled prefix
     * **CRITICAL:** Resolution must contain "Won", "Lost", or "Settled" for Win Rate formulas to work
     * Options: "Won - Resolved favorably", "Won - Full remedy granted", "Lost - No violation found", "Lost - Withdrawn by member", "Settled - Partial remedy", "Settled - Compromise reached"
8. Progress toasts every 500 records
9. After completion: calls `updateMemberDirectorySnapshots()`
10. Error handling with retry logic

**Critical Fix (Win Rate Bug):**
Resolution text MUST include "Won" prefix for Win Rate formulas to work correctly:
```javascript
const resolution = isClosed ? [
  "Won - Resolved favorably",
  "Won - Full remedy granted",
  "Lost - No violation found",
  "Lost - Withdrawn by member",
  "Settled - Partial remedy",
  "Settled - Compromise reached"
][Math.floor(Math.random() * 6)] : "";
```

**Batch Size:** 500 rows per write

**Performance:** Loads member data once (not in loop) - critical for speed

---

### updateMemberDirectorySnapshots()

**Purpose:** Update Member Directory columns Z, AA, AB, AC, AD, AE with grievance data

**Process:**
1. Reads all member IDs
2. Reads all grievance data (all 28 columns)
3. Builds memberSnapshots object mapping member ID to:
   - Status (prioritizes Open/Filed/Pending)
   - Next deadline (earliest upcoming deadline)
   - Steward who contacted
4. For each member:
   - Sets grievance status, next deadline
   - Generates random recent contact date (last 14 days)
   - Random contact notes
5. Writes all updates in single batch (rows 2 to last member)

**Columns Updated:**
- Column 25 (Y): Grievance status snapshot
- Column 26 (Z): Next grievance deadline
- Column 27 (AA): Most recent steward contact date
- Column 28 (AB): Steward who contacted
- Column 29 (AC): Notes from steward contact

**Called By:** SEED_5K_GRIEVANCES() after seeding complete

---

### clearAllData()

**Purpose:** Delete all member and grievance data (keeps structure)

**Process:**
1. Prompts for confirmation
2. Clears Member Directory rows 2 to last row
3. Clears Grievance Log rows 2 to last row
4. Keeps headers intact
5. Toast notification

---

### nukeSeedData()

**Purpose:** Complete nuclear option - delete all seed data

**Implementation:** Currently calls clearAllData()

**Future Enhancement:** Could also reset Config, clear Analytics cache, etc.

---

## Color Scheme

### COLORS Constant

```javascript
const COLORS = {
  // Primary brand colors
  PRIMARY_BLUE: "#7EC8E3",      // Light blue for general use
  PRIMARY_PURPLE: "#7C3AED",    // Main brand purple
  UNION_GREEN: "#059669",       // Union/success green
  SOLIDARITY_RED: "#DC2626",    // Alert/urgent red

  // Accent colors
  ACCENT_TEAL: "#14B8A6",       // Teal for variety
  ACCENT_PURPLE: "#7C3AED",     // Duplicate of primary (consolidate?)
  ACCENT_ORANGE: "#F97316",     // Orange for warnings/attention
  ACCENT_YELLOW: "#FCD34D",     // Yellow for highlights

  // Neutral colors
  WHITE: "#FFFFFF",
  LIGHT_GRAY: "#F3F4F6",        // Backgrounds
  BORDER_GRAY: "#D1D5DB",       // Borders
  TEXT_GRAY: "#6B7280",         // Secondary text
  TEXT_DARK: "#1F2937",         // Primary text

  // Specialty colors
  CARD_BG: "#FAFAFA",           // Card backgrounds
  INFO_LIGHT: "#E0E7FF",        // Info messages
  SUCCESS_LIGHT: "#D1FAE5",     // Success messages
  HEADER_BLUE: "#3B82F6",       // Alternative header
  HEADER_GREEN: "#10B981"       // Alternative header
};
```

### Usage Guidelines

**Sheet Tab Colors:**
- Config: Blue (#2563EB)
- Member Directory: Green (#059669)
- Grievance Log: Red (#DC2626)
- Dashboard: Purple (#7C3AED)
- Executive Dashboard: Purple (#7C3AED)
- KPI Performance: Green (#059669)
- Steward Workload: Purple (#7C3AED)
- Trends: Green (#059669)
- Location: Teal (#14B8A6)
- Type Analysis: Blue (#7EC8E3)
- Member Engagement: Purple (#7C3AED)
- Cost Impact: Red (#DC2626)
- Feedback: Purple (#7C3AED)
- Archive: Gray (#6B7280)
- Diagnostics: Red (#DC2626)

**Header Colors:**
- Critical/Alert: Red (SOLIDARITY_RED)
- Success/Positive: Green (UNION_GREEN)
- Informational: Purple (PRIMARY_PURPLE)
- Warning: Orange (ACCENT_ORANGE)
- Neutral: Gray (LIGHT_GRAY)

---

## Column Mapping System

**🔴 CRITICAL: This system is MANDATORY for ALL column references**

### MEMBER_COLS Constant

**Purpose:** Single source of truth for all Member Directory column positions (31 columns)

**Implementation (from Constants.gs):**

```javascript
const MEMBER_COLS = {
  // 31 columns total - Added: COMMITTEES, HOME_TOWN, PREFERRED_COMM, BEST_TIME
  MEMBER_ID: 1,                    // A
  FIRST_NAME: 2,                   // B
  LAST_NAME: 3,                    // C
  JOB_TITLE: 4,                    // D
  WORK_LOCATION: 5,                // E
  UNIT: 6,                         // F
  OFFICE_DAYS: 7,                  // G
  EMAIL: 8,                        // H
  PHONE: 9,                        // I
  IS_STEWARD: 10,                  // J
  COMMITTEES: 11,                  // K - Multi-select: which committees steward is in
  SUPERVISOR: 12,                  // L
  MANAGER: 13,                     // M
  ASSIGNED_STEWARD: 14,            // N
  PREFERRED_COMM: 15,              // O - Multi-select: preferred communication methods
  BEST_TIME: 16,                   // P - Multi-select: best times to reach member
  LAST_VIRTUAL_MTG: 17,            // Q
  LAST_INPERSON_MTG: 18,           // R
  OPEN_RATE: 19,                   // S
  VOLUNTEER_HOURS: 20,             // T
  INTEREST_LOCAL: 21,              // U
  HOME_TOWN: 22,                   // V - Member's home town
  INTEREST_CHAPTER: 23,            // W
  INTEREST_ALLIED: 24,             // X
  HAS_OPEN_GRIEVANCE: 25,          // Y
  GRIEVANCE_STATUS: 26,            // Z
  NEXT_DEADLINE: 27,               // AA
  RECENT_CONTACT_DATE: 28,         // AB
  CONTACT_STEWARD: 29,             // AC
  CONTACT_NOTES: 30,               // AD
  START_GRIEVANCE: 31              // AE - Checkbox to start grievance with prepopulated member info
};
```

### GRIEVANCE_COLS Constant

**Purpose:** Single source of truth for all Grievance Log column positions (28 columns)

**Why This Exists:**
- Hardcoded column letters (AB:AB, Y:Y, etc.) break if columns are reordered
- Formulas scattered throughout codebase would all need manual updates
- Dynamic system allows updating one constant to fix all formulas

**Implementation:**

```javascript
const GRIEVANCE_COLS = {
  GRIEVANCE_ID: 1,      // A
  MEMBER_ID: 2,         // B
  FIRST_NAME: 3,        // C
  LAST_NAME: 4,         // D
  STATUS: 5,            // E
  CURRENT_STEP: 6,      // F
  INCIDENT_DATE: 7,     // G
  FILING_DEADLINE: 8,   // H
  DATE_FILED: 9,        // I
  STEP1_DUE: 10,        // J
  STEP1_RCVD: 11,       // K
  STEP2_APPEAL_DUE: 12, // L
  STEP2_APPEAL_FILED: 13, // M
  STEP2_DUE: 14,        // N
  STEP2_RCVD: 15,       // O
  STEP3_APPEAL_DUE: 16, // P
  STEP3_APPEAL_FILED: 17, // Q
  DATE_CLOSED: 18,      // R
  DAYS_OPEN: 19,        // S
  NEXT_ACTION_DUE: 20,  // T
  DAYS_TO_DEADLINE: 21, // U
  ARTICLES: 22,         // V
  ISSUE_CATEGORY: 23,   // W
  MEMBER_EMAIL: 24,     // X
  UNIT: 25,             // Y
  LOCATION: 26,         // Z
  STEWARD: 27,          // AA
  RESOLUTION: 28        // AB
};
```

### getColumnLetter() Helper Function

**Purpose:** Convert column number to letter notation

**Implementation:**
```javascript
function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}
```

**Examples:**
- `getColumnLetter(1)` → "A"
- `getColumnLetter(26)` → "Z"
- `getColumnLetter(27)` → "AA"
- `getColumnLetter(28)` → "AB"
- `getColumnLetter(52)` → "AZ"

### Dynamic Formula Examples

**Example 1: Member Directory (Old vs New)**

**❌ Old (Hardcoded - NEVER DO THIS):**
```javascript
["Active Stewards", "=COUNTIF('Member Directory'!J:J,\"Yes\")"]
```

**✅ New (Dynamic - ALWAYS DO THIS):**
```javascript
const isStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
["Active Stewards", `=COUNTIF('Member Directory'!${isStewardCol}:${isStewardCol},"Yes")`]
```

**Example 2: Grievance Log (Old vs New)**

**❌ Old (Hardcoded - NEVER DO THIS):**
```javascript
["Win Rate", "=COUNTIFS('Grievance Log'!AB:AB,\"*Won*\")"]
```

**✅ New (Dynamic - ALWAYS DO THIS):**
```javascript
const resolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
["Win Rate", `=COUNTIFS('Grievance Log'!${resolutionCol}:${resolutionCol},"*Won*")`]
```

### Where Dynamic Columns Are Used (100% Coverage)

**✅ Member Directory (All References Dynamic):**
- Main Dashboard: MEMBER_ID, IS_STEWARD, OPEN_RATE, VOLUNTEER_HOURS, LAST_VIRTUAL_MTG, LAST_INPERSON_MTG, INTEREST_LOCAL, INTEREST_CHAPTER
- Executive Dashboard: MEMBER_ID, IS_STEWARD
- Analytics Data Sheet: All member-related aggregations
- Verification: `grep "'Member Directory'![A-Z]:[A-Z]" *.gs` → 0 matches ✅

**✅ Grievance Log (All References Dynamic):**
- Analytics Data Sheet: UNIT, STEWARD, STATUS
- Executive Dashboard: RESOLUTION, STATUS, DAYS_OPEN, DAYS_TO_DEADLINE
- Main Dashboard: STATUS, DATE_CLOSED, DAYS_OPEN
- All QUERY formulas use dynamic column ranges
- Verification: `grep "'Grievance Log'![A-Z]:[A-Z]" *.gs` → 0 matches ✅

**Status: 100% Dynamic - NO hardcoded column references exist**

### Benefits

1. **Future-proof:** Add/remove/reorder ANY columns by updating MEMBER_COLS or GRIEVANCE_COLS
2. **No hunting:** All column positions in two centralized constants
3. **Self-documenting:** Clear mapping of column names to positions (MEMBER_ID vs "A")
4. **Error-proof:** No typos in column letters, compiler catches missing constants
5. **Maintainable:** Change once, fixes everywhere automatically
6. **100% Coverage:** Every single column reference uses this system (verified with grep)

---

## File Architecture

### Project Structure

```
509-dashboard/
├── Code.gs                      # Modular version (primary)
├── Complete509Dashboard.gs      # Single-file version (backup/deployment)
├── ColumnToggles.gs            # Column visibility functions
├── InteractiveDashboard.gs     # Interactive dashboard logic
├── ADHDEnhancements.gs         # ADHD-friendly features
├── MenuItems.gs                # Menu creation (if separate)
├── SeedData.gs                 # Seed functions (if separate)
└── FEATURES.md                 # This document
```

### Code.gs vs Complete509Dashboard.gs

**Code.gs:**
- Modular architecture (functions may be split across files)
- Primary development version
- Easier to navigate and maintain
- Recommended for development

**Complete509Dashboard.gs:**
- Single-file version containing all code
- Used for deployment to Google Sheets
- Backup/reference version
- Should be kept in sync with Code.gs

**Sync Strategy:**
- Develop in Code.gs + separate module files
- Before deployment, sync changes to Complete509Dashboard.gs
- Both files should have identical functionality

### Key Functions by File

**Code.gs / Complete509Dashboard.gs:**
- `CREATE_509_DASHBOARD()` - Main setup function
- `DIAGNOSE_SETUP()` - Health check function
- All sheet creation functions (createMemberDirectory, createGrievanceLog, etc.)
- `setupDataValidations()` - Apply all validations
- `setupFormulasAndCalculations()` - Set formulas for first 100 rows
- `SEED_20K_MEMBERS()` - Generate member data
- `SEED_5K_GRIEVANCES()` - Generate grievance data
- `updateMemberDirectorySnapshots()` - Update member columns from grievances
- `onOpen()` - Create menu system

**ColumnToggles.gs:**
- `toggleGrievanceColumns()` - [DISABLED] Show/hide grievance columns
- `toggleLevel2Columns()` - Show/hide Level 2 engagement columns
- `showAllMemberColumns()` - Unhide all columns in Member Directory

**InteractiveDashboard.gs:**
- `createInteractiveDashboardSheet()` - Build interactive dashboard
- `setupInteractiveDashboardControls()` - Add dropdown validations
- `rebuildInteractiveDashboard()` - Refresh based on user selections

**ADHDEnhancements.gs:**
- `hideAllGridlines()` - Focus mode
- `showAllGridlines()` - Show gridlines
- `reorderSheetsLogically()` - Reorder sheets in logical order
- `setupADHDDefaults()` - Apply ADHD-friendly defaults

---

## Known Issues & Limitations

### Current Issues

**1. Toggle Grievance Columns - DISABLED**
- Function expects columns 12-21 to contain grievance tracking data
- Current Member Directory structure has engagement data in those positions
- Function now shows informative alert instead of trying to hide wrong columns
- Users can manually hide/show columns via right-click

**2. Member Directory Column Groups**
- Old versions may have residual column groups from previous structures
- Solution: createMemberDirectory() now deletes and recreates sheet (not just clears)
- This ensures clean slate without legacy formatting

**3. Win Rate Formula Dependency**
- Win Rate formulas depend on Resolution Summary containing "Won", "Lost", or "Settled"
- If seed data changes resolution format, Win Rate will break
- Solution: SEED_5K_GRIEVANCES() now generates proper resolution prefixes

### Limitations

**Performance:**
- Seed functions can take 2-3 minutes for 20k members + 5k grievances
- Google Sheets has 6-minute execution limit
- Batch processing (1000/500 rows) prevents timeout

**Formula Rows:**
- Auto-calculated formulas only set for first 100 rows
- Manually add formulas if more than 100 open grievances/members
- Future: Could extend to 1000 rows or use ARRAYFORMULA

**Google Sheets Limits:**
- Max 10 million cells per spreadsheet
- Max 50,000 characters per cell
- Max 40,000 new rows per day (via API)

---

## Future Enhancements

### Planned Features

**1. Real-Time Notifications**
- Email alerts for approaching deadlines
- Slack/Teams integration
- SMS notifications for critical cases

**2. Advanced Analytics**
- Predictive modeling (case outcome prediction)
- Trend analysis with machine learning
- Sentiment analysis on member feedback

**3. Mobile App**
- Native iOS/Android apps
- Offline mode
- Push notifications

**4. Integration**
- Union dues payment system integration
- Document management system (contracts, forms)
- Calendar integration for deadlines

**5. Automation**
- Auto-assign stewards based on workload
- Auto-generate grievance letters
- Auto-update member engagement scores

**6. Enhanced Member Engagement**
- Member portal (view own grievances)
- Survey builder and distribution
- Event registration system

### Technical Debt

**1. Extend Formula Rows**
- Currently only first 100 rows have formulas
- Should extend to 1000 rows or use ARRAYFORMULA

**2. Member Directory Columns**
- Add actual grievance tracking columns (Total Grievances, Active Grievances, etc.)
- Re-enable toggleGrievanceColumns() function
- Requires adding 10 new calculated columns

**3. Error Logging**
- Implement comprehensive error logging to Diagnostics sheet
- Track all seed operations, formula errors, validation failures

**4. Performance Optimization**
- Cache Analytics Data calculations
- Lazy-load dashboard charts
- Implement pagination for large data views

---

## Reconstruction Guide

### How to Rebuild This Exact System

**Prerequisites:**
1. Google account with Google Sheets access
2. Apps Script editor access
3. This FEATURES.md document
4. All .gs files from repository

**Step-by-Step:**

1. **Create New Google Sheet**
   - Name: "509 Dashboard"
   - Open Tools → Script Editor

2. **Copy All Code**
   - Copy entire contents of Code.gs
   - Paste into Code.gs in Apps Script editor
   - Copy ColumnToggles.gs → Create new file, paste
   - Copy InteractiveDashboard.gs → Create new file, paste
   - Copy ADHDEnhancements.gs → Create new file, paste

3. **Save and Deploy**
   - Save all files
   - Run function: CREATE_509_DASHBOARD()
   - Grant permissions when prompted
   - Wait for setup to complete (2-3 minutes)

4. **Verify Setup**
   - Run function: DIAGNOSE_SETUP()
   - Check that all 21 sheets exist
   - Verify menu appears: "📊 509 Dashboard"

5. **Seed Test Data**
   - Menu → Admin → Seed 20k Members (wait 2-3 min)
   - Menu → Admin → Seed 5k Grievances (wait 1-2 min)

6. **Test Functionality**
   - Open Main Dashboard - verify metrics calculate
   - Open Executive Dashboard - verify Win Rate shows correctly
   - Open Interactive Dashboard - verify dropdowns work
   - Test menu items

7. **Customize** (Optional)
   - Update Config sheet with real job titles, locations, etc.
   - Clear seed data: Menu → Admin → Clear All Data
   - Import real member/grievance data

### Verification Checklist

- [ ] 21 sheets exist with correct names
- [ ] Config sheet has 13 columns with sample data
- [ ] Member Directory has 31 columns with correct headers
- [ ] Grievance Log has 28 columns with correct headers
- [ ] Dashboard shows live metrics (after seeding)
- [ ] Menu "📊 509 Dashboard" appears
- [ ] All menu items work without errors
- [ ] Data validations work (try entering invalid data)
- [ ] Formulas calculate correctly (check Win Rate)
- [ ] Executive Dashboard shows Quick Stats + Detailed KPIs
- [ ] KPI Performance Dashboard has 12 columns
- [ ] Seed functions complete without errors
- [ ] Win Rate shows >0% after seeding (not 0%)
- [ ] **🔴 CRITICAL:** 100% dynamic column references verified:
  - [ ] `grep "'Member Directory'![A-Z]:[A-Z]" *.gs` → 0 matches
  - [ ] `grep "'Grievance Log'![A-Z]:[A-Z]" *.gs` → 0 matches
  - [ ] MEMBER_COLS constant defined with all 31 columns
  - [ ] GRIEVANCE_COLS constant defined with all 28 columns

---

## Appendix: Constants Reference

### SHEETS Constant

```javascript
const SHEETS = {
  CONFIG: "Config",
  MEMBER_DIR: "Member Directory",
  GRIEVANCE_LOG: "Grievance Log",
  DASHBOARD: "Dashboard",
  ANALYTICS: "Analytics Data",
  FEEDBACK: "Feedback & Development",
  MEMBER_SATISFACTION: "Member Satisfaction",
  INTERACTIVE_DASHBOARD: "🎯 Interactive (Your Custom View)",
  STEWARD_WORKLOAD: "👨‍⚖️ Steward Workload",
  TRENDS: "📈 Trends & Timeline",
  LOCATION: "🗺️ Location Analytics",
  TYPE_ANALYSIS: "📊 Type Analysis",
  EXECUTIVE_DASHBOARD: "💼 Executive Dashboard",
  KPI_PERFORMANCE: "📊 KPI Performance Dashboard",
  MEMBER_ENGAGEMENT: "👥 Member Engagement",
  COST_IMPACT: "💰 Cost Impact",
  ARCHIVE: "📦 Archive",
  DIAGNOSTICS: "🔧 Diagnostics"
};
```

**Total Sheets:** 21 (includes hidden Analytics Data)
**Recent Changes:** Consolidated from 25 to 21 sheets

---

## Changelog

### Version 2.4 (Current)

**🔄 CONFIG & DOCUMENTATION SYNC: Constants Updated**

**Changes:**
- Updated AI_REFERENCE.md to match current Constants.gs structure
- Config sheet now documented with all 32 columns (was showing old 13-column structure)
- MEMBER_COLS now includes COMMITTEES (K), HOME_TOWN (V), START_GRIEVANCE (AE)
- Fixed incorrect help text in Code.gs referencing wrong Config columns:
  - Committees: Changed "column AD" to "column I" (STEWARD_COMMITTEES)
  - Preferred Comm: Changed "column M" to "column N" (COMM_METHODS)
- Updated Data Validation Rules documentation to use CONFIG_COLS/MEMBER_COLS constants
- Documented multi-status "Has Open Grievance" formula (Open, Pending Info, Appealed, In Arbitration)

**Files Changed:**
- Code.gs: Fixed help text in setupDataValidations()
- AI_REFERENCE.md: Updated Config, Member Directory, and Data Validation sections

---

### Version 2.3

**🔴 CRITICAL BUG FIX: hideGridlines() TypeError Resolved**

**Issue:**
- Runtime error: `TypeError: sheet.hideGridlines is not a function`
- Affected CREATE_509_DASHBOARD and all gridline-related functions
- Method `hideGridlines()` does not exist in Google Apps Script API

**Fix:**
- Replaced all `sheet.hideGridlines()` calls with `sheet.setHiddenGridlines(true)`
- This is the correct Google Apps Script method for hiding gridlines
- Fixed 9 occurrences across 3 files

**Files Changed:**
- Complete509Dashboard.gs: Fixed 3 instances (lines 4446, 4610, 4765)
- ADHDEnhancements.gs: Fixed 3 instances (lines 29, 189, 344)
- ConsolidatedDashboard.gs: Fixed 3 instances (lines 11709, 11869, 12024)
- AI_REFERENCE.md: Updated version and changelog

**Commit:**
- Fix hideGridlines TypeError - use setHiddenGridlines(true) instead

---

### Version 2.2

**🔴 CRITICAL UPDATE: All Runtime Errors Fixed**

**Comprehensive Code Review Completed:**
- Reviewed entire codebase for stubs, dead ends, and errors
- Found and fixed 9 critical runtime errors that would cause crashes
- Fixed 6 wrong SHEETS constant references
- Replaced mock data with real calculations
- Removed disabled/broken menu features

**Critical Fixes:**

1. **Missing Functions Added (3):**
   - `recalcGrievanceRow()` in GrievanceWorkflow.gs
   - `recalcMemberRow()` in GrievanceWorkflow.gs
   - `rebuildDashboard()` in SeedNuke.gs

2. **SHEETS Constants Fixed (6 wrong references):**
   - `SHEETS.EXECUTIVE` → `SHEETS.EXECUTIVE_DASHBOARD`
   - `SHEETS.KPI_BOARD` → `SHEETS.KPI_PERFORMANCE`
   - Removed: `SHEETS.PERFORMANCE`, `SHEETS.QUICK_STATS`, `SHEETS.FUTURE_FEATURES`, `SHEETS.PENDING_FEATURES`
   - Added missing: `SHEETS.MEMBER_SATISFACTION` to reorder list

3. **Hardcoded Column Reference Fixed:**
   - Replaced hardcoded column 21 with named constant `CONFIG_STEWARD_INFO_COL`

4. **Menu Cleanup:**
   - Removed non-functional "Toggle Advanced Grievance Columns" menu item

5. **Mock Data Replaced:**
   - UnifiedOperationsMonitor.gs now uses real win/loss calculations instead of 75% fake data

6. **UnifiedOperationsMonitor.gs Made Fully Dynamic (HIGH PRIORITY):**
   - Fixed 61 hardcoded grievance array indices: g[4], g[22], g[27], etc.
   - Fixed 7 hardcoded member array indices: m[0], m[9], m[20], etc.
   - All array access now uses `MEMBER_COLS - 1` / `GRIEVANCE_COLS - 1`
   - Example: `g[4]` → `g[GRIEVANCE_COLS.STATUS - 1]`

7. **Executive Dashboard Column References (2 instances):**
   - Fixed hardcoded 'Member Directory'!A2:A → dynamic execMemberIdCol
   - Fixed hardcoded 'Grievance Log'!A2:A → dynamic grievanceIdCol

**Files Changed:**
- ADHDEnhancements.gs: Fixed SHEETS constants
- GrievanceWorkflow.gs: Added missing functions, fixed hardcoded column, made formulas fully dynamic
- SeedNuke.gs: Added rebuildDashboard() function
- UnifiedOperationsMonitor.gs: Replaced mock data, fixed ALL 68 hardcoded array indices
- Code.gs: Removed disabled menu item, fixed Executive Dashboard hardcoded columns
- Complete509Dashboard.gs: Removed disabled menu item, fixed Executive Dashboard (parity maintained)
- AI_REFERENCE.md: Added Code Quality section, updated changelog

**Commits:**
- cb36266: Fix all critical code issues from comprehensive review
- 083d253: Eliminate ALL hardcoded column references - 100% dynamic
- cdf34d8: Fix UnifiedOperationsMonitor.gs - 100% dynamic array access

---

### Version 2.1

**🔴 CRITICAL UPDATE: 100% Dynamic Column System Complete**

**Major Changes:**
- ✅ Added MEMBER_COLS constant with all 31 Member Directory columns
- ✅ Converted ALL remaining hardcoded column references to dynamic
- ✅ 100% dynamic coverage achieved - ZERO hardcoded references remain
- ✅ Both Code.gs and Complete509Dashboard.gs updated
- ✅ Test suite added from testing branch

**Files Changed:**
- Code.gs: Added MEMBER_COLS constant, updated all Member Directory formulas
- Complete509Dashboard.gs: Same changes (100% parity maintained)
- AI_REFERENCE.md: Added comprehensive dynamic column documentation

**Verification:**
- `grep "'Member Directory'![A-Z]:[A-Z]" *.gs` → 0 matches ✅
- `grep "'Grievance Log'![A-Z]:[A-Z]" *.gs` → 0 matches ✅
- All 31 Member Directory columns now use MEMBER_COLS
- All 28 Grievance Log columns use GRIEVANCE_COLS

**Commits:**
- 58859ed: Make all column references fully dynamic with MEMBER_COLS
- b8e823f: Add comprehensive test suite from testing branch

---

### Version 2.0

**Major Changes:**
- Consolidated 25 sheets → 21 sheets
- ✅ Dynamic column mapping system for GRIEVANCE_COLS
- Fixed critical Win Rate formula bug
- Disabled legacy toggleGrievanceColumns() function
- Fixed column group error (delete/recreate Member Directory)
- Merged Feedback + Future Features + Pending Features
- Merged Executive Summary + Quick Stats → Executive Dashboard
- Merged Performance Metrics + KPI Board → KPI Performance Dashboard

**Files Changed:**
- Code.gs: Added GRIEVANCE_COLS, getColumnLetter(), Grievance formulas dynamic
- Complete509Dashboard.gs: Same changes for consistency (100% parity)
- ColumnToggles.gs: Disabled toggleGrievanceColumns()

**Bug Fixes:**
- Win Rate now shows correctly (resolution text includes "Won"/"Lost"/"Settled")
- Column group error fixed (delete/recreate sheet)
- ✅ Grievance Log column references now dynamic

**Dynamic Column Implementation (COMPLETE):**

Commit f1b28a9 completed the dynamic column conversion. ALL formulas now use dynamic references:

**Main Dashboard:**
- Grievance Metrics: STATUS (E), DATE_CLOSED (R), DAYS_OPEN (S)
- QUERY formula: All column letters dynamic (A, C, T, U, E)
- Range: A:U → dynamic first:last column

**Analytics Data Sheet:**
- Status section: STATUS (E) - fully dynamic
- Unit section: UNIT (Y) - fully dynamic
- Steward section: STEWARD (AA), STATUS (E) - fully dynamic

**Member Directory:**
- Has Open Grievance: MEMBER_ID (B), STATUS (E)
- Grievance Status Snapshot: STATUS (E), MEMBER_ID (B)
- Next Grievance Deadline: NEXT_ACTION_DUE (T), MEMBER_ID (B)

**Executive Dashboard:**
- Quick Stats: RESOLUTION (AB), STATUS (E), DAYS_OPEN (S), DAYS_TO_DEADLINE (U)

**Columns Using Dynamic References:**
- GRIEVANCE_ID (A)
- MEMBER_ID (B)
- FIRST_NAME (C)
- STATUS (E)
- DATE_CLOSED (R)
- DAYS_OPEN (S)
- NEXT_ACTION_DUE (T)
- DAYS_TO_DEADLINE (U)
- UNIT (Y)
- STEWARD (AA)
- RESOLUTION (AB)

**Verification:** `grep "'Grievance Log'![A-Z]+:[A-Z]+"` returns ZERO hardcoded references.

**How to Reorder Columns:**
1. Update GRIEVANCE_COLS constant with new column positions
2. All formulas automatically adjust - no manual updates needed
3. This is the ONLY place column positions are defined

---

## Code Quality & Known Issues

### Recent Code Review (Version 2.2)

A comprehensive code review was conducted covering stubs, dead ends, and errors. All **critical** and **high-priority** issues have been resolved.

**✅ Fixed Issues (Resolved in Version 2.2):**

1. **Missing Function Definitions (CRITICAL)** - FIXED
   - Added `recalcGrievanceRow()` to GrievanceWorkflow.gs
   - Added `recalcMemberRow()` to GrievanceWorkflow.gs
   - Added `rebuildDashboard()` to SeedNuke.gs

2. **Wrong SHEETS Constants (CRITICAL)** - FIXED
   - Fixed `SHEETS.EXECUTIVE` → `SHEETS.EXECUTIVE_DASHBOARD`
   - Fixed `SHEETS.KPI_BOARD` → `SHEETS.KPI_PERFORMANCE`
   - Removed references to non-existent `SHEETS.PERFORMANCE`, `SHEETS.QUICK_STATS`, `SHEETS.FUTURE_FEATURES`, `SHEETS.PENDING_FEATURES`
   - Added `SHEETS.MEMBER_SATISFACTION` to reorder list

3. **Hardcoded Config Column** - FIXED
   - Replaced hardcoded column 21 with named constant `CONFIG_STEWARD_INFO_COL`

4. **Disabled Menu Feature** - FIXED
   - Removed non-functional "Toggle Advanced Grievance Columns" from menu

5. **Mock Data** - FIXED
   - Replaced 75% fake win rate with real calculations in UnifiedOperationsMonitor.gs

6. **UnifiedOperationsMonitor.gs Hardcoded Array Indices (HIGH PRIORITY)** - FIXED
   - Replaced 61 hardcoded grievance array indices (g[4], g[22], g[27], etc.)
   - Replaced 7 hardcoded member array indices (m[0], m[9], m[20], etc.)
   - All array access now uses `MEMBER_COLS - 1` and `GRIEVANCE_COLS - 1` pattern
   - Example: `g[4]` → `g[GRIEVANCE_COLS.STATUS - 1]`

**⚠️ Known Technical Debt (Non-Critical):**

None - All issues resolved!

**Verification Commands:**

```bash
# Verify no hardcoded sheet column references (should return 0)
grep "'Member Directory'![A-Z]:[A-Z]" *.gs | wc -l
grep "'Grievance Log'![A-Z]:[A-Z]" *.gs | wc -l

# Verify no hardcoded array indices (should return 0)
grep "g\[[0-9]\+\]\|m\[[0-9]\+\]" UnifiedOperationsMonitor.gs | wc -l
```

---

## Support & Maintenance

### How to Debug Issues

**1. Run Diagnostics:**
```
Menu → Help & Support → 🔧 Diagnose Setup
```
This checks all 21 sheets exist and reports missing ones.

**2. Check Execution Log:**
```
Apps Script Editor → View → Logs
```
All errors logged here with timestamps.

**3. Common Errors:**

**"A column group does not exist..."**
- Solution: Delete Member Directory sheet, run CREATE_509_DASHBOARD()

**Win Rate shows 0%:**
- Check Resolution Summary column contains "Won", "Lost", or "Settled"
- Re-seed grievances with updated SEED_5K_GRIEVANCES()

**Formulas show #REF! error:**
- Check sheet names match SHEETS constant exactly
- Verify columns haven't been deleted/moved

**Dropdown validations not working:**
- Run setupDataValidations() again
- Check Config sheet has data in all columns

**Functions not found errors:**
- Ensure all .gs files are deployed together
- Check that SHEETS constants match actual sheet names

### Getting Help

1. Check this AI_REFERENCE.md document first
2. Run DIAGNOSE_SETUP() to identify issues
3. Check Apps Script logs for errors
4. Review recent commits for changes
5. Create GitHub issue with:
   - Error message
   - Steps to reproduce
   - Screenshots
   - Execution log excerpt

---

**Document Version:** 2.4
**Last Updated:** 2025-12-05
**Maintained By:** Claude (AI Assistant)
**Repository:** [Add GitHub URL]

---

*This document serves as the complete technical specification and reconstruction guide for the 509 Dashboard system. Keep it updated as features are added or changed.*
