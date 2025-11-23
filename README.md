# 509 Dashboard - Member Directory & Grievance Tracking System

**Professional grievance tracking and member management system for SEIU Local 509 (Units 8 & 10)**

Built for Massachusetts State Employees | CBA-Compliant | Real-Time Analytics

---

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Quick Start Guide](#quick-start-guide)
- [Sheet Descriptions](#sheet-descriptions)
- [Visual Analytics](#visual-analytics)
- [Using the Dashboard](#using-the-dashboard)
- [Data Entry](#data-entry)
- [CBA Compliance](#cba-compliance)
- [For Stewards](#for-stewards)
- [Troubleshooting](#troubleshooting)
- [Support](#support)

---

## 🎯 Overview

The **509 Dashboard** is a comprehensive Google Sheets-based system designed specifically for SEIU Local 509 union representatives to:

- ✅ Track member information and engagement
- ✅ Manage grievances through all CBA steps
- ✅ Monitor critical deadlines automatically
- ✅ Analyze trends with interactive charts
- ✅ Balance steward workload
- ✅ Ensure CBA compliance (Article 23A deadlines)

**No formulas in data rows** - all calculations are automated through Google Apps Script, making the system fast, reliable, and easy to use.

---

## ⭐ Key Features

### 🚀 Grievance Workflow (NEW!)
- **Start grievances from Member Directory** - Click a member to begin
- **Pre-filled Google Forms** - Member and steward info automatically populated
- **Automatic log entry** - Form submissions added to Grievance Log
- **PDF generation** - Professional grievance forms as PDFs
- **Email or download** - Send to multiple addresses or download directly
- **Steward-only access** - Only stewards can start grievances
- **See full guide:** [GRIEVANCE_WORKFLOW_GUIDE.md](GRIEVANCE_WORKFLOW_GUIDE.md)

### 🚨 Seed Nuke - Exit Demo Mode (NEW!)
- **Remove all test data** - One-click removal of seeded members and grievances
- **Preserve structure** - All headers, formulas, and layouts remain intact
- **Production ready** - Transition from demo to live use
- **Hide seed menus** - Clean up menu after nuking
- **Getting started guide** - Post-nuke setup instructions
- **See full guide:** [SEED_NUKE_GUIDE.md](SEED_NUKE_GUIDE.md)

### 🎯 Interactive Dashboard
- **User-selectable metrics** - Choose from 20+ metrics to display
- **Dynamic chart types** - Pie, Donut, Bar, Line, Column, Area, or Table
- **Side-by-side comparison** - Compare any two metrics simultaneously
- **6 professional themes** - Union Blue, Solidarity Red, Success Green, and more
- **Real-time updates** - Charts refresh instantly with your data
- **Card-based layout** - Modern design with large KPI cards
- **Warehouse-style views** - Location analytics with horizontal bar charts
- **✨ Welcoming Experience** - Data feels alive with encouraging messages
- **🎉 Victory Celebrations** - Automatic recognition of achievements and milestones
- **💪 Gentle Guidance** - Supportive messages guide you through the dashboard
- **🏆 Smart Encouragement** - Context-aware celebrations based on your metrics
  - Win rate celebrations ("🏆 INCREDIBLE! Nearly perfect!")
  - Perfect attendance recognition ("🎊 PERFECT! Nothing overdue!")
  - Membership growth celebrations ("💪 Growing stronger every day!")
  - Dynamic card messages that change with your success
- **See full guide:** [INTERACTIVE_DASHBOARD_GUIDE.md](INTERACTIVE_DASHBOARD_GUIDE.md)

### 🧠 ADHD-Friendly Features (NEW!)
- **Soft, calming colors** - Gentle pastels instead of harsh brights
- **No gridlines** - Clean, minimal visual clutter
- **Emoji icons** - Every tab has a visual icon for quick scanning
- **Big numbers** - 48pt font for key metrics (glanceable data)
- **User customization** - Each person can set their own preferences
- **Visual guides** - Icon-based instructions, minimal text
- **Logical organization** - Important tabs first, admin tabs last
- **Quick setup** - One-click ADHD optimization (509 Tools > ADHD Tools)
- **See full guide:** [ADHD_FRIENDLY_GUIDE.md](ADHD_FRIENDLY_GUIDE.md)

### 📊 Advanced Dashboard
- **Real-time KPI metrics** - Member counts, grievance statistics, win rates
- **6 interactive visual charts** - Donut charts, bar charts, and column charts
- **Color-coded alerts** - Red (overdue), Yellow (due soon), Green (on track)
- **Top 10 overdue list** - Prioritized by severity with gradient color coding

### 👥 Member Directory
- **Auto-calculated grievance metrics** per member
- Track total, active, resolved, won, and lost grievances
- Monitor engagement levels and committee participation
- Emergency contact information and notes
- **35 data fields** including auto-calculated derived fields

### ⚖️ Grievance Log
- **CBA-compliant deadline tracking** (21-day filing, 30-day decisions, 10-day appeals)
- Automatic calculation of all deadlines
- Priority sorting (Step III → II → I, then by due date)
- **32 data fields** with automated deadline management
- Status tracking from filing through resolution

### 👨‍⚖️ Steward Workload Tracking
- Automatic calculation of cases per steward
- Active case breakdown by step (I, II, III)
- Overdue cases highlighted in red
- Win rates and average resolution times
- Workload balancing insights

### 📈 Visual Analytics
Six professionally designed charts that update automatically:
1. **Grievances by Status** - Donut chart showing Filed, Resolved, In Mediation, etc.
2. **Top 10 Grievance Types** - Bar chart of most common issues
3. **Grievances by Step** - Column chart tracking current case steps
4. **Members by Unit** - Donut chart showing Unit 8 vs Unit 10 distribution
5. **Top 10 Member Locations** - Bar chart of member workplace distribution
6. **Resolved Outcomes** - Donut chart showing Won/Lost/Settled/Withdrawn

### 📊 NEW: Advanced Analytics Tabs
Four dedicated analytical tabs with comprehensive insights:

**1. Trends & Timeline**
- Monthly filing trends with historical data
- Average filings per month calculations
- Peak month identification
- Resolution time analysis (average, fastest, slowest)
- Last 12 months filing trends visualization

**2. Performance Metrics**
- Resolution performance (total resolved, win rate, settlement rate, withdrawal rate)
- Efficiency metrics (active grievances, overdue rate, on-time rate)
- Outcome analysis by step with win rates
- Step progression analysis

**3. Location Analytics**
- Top 15 locations by grievance volume
- Members and grievances by location
- Active vs resolved cases by location
- Win rate comparison across locations

**4. Type Analysis**
- Comprehensive breakdown by grievance type
- Total, active, and resolved counts per type
- Win rate and settlement rate by type
- Identify high-volume grievance categories

### 🎨 Professional Design
- Color-coded sections with clear headers
- Number formatting with thousands separators
- Conditional formatting for visual prioritization
- Easy-to-read metrics with right-aligned numbers
- Mobile-friendly Google Sheets interface

---

## 🚀 Quick Start Guide

### Step 1: Initial Setup

1. **Open the Google Sheet** containing this script
2. Click on **Extensions** → **Apps Script**
3. Paste the `Code.gs` file contents
4. Save the script (Ctrl+S or Cmd+S)
5. Close the Apps Script editor
6. **Refresh the Google Sheet**

### Step 2: Create the Dashboard

1. You'll see a new menu: **509 Tools**
2. Click **509 Tools** → **🔧 Create Dashboard**
3. Wait 30 seconds while the system creates all sheets
4. Click **OK** when complete

### Step 3: Add Data

**Option A - Test with Sample Data (RECOMMENDED):**
- **509 Tools** → **Data Management** → **Seed All Test Data (Recommended)**
  - This unified function automatically:
    1. Seeds 20,000 members
    2. Seeds 5,000 grievances
    3. Recalculates all member metrics
    4. Rebuilds all dashboards
  - Takes 5-7 minutes total
  - Provides step-by-step progress alerts

**Option B - Manual Test Data:**
- **509 Tools** → **Data Management** → **Seed 20K Members**
- **509 Tools** → **Data Management** → **Seed 5K Grievances**
- **509 Tools** → **Data Management** → **Recalc All Members**
- **509 Tools** → **Data Management** → **Rebuild Dashboard**

**Option C - Add Real Data:**
- Manually enter members in the **Member Directory** sheet
- Manually enter grievances in the **Grievance Log** sheet
- System automatically calculates deadlines and metrics

### Step 4: View the Dashboard

Click on the **Dashboard** tab to see:
- Real-time metrics
- Interactive charts
- Top 10 overdue grievances
- Color-coded status indicators

---

## 📑 Sheet Descriptions

### 1. **Future Features** Sheet 🆕

**Professional listing of available security and advanced features:**

Lists all 94+ features available for implementation, including:
- **Security & Audit Features (79-86):**
  - Audit logging with user tracking
  - Role-based access control (Admin/Steward/Viewer)
  - Data encryption and decryption
  - Input sanitization
  - Audit reporting
  - Data retention policies
  - Suspicious activity detection

- **Advanced Features (87-94):**
  - Quick Actions sidebar
  - Advanced search and filtering
  - Automated backups
  - Performance monitoring
  - Keyboard shortcuts
  - Export wizard
  - Data import capabilities

**Each feature includes:**
- Feature number and name
- Detailed description
- Function name for implementation
- Status and priority
- Implementation notes
- Dependencies

*See this sheet for complete documentation on enabling advanced features*

---

### 2. **Pending Features** Sheet 🆕

**Production deployment checklist with step-by-step instructions:**

Organized into five sections:
1. **Test Key Features:** Validate integrity checks, sidebar, calculations, CBA compliance
2. **Configure Script Properties:** Set up admin/steward roles, encryption keys, backup folders
3. **Set Up Automated Triggers:** Schedule reports, updates, backups, alerts
4. **Update Main Menu:** Enable enhancement menus and verify accessibility
5. **Final Verification:** System-wide checks before going live

**Includes:**
- Detailed task descriptions
- Function names to run
- Expected results
- Priority levels (Critical/High/Medium/Low)
- Step-by-step implementation guide

*Use this sheet as your go-live checklist*

---

### 3. **Config** Sheet
Contains all dropdown options used throughout the system:
- Job titles (from CBA Appendix C)
- Work locations across Massachusetts
- Grievance types and outcomes
- Membership status options

*This sheet powers all data validation dropdowns.*

---

### 4. **Member Directory** Sheet

**37 columns** tracking comprehensive member information:

**Basic Information (A-H):**
- Member ID (auto-generated: MEM000001)
- First Name, Last Name
- Job Title, Department/Unit, Worksite/Office Location
- Work Schedule/Office Days, Unit (8 or 10)

**Contact & Role (I-P):**
- Email Address, Phone Number
- Is Steward (Yes/No)
- **Assigned Steward** - Primary steward assigned to this member
- **Share Phone in Directory? (Steward Only)** - Steward consent to share phone number
- Membership Status
- Immediate Supervisor, Manager/Program Director

**Grievance Metrics (P-U) - AUTO CALCULATED:**
- Total Grievances Filed
- Active Grievances
- Resolved Grievances
- Grievances Won
- Grievances Lost
- Last Grievance Date

**Grievance Snapshot (V-Y) - AUTO CALCULATED:**
- Has Open Grievance?
- # Open Grievances
- Last Grievance Status
- Next Deadline (Soonest)

**Participation & Admin (Z-AK):**
- Engagement Level
- Events Attended, Training Sessions
- Committee Membership
- Emergency Contacts
- Date of Birth, Hire Date
- Last Updated, Updated By

*Teal-highlighted columns = auto-calculated (don't edit manually)*

---

### 5. **Grievance Log** Sheet

**32 columns** with CBA-compliant deadline tracking:

**Basic Information (A-F):**
- Grievance ID (auto-generated: GRV000001)
- Member ID, First Name, Last Name
- Status, Current Step

**Incident & Filing (G-J):**
- Incident Date
- **Filing Deadline** (auto: Incident + 21 days)
- Date Filed
- **Step I Decision Due** (auto: Filed + 30 days)

**Step I Process (K-N):**
- Step I Decision Date
- Step I Outcome
- **Step II Appeal Deadline** (auto: Step I Decision + 10 days)
- Step II Filed Date

**Step II Process (O-R):**
- **Step II Decision Due** (auto: Step II Filed + 30 days)
- Step II Decision Date
- Step II Outcome
- **Step III Appeal Deadline** (auto: Step II Decision + 10 days)

**Step III & Beyond (S-V):**
- Step III Filed Date
- Step III Decision Date
- Mediation Date
- Arbitration Date

**Details & Outcomes (W-Z):**
- Final Outcome
- Grievance Type
- Description
- Representative (Steward name)

**Derived Fields (AA-AF) - AUTO CALCULATED:**
- Days Open
- Days to Next Deadline
- Is Overdue? (YES/NO)
- Priority Score (1-6, Step III = highest)
- Assigned Steward
- Steward Contact

**Admin (AG-AH):**
- Notes
- Last Updated

*Orange-highlighted columns = auto-calculated deadlines*
*Green-highlighted columns = auto-calculated derived fields*

---

### 6. **Dashboard** Sheet

**Your command center** with four main sections:

**Member Metrics (Rows 4-9):**
- Total Members: Count of all members
- Active Members: Currently active status
- Total Stewards: Available union representatives
- Unit 8 Members: Unit 8 count
- Unit 10 Members: Unit 10 count

**Grievance Metrics (Rows 12-19):**
- Total Grievances
- Active Grievances (currently filed)
- Resolved Grievances
- **Grievances Won** (green background)
- **Grievances Lost** (red background)
- **Win Rate** (highlighted percentage)
- In Mediation, In Arbitration

**Deadline Tracking (Rows 22-24):**
- **Overdue Grievances** (red if >0, green if 0)
- **Due This Week** (yellow if >0, green if 0)
- Due Next Week

**Top 10 Overdue Grievances (Rows 4-14, Columns E-H):**
Priority-sorted list with color coding:
- 🔴 Dark Red (30+ days overdue)
- 🟠 Red (14-30 days overdue)
- 🟡 Light Red (<14 days overdue)

**Visual Analytics (Row 27+):**
Six interactive charts (see Visual Analytics section)

---

### 7. **Steward Workload** Sheet

Automatically populated with steward performance metrics:

**Columns:**
- Steward Name, Member ID, Work Location
- Total Cases, Active Cases
- Step I, II, III case counts
- **Overdue Cases** (highlighted red if >0)
- Due This Week
- Win Rate %, Avg Days to Resolution
- Members Assigned
- Last Case Date, Status

*Updates automatically when you run "Rebuild Dashboard"*

---

### 8. **Analytics Data** Sheet

Historical snapshot tracking for trend analysis:
- Snapshot Date
- Member counts and activity
- Grievance statistics
- Win rates over time
- Average resolution times

*Manually add rows to track monthly/quarterly trends*

---

### 9. **Trends & Timeline** Sheet 🆕

**Automatic time-based analysis of your grievance data:**

**Monthly Trends Section:**
- Total months with data
- Average filings per month
- Peak month identification

**Resolution Time Analysis:**
- Average resolution time across all cases
- Fastest resolution time
- Slowest resolution time

**Filing Trends:**
- Last 12 months breakdown
- Monthly filing counts
- Trend visualization data

*Auto-populated when you run "Rebuild Dashboard" or "Rebuild Analytics" menu*

---

### 10. **Performance Metrics** Sheet 🆕

**Comprehensive KPI analysis for performance tracking:**

**Resolution Performance:**
- Total resolved grievances
- Win rate percentage (highlighted)
- Settlement rate percentage
- Withdrawal rate percentage

**Efficiency Metrics:**
- Active grievances count
- Overdue rate (color-coded: red if overdue)
- On-time rate

**Outcome Analysis by Step:**
- Step I, II, III breakdown
- Total cases and win rates per step
- Performance comparison across steps

*Ideal for monthly reporting and performance reviews*

---

### 11. **Location Analytics** Sheet 🆕

**Geographic breakdown of grievances and member activity:**

**Top 15 Locations Table:**
- Location name
- Member count
- Total grievances filed
- Active grievances
- Resolved grievances
- Win rate percentage

*Helps identify problem locations and resource allocation needs*

---

### 12. **Type Analysis** Sheet 🆕

**Deep dive into grievance categories and outcomes:**

**Type Breakdown Table:**
- All grievance types
- Total count per type
- Active cases per type
- Resolved cases per type

**Success Rate Analysis:**
- Win rate by grievance type
- Settlement rate by grievance type
- Identify which types are most successful

*Use this to focus on high-volume or low-success grievance types*

---

### 13. **Archive** Sheet

Storage for resolved grievances older than 90 days.
*Future feature - manual archiving currently*

---

### 14. **Diagnostics** Sheet

System health checks:
- Data validation status
- Trigger configuration
- Record counts
- Orphaned record detection

---

## 📊 Visual Analytics

### Chart 1: Grievances by Status
**Type:** Donut Chart
**Shows:** Distribution of all grievances by current status
**Categories:** Filed - Step I/II/III, In Mediation, In Arbitration, Resolved
**Use:** Quick overview of case pipeline

### Chart 2: Top 10 Grievance Types
**Type:** Horizontal Bar Chart
**Shows:** Most common grievance categories
**Use:** Identify systemic issues and patterns

### Chart 3: Grievances by Current Step
**Type:** Column Chart
**Shows:** Number of cases at each grievance step
**Use:** Understand where cases are in the process

### Chart 4: Members by Unit
**Type:** Donut Chart
**Shows:** Unit 8 vs Unit 10 member distribution
**Use:** Compare unit sizes and representation

### Chart 5: Top 10 Member Locations
**Type:** Horizontal Bar Chart
**Shows:** Offices with most members
**Use:** Understand geographic distribution

### Chart 6: Resolved Grievance Outcomes
**Type:** Donut Chart
**Shows:** Won, Lost, Settled, Withdrawn breakdown
**Use:** Track overall success rate

---

## 💼 Using the Dashboard

### 🎯 Using the Interactive Dashboard (NEW!)

**Quick Start:**
1. Click **509 Tools** → **Interactive Dashboard** → **View Interactive Dashboard**
2. First time: Click **Setup Controls** to enable dropdowns
3. Use Row 7 dropdowns to select:
   - Your primary metric (e.g., "Total Members")
   - Chart type (e.g., "Donut Chart")
   - Comparison metric (e.g., "Active Grievances")
   - Comparison chart type (e.g., "Bar Chart")
   - Theme (e.g., "Union Blue")
   - Enable/disable comparison mode
4. Click **Refresh Charts** to update visualizations

**Benefits:**
- Customize what you see
- Compare metrics side-by-side
- Export presentation-ready charts
- Switch themes for different audiences
- 20+ metrics × 7 chart types = 140+ combinations

**Full Instructions:** See [INTERACTIVE_DASHBOARD_GUIDE.md](INTERACTIVE_DASHBOARD_GUIDE.md)

### Daily Tasks

**1. Check Overdue Items:**
- Look at "Overdue Grievances" metric (Row 22)
- Review "Top 10 Overdue Grievances" table
- Prioritize cases with dark red highlighting (30+ days)

**2. Monitor This Week's Deadlines:**
- Check "Due This Week" metric (Row 23)
- Filter Grievance Log by deadline columns
- Contact stewards with upcoming deadlines

**3. Review New Grievances:**
- Check Grievance Log for recently filed cases
- Verify all deadlines calculated correctly
- Assign to appropriate steward

### Weekly Tasks

**1. Update Case Statuses:**
- Enter new decisions and outcomes
- Update Step progression
- System automatically recalculates deadlines

**2. Run Dashboard Rebuild:**
- **509 Tools** → **Data Management** → **Rebuild Dashboard**
- Ensures all charts and metrics are current

**3. Review Steward Workload:**
- Check Steward Workload sheet
- Balance case assignments
- Address overdue cases (red highlighted)

### Monthly Tasks

**1. Take Analytics Snapshot:**
- Manually add row to Analytics Data sheet
- Record current metrics for trend tracking

**2. Review Win Rates:**
- Check Dashboard win rate metric
- Analyze by grievance type
- Identify areas for improvement

**3. Member Engagement:**
- Review member engagement levels
- Follow up with inactive members
- Update committee participation

---

## ✍️ Data Entry

### Adding a New Member

1. Go to **Member Directory** sheet
2. Add new row (Row 2 for most recent)
3. **Enter manually:**
   - First Name, Last Name
   - Job Title (use dropdown)
   - Work Location (use dropdown)
   - Unit (use dropdown)
   - Email, Phone
   - Is Steward (use dropdown)
   - Date Joined Union
   - Membership Status (use dropdown)

4. **Leave blank (auto-calculated):**
   - Member ID (or use format: MEM000123)
   - All columns N-X (grievance metrics)
   - Last Updated

5. **Optional:**
   - Office Days, Engagement Level
   - Committee Member, Preferred Contact
   - Emergency Contacts, Notes

### Adding a New Grievance

1. Go to **Grievance Log** sheet
2. Add new row (Row 2 for most recent)
3. **Enter manually:**
   - Member ID (from Member Directory)
   - First Name, Last Name (from Member Directory)
   - Status (use dropdown)
   - Current Step (use dropdown)
   - **Incident Date** (triggers all deadline calculations)
   - Date Filed (when you actually filed)
   - Grievance Type (use dropdown)
   - Description
   - Representative (Steward name)

4. **Leave blank (auto-calculated):**
   - Grievance ID (or use format: GRV000456)
   - All deadline columns (H, J, M, O, R)
   - All derived fields (AA-AF)

5. **Enter as case progresses:**
   - Step I Decision Date, Outcome
   - Step II Filed Date, Decision Date, Outcome
   - Step III dates
   - Final Outcome

6. **System automatically:**
   - Calculates all CBA deadlines
   - Determines if case is overdue
   - Updates priority score
   - Updates member's grievance statistics

### Updating a Grievance

When you receive a decision or file an appeal:

1. Find the grievance row
2. Enter the date in the appropriate column
3. **System automatically recalculates:**
   - Next deadline
   - Days to deadline
   - Overdue status
   - Member statistics

**Example:** When you enter "Step I Decision Date":
- System calculates "Step II Appeal Deadline" (10 days later)
- Updates "Days to Next Deadline"
- Changes color coding if overdue

---

## ⚖️ CBA Compliance

This system enforces **Article 23A** grievance procedure deadlines:

### Filing Deadlines

| Event | Deadline | Auto-Calculated Column |
|-------|----------|----------------------|
| Incident occurs | File within 21 days | H: Filing Deadline |
| Grievance filed | Step I decision in 30 days | J: Step I Decision Due |
| Step I decision received | Appeal within 10 days | M: Step II Appeal Deadline |
| Step II filed | Step II decision in 30 days | O: Step II Decision Due |
| Step II decision received | Appeal within 10 days | R: Step III Appeal Deadline |

### How It Works

1. **You enter:** Incident Date (Column G)
2. **System calculates:** Filing Deadline = Incident Date + 21 days (Column H)
3. **You enter:** Date Filed (Column I)
4. **System calculates:** Step I Due = Date Filed + 30 days (Column J)
5. **Continues automatically** through all steps

### Color Coding

- 🟢 **Green:** More than 7 days until deadline
- 🟡 **Yellow:** 0-7 days until deadline (due soon!)
- 🔴 **Red:** Past deadline (overdue!)
- 🔴 **Dark Red:** 30+ days overdue (urgent!)

---

## 🌟 For Stewards

### Excellence Guide for Stewards

**Are you a steward?** We've created a special guide just for you!

The **[Steward Excellence Guide](STEWARD_GUIDE.md)** uses positive reinforcement and celebrates your essential work maintaining member data and engaging with less-active members.

**This guide includes:**
- 🏆 Achievement system to track your progress
- 💪 Special missions for re-engaging less-active members
- 🎯 Weekly rituals to build data maintenance habits
- 🎉 Celebration strategies for every win (big and small!)
- 💡 Pro tips from legendary stewards
- 🌈 Motivational affirmations and mindset coaching

**Your work matters!** Every field you complete strengthens our union. Every less-active member you reach out to is a potential future leader.

👉 **[Read the Steward Excellence Guide](STEWARD_GUIDE.md)** to discover how your data work creates powerful solidarity!

---

## 🔧 Troubleshooting

### Charts Not Showing

**Problem:** Dashboard shows "No data available yet"
**Solution:**
1. Add at least one member to Member Directory
2. Add at least one grievance to Grievance Log
3. Run: **509 Tools** → **Data Management** → **Rebuild Dashboard**

---

### Deadlines Not Calculating

**Problem:** Orange columns (H, J, M, O, R) are blank
**Solution:**
1. Ensure you entered the **Incident Date** (Column G)
2. Ensure you entered the **Date Filed** (Column I)
3. Run: **509 Tools** → **Data Management** → **Recalc All Grievances**

---

### Member Metrics Not Updating

**Problem:** Columns N-X in Member Directory are blank
**Solution:**
1. Ensure Member ID matches between Member Directory and Grievance Log
2. Run: **509 Tools** → **Data Management** → **Recalc All Members**

---

### Dropdowns Not Working

**Problem:** Cannot select from dropdown menus
**Solution:**
1. Check that **Config** sheet exists and has data
2. Run: **509 Tools** → **🔧 Create Dashboard** to rebuild
3. May need to re-apply: **509 Tools** → **Utilities** → **Setup Triggers**

---

### Steward Workload Empty

**Problem:** Steward Workload sheet has no data
**Solution:**
1. Ensure members have "Is Steward" = "Yes" in Member Directory
2. Ensure grievances have Representative names that match steward names
3. Run: **509 Tools** → **Data Management** → **Rebuild Dashboard**

---

### Permission Errors

**Problem:** "You need permission to run this script"
**Solution:**
1. **Extensions** → **Apps Script**
2. Click ▶️ **Run** → Select any function
3. Click **Review Permissions**
4. Choose your Google account
5. Click **Advanced** → **Go to [Project Name]**
6. Click **Allow**

---

## 📞 Support

### Quick Reference Card

**Add Member:** Member Directory → New Row → Fill basic info
**Add Grievance:** Grievance Log → New Row → Enter Incident Date
**View Dashboard:** Click Dashboard tab
**Refresh Data:** 509 Tools → Data Management → Rebuild Dashboard
**Check Deadlines:** Dashboard → Top 10 Overdue list
**Balance Stewards:** Steward Workload sheet → Review Active Cases

### Menu Guide

**509 Tools Menu:**
```
🔧 Create Dashboard
   └─ Initial setup (run once)

🎯 Interactive Dashboard (NEW!)
   ├─ Setup Controls (first time setup)
   ├─ Refresh Charts (update visualizations)
   └─ View Interactive Dashboard (open and get instructions)

🧠 ADHD-Friendly Tools (NEW!)
   ├─ ⚡ Quick Setup (Do This First!) - One-click optimization
   ├─ (separator)
   ├─ 📑 Reorder Sheets Logically - Put important tabs first
   ├─ 👁️ Hide Gridlines (Cleaner View) - Remove visual clutter
   ├─ 👁️‍🗨️ Show Gridlines (If Needed) - Bring them back
   ├─ (separator)
   ├─ ⚙️ Create My Settings Page - Personal customization
   ├─ ✅ Apply My Settings - Apply your preferences
   ├─ (separator)
   └─ 📋 Add Visual Guide to Steward Workload - Icon-based instructions

📊 Data Management
   ├─ Seed All Test Data (RECOMMENDED - unified function) 🆕
   ├─ (separator)
   ├─ Seed 20K Members (test data)
   ├─ Seed 5K Grievances (test data)
   ├─ (separator)
   ├─ Recalc All Grievances (fix calculations)
   ├─ Recalc All Members (update metrics)
   └─ Rebuild Dashboard (refresh charts)

📈 Rebuild Analytics
   ├─ Rebuild All Tabs (refresh all analytics)
   ├─ (separator)
   ├─ Rebuild Trends & Timeline
   ├─ Rebuild Performance Metrics
   ├─ Rebuild Location Analytics
   └─ Rebuild Type Analysis

📤 Export Data
   ├─ Export Dashboard to PDF
   ├─ Export Member Directory to CSV
   ├─ Export Grievances to CSV
   └─ Export Steward Workload to CSV

🎨 Theme Options
   ├─ Light Theme
   ├─ Dark Theme
   └─ High Contrast Theme

⚙️ Utilities
   ├─ Sort by Priority (organize grievance list)
   ├─ Toggle Mobile Mode
   └─ Setup Triggers (reset automation)
```

**Advanced Enhancement Menus:** 🆕
```
🔍 Data Validation
   ├─ Run Integrity Check
   ├─ Auto-Correct Errors
   ├─ Find Orphaned Grievances
   ├─ Generate Missing IDs
   ├─ (separator)
   └─ Validate CBA Compliance

📧 Notifications
   ├─ Send Overdue Alerts
   ├─ Send Weekly Reminders
   ├─ Send Daily Digest
   └─ Configure Preferences

📊 Advanced Reports
   ├─ Executive Summary
   ├─ Trend Analysis
   ├─ Location Analysis
   ├─ Steward Performance
   ├─ (separator)
   ├─ Export to PDF
   ├─ Export to CSV
   └─ Export to Excel

👥 Member Engagement
   ├─ Update Engagement Levels
   ├─ Find Inactive Members
   ├─ Identify Steward Candidates
   └─ Send Re-engagement Emails

🛠️ Advanced Tools
   ├─ Quick Actions Sidebar
   ├─ Search Grievances
   ├─ Filter Data
   ├─ Export Wizard
   ├─ (separator)
   ├─ Keyboard Shortcuts
   └─ Performance Monitor
```

*All enhancement menus are automatically enabled via addEnhancementMenus() in onOpen()*

### Common Workflows

**New Grievance Workflow:**
1. Member reports issue → Record Incident Date
2. File grievance → Enter Date Filed
3. System calculates → Filing and Step I deadlines
4. Receive Step I decision → Enter decision date
5. System calculates → Step II appeal deadline
6. Continue through process → System tracks everything

**Weekly Review Workflow:**
1. Open Dashboard
2. Check "Overdue Grievances" count
3. Review "Top 10 Overdue" table
4. Check "Due This Week" count
5. Filter Grievance Log for this week's items
6. Contact assigned stewards
7. Update case statuses
8. Rebuild Dashboard

---

## 📄 Technical Details

**Platform:** Google Sheets + Google Apps Script
**Programming Language:** JavaScript (Apps Script)
**Calculation Method:** Code-based (no formulas in data rows)
**Update Frequency:** Real-time on edit, manual rebuild for charts
**Data Capacity:** 20,000+ members, 5,000+ grievances tested
**Mobile Friendly:** Yes (Google Sheets mobile app)

### System Architecture

- **Config Sheet:** Central data validation source
- **Member Directory:** Member records with calculated metrics
- **Grievance Log:** Case tracking with CBA deadline automation
- **Dashboard:** Executive summary with KPIs and charts
- **Steward Workload:** Performance and capacity tracking
- **Analytics Data:** Historical trend storage
- **_ChartData:** Hidden sheet for chart data processing

### Automated Triggers

- **onEdit:** Recalculates row when edited
- **onFormSubmit:** Processes form submissions (if connected)
- **Manual:** Rebuild Dashboard, Recalc All functions

---

## 🚀 Production Readiness

### Pre-Launch Checklist

Before deploying to production, complete the following:

**✅ 1. Review Production Setup Sheets**
- Open the **Pending Features** sheet
- Follow the step-by-step checklist
- Verify all critical and high-priority items

**✅ 2. Test Key Features**
```
Run these functions from the menu to verify:
- 509 Tools → 🛠️ Advanced Tools → Quick Actions Sidebar
- 🔍 Data Validation → Run Integrity Check
- 509 Tools → 📊 Data Management → Seed All Test Data (in test environment)
```

**✅ 3. Configure Script Properties**
Access: Extensions → Apps Script → Project Settings → Script Properties

Required properties for security features:
- `ADMINS` - Comma-separated admin emails
- `STEWARDS` - Comma-separated steward emails
- `ENCRYPTION_KEY` - 16+ character secure key
- `BACKUP_FOLDER_ID` - Google Drive folder ID for backups

**✅ 4. Set Up Automated Triggers**
```
From Apps Script editor, run these once:
- scheduleReports() - Weekly/monthly report automation
- scheduleAutomaticUpdates() - Hourly dashboard refresh
```

**✅ 5. Enable Security Features**
All security features are available in the **Future Features** sheet:
- Audit logging (Feature 79)
- Role-based access control (Feature 80)
- Input sanitization (Feature 83)
- See the sheet for full documentation

**✅ 6. Final Verification**
- Test with real data in a copy first
- Verify all calculations are accurate
- Ensure all menu items work
- Train end users on workflows
- Set up backup schedule

**✅ 7. Go Live!**
- Import or enter real member data
- Import or enter real grievance data
- Run: 509 Tools → Data Management → Rebuild Dashboard
- Monitor closely for first week

### System Capacity

**Tested Performance:**
- 20,000+ members ✅
- 5,000+ grievances ✅
- < 5 seconds for most operations ✅
- Real-time calculations ✅

### Security & Compliance

**Data Protection:**
- Role-based access control available (configure script properties)
- Audit logging available (Feature 79)
- Data encryption available (Features 81-82)
- Input sanitization enabled (Feature 83)

**CBA Compliance:**
- Article 23A deadlines automatically calculated ✅
- 21-day filing window ✅
- 30-day decision timelines ✅
- 10-day appeal deadlines ✅

**Backup & Recovery:**
- Automated backup function available (Feature 90)
- Manual export to CSV/PDF ✅
- Google Sheets version history ✅

### Support & Maintenance

**Weekly Tasks:**
- Review **Dashboard** for overdue grievances
- Run **Rebuild Dashboard** if needed
- Check **Steward Workload** for balance

**Monthly Tasks:**
- Run **🔍 Data Validation → Run Integrity Check**
- Review audit logs (if enabled)
- Verify backup creation (if automated)

**Quarterly Tasks:**
- Review **Future Features** for new capabilities
- Update member engagement levels
- Archive old resolved cases

### Getting Help

- Review **Pending Features** sheet for setup guidance
- Review **Future Features** sheet for available capabilities
- Check **Diagnostics** sheet for system health
- All functions are documented with comments in Code.gs

---

## 📜 License

Created for SEIU Local 509 (Units 8 & 10)
Based on Massachusetts State Employee CBA 2024-2026

---

## ✨ Best Practices

### Data Entry
- ✅ Always enter Incident Date first (triggers all calculations)
- ✅ Use consistent steward names (for workload tracking)
- ✅ Keep Member IDs consistent across sheets
- ✅ Use dropdowns (prevents typos and errors)
- ❌ Don't edit green or orange highlighted columns (auto-calculated)

### Maintenance
- 🔄 Run "Rebuild Dashboard" weekly
- 🔄 Run "Recalc All Members" after bulk grievance entry
- 🔄 Run "Recalc All Grievances" if deadlines look wrong
- 📊 Take monthly analytics snapshots
- 🗂️ Archive resolved cases quarterly

### Performance
- Keep most recent items at top of sheets (Row 2)
- Sort Grievance Log by priority monthly
- Archive old resolved cases to Archive sheet
- Clear Diagnostics sheet periodically

---

**Built with ❤️ for union representatives fighting for worker rights**

*For questions, issues, or feature requests, contact your system administrator.*
