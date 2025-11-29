# 🚀 Quick Deploy - SEIU 509 Dashboard

## One-File Deployment (Easiest Method)

This guide shows you how to deploy the **entire 509 Dashboard** using a **single file**.

---

## 📦 What You Get

- ✅ **All 8 modules** in one file (5,863 lines)
- ✅ **Complete functionality** - nothing missing
- ✅ **Copy/paste deployment** - no complex setup
- ✅ **20k members + 5k grievances** capacity
- ✅ **Terminal dashboard** with 26+ real-time metrics
- ✅ **ADHD-friendly** design and features

---

## 🔧 5-Minute Setup

### **Step 1: Get the File**

Download `Complete509Dashboard.gs` from this repository:
```
https://github.com/Woop91/509-dashboard/blob/claude/final-pass-review-015wMPj45ZdPiVcrdm8Br4zq/Complete509Dashboard.gs
```

Or use git:
```bash
git clone https://github.com/Woop91/509-dashboard.git
cd 509-dashboard
git checkout claude/final-pass-review-015wMPj45ZdPiVcrdm8Br4zq
```

### **Step 2: Create a New Google Sheet**

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **"+ Blank"** to create a new spreadsheet
3. Name it **"SEIU 509 Dashboard"**

### **Step 3: Open Apps Script Editor**

1. In your Google Sheet, click **Extensions → Apps Script**
2. You'll see a default `Code.gs` file with some starter code
3. **Delete all the default code** (select all and delete)

### **Step 4: Paste the Complete File**

1. Open `Complete509Dashboard.gs` in your text editor
2. **Select ALL** (Ctrl+A / Cmd+A)
3. **Copy** (Ctrl+C / Cmd+C)
4. **Paste** into the Apps Script editor (Ctrl+V / Cmd+V)
5. **Save** the project (Ctrl+S / Cmd+S or click 💾 icon)

### **Step 5: Authorize the Script**

1. In the Apps Script editor toolbar, find the function dropdown
2. Select **`CREATE_509_DASHBOARD`** from the dropdown
3. Click the **▶️ Run** button
4. A dialog will appear asking for permissions:
   - Click **"Review permissions"**
   - Choose your Google account
   - Click **"Advanced"**
   - Click **"Go to 509 Dashboard Scripts (unsafe)"**
   - Click **"Allow"**

5. The script will run and create all sheets (~10 seconds)

### **Step 6: Close Apps Script & Refresh**

1. Close the Apps Script editor tab
2. Go back to your Google Sheet
3. **Refresh the page** (F5 or Ctrl+R / Cmd+R)
4. You should see a new menu: **📊 509 Dashboard**

### **Step 7: Seed Test Data**

1. Click **📊 509 Dashboard → ⚙️ Admin**
2. Click **"Seed 20k Members"** (takes ~2 minutes)
3. Wait for completion toast message
4. Click **"Seed 5k Grievances"** (takes ~1-2 minutes)
5. Wait for completion toast message

### **Step 8: Open the Terminal Dashboard**

1. Click **📊 509 Dashboard → 🎯 Unified Operations Monitor**
2. The terminal-themed dashboard opens in a new dialog
3. Explore the 7 sections:
   - Executive Status & Alerts
   - Process Efficiency
   - Network Health
   - Active Grievance Log
   - Follow-up Radar
   - Predictive Alerts
   - Systemic Risk Monitor

---

## ✅ You're Done!

Your dashboard is fully operational with:
- ✅ 20,000 test members
- ✅ 5,000 test grievances
- ✅ Real-time analytics
- ✅ Terminal operations monitor
- ✅ Interactive customizable views
- ✅ All ADHD-friendly features

---

## 📚 Next Steps

### **Explore the Features:**

| Menu Item | What It Does |
|-----------|-------------|
| 🔄 Refresh All | Update all calculations |
| 🎯 Unified Operations Monitor | Open terminal dashboard |
| 📊 Dashboard | Go to main dashboard sheet |
| ❓ Help | Show help dialog |
| ⚙️ Admin → Clear All Data | Reset (for testing) |

### **Customize Your Data:**

1. Edit **Config** tab to change dropdown options
2. Manually add/edit members in **Member Directory**
3. Manually add/edit grievances in **Grievance Log**
4. Use **Interactive Dashboard** for custom views

### **Read the Documentation:**

- `README.md` - Complete feature documentation
- `ADHD_FRIENDLY_GUIDE.md` - Accessibility features
- `STEWARD_GUIDE.md` - Guide for union stewards
- `INTERACTIVE_DASHBOARD_GUIDE.md` - Custom views
- `GRIEVANCE_WORKFLOW_GUIDE.md` - Workflow automation

---

## 🛠️ Troubleshooting

### **Menu doesn't appear?**
- Refresh the page (F5)
- Or manually run `onOpen()` from Apps Script editor

### **Authorization error?**
- Make sure you completed Step 5 fully
- Close and reopen the Google Sheet
- Try running any function from Apps Script editor again

### **Seeding takes forever?**
- Normal for large datasets (2-3 minutes total)
- Don't close the sheet while seeding
- Check Apps Script logs if it fails: View → Execution log

### **Dashboard shows "CONNECTION ERROR"?**
- Make sure you opened it via menu, not directly
- Refresh the sheet and try again
- Verify `getUnifiedDashboardData()` function exists in your code

### **Need help?**
- Click **📊 509 Dashboard → ❓ Help** in your sheet
- Check the GitHub issues page
- Review documentation files

---

## 📁 What's Included in Complete509Dashboard.gs?

The single file contains these modules:

1. **Core System** (Code.gs)
   - Configuration constants
   - Sheet creation functions
   - Data seeding (20k members, 5k grievances)
   - Menu system

2. **Unified Operations Monitor**
   - Terminal-themed dashboard
   - 26+ backend calculation functions
   - 7 analytics sections
   - Real-time metrics

3. **Interactive Dashboard**
   - Member-customizable views
   - Dynamic charts
   - Filtering and sorting

4. **ADHD Enhancements**
   - Color coding
   - Visual hierarchy
   - Focus-friendly design

5. **Grievance Workflow**
   - Process automation
   - Deadline tracking
   - Status management

6. **Getting Started & FAQ**
   - Help system
   - Documentation
   - Tutorials

7. **Seed & Nuke Utilities**
   - Data generation
   - Cleanup functions
   - Testing tools

8. **Column Toggles**
   - UI controls
   - Column visibility
   - Layout management

---

## 🎉 Success!

You now have a fully functional union management system with:
- Real-time grievance tracking
- Comprehensive analytics dashboard
- Member engagement monitoring
- ADHD-friendly interface
- Deadline management
- And much more!

**Questions?** Open an issue on GitHub or check the documentation files.

---

**Version:** Final Pass Review
**Last Updated:** 2025-11-26
**Branch:** claude/final-pass-review-015wMPj45ZdPiVcrdm8Br4zq
**GitHub:** https://github.com/Woop91/509-dashboard
