# Seed Function Timeout Issue Explained

## Problem
You're seeing only ~5,002 members and ~500 grievances instead of 20,000 and 5,000.

## Root Cause
**Google Apps Script Execution Time Limits:**
- Menu-triggered functions (like your seed functions): **6 minutes maximum**
- The seed functions are timing out before completing

## Why This Happens
Creating 20,000 rows × 32 columns = 640,000 cells takes significant time:
- Data generation
- Batch writes to Google Sheets
- SpreadsheetApp.flush() calls
- API rate limiting

## Current Performance
Based on ~5,002 members in 6 minutes:
- Rate: ~833 members/minute
- For 20,000 members: Would need ~24 minutes (4x the limit!)

## Solutions

### Option 1: Use Realistic Test Data (RECOMMENDED)
Modify the seed functions to use smaller, more realistic numbers:
- Change 20,000 → 5,000 members
- Change 5,000 → 1,000 grievances

This will:
- Complete within 6-minute limit
- Still provide robust testing
- Faster iteration during development

### Option 2: Use Installable Triggers
Create time-based or manual triggers with 30-minute limit:
1. Go to Extensions → Apps Script
2. Click "Triggers" (clock icon)
3. Add trigger → Choose function → Time-driven
4. Set to run manually or on schedule

### Option 3: Optimize Batch Processing
- Increase BATCH_SIZE from 1000 to 2000
- Remove some SpreadsheetApp.flush() calls
- Reduce toast notifications

## Recommendation
For your use case, **5,000 members and 1,000 grievances** is more than sufficient for:
- Testing all features
- Demonstrating dashboards
- Performance testing
- Real-world scenarios

Most union locals don't have 20,000 active members anyway!

## How to Change
Edit the seed function calls:
```javascript
// In Code.gs or Complete509Dashboard.gs
for (let i = 1; i <= 5000; i++) {  // Changed from 20000
  // ... member generation
}

// And for grievances:
for (let i = 1; i <= 1000; i++) {  // Changed from 5000
  // ... grievance generation
}
```

Or update the menu items and alerts to reflect the new numbers.
