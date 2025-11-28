# Installable Triggers Guide for Google Apps Script

## What Are Installable Triggers?

**Installable triggers** let your Google Apps Script functions run with special permissions and longer execution time limits.

### Comparison: Simple Triggers vs Installable Triggers

| Feature | Simple Triggers | Installable Triggers |
|---------|----------------|---------------------|
| **Execution Time** | 6 minutes max | 30 minutes max |
| **How to Create** | Automatic (like `onOpen()`) | Manual setup required |
| **Permissions** | Limited | Full access |
| **Best For** | Menu items, quick actions | Long-running processes, scheduled tasks |

## Why You Need Installable Triggers

Your current issue: Seed functions timeout at ~5,000 members (6-minute limit)

**With installable triggers:**
- Can run for up to **30 minutes**
- Can seed 20,000+ members without timeout
- Can schedule automatic daily/weekly tasks
- Better for production workloads

## How to Set Up Installable Triggers

### Method 1: Manual Trigger (For Seed Functions)

**Step 1: Open Apps Script Editor**
1. Open your Google Sheet
2. Go to **Extensions → Apps Script**

**Step 2: Create the Trigger**
1. Click the **clock icon** ⏰ on the left sidebar (Triggers)
2. Click **+ Add Trigger** (bottom right)

**Step 3: Configure the Trigger**
Fill in the settings:

```
Choose which function to run: SEED_20K_MEMBERS
Choose which deployment should run: Head
Select event source: From spreadsheet
Select event type: On open
```

OR for manual execution:
```
Choose which function to run: SEED_20K_MEMBERS
Choose which deployment should run: Head
Select event source: Time-driven
Select type of time based trigger: Minutes timer
Select minute interval: Every minute
```

Then delete the trigger after it runs once.

**Step 4: Save**
- Click **Save**
- Authorize the trigger when prompted

### Method 2: Time-Based Trigger (For Scheduled Tasks)

Use this for automatic daily backups, weekly reports, etc.

**Configuration:**
```
Choose which function to run: createAutomatedBackup
Choose which deployment should run: Head
Select event source: Time-driven
Select type of time based trigger: Day timer
Select time of day: 2am to 3am
```

**Common Schedules:**
- **Daily backups**: Day timer, 2am-3am
- **Weekly reports**: Week timer, Every Monday, 9am-10am
- **Monthly summaries**: Month timer, 1st of month, 8am-9am

### Method 3: Programmatic Trigger Creation

Add this function to your script:

```javascript
function setupSeedTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'SEED_20K_MEMBERS') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new one-time trigger
  ScriptApp.newTrigger('SEED_20K_MEMBERS')
    .timeBased()
    .after(10000)  // Run after 10 seconds
    .create();

  SpreadsheetApp.getUi().alert('Trigger set! Function will run in 10 seconds with 30-minute timeout.');
}
```

Then run `setupSeedTrigger()` from the menu instead of `SEED_20K_MEMBERS()`.

## Recommended Setup for Your Dashboard

### Option A: Keep Small Dataset (Current - RECOMMENDED ✅)
- 5,000 members, 1,000 grievances
- Works with menu triggers (6-minute limit)
- No setup required
- Fast and sufficient for testing

### Option B: Large Dataset with Installable Triggers
If you really need 20,000 members:

1. **Change Code.gs loops back to:**
```javascript
for (let i = 1; i <= 20000; i++) {  // Members
for (let i = 1; i <= 5000; i++) {   // Grievances
```

2. **Create manual triggers:**
- Trigger 1: `SEED_20K_MEMBERS` → Time-driven → After 10 seconds
- Trigger 2: `SEED_5K_GRIEVANCES` → Time-driven → After 5 minutes

3. **Run sequence:**
- Activate SEED_20K_MEMBERS trigger → Wait 3-5 minutes for completion
- Activate SEED_5K_GRIEVANCES trigger → Wait 2-3 minutes for completion

## Managing Triggers

### View All Triggers
1. Apps Script Editor → Clock icon ⏰
2. See list of all active triggers

### Delete a Trigger
1. Click the three dots **⋮** next to trigger
2. Select **Delete trigger**

### Monitor Execution
1. Apps Script Editor → **Executions** icon (list with play button)
2. See all trigger executions, duration, and status

## Troubleshooting

### "Authorization Required"
- **Solution**: Click the trigger → Authorize → Select your Google account → Allow

### "Timeout after 30 minutes"
- **Solution**: Break data into smaller chunks or optimize batch processing

### "Trigger not running"
- **Check**: Executions log for errors
- **Fix**: Delete and recreate trigger

### "Too many triggers"
- **Limit**: 20 triggers per user per script
- **Solution**: Delete unused triggers

## Best Practices

1. **Test first**: Test with small dataset before creating large triggers
2. **Clean up**: Delete triggers after one-time use
3. **Monitor**: Check Executions log regularly
4. **Document**: Note which triggers are active and why
5. **Backup**: Always backup data before large operations

## For Your Current Situation

**RECOMMENDED APPROACH:**
Keep the current 5,000 member / 1,000 grievance setup. It's:
- Fast (completes in 1 minute)
- Reliable (no timeout issues)
- Sufficient for testing and demos
- No trigger setup required

**Only use installable triggers if:**
- Production deployment with real data
- Actually need 20,000+ member records
- Setting up automated scheduled tasks (backups, reports)

## Additional Resources

- [Google Apps Script Triggers Documentation](https://developers.google.com/apps-script/guides/triggers)
- [Installable Triggers Reference](https://developers.google.com/apps-script/guides/triggers/installable)

---

**Summary**: Installable triggers give you 30 minutes instead of 6, but require setup. For your dashboard, the current 5,000/1,000 dataset with menu triggers is the best choice unless you specifically need larger datasets.
