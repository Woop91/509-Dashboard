# Data Integrity & Cross-Population Guide

## Overview

The 509 Dashboard maintains strict data integrity between the **Member Directory** and **Grievance Log** sheets. This document explains how member and grievance data are linked and synchronized.

---

## Member ID System

### Random ID Generation

**Member IDs are randomly generated**, not sequential.

- **Format:** `MEM######` (e.g., `MEM847219`, `MEM092847`)
- **Range:** 000000 - 999999 (1 million possible IDs)
- **Uniqueness:** Guaranteed - system checks for duplicates during generation
- **Function:** `generateRandomMemberId(existingIds)`

**Why Random IDs?**
- Better security - sequential IDs can leak organizational size
- Prevents prediction of other member IDs
- Industry best practice for member identification systems

### Example Member IDs

```
MEM482719
MEM038471
MEM928374
MEM104728
MEM592837
```

---

## Grievance ↔ Member Relationship

### Data Flow

```
┌──────────────────────┐
│  Member Directory    │
│  MEM482719          │
│  John Smith         │
│  Job Title          │
│  Location           │
└──────────────────────┘
         ↑
         │ Cross-populates
         │ grievance stats
         │
┌──────────────────────┐
│  Grievance Log       │
│  GRV000001          │
│  MEM482719 ←────────┼─── MUST match Member Directory
│  John Smith ←───────┼─── MUST match Member Directory
│  Smith      ←───────┼─── MUST match Member Directory
│  Status: Filed      │
└──────────────────────┘
```

### Key Rules

1. **All grievances MUST use actual Member IDs from the Member Directory**
   - Column B in Grievance Log = Member ID
   - Must exactly match a Member ID in Member Directory Column A

2. **Member information must match between sheets**
   - First Name (Grievance Col C) = First Name (Member Directory Col B)
   - Last Name (Grievance Col D) = Last Name (Member Directory Col C)

3. **Random Member IDs are used throughout**
   - No sequential numbering
   - Each member gets a unique random ID when created

---

## Cross-Population: Grievance Stats → Member Directory

### What Gets Cross-Populated?

When you run **"Recalc All Members"**, the system automatically calculates and populates these fields in the Member Directory:

| Column | Field Name | Description | Calculation Source |
|--------|-----------|-------------|-------------------|
| M | Total Grievances | Total grievances filed by member | Count of all grievances with matching Member ID |
| N | Active Grievances | Currently active grievances | Count where Status starts with "Filed" |
| O | Resolved Grievances | Completed grievances | Count where Status starts with "Resolved" |
| P | Grievances Won | Successfully resolved | Count where Status = "Resolved - Won" |
| Q | Grievances Lost | Unsuccessfully resolved | Count where Status = "Resolved - Lost" |
| R | Last Grievance Date | Most recent grievance | Latest Incident Date from all grievances |

### How Cross-Population Works

The `recalcMemberRow()` function:

1. Takes a Member ID (e.g., `MEM482719`)
2. Searches the entire Grievance Log for matching Member IDs
3. Calculates aggregate statistics
4. Writes the results to the Member Directory columns M-R
5. Updates "Last Updated" timestamp

```javascript
// Simplified example
const memberId = "MEM482719";
const allGrievances = getGrievanceLog();
const memberGrievances = allGrievances.filter(g => g.memberId === memberId);

totalGrievances = memberGrievances.length;
activeGrievances = memberGrievances.filter(g => g.status.startsWith("Filed")).length;
resolvedGrievances = memberGrievances.filter(g => g.status.startsWith("Resolved")).length;
wonGrievances = memberGrievances.filter(g => g.status === "Resolved - Won").length;
// etc.
```

### When to Run Cross-Population

Run **"509 Tools" → "Data Management" → "Recalc All Members"** when:

- ✅ After adding new grievances (individually or in bulk)
- ✅ After updating grievance statuses
- ✅ After seeding test data
- ✅ When member grievance counts look incorrect
- ✅ Weekly as part of dashboard maintenance

**Note:** This does NOT need to be run after adding new members - only after grievance changes.

---

## Data Validation Functions

### Available Helper Functions

#### 1. `validateAndGetMemberData(memberId)`

Checks if a Member ID exists and returns the member's data.

**Usage:**
```javascript
const memberData = validateAndGetMemberData("MEM482719");

if (memberData) {
  // Member exists
  const [memberId, firstName, lastName, jobTitle] = memberData;
  console.log(`Found: ${firstName} ${lastName}`);
} else {
  // Member ID not found
  console.log("Invalid Member ID");
}
```

**Returns:**
- `[MemberId, FirstName, LastName, JobTitle]` if found
- `null` if not found

#### 2. `getAllMemberIds()`

Gets all valid Member IDs in the system.

**Usage:**
```javascript
const validIds = getAllMemberIds();
console.log(`Total members: ${validIds.length}`);
// ["MEM482719", "MEM038471", "MEM928374", ...]
```

**Returns:**
- Array of all Member IDs
- Empty array `[]` if no members exist

#### 3. `generateRandomMemberId(existingIds)`

Generates a new random unique Member ID.

**Usage:**
```javascript
const existingIds = new Set(getAllMemberIds());
const newId = generateRandomMemberId(existingIds);
console.log(`New Member ID: ${newId}`);
// "MEM583920"
```

**Parameters:**
- `existingIds` - Set of existing IDs to avoid duplicates

**Returns:**
- Random Member ID string (e.g., `"MEM583920"`)

---

## Seeding Test Data

### Seeding Members (with Random IDs)

**Menu:** `509 Tools → Data Management → Seed 20K Members`

Creates 20,000 test members with:
- ✅ Random Member IDs (not sequential)
- ✅ Realistic names, job titles, locations
- ✅ 10% are marked as stewards
- ✅ Random dates, phone numbers, emails

**Processing Time:** 3-5 minutes

### Seeding Grievances (with Actual Member IDs)

**Menu:** `509 Tools → Data Management → Seed 5K Grievances`

Creates 5,000 test grievances that:
- ✅ Use ACTUAL Member IDs from Member Directory
- ✅ Pull matching First Name and Last Name from members
- ✅ Ensure complete data integrity
- ✅ Calculate CBA deadlines automatically

**Processing Time:** 2-3 minutes

**IMPORTANT:** Always seed members FIRST, then grievances!

```
Order:
1. Seed 20K Members
2. Seed 5K Grievances
3. Recalc All Members
4. Rebuild Dashboard
```

---

## Manual Data Entry Best Practices

### Adding a New Member

1. Go to **Member Directory** sheet
2. Insert new row at row 2
3. Generate a random Member ID or use format `MEM######`
4. Enter: First Name, Last Name, Job Title, Location, Unit
5. Leave columns M-R blank (auto-calculated)
6. After adding, Member ID is now available for grievances

### Adding a New Grievance

1. Go to **Grievance Log** sheet
2. Insert new row at row 2
3. **CRITICAL:** Enter a valid Member ID from Member Directory
4. Copy the member's First Name and Last Name exactly as shown in Member Directory
5. Enter grievance details (Incident Date, Status, Type, etc.)
6. System auto-calculates deadlines
7. Run **"Recalc All Members"** to update member's statistics

---

## Troubleshooting

### Problem: "Member not found" errors

**Solution:**
- Verify Member ID exists in Member Directory Column A
- Check for typos (case-sensitive)
- Ensure format is `MEM######`

### Problem: Member grievance counts are zero

**Solution:**
- Run **"Recalc All Members"**
- Verify grievances use correct Member ID format
- Check that Member IDs match between sheets

### Problem: Duplicate Member IDs

**Solution:**
- Random generation prevents this automatically
- If manually entering, use `getAllMemberIds()` to check for duplicates
- Member IDs must be unique system-wide

### Problem: Names don't match between sheets

**Solution:**
- When creating a grievance, copy name exactly from Member Directory
- Use `validateAndGetMemberData()` to auto-fill names
- Run data cleanup script to sync names

---

## System Architecture

### Data Integrity Checks

The system maintains integrity through:

1. **Reference Lookups**
   - Grievances reference members by ID
   - Cross-validation before seeding

2. **Automatic Calculations**
   - Deadline calculations from CBA rules
   - Aggregate statistics from grievance data

3. **Update Triggers**
   - Manual recalculation commands
   - Scheduled updates (if configured)

4. **Validation Functions**
   - Member ID existence checks
   - Data format validation

### Column Mapping Reference

**Member Directory:**
- A: Member ID (Random, e.g., MEM482719)
- B: First Name
- C: Last Name
- D: Job Title
- E: Work Location
- F: Unit
- M: Total Grievances (AUTO-CALCULATED)
- N: Active Grievances (AUTO-CALCULATED)
- O: Resolved Grievances (AUTO-CALCULATED)
- P: Grievances Won (AUTO-CALCULATED)
- Q: Grievances Lost (AUTO-CALCULATED)
- R: Last Grievance Date (AUTO-CALCULATED)

**Grievance Log:**
- A: Grievance ID (Sequential, e.g., GRV000001)
- B: Member ID (MUST match Member Directory)
- C: First Name (MUST match Member Directory)
- D: Last Name (MUST match Member Directory)
- E: Status
- F: Current Step
- G: Incident Date
- H-Z: Dates, outcomes, details, etc.

---

## Best Practices Summary

✅ **DO:**
- Use random Member IDs for all new members
- Always validate Member ID exists before creating grievance
- Run "Recalc All Members" after bulk grievance updates
- Keep First/Last names consistent between sheets
- Use provided validation functions

❌ **DON'T:**
- Use sequential Member IDs (e.g., MEM000001, MEM000002)
- Create grievances with non-existent Member IDs
- Manually edit auto-calculated columns (M-R in Member Directory)
- Modify Member ID after grievances are linked
- Skip the "Recalc All Members" step

---

**Last Updated:** 2025-11-22
**System Version:** 509 Dashboard v2.0
