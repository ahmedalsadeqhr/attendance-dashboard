# Attendance Dashboard - User Guide
## For HR & Payroll Staff

---

## Getting Started

### Step 1: Open the Application

Double-click on **AttendanceDashboard.exe** to start the application.

```
[Double-click this icon]
    AttendanceDashboard.exe
```

A window will appear with the title "Attendance Dashboard v2.3"

---

## Loading Your Files

You need to load **3 files** before generating a report:

### Step 2: Load Attendance Files

1. Click the **"Select Attendance Files"** button
2. Navigate to your attendance files folder
3. Select one or more `.xlsx` files (hold Ctrl to select multiple)
4. Click **"Open"**

```
+----------------------------------+
|  [Select Attendance Files]       |  <-- Click here
|  Status: 3 files selected        |
+----------------------------------+
```

**What are attendance files?**
- Excel files exported from your attendance machine
- Contains clock in/out times for employees

---

### Step 3: Load Master Data

1. Click the **"Select Master Data"** button
2. Select your employee master data Excel file
3. Click **"Open"**

```
+----------------------------------+
|  [Select Master Data]            |  <-- Click here
|  Status: Master data loaded      |
+----------------------------------+
```

**What is master data?**
- Excel file with employee information
- Contains: CRM, Name, Department, National ID, Vendor, PS ID, Join Date

---

### Step 4: Load Leave Sheet (Optional but Recommended)

1. Click the **"Select Leave Sheet"** button
2. Select your leave records Excel file
3. Click **"Open"**

```
+----------------------------------+
|  [Select Leave Sheet]            |  <-- Click here
|  Status: Leave data loaded       |
+----------------------------------+
```

**What is the leave sheet?**
- Excel file with employee leave records
- Contains monthly sheets (Jan, Feb, Mar, etc.)
- Shows who took Annual Leave, Sick Leave, etc.

---

## Using Employee Filters

After loading Master Data, you can filter which employees appear in the report.

### Filter by Department

```
+----------------------------------+
|  Department Filter               |
|  [x] Select All                  |
|  [x] IT Department               |
|  [x] HR Department               |
|  [ ] Finance Department          |  <-- Unchecked = excluded
+----------------------------------+
```

- **Check** the box to include that department
- **Uncheck** to exclude
- Use **"Select All"** to quickly select/deselect all

### Filter by CRM

```
+----------------------------------+
|  CRM Filter                      |
|  [x] Select All                  |
|  [x] CRM001                      |
|  [x] CRM002                      |
|  [ ] CRM003                      |  <-- Unchecked = excluded
+----------------------------------+
```

### Reset Filters

Click the **"Reset"** button to restore all selections.

```
Showing 45 of 50 employees    [Reset]
```

---

## Generating the Report

### Step 5: Generate Report

1. Make sure all required files are loaded (green checkmarks)
2. Click the **"Generate Report"** button
3. Choose where to save the Excel file
4. Wait for processing to complete

```
+----------------------------------+
|                                  |
|     [Generate Report]            |  <-- Click when ready
|                                  |
+----------------------------------+

Processing... Please wait...
============================
Report generated successfully!
```

---

## Understanding the Generated Excel Report

The report contains **4 sheets**:

### Sheet 1: Summary Report

This is your main working sheet.

```
+-------+--------+----------+----------+----------+----------+
| CRM   | Name   | 21-Dec   | 22-Dec   | 23-Dec   | 24-Dec   |
+-------+--------+----------+----------+----------+----------+
| CRM001| Ahmed  | Normal   | Late     | Normal   | Weekend  |
| CRM002| Sara   | Normal   | Normal   | Absent   | Weekend  |
+-------+--------+----------+----------+----------+----------+
```

**Color Guide:**
| Color | Meaning |
|-------|---------|
| Green | No penalty (Normal, Approved, Leaves) |
| Yellow | Late (with penalty) |
| Red | Absent (2 days deduction) |
| Pink | Missing Punch |
| Orange | Half Day / Early Departure |
| Light Blue | Sick Leave / Unpaid Leave |
| Purple | Backdated Leave (from previous month) |
| Gray | Weekend / OFF day |

---

### Sheet 2: Individual Analytics

Shows attendance statistics for each employee.

```
+-------+--------+--------+--------+--------+
| CRM   | Name   | Normal | Late   | Absent |
+-------+--------+--------+--------+--------+
| CRM001| Ahmed  | 18     | 2      | 1      |
| CRM002| Sara   | 19     | 1      | 1      |
+-------+--------+--------+--------+--------+
```

---

### Sheet 3: Alerts

Shows employees who need attention:
- High absences
- Frequent late arrivals
- Many missing punches

```
+-------+--------+------------------+
| CRM   | Name   | Alert            |
+-------+--------+------------------+
| CRM003| Mohamed| 5 Absences       |
| CRM007| Fatima | 6 Late arrivals  |
+-------+--------+------------------+
```

---

### Sheet 4: Penalties

This is your payroll deduction sheet.

```
+-------+--------+-------+--------+--------+---------+
| CRM   | Name   | Late  | Late   | Absent | Absent  |
|       |        | Count | Penalty| Count  | Deduct  |
+-------+--------+-------+--------+--------+---------+
| CRM001| Ahmed  | 2     | 300 EGP| 1      | 2 days  |
| CRM002| Sara   | 1     | 100 EGP| 1      | 2 days  |
+-------+--------+-------+--------+--------+---------+
```

**Penalties Sheet Columns:**
| Column | What it Shows |
|--------|---------------|
| A | CRM |
| B | Name |
| C | National ID |
| D | Vendor |
| E | PS ID |
| F | Department |
| G | Join Date |
| H | Late Count |
| I | Late Penalty (EGP) |
| J | Missing Punches |
| K | Punch Deduction (days) |
| L | Absences |
| M | Absence Deduction (days) |
| N | Early Departure Count |
| O | Early Departure Deduction |
| P | Half Day Count |
| Q | Half Day Deduction |
| R | Sick Leave Count |
| S | Sick Leave Deduction |
| T | Unpaid Leave Count |
| U | Unpaid Leave Deduction |
| V | Total Penalty (EGP) |
| W | Total Deduction (days) |
| X | Warnings |
| Y | Backdated Leaves |

---

## Changing Employee Status (Justifications)

You can change an employee's status directly in Excel!

### How to Change Status

1. Open the generated Excel file
2. Go to **Summary Report** sheet
3. Click on any date cell (e.g., "Late" or "Absent")
4. A **dropdown arrow** appears
5. Click the arrow and select new status

```
+----------+
| Late   ▼ |  <-- Click the arrow
+----------+
    |
    v
+------------------+
| Normal           |
| Late (Approved)  |  <-- Select to remove penalty
| Late             |
| Absent           |
| ...              |
+------------------+
```

### Important: Penalties Update Automatically!

When you change a status in Summary Report:
- The **Penalties sheet updates automatically**
- No need to recalculate anything

**Example:**
- Change "Late" to "Late (Approved)"
- Late count decreases by 1
- Late penalty recalculates automatically

---

## Dropdown Status Options

### No Penalty (Green)
| Status | Use When |
|--------|----------|
| Normal | Regular attendance, no issues |
| Late (Approved) | Late arrival was approved by manager |
| Early Departure (Approved) | Early leave was approved |
| Annual Leave | Planned vacation leave |
| Casual Leave | Personal/emergency leave |
| Marriage Leave | Wedding leave |
| Paternity Leave | Father's leave for new baby |
| Maternity Leave | Mother's leave for new baby |
| Bereavement Leave | Leave for family death |
| Military Call Leave | Called for military service |

### With Penalty
| Status | Penalty |
|--------|---------|
| Late | 100/200/500 EGP (see scale below) |
| Absent | 2 days deduction |
| Missing Punch In | 0.5 day after 3 occurrences |
| Missing Punch Out | 0.5 day after 3 occurrences |
| Early Departure | 0.5 day deduction |
| Half Day | 0.5 day deduction |
| Early Leave (HD) | 0.5 day deduction |
| Sick Leave | 0.25 day deduction |
| Unpaid Leave | 1 day deduction |

### Backdated Leaves (Purple)
Use these for leaves from **previous months**:
| Status | Penalty |
|--------|---------|
| Annual Leave (BD) | No deduction |
| Casual Leave (BD) | No deduction |
| Sick Leave (BD) | 0.25 day |
| Unpaid Leave (BD) | 1 day |
| Half Day (BD) | 0.5 day |
| Early Departure (BD) | 0.5 day |
| Early Leave (HD) (BD) | 0.5 day |

---

## Penalty Rules

### Late Penalty Scale

```
1st Late  =  100 EGP
2nd Late  =  200 EGP
3rd Late  =  500 EGP + Warning
4th Late  =  500 EGP
5th Late  =  500 EGP
... and so on
```

**Note:** "Late (Approved)" does NOT count toward late penalties!

---

### Missing Punch Rules

```
1st Missing Punch  =  No deduction
2nd Missing Punch  =  No deduction
3rd Missing Punch  =  No deduction
4th Missing Punch  =  0.5 day deduction
5th Missing Punch  =  0.5 day deduction
6th+ Missing Punch =  0.5 day + Warning
```

**Note:** All types count toward the limit:
- Missing Punch In
- Missing Punch In (Justified)
- Missing Punch Out
- Missing Punch Out (Justified)

---

## Adding Backdated Leaves

### What are Backdated Leaves?

Leaves from a **previous month** that weren't recorded at the time.

**Example:**
- Employee took sick leave in **November**
- You're processing **January payroll** (Dec 21 - Jan 20)
- You need to add this to the current payroll period

### How to Add Backdated Leave

**Method 1: In the Leave Sheet (Before Report)**

1. Open your Leave Sheet Excel file
2. Go to the current month's sheet (e.g., "Jan")
3. Find the employee's row
4. Select any date cell in the payroll period
5. Choose the (BD) version from dropdown:
   - "Sick Leave (BD)" instead of "Sick Leave"
   - "Annual Leave (BD)" instead of "Annual Leave"

**Method 2: In the Generated Report (After Report)**

1. Open the generated Excel report
2. Go to **Summary Report** sheet
3. Click on any date cell for that employee
4. Select the (BD) version from dropdown

### How to Identify Backdated Leaves

- **Purple color** in Summary Report
- **Column Y** in Penalties sheet shows count
- HR can review and verify these entries

---

## Tips & Tricks

### Tip 1: Force Recalculate

If penalties don't update automatically:
- Press **Ctrl + Shift + F9** in Excel

### Tip 2: Remove Late Penalty

To remove a late penalty:
- Change "Late" to "Late (Approved)"

### Tip 3: Remove Early Departure Penalty

To remove early departure penalty:
- Change "Early Departure" to "Early Departure (Approved)"

### Tip 4: Check Before Saving

Always check these before finalizing:
1. Penalties sheet totals
2. Backdated Leaves column
3. Warning count

### Tip 5: Save Original

Keep a copy of the original generated report before making changes.

---

## Troubleshooting

### Problem: Application won't open

**Solution:**
- Make sure you have the .exe file, not a shortcut
- Try running as Administrator (right-click > Run as administrator)

---

### Problem: "File not found" error

**Solution:**
- Make sure the Excel files are not open in another program
- Close Excel and try again

---

### Problem: Filters not showing

**Solution:**
- Load the Master Data file first
- Filters appear after Master Data is loaded

---

### Problem: Leaves not appearing in report

**Solution:**
- Make sure the CRM in Leave Sheet matches Master Data exactly
- Check that the leave date is within the attendance date range

---

### Problem: Penalties not updating when I change status

**Solution:**
1. Press **Ctrl + Shift + F9** to force recalculate
2. Make sure you're changing cells in the Summary Report sheet
3. Save and reopen the file

---

### Problem: Report shows extra months

**Solution:**
- This was fixed in v2.3
- Make sure you're using the latest version of the application

---

### Problem: Backdated leaves not counted in penalties

**Solution:**
- Use the (BD) suffix versions (e.g., "Sick Leave (BD)")
- Check Column Y in Penalties sheet for the count
- Press Ctrl + Shift + F9 to recalculate

---

## Quick Reference Card

Print this section for your desk!

```
+--------------------------------------------------+
|           ATTENDANCE DASHBOARD v2.3              |
|              QUICK REFERENCE                     |
+--------------------------------------------------+

STEP-BY-STEP:
1. Open AttendanceDashboard.exe
2. Click "Select Attendance Files" → choose files
3. Click "Select Master Data" → choose file
4. Click "Select Leave Sheet" → choose file
5. Apply filters if needed
6. Click "Generate Report"
7. Save Excel file

TO CHANGE STATUS IN EXCEL:
1. Open generated Excel
2. Go to "Summary Report" sheet
3. Click date cell → click dropdown arrow
4. Select new status
5. Penalties update automatically!

LATE PENALTIES:
• 1st = 100 EGP
• 2nd = 200 EGP
• 3rd+ = 500 EGP

DEDUCTIONS (DAYS):
• Absent = 2 days
• Missing Punch (4th+) = 0.5 day
• Early Departure = 0.5 day
• Half Day = 0.5 day
• Sick Leave = 0.25 day
• Unpaid Leave = 1 day

COLORS:
• Green = No penalty
• Yellow = Late
• Red = Absent
• Pink = Missing Punch
• Orange = Half Day
• Purple = Backdated (BD)
• Gray = Weekend

SHORTCUT: Ctrl+Shift+F9 = Recalculate

+--------------------------------------------------+
```

---

## Contact Support

If you encounter issues not covered in this guide, please contact:
- Your IT Department
- The Development Team

---

*User Guide - Attendance Dashboard v2.3*
*Last Updated: January 2026*
