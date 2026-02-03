# Attendance Dashboard v2.3 - Documentation

## Overview

The Attendance Dashboard is a Python-based application that processes attendance data and generates comprehensive Excel reports with penalty calculations based on company attendance policies.

---

## Table of Contents

1. [Features](#features)
2. [Installation & Requirements](#installation--requirements)
3. [Master Data Format](#master-data-format)
4. [Leave Sheet Format](#leave-sheet-format)
5. [Employee Filters](#employee-filters)
6. [Justification System](#justification-system)
7. [Backdated Leaves](#backdated-leaves)
8. [Penalties Calculation](#penalties-calculation)
9. [Generated Reports](#generated-reports)
10. [Troubleshooting](#troubleshooting)

---

## Features

### Core Features
- Process multiple attendance files simultaneously
- Load employee master data with automatic column detection
- Load leave records for accurate status tracking
- Generate comprehensive Excel reports

### New Features (v2.3)
- **Multi-Sheet Leave Format**: Supports leave files with Master + monthly sheets (Jan-Dec)
- **Backdated Leaves (BD)**: Track leaves transferred from previous months with (BD) suffix
- **Backdated Leaves Column**: New column in Penalties sheet to track backdated entries
- **Leave Sheet Data Validation**: Dropdown lists in leave sheets for easy data entry

### Features (v2.2)
- **Employee Filters**: Filter by Department and/or CRM before generating reports
- **Justification Dropdowns**: Each date cell in Summary Report has a dropdown menu for status justification
- **Enhanced Penalties Sheet**:
  - Additional columns: Vendor, PS ID, National ID, Join Date
  - New deduction types: Early Departure, Half Day, Sick Leave, Unpaid Leave
  - Dynamic linking to Summary Report (changes auto-update penalties)
- **Missing Punch Differentiation**: Distinguishes between "Missing Punch In" and "Missing Punch Out"

---

## Installation & Requirements

### Required Python Packages
```
pandas
openpyxl
customtkinter
```

### Installation
```bash
pip install pandas openpyxl customtkinter
```

### Running the Application
```bash
python attendance_dashboard_app.py
```

---

## Master Data Format

The application automatically detects columns from your Master Data Excel file.

### Required Columns
| Column | Detected Names |
|--------|---------------|
| Employee ID | AC-No., Ac-No., AC No, PS ID |
| CRM | CRM |
| Name | Name |

### Optional Columns
| Column | Detected Names |
|--------|---------------|
| Department | Department, Dept |
| Position | Position, Title |
| National ID | Identity Number, Idnetity Number, National ID, NID |
| Vendor | Vendor |
| PS ID | PS ID, PS Id, PSID, PS-ID |
| Join Date | Join Date, Joining Date, Date of Joining, Hire Date, DOJ |

### Notes
- Column names with newlines are supported (e.g., "Join Date\n(yyyy/mm/dd)")
- Excel serial dates are automatically converted to readable format
- Column detection is case-insensitive

---

## Leave Sheet Format

The application supports two leave sheet formats:

### Multi-Sheet Format (Recommended)
A workbook with multiple sheets:

| Sheet | Content |
|-------|---------|
| Master | Employee master data (optional, skipped) |
| Jan | January leave data |
| Feb | February leave data |
| ... | ... |
| Dec | December leave data |

Each monthly sheet structure:
- **Columns 1-30**: Employee info (CRM, Name, balances, etc.)
- **Columns 31+**: Date columns as datetime headers
- **Cell values**: Leave type (e.g., "Annual Leave", "Sick Leave")

### Single-Sheet Format (Legacy)
A single sheet with:
- CRM column
- Date columns as strings (e.g., "1-Jan", "2-Jan")

### Leave Types Available
| Leave Type | Deduction |
|------------|-----------|
| Annual Leave | No deduction |
| Casual Leave | No deduction |
| Sick Leave | 0.25 day |
| Unpaid Leave | 1 day |
| Early Leave (HD) | 0.5 day |
| Half Day | 0.5 day |
| Marriage Leave | No deduction |
| Paternity Leave | No deduction |
| Maternity Leave | No deduction |
| Bereavement Leave | No deduction |
| Military Call Leave | No deduction |
| Pilgrimage Leave | No deduction |

---

## Employee Filters

### How to Use
1. Load Master Data file
2. Filter section appears automatically with Department and CRM options
3. Check/uncheck departments and CRMs to include/exclude
4. Use "Select All" checkbox for quick selection
5. Click "Reset" to restore all selections
6. Status shows "Showing X of Y employees"

### Filter Behavior
- Employees must match BOTH selected department AND selected CRM
- Unmatched employees (not in master data) are excluded when filters are active
- Filters apply to the generated report only

---

## Justification System

### Dropdown Options

Each date cell in the Summary Report has a dropdown with these options:

| Justification | Deduction | Color |
|--------------|-----------|-------|
| Normal | No Deduction | Green |
| Late (Approved) | No Deduction | Green |
| Late | Late Penalty (EGP) | Yellow |
| Absent | 2 days | Red |
| Missing Punch In | 0.5 day after 3 occurrences | Pink |
| Missing Punch In (Justified) | 0.5 day after 3 occurrences | Pink |
| Missing Punch Out | 0.5 day after 3 occurrences | Pink |
| Missing Punch Out (Justified) | 0.5 day after 3 occurrences | Pink |
| Early Departure (Approved) | No Deduction | Green |
| Early Departure | 0.5 day | Orange |
| Half Day | 0.5 day | Orange |
| Sick Leave | 0.25 day | Light Blue |
| Annual Leave | No Deduction | Green |
| Casual Leave | No Deduction | Green |
| Marriage Leave | No Deduction | Green |
| Paternity Leave | No Deduction | Green |
| Maternity Leave | No Deduction | Green |
| Bereavement Leave | No Deduction | Green |
| Military Call Leave | No Deduction | Green |
| Unpaid Leave | 1 day | Light Blue |
| Weekend | N/A | Gray |

### Backdated Leave Options (BD)
All leave types also have backdated variants with `(BD)` suffix:

| Justification | Deduction | Color |
|--------------|-----------|-------|
| Annual Leave (BD) | No Deduction | Purple |
| Casual Leave (BD) | No Deduction | Purple |
| Sick Leave (BD) | 0.25 day | Purple |
| Unpaid Leave (BD) | 1 day | Purple |
| Half Day (BD) | 0.5 day | Purple |
| Early Departure (BD) | 0.5 day | Purple |
| Marriage Leave (BD) | No Deduction | Purple |
| Paternity Leave (BD) | No Deduction | Purple |
| Maternity Leave (BD) | No Deduction | Purple |
| Bereavement Leave (BD) | No Deduction | Purple |

### Automatic Status Detection
- **No Clock In + No Clock Out** = Absent
- **No Clock In only** = Missing Punch In
- **No Clock Out only** = Missing Punch Out
- **Clock In after threshold** = Late
- **OFF day with no punches** = Weekend
- **OFF day with punches** = Worked on Day Off

---

## Backdated Leaves

### What are Backdated Leaves?
Backdated leaves are leave records transferred from previous months that need to be recorded in the current payroll period.

### Example Scenario
- Employee took leave in **November** but it wasn't recorded
- You're processing **January payroll** (Dec 21 - Jan 20)
- You need to record this leave in the current period

### How to Add Backdated Leaves

1. Open the **Leave Sheet** (e.g., `Updated Leaves.xlsx`)
2. Go to the **Jan** sheet (or current month)
3. Find the employee's row by CRM
4. In any date column within the payroll period, select from dropdown:
   - `Annual Leave (BD)` instead of `Annual Leave`
   - `Sick Leave (BD)` instead of `Sick Leave`
   - etc.

### Backdated Leave Rules
- **Same deduction** as regular leave types
- **Purple color** in Summary Report for easy identification
- **Separate column** in Penalties sheet showing backdated count
- **HR Review**: Backdated column helps HR identify transferred records

### Payroll Cycle Note
If your payroll cycle is Dec 21 - Jan 20:
- Add backdated leaves to a date within this range (e.g., Jan 1)
- The leave will count toward January's penalties
- The (BD) suffix identifies it as transferred from another month

---

## Penalties Calculation

### Late Penalties (EGP)
| Occurrence | Penalty |
|------------|---------|
| 1st Late | 100 EGP |
| 2nd Late | 200 EGP |
| 3rd Late | 500 EGP + Warning |
| 4th+ Late | 500 EGP each |

**Note**: "Late (Approved)" does NOT count toward late penalties.

### Missing Punch Rules
- First 3 occurrences: No deduction
- 4th+ occurrence: 0.5 day deduction each
- 6+ occurrences: Warning issued
- All 4 types count: Missing Punch In, Missing Punch In (Justified), Missing Punch Out, Missing Punch Out (Justified)

### Other Deductions
| Type | Deduction per Occurrence |
|------|-------------------------|
| Absent | 2 days |
| Early Departure | 0.5 day |
| Half Day | 0.5 day |
| Sick Leave | 0.25 day |
| Unpaid Leave | 1 day |

### No Deduction Types
- Normal
- Late (Approved)
- Early Departure (Approved)
- Annual Leave
- Casual Leave
- Marriage Leave
- Paternity Leave
- Maternity Leave
- Bereavement Leave
- Military Call Leave
- Weekend

---

## Generated Reports

### Excel Workbook Structure

#### Sheet 1: Summary Report
- Employee attendance matrix with dates as columns
- Dropdown menus for justification on each date cell
- Color-coded cells based on status
- Normal Days and Abnormal Days counts

#### Sheet 2: Individual Analytics
- Per-employee attendance statistics
- Breakdown of status counts

#### Sheet 3: Alerts
- Employees with high absences
- Employees with frequent late arrivals
- Employees with many missing punches

#### Sheet 4: Penalties
25 columns with dynamic formulas linked to Summary Report:

| Column | Description |
|--------|-------------|
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
| Y | **Backdated Leaves** (Purple) |

### Dynamic Linking
- Penalties sheet uses COUNTIF formulas referencing Summary Report
- When you change a status in Summary Report dropdown, penalties auto-update
- Example: Change "Late" to "Late (Approved)" → Late count decreases by 1

---

## Troubleshooting

### Join Date Shows 1970-01-01
- **Cause**: Excel serial dates not being parsed correctly
- **Solution**: Fixed in v2.2 - supports Excel serial dates and column names with newlines

### Filters Not Showing
- **Cause**: Master data not loaded yet
- **Solution**: Load master data file first, filters appear automatically

### Column Not Detected
- **Cause**: Column name doesn't match expected patterns
- **Solution**: Check the column name matches one of the detected names listed above

### Missing Punch Shows Wrong Type
- **Cause**: Old version used generic "Missing Punch"
- **Solution**: v2.2 now differentiates between "Missing Punch In" and "Missing Punch Out"

### Penalties Not Updating
- **Cause**: Excel formulas need recalculation
- **Solution**: Press Ctrl+Shift+F9 in Excel to force recalculation

### Leaves Not Appearing in Report
- **Cause**: (v2.3 Fixed) Leaves were only applied to employees with attendance records
- **Solution**: v2.3 now uses master data CRMs, so leaves apply even without attendance

### Summary Report Shows Extra Months
- **Cause**: (v2.3 Fixed) Leaves outside attendance date range were expanding the report
- **Solution**: v2.3 strictly filters leaves to attendance date range only

### Backdated Leaves Not Counted in Penalties
- **Cause**: (v2.3 Fixed) COUNTIF formulas didn't include BD variants
- **Solution**: v2.3 penalty formulas now count both regular and (BD) variants

### App Running Slowly
- **Cause**: Loading all 12 monthly sheets from leave file
- **Solution**: v2.3 optimizes by only loading relevant months (current + previous)

---

## Configuration

### Default Settings (config.json)
```json
{
  "work_start_time": "9:00 AM",
  "work_end_time": "5:00 PM",
  "late_threshold": "12:00 PM",
  "off_days": [4],
  "penalties": {
    "currency": "EGP",
    "late_1st": 100,
    "late_2nd": 200,
    "late_3rd": 500,
    "late_4th_plus": 500,
    "missing_punch_threshold": 3,
    "missing_punch_deduction": 0.5,
    "absence_deduction": 2
  }
}
```

### OFF Days
- 0 = Monday
- 1 = Tuesday
- 2 = Wednesday
- 3 = Thursday
- 4 = Friday (default)
- 5 = Saturday
- 6 = Sunday

---

## Version History

### v2.3 (Current)
- **Multi-Sheet Leave Format**: Support for leave files with Master + Jan-Dec sheets
- **Backdated Leaves (BD)**: New suffix to track leaves transferred from previous months
- **Backdated Leaves Column**: New column Y in Penalties sheet (now 25 columns)
- **Purple Color Coding**: Backdated leaves highlighted in purple
- **Leave Sheet Data Validation**: Dropdown lists added to leave sheet date columns
- **Named Range Validation**: Uses hidden LeaveTypes sheet for dropdown options

### v2.2
- Added Employee Filters (Department/CRM)
- Added Justification dropdown system
- Enhanced Penalties sheet with new columns
- Added Vendor, PS ID, National ID, Join Date
- Added Early Departure, Half Day, Sick Leave, Unpaid Leave deductions
- Dynamic formula linking between sheets
- Differentiated Missing Punch In vs Out
- Fixed Excel serial date parsing
- Fixed column detection with newlines

### v2.1
- Basic attendance processing
- Simple penalties calculation
- Leave record integration

---

## Support

For issues or feature requests, please contact the development team.

---

*Documentation last updated: January 2026 (v2.3)*
