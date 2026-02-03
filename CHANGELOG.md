# Changelog - Attendance Dashboard

All notable changes to the Attendance Dashboard application.

---

## [v2.3] - January 2026

### Added

#### Multi-Sheet Leave Format Support
- Automatically detects leave files with Master + monthly sheets (Jan-Dec)
- Processes each monthly sheet separately
- Skips Master sheet (employee data only)
- Parses datetime column headers for leave dates

#### Backdated Leaves (BD) Feature
- New `(BD)` suffix for leaves transferred from previous months
- 10 backdated leave variants added to dropdown options:
  - Annual Leave (BD)
  - Casual Leave (BD)
  - Sick Leave (BD)
  - Unpaid Leave (BD)
  - Half Day (BD)
  - Early Departure (BD)
  - Marriage Leave (BD)
  - Paternity Leave (BD)
  - Maternity Leave (BD)
  - Bereavement Leave (BD)

#### Backdated Leaves Column in Penalties Sheet
- New Column Y: "Backdated Leaves"
- Counts all leaves with (BD) suffix using COUNTIF formula
- Purple highlighting for easy identification
- Penalties sheet now has 25 columns (was 24)

#### Purple Color Coding
- All (BD) suffixed statuses display in light purple (#E1D5E7)
- Consistent across Summary Report and Penalties sheet
- Applied before other color rules for priority

#### Leave Sheet Data Validation
- Added dropdown lists to all date columns in leave sheets
- Hidden "LeaveTypes" sheet stores option list
- Named range "LeaveTypeList" for validation
- Bypasses Excel's 255 character limit for inline lists

### Changed
- `load_leave_data()` now detects multi-sheet vs single-sheet format
- `_apply_status_color()` checks for (BD) suffix first
- `create_penalties_sheet()` expanded to 25 columns
- Policy legend updated with backdated leaves explanation
- `fill_leave_records()` now uses master data CRMs instead of attendance CRMs
- Penalty formulas now count both regular and (BD) variants

### Fixed
- **Critical: Leaves not applied** - Fixed issue where leaves were only applied to employees with existing attendance records. Now uses master data CRMs.
- **Critical: Extra months in Summary** - Fixed date range expansion caused by leaves outside attendance period. Now strictly filters by attendance date range.
- **Critical: BD leaves not counted** - Fixed COUNTIF formula (`"*(BD)"` instead of `"*(BD)*"`) and updated all penalty formulas to include BD variants.
- Leave file with multi-sheet format no longer causes crash
- Date range calculation now works correctly with monthly sheets
- Added `Early Leave (HD)` support with proper color coding (Orange)
- Added lowercase `Unpaid leave` variant support
- Optimized leave parsing for better performance (pre-identifies date columns)
- Only loads relevant monthly sheets (current + previous month) for faster processing

---

## [v2.2] - January 2026

### Added

#### Employee Filters
- New filter section appears after loading Master Data
- Filter by Department with checkboxes
- Filter by CRM with checkboxes
- "Select All" checkbox for each filter category
- "Reset" button to restore all selections
- Live count display: "Showing X of Y employees"
- Window size increased from 600x560 to 600x700

#### Justification Dropdown System
- 21 justification options available in Summary Report
- Dropdown menu on every date cell
- Color-coded cells based on justification type:
  - Green: No deduction statuses
  - Yellow: Late (with deduction)
  - Red: Absent
  - Pink: Missing Punch types
  - Orange: Early Departure, Half Day
  - Light Blue: Sick Leave, Unpaid Leave
  - Gray: Weekend

#### Enhanced Penalties Sheet
- Expanded from 16 to 24 columns
- New employee info columns:
  - Vendor (from Column B of Master Data)
  - PS ID (from Column C of Master Data)
  - Join Date (from Column M of Master Data)
- New deduction columns:
  - Early Departure Count & Deduction (0.5 days)
  - Half Day Count & Deduction (0.5 days)
  - Sick Leave Count & Deduction (0.25 days)
  - Unpaid Leave Count & Deduction (1 day)
- Dynamic Excel formulas linked to Summary Report
- Auto-updating totals with SUM formulas

#### Missing Punch Differentiation
- "Missing Punch In" - when no clock in time
- "Missing Punch Out" - when no clock out time
- Replaces generic "Missing Punch" status

### Fixed

#### Join Date Parsing
- Fixed Excel serial date conversion (was showing 1970-01-01)
- Added support for numpy int64 type
- Added support for column names with newlines (e.g., "Join Date\n(yyyy/mm/dd)")

#### Column Detection
- Improved detection for columns with special characters
- Added more column name variations for Join Date
- Direct search for columns with newlines

### Changed

- Late Count formula now excludes "Late (Approved)"
- Missing Punches formula counts all 4 types:
  - Missing Punch In
  - Missing Punch In (Justified)
  - Missing Punch Out
  - Missing Punch Out (Justified)
- Updated policy legend with all deduction rules

---

## [v2.1] - Previous Version

### Features
- Basic attendance file processing
- Master data loading
- Leave record integration
- Simple penalties calculation
- Excel report generation with 4 sheets:
  - Summary Report
  - Individual Analytics
  - Alerts
  - Penalties

---

## Deduction Rules Reference

### Late Penalties (EGP)
| Count | Penalty |
|-------|---------|
| 1st | 100 |
| 2nd | 200 |
| 3rd+ | 500 + Warning |

### Day Deductions
| Type | Days |
|------|------|
| Absent | 2.0 |
| Missing Punch (after 3) | 0.5 |
| Early Departure | 0.5 |
| Half Day | 0.5 |
| Sick Leave | 0.25 |
| Unpaid Leave | 1.0 |

### No Deduction
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

## File Changes Summary

### Modified Files
- `attendance_dashboard_app.py` - Main application (~300 lines added/modified)

### New Files
- `DOCUMENTATION.md` - Comprehensive user documentation
- `CHANGELOG.md` - Version history and changes

### Key Code Changes

#### New Methods Added
- `create_filter_section()` - Creates filter UI
- `populate_filters()` - Populates filter checkboxes
- `toggle_all_departments()` - Select all departments
- `toggle_all_crms()` - Select all CRMs
- `update_filter_counts()` - Updates filter status
- `get_filtered_employee_mapping()` - Returns filtered employees
- `reset_filters()` - Resets all filters
- `_apply_status_color()` - Applies color to status cells

#### Modified Methods
- `__init__` - Added filter variables, increased window size
- `create_upload_section()` - Stored container reference
- `select_master_data()` - Calls populate_filters()
- `load_master_data()` - Added Vendor, PS ID, Join Date extraction
- `determine_status()` - Differentiated Missing Punch In/Out
- `process_single_attendance_file()` - Uses filtered mapping
- `generate_report()` - Added filter validation
- `calculate_penalties()` - Added new fields
- `create_summary_sheet()` - Added dropdown validation
- `create_penalties_sheet()` - Expanded to 24 columns with formulas
- `clear_all()` - Clears filter state
