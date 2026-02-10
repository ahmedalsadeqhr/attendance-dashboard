# Attendance Dashboard - Project Overview

Professional attendance processing with automated penalty calculation for 51Talk HR teams.

---

## Available Versions

| Version | Type | Status | Best For |
|---------|------|--------|----------|
| **v4.0** | Pure HTML/JS | **Recommended** | Sharing with teammates |
| v3.0 | Flask + HTML | Working | Local web interface |
| v2.2 | CustomTkinter | Working | Desktop app with EXE |

---

## v4.0 - Pure HTML/JavaScript (Recommended)

**File:** `attendance_dashboard.html` (root folder)

**Why v4.0?**
- No installation required - just open in browser
- Share single file via email/SharePoint/Google Drive
- Works on any OS (Windows, Mac, Linux)
- Modern UI with animations and 51Talk branding

**How to Use:**
1. Double-click `attendance_dashboard.html`
2. Select Master Data (employee list)
3. Select Attendance file(s)
4. (Optional) Select Leave sheet
5. Click "Generate Report"
6. Excel downloads automatically

**Features:**
- Embedded 51Talk logo (no external files needed)
- Gradient headers and animated buttons
- Drag & drop file selection
- Animated progress bar with percentage
- Color-coded status cells
- 4-sheet Excel output
- Keyboard shortcuts (Ctrl+M/A/L/G/R)
- Settings saved in browser

---

## v3.0 - Flask + HTML

**Folder:** `attendance_v3/`

**How to Use:**
1. Run `RUN_APP.bat`
2. Opens browser at http://127.0.0.1:5000
3. Same workflow as v4.0

**When to Use:**
- Need server-side processing for large files
- Prefer Python backend

---

## v2.2 - CustomTkinter Desktop

**Folder:** `files (1)/`

**How to Use:**
1. Run `RUN_APP.bat`
2. Desktop app opens
3. Same workflow as other versions

**When to Use:**
- Need standalone EXE (run `CREATE_EXE.bat`)
- Prefer desktop application

---

## Input Files

### Master Data (Required)
| Column | Description |
|--------|-------------|
| PS ID / AC-No. | Employee ID |
| CRM | Employee code |
| Name | Employee name |

### Attendance Files (Required)
| Column | Description |
|--------|-------------|
| AC-No. | Employee ID |
| Date | Record date |
| Clock In/Out 1-5 | Time records |

### Leave Sheet (Optional)
Matrix or vertical format with CRM and leave types.

---

## Output Report (4 Sheets)

### Sheet 1: Summary Report
Matrix view - employees vs dates with color-coded status.

### Sheet 2: Individual Analytics
Per-employee statistics and attendance rates.

### Sheet 3: Alerts & Warnings
Flagged issues (high absences, frequent late, etc.)

### Sheet 4: Penalties
Automated penalty calculations based on Policy 2026.

---

## Status Types

| Status | Color | Description |
|--------|-------|-------------|
| Normal | Green | On time |
| Late | Yellow | After 12:00 PM |
| Absent | Red | No record |
| Missing Punch | Pink | Partial record |
| Weekend | Gray | Friday (OFF) |
| Leave | Blue | Approved leave |

---

## Penalty Rules (Policy 2026)

| Violation | Penalty |
|-----------|---------|
| Late (1st) | EGP 100 |
| Late (2nd) | EGP 200 |
| Late (3rd+) | EGP 500 + warning |
| Missing Punch | 0.5 day after 3x |
| Absence | 2 days per occurrence |

---

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| Ctrl+M | Select Master Data |
| Ctrl+A | Select Attendance |
| Ctrl+L | Select Leave Sheet |
| Ctrl+G | Generate Report |
| Ctrl+R | Clear All |

---

## Folder Structure

```
V.5/
├── attendance_dashboard.html    <- v4.0 (RECOMMENDED)
├── attendance_v3/               <- v3.0 Flask version
│   ├── main.py
│   ├── RUN_APP.bat
│   └── web/
└── files (1)/                   <- v2.2 Desktop version
    ├── attendance_dashboard_app.py
    ├── RUN_APP.bat
    ├── CREATE_EXE.bat
    ├── dist/                    <- EXE output
    ├── docs/                    <- Documentation
    ├── test data/               <- Sample files
    └── Policy/                  <- HR policy docs
```

---

## Quick Start Guide

**For Most Users (v4.0):**
```
1. Open: attendance_dashboard.html
2. Select files
3. Click Generate
4. Done!
```

**For Desktop Users (v2.2):**
```
1. Run: files (1)/RUN_APP.bat
2. Select files
3. Click Generate
4. Done!
```

---

v4.0 | January 2026 | 51Talk HR Team
