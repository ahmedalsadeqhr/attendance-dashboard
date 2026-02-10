# Attendance Dashboard - Changelog

Complete version history for all releases.

---

## v4.0 - Pure HTML/JavaScript (January 2026)

**Status:** Complete | **Type:** Shareable HTML

### New in v4.0
- **Standalone HTML file** - No installation, no server, just open in browser
- **Cross-platform** - Works on Windows, Mac, Linux
- **Shareable** - Single file to email or upload

### UI Enhancements
- Embedded 51Talk logo (base64 encoded)
- Gradient header with glow effects
- Animated generate button with shimmer
- File type icons (users, clock, calendar)
- Animated progress bar with percentage
- Modal slide-in animations
- Success checkmark animation
- Footer with brand gradient

### Technical
- Tailwind CSS (CDN)
- xlsx-js-style for Excel output
- localStorage for settings
- ~2200 lines total

---

## v3.0 - Flask + HTML (January 2026)

**Status:** Working | **Type:** Local Web Server

### New in v3.0
- **Web interface** - Modern HTML/Tailwind UI
- **Flask backend** - Python processing power
- **Same features** as v2.2 with better UI

### Technical
- Flask web framework
- Tailwind CSS styling
- RESTful API endpoints
- File dialog integration

---

## v2.2 - CustomTkinter Desktop (January 2026)

**Status:** Production | **Type:** Desktop Application

### New in v2.2
- **Modern UI** - CustomTkinter with rounded corners
- **Penalty Sheet** - Automated Policy 2026 calculations
- **Keyboard Shortcuts** - Ctrl+M/A/L/G/R
- **Friday = Weekend** - Always marked as OFF

### Penalty Calculations
| Violation | Penalty |
|-----------|---------|
| Late (1st) | EGP 100 |
| Late (2nd) | EGP 200 |
| Late (3rd+) | EGP 500 + warning |
| Missing Punch | 0.5 day after 3x |
| Absence | 2 days per occurrence |

### 4-Sheet Report
1. Summary - Color-coded attendance matrix
2. Analytics - Per-employee statistics
3. Alerts - Flagged issues & warnings
4. Penalties - Automated calculations

---

## v2.1 - Bug Fixes (December 2025)

### Fixes
- Settings dialog with persistence
- Rotating log files
- File validation improvements
- Column detection (PS ID / AC-No.)
- .xls file support (xlrd)
- CRM matching (ID normalization)
- Full date range generation
- Blank cells show "Absent"

---

## v2.0 - Initial Release (December 2025)

### Features
- Basic tkinter UI
- Excel file processing
- 3-sheet report output
- Color-coded status cells

---

## Status Color Legend

| Status | Color | Hex |
|--------|-------|-----|
| Normal | Green | #C6EFCE |
| Late | Yellow | #FFEB9C |
| Absent | Red | #FFC7CE |
| Missing Punch | Pink | #FFB6C1 |
| Weekend | Gray | #D9D9D9 |
| Leave | Blue | #BDD7EE |

---

## Dependencies by Version

### v4.0 (Browser)
- Tailwind CSS (CDN)
- xlsx-js-style (CDN)

### v3.0 (Flask)
- Flask
- pandas
- openpyxl

### v2.2 (Desktop)
- customtkinter>=5.2.0
- pandas>=1.3.0
- openpyxl>=3.0.9
- xlsxwriter>=3.0.0
- xlrd>=2.0.0

---

January 2026 | 51Talk HR Team
