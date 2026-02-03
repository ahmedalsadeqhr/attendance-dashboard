# Attendance Dashboard v2.2

Professional attendance processing application with both Desktop (CustomTkinter) and Web (Streamlit) interfaces.

## Features

- **Multi-file Processing**: Load master data, multiple attendance files, and leave sheets
- **Intelligent Filtering**: Filter by department and CRM
- **Resigned Employee Detection**: Automatically detects exit dates and marks resigned employees
- **Leave vs Attendance Conflict Detection**: Identifies conflicts in a dedicated "Duplicates" sheet
- **Automated Penalty Calculation**: Based on configurable attendance policy
- **Color-coded Reports**: Easy-to-read Excel output with status highlighting

## Versions

### Web Version (Streamlit)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

```bash
pip install -r requirements.txt
streamlit run attendance_dashboard_streamlit.py
```

### Desktop Version (CustomTkinter)

```bash
pip install -r requirements.txt
python attendance_dashboard_app.py
```

Or double-click `RUN_APP.bat`

## Requirements

```
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.0.0
xlrd>=2.0.0
customtkinter>=5.2.0  # Desktop only
streamlit>=1.28.0     # Web only
```

## Input Files

1. **Master Data** (Excel) - Employee PS ID, CRM, Name, Department, Exit Date (optional)
2. **Attendance Files** (Excel) - Daily clock in/out records
3. **Leave Sheet** (Optional) - Leave records with CRM, Date, Type

## Output (5-Sheet Report)

1. **Summary Report** - Color-coded attendance matrix with dropdown justifications
2. **Individual Analytics** - Per-employee statistics
3. **Alerts** - Flagged issues & warnings
4. **Penalties** - Automated penalty calculations linked to summary
5. **Duplicates** - Leave vs attendance conflicts

## Status Color Coding

| Status | Color |
|--------|-------|
| Normal | Green |
| Late | Yellow |
| Absent | Red |
| Missing Punch | Pink |
| Weekend | Gray |
| Resigned | Light Gray |
| Leave (various) | Blue/Light Blue |
| Backdated (BD) | Purple |

## Keyboard Shortcuts (Desktop)

- `Ctrl+M` - Select Master
- `Ctrl+A` - Select Attendance
- `Ctrl+L` - Select Leave
- `Ctrl+G` - Generate Report
- `Ctrl+R` - Clear All

## License

MIT License

---

v2.2 | February 2026
