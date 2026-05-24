"""
Self-contained tests for the new iTalent HR leave export parser.
Covers: format detection, Passed-only filter, date-range expansion,
        half-day string parsing, Employee ID -> CRM lookup, and
        integration with fill_leave_records.
"""
import io
import sys
import os
from datetime import datetime, timedelta
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from attendance_dashboard_streamlit import AttendanceProcessor, DEFAULT_CONFIG

PASS = 0
FAIL = 0


def check(label, condition, detail=""):
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f"  [+] {label}" + (f"  ({detail})" if detail else ""))
    else:
        FAIL += 1
        print(f"  [X] {label}" + (f"  => {detail}" if detail else ""))


def section(title):
    print(f"\n{'='*60}\n  {title}\n{'='*60}")


def make_italent_excel(rows, include_label_row=True):
    """Build an in-memory iTalent-format Excel file (BytesIO).

    Row 0 is the human-readable label row (mimicking the real export).
    Row 1+ are data rows.
    """
    label_row = {
        "StaffId-ExportName": "Employee",
        "StaffId-Email": "Employee Email",
        "OIdVacationType": "Leave Project",
        "LookupPrefix_StaffId_OIdDepartment": "Department",
        "AttendanceOrgId": "Attendance Department",
        "LookupPrefix_StaffId_Gender": "Gender",
        "LookupPrefix_StaffId_EmploymentType": "Employee Type",
        "LookupPrefix_StaffId_EmployeeStatus": "Employee Status",
        "LookupPrefix_StaffId_JobNumber": "Employee ID",
        "VacationStartDateTime": "Start Time",
        "VacationStopDateTime": "End Time",
        "VacationDurationIncludeUnit": "Total Duration of Leave",
        "CreatedTime": "Application Time",
        "DocumentType": "Document Type",
        "Reason": "Reason",
        "Attachment": "Attachments",
        "ApproveStatus": "Approval Status",
        "ErrorLog": "Reason for deletion failure",
    }

    col_names = list(label_row.keys())
    friendly = list(label_row.values())

    data_rows = []
    if include_label_row:
        data_rows.append(friendly)

    for row in rows:
        data_rows.append([
            row.get("Employee", ""),
            row.get("Email", ""),
            row.get("Leave Project", "Annual Leave"),
            row.get("Department", ""),
            row.get("Department", ""),
            row.get("Gender", ""),
            "",
            row.get("Status", "Formal"),
            row.get("Employee ID", 0),
            row.get("Start Time", ""),
            row.get("End Time", ""),
            row.get("Duration", "1 Day(s)"),
            "",
            "Application",
            "",
            "",
            row.get("Approval Status", "Pending"),
            "",
        ])

    df = pd.DataFrame(data_rows, columns=col_names)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def make_processor_with_employees(employees):
    p = AttendanceProcessor()
    for emp in employees:
        ac_no = str(emp["ac_no"])
        p.employee_mapping[ac_no] = {
            "crm": emp["crm"],
            "name": emp.get("name", emp["crm"]),
            "department": emp.get("department", ""),
            "position": "",
            "national_id": "",
            "vendor": "",
            "ps_id": ac_no,
            "join_date": "",
            "exit_date": "",
        }
    p.crm_to_info = {
        info["crm"]: info
        for info in p.employee_mapping.values()
        if info["crm"]
    }
    return p


# ── TEST 1: Format detection ──────────────────────────────────────────────────
section("TEST 1: iTalent format detection")

buf = make_italent_excel([
    {"Employee ID": 100, "Approval Status": "Passed",
     "Start Time": "2026-05-10 12:00:00", "End Time": "2026-05-10 21:00:00"},
])
xlsx = pd.ExcelFile(buf)
buf.seek(0)
df_probe = pd.read_excel(buf, nrows=2, skiprows=1)
cols_lower = {str(c).strip().lower() for c in df_probe.columns}

check("probe detects 'approval status'", "approval status" in cols_lower)
check("probe detects 'employee id'", "employee id" in cols_lower)
check("probe detects 'start time'", "start time" in cols_lower)
check("all three required cols present", all(c in cols_lower for c in ["approval status", "employee id", "start time"]))


# ── TEST 2: Only Passed records loaded ────────────────────────────────────────
section("TEST 2: Only 'Passed' records are loaded")

p2 = make_processor_with_employees([
    {"ac_no": 200, "crm": "EMP200"},
    {"ac_no": 201, "crm": "EMP201"},
    {"ac_no": 202, "crm": "EMP202"},
    {"ac_no": 203, "crm": "EMP203"},
    {"ac_no": 204, "crm": "EMP204"},
])

buf2 = make_italent_excel([
    {"Employee ID": 200, "Approval Status": "Passed",
     "Start Time": "2026-05-10 12:00:00", "End Time": "2026-05-10 21:00:00"},
    {"Employee ID": 201, "Approval Status": "Pending",
     "Start Time": "2026-05-11 12:00:00", "End Time": "2026-05-11 21:00:00"},
    {"Employee ID": 202, "Approval Status": "Failed",
     "Start Time": "2026-05-12 12:00:00", "End Time": "2026-05-12 21:00:00"},
    {"Employee ID": 203, "Approval Status": "Returned",
     "Start Time": "2026-05-13 12:00:00", "End Time": "2026-05-13 21:00:00"},
    {"Employee ID": 204, "Approval Status": "Draft",
     "Start Time": "2026-05-14 12:00:00", "End Time": "2026-05-14 21:00:00"},
])

ok2 = p2.load_leave_data(buf2, "test_leave.xlsx")
check("load_leave_data returns True", ok2)
check("exactly 1 leave record loaded (only Passed)", len(p2.leave_records) == 1,
      f"got {len(p2.leave_records)}")
check("loaded record belongs to EMP200", p2.leave_records[0]["crm"] == "EMP200",
      f"got {p2.leave_records[0]['crm'] if p2.leave_records else 'none'}")


# ── TEST 3: Date range expansion (multi-day leave) ────────────────────────────
section("TEST 3: Date range expanded into individual days")

p3 = make_processor_with_employees([{"ac_no": 300, "crm": "EMP300"}])

buf3 = make_italent_excel([
    {"Employee ID": 300, "Approval Status": "Passed",
     "Start Time": "2026-05-11 12:00:00", "End Time": "2026-05-13 21:00:00",
     "Duration": "3 Day(s)", "Leave Project": "Annual Leave"},
])

p3.load_leave_data(buf3, "test_leave.xlsx")

dates_loaded = sorted(r["date"].date() for r in p3.leave_records)
expected_dates = [
    datetime(2026, 5, 11).date(),
    datetime(2026, 5, 12).date(),
    datetime(2026, 5, 13).date(),
]

check("3 calendar days expanded", len(p3.leave_records) == 3,
      f"got {len(p3.leave_records)}")
check("correct dates: May 11-13", dates_loaded == expected_dates,
      f"got {dates_loaded}")
check("leave_type is Annual Leave",
      all(r["leave_type"] == "Annual Leave" for r in p3.leave_records))


# ── TEST 4: Half-day string parsing ──────────────────────────────────────────
section("TEST 4: Half-day strings parsed correctly")

p4 = AttendanceProcessor()
parse = p4._parse_italent_datetime

# Normal datetime string
check("normal datetime parsed",
      parse("2026-05-10 12:00:00") == pd.Timestamp("2026-05-10 12:00:00"))

# Half-day with non-breaking space
check("First half day parsed",
      parse("2026/05/14 First\xa0half\xa0day").date() == datetime(2026, 5, 14).date())

check("Second half day parsed",
      parse("2026/05/09 Second\xa0half\xa0day").date() == datetime(2026, 5, 9).date())

# Regular half-day with normal space
check("First half day (normal space) parsed",
      parse("2026/05/14 First half day").date() == datetime(2026, 5, 14).date())

# datetime object passed through
ts = pd.Timestamp("2026-06-01 12:00:00")
check("Timestamp passed through unchanged", parse(ts) == ts)

# NaN/None returns None
check("NaN returns None", parse(float("nan")) is None)
check("None returns None", parse(None) is None)


# ── TEST 5: Half-day records load correctly end-to-end ───────────────────────
section("TEST 5: Half-day leave records load without skipping")

p5 = make_processor_with_employees([{"ac_no": 500, "crm": "EMP500"}])

# Build a DataFrame directly to use the half-day string format
df5_label = pd.DataFrame([[
    "Employee", "Employee Email", "Leave Project", "Department",
    "Attendance Department", "Gender", "Employee Type", "Employee Status",
    "Employee ID", "Start Time", "End Time", "Total Duration of Leave",
    "Application Time", "Document Type", "Reason", "Attachments",
    "Approval Status", "Reason for deletion failure",
]])
df5_data = pd.DataFrame([[
    "Test Employee", "test@example.com", "Annual Leave", "Dept A",
    "Dept A", "Male", "", "Formal",
    500,
    "2026/05/14 First\xa0half\xa0day",
    "2026/05/14 Second\xa0half\xa0day",
    "0.5 Day(s)",
    "", "Application", "", "",
    "Passed", "",
]])

df5 = pd.concat([df5_label, df5_data], ignore_index=True)
df5.columns = [
    "StaffId-ExportName", "StaffId-Email", "OIdVacationType",
    "LookupPrefix_StaffId_OIdDepartment", "AttendanceOrgId",
    "LookupPrefix_StaffId_Gender", "LookupPrefix_StaffId_EmploymentType",
    "LookupPrefix_StaffId_EmployeeStatus", "LookupPrefix_StaffId_JobNumber",
    "VacationStartDateTime", "VacationStopDateTime",
    "VacationDurationIncludeUnit", "CreatedTime", "DocumentType",
    "Reason", "Attachment", "ApproveStatus", "ErrorLog",
]
buf5 = io.BytesIO()
df5.to_excel(buf5, index=False)
buf5.seek(0)

ok5 = p5.load_leave_data(buf5, "test_leave.xlsx")
check("load_leave_data returns True", ok5)
check("half-day record loaded (not skipped)", len(p5.leave_records) == 1,
      f"got {len(p5.leave_records)}")
if p5.leave_records:
    check("date is 2026-05-14",
          p5.leave_records[0]["date"].date() == datetime(2026, 5, 14).date(),
          f"got {p5.leave_records[0]['date'].date()}")


# ── TEST 6: Unknown Employee ID is skipped with a warning ────────────────────
section("TEST 6: Unmatched Employee ID skipped with warning logged")

p6 = make_processor_with_employees([{"ac_no": 600, "crm": "EMP600"}])
buf6 = make_italent_excel([
    {"Employee ID": 600, "Approval Status": "Passed",
     "Start Time": "2026-05-10 12:00:00", "End Time": "2026-05-10 21:00:00"},
    {"Employee ID": 999, "Approval Status": "Passed",   # no match in master
     "Start Time": "2026-05-11 12:00:00", "End Time": "2026-05-11 21:00:00"},
])

p6.load_leave_data(buf6, "test_leave.xlsx")
check("matched record loaded", any(r["crm"] == "EMP600" for r in p6.leave_records))
check("unmatched ID not loaded", not any(r["crm"] == "EMP999" for r in p6.leave_records))
check("warning logged for unmatched ID",
      any("no matching employee" in l["message"].lower() for l in p6.logs))


# ── TEST 7: Integration with fill_leave_records ───────────────────────────────
section("TEST 7: Leave records applied correctly by fill_leave_records")

p7 = make_processor_with_employees([{"ac_no": 700, "crm": "EMP700", "name": "Test User"}])

# Simulate attendance: EMP700 was Absent on 2026-05-10 (a Sunday - workday)
monday = datetime(2026, 5, 11)  # Monday
p7.processed_data = [
    {
        "ac_no": "700", "crm": "EMP700", "name": "Test User",
        "department": "", "position": "",
        "date": monday, "day": "Monday",
        "clock_in": None, "clock_out": None,
        "in_status": "Absent", "out_status": "Absent",
        "status": "Absent",
    }
]

buf7 = make_italent_excel([
    {"Employee ID": 700, "Approval Status": "Passed",
     "Start Time": "2026-05-11 12:00:00", "End Time": "2026-05-11 21:00:00",
     "Leave Project": "Annual Leave"},
])
p7.load_leave_data(buf7, "test_leave.xlsx")
check("leave record loaded for EMP700", len(p7.leave_records) == 1,
      f"got {len(p7.leave_records)}")

p7.fill_leave_records()
record = next((r for r in p7.processed_data
               if r["crm"] == "EMP700" and r["date"].date() == monday.date()), None)

check("Absent record updated with leave type",
      record is not None and record["status"] == "Annual Leave",
      f"status={record['status'] if record else 'not found'}")
check("in_status set to On Leave",
      record is not None and record["in_status"] == "On Leave",
      f"in_status={record['in_status'] if record else 'not found'}")


# ── TEST 8: Off-days filtered by fill_leave_records ──────────────────────────
section("TEST 8: Leave on Friday (off-day) is ignored by fill_leave_records")

p8 = make_processor_with_employees([{"ac_no": 800, "crm": "EMP800"}])
friday = datetime(2026, 5, 15)  # Friday = weekday 4

buf8 = make_italent_excel([
    {"Employee ID": 800, "Approval Status": "Passed",
     "Start Time": "2026-05-15 12:00:00", "End Time": "2026-05-15 21:00:00"},
])
p8.load_leave_data(buf8, "test_leave.xlsx")

# Friday leave is loaded (expansion doesn't filter off days)
check("leave record loaded for Friday date", len(p8.leave_records) == 1)

# Simulate attendance has a Weekend record for that Friday
p8.processed_data = [
    {
        "ac_no": "800", "crm": "EMP800", "name": "Test",
        "department": "", "position": "",
        "date": friday, "day": "Friday",
        "clock_in": None, "clock_out": None,
        "in_status": "Weekend", "out_status": "Weekend",
        "status": "Weekend",
    }
]

p8.fill_leave_records()
record8 = next((r for r in p8.processed_data
                if r["crm"] == "EMP800"), None)
check("Friday status remains Weekend (leave not applied on off-day)",
      record8 is not None and record8["status"] == "Weekend",
      f"status={record8['status'] if record8 else 'not found'}")


# ── FINAL RESULTS ─────────────────────────────────────────────────────────────
print(f"\n{'='*60}")
print("  FINAL RESULTS")
print(f"{'='*60}\n")
print(f"  Passed : {PASS}/{PASS+FAIL}")
print(f"  Failed : {FAIL}/{PASS+FAIL}")
print()
if FAIL == 0:
    print("  ALL TESTS PASSED")
else:
    print("  SOME TESTS FAILED — review output above")
print()
sys.exit(0 if FAIL == 0 else 1)
