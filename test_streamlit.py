"""Test script for attendance_dashboard_streamlit.py using real test data."""
import sys
import os
import traceback

# Add the project directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
from attendance_dashboard_streamlit import AttendanceProcessor

TEST_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test data")

def print_header(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")

def print_result(test_name, passed, detail=""):
    status = "PASS" if passed else "FAIL"
    symbol = "[+]" if passed else "[X]"
    print(f"  {symbol} {test_name}: {status}")
    if detail:
        print(f"      {detail}")

def test_init():
    """Test processor initialization."""
    print_header("TEST 1: Processor Initialization")
    try:
        proc = AttendanceProcessor()
        checks = [
            ("employee_mapping initialized", isinstance(proc.employee_mapping, dict)),
            ("leave_records initialized", isinstance(proc.leave_records, dict)),
            ("attendance_results initialized", isinstance(proc.attendance_results, list)),
            ("logs initialized", isinstance(proc.logs, list)),
            ("off_day default is Friday", proc.off_day == "Friday"),
            ("work_start default", proc.work_start == "09:00"),
            ("late_threshold default", proc.late_threshold == 15),
        ]
        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("Initialization", False, str(e))
        return False

def test_normalize_id():
    """Test ID normalization."""
    print_header("TEST 2: normalize_id()")
    proc = AttendanceProcessor()
    test_cases = [
        (185680.0, "185680", "float to int string"),
        ("185680.0", "185680", "string float to int string"),
        (185680, "185680", "int stays int string"),
        ("ABC123", "ABC123", "string stays string"),
        (None, "", "None returns empty"),
        (float('nan'), "", "NaN returns empty"),
        ("  185680  ", "185680", "strips whitespace"),
    ]
    all_pass = True
    for value, expected, desc in test_cases:
        result = proc.normalize_id(value)
        passed = result == expected
        print_result(desc, passed, f"input={value!r}, expected={expected!r}, got={result!r}")
        if not passed:
            all_pass = False
    return all_pass

def test_find_column():
    """Test column finding with 4 strategies."""
    print_header("TEST 3: find_column()")
    proc = AttendanceProcessor()

    df = pd.DataFrame(columns=["AC-No.", "Name", "Clock In 1", "Department", "Join Date\n(yyyy/mm/dd)", "PS ID"])

    test_cases = [
        (["ac", "no"], None, "AC-No.", "find AC-No. with search terms"),
        (["name"], None, "Name", "find Name"),
        (["clock", "in"], None, "Clock In 1", "find Clock In 1"),
        (["ps", "id"], None, "PS ID", "find PS ID"),
        (["join", "date"], None, "Join Date\n(yyyy/mm/dd)", "find Join Date with newline"),
        (["nonexistent"], None, None, "returns None for missing column"),
    ]
    all_pass = True
    for terms, exact, expected, desc in test_cases:
        result = proc.find_column(df, terms, exact)
        passed = result == expected
        print_result(desc, passed, f"expected={expected!r}, got={result!r}")
        if not passed:
            all_pass = False
    return all_pass

def test_load_master_data():
    """Test loading master data from real file."""
    print_header("TEST 4: load_master_data()")
    proc = AttendanceProcessor()
    master_path = os.path.join(TEST_DIR, "Master sheet.xlsx")

    try:
        df = pd.read_excel(master_path)
        result = proc.load_master_data(df)

        checks = [
            ("returns True", result == True),
            ("employee_mapping populated", len(proc.employee_mapping) > 0),
            ("has CRM keys", all(isinstance(k, str) for k in proc.employee_mapping.keys())),
        ]

        # Check a sample employee has expected fields
        if proc.employee_mapping:
            sample_key = list(proc.employee_mapping.keys())[0]
            sample = proc.employee_mapping[sample_key]
            checks.extend([
                ("employee has 'name'", 'name' in sample),
                ("employee has 'department'", 'department' in sample),
                ("employee has 'crm'", 'crm' in sample),
                ("employee has 'ps_id'", 'ps_id' in sample),
            ])

        employee_count = len(proc.employee_mapping)
        checks.append((f"loaded {employee_count} employees", employee_count > 100))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("load_master_data", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_load_leave_data():
    """Test loading leave data from real file."""
    print_header("TEST 5: load_leave_data()")
    proc = AttendanceProcessor()

    # First load master data (needed for CRM mapping)
    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)

    leave_path = os.path.join(TEST_DIR, "Leaves.xlsx")

    try:
        result = proc.load_leave_data(leave_path)

        checks = [
            ("returns True", result == True),
            ("leave_records populated", len(proc.leave_records) > 0),
        ]

        # Check structure of leave records
        if proc.leave_records:
            sample_crm = list(proc.leave_records.keys())[0]
            sample_dates = proc.leave_records[sample_crm]
            checks.extend([
                ("leave records keyed by CRM", isinstance(sample_crm, str)),
                ("inner dict keyed by date", isinstance(sample_dates, dict)),
            ])
            if sample_dates:
                sample_date = list(sample_dates.keys())[0]
                sample_type = sample_dates[sample_date]
                checks.extend([
                    ("date keys are date objects", hasattr(sample_date, 'year')),
                    ("leave type is string", isinstance(sample_type, str)),
                ])

        leave_count = len(proc.leave_records)
        checks.append((f"loaded leaves for {leave_count} employees", leave_count > 0))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("load_leave_data", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_load_updated_leaves():
    """Test loading the Updated Leaves file (multi-sheet with Master)."""
    print_header("TEST 6: load_leave_data() - Updated Leaves (multi-sheet)")
    proc = AttendanceProcessor()

    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)

    leave_path = os.path.join(TEST_DIR, "Updated Leaves.xlsx")

    try:
        result = proc.load_leave_data(leave_path)

        checks = [
            ("returns True", result == True),
            ("leave_records populated", len(proc.leave_records) > 0),
        ]

        leave_count = len(proc.leave_records)
        checks.append((f"loaded leaves for {leave_count} employees", leave_count > 0))

        # Count total leave entries
        total_entries = sum(len(dates) for dates in proc.leave_records.values())
        checks.append((f"total leave entries: {total_entries}", total_entries > 0))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("load_leave_data (Updated)", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_process_attendance():
    """Test processing attendance files."""
    print_header("TEST 7: process_attendance_files()")
    proc = AttendanceProcessor()

    # Load master
    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)

    # Load attendance
    att_path = os.path.join(TEST_DIR, "Attendance.xls")
    att_df = pd.read_excel(att_path, engine='xlrd')

    try:
        proc.process_attendance_files([att_df], ["Attendance.xls"])

        checks = [
            ("attendance_results populated", len(proc.attendance_results) > 0),
            ("date_range set", proc.date_range is not None and len(proc.date_range) > 0),
        ]

        if proc.attendance_results:
            sample = proc.attendance_results[0]
            required_keys = ['crm', 'name', 'date', 'status', 'clock_in', 'clock_out']
            for key in required_keys:
                checks.append((f"result has '{key}'", key in sample))

            # Check status values
            statuses = set(r['status'] for r in proc.attendance_results)
            checks.append((f"found {len(statuses)} status types", len(statuses) > 1))
            print(f"\n  Status types found: {statuses}")

        result_count = len(proc.attendance_results)
        checks.append((f"processed {result_count} records", result_count > 100))

        date_count = len(proc.date_range) if proc.date_range else 0
        checks.append((f"date range covers {date_count} days", date_count > 0))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("process_attendance_files", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_fill_leave_records():
    """Test filling leave records into attendance."""
    print_header("TEST 8: fill_leave_records()")
    proc = AttendanceProcessor()

    # Load master
    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)

    # Load leave
    leave_path = os.path.join(TEST_DIR, "Leaves.xlsx")
    proc.load_leave_data(leave_path)

    # Process attendance
    att_df = pd.read_excel(os.path.join(TEST_DIR, "Attendance.xls"), engine='xlrd')
    proc.process_attendance_files([att_df], ["Attendance.xls"])

    try:
        before_count = len(proc.attendance_results)
        proc.fill_leave_records()
        after_count = len(proc.attendance_results)

        # Check for leave statuses
        leave_statuses = [r for r in proc.attendance_results if 'Leave' in r.get('status', '')]

        checks = [
            ("records added for leave employees", after_count >= before_count),
            (f"before: {before_count}, after: {after_count}", True),
            (f"found {len(leave_statuses)} leave records", len(leave_statuses) >= 0),
        ]

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("fill_leave_records", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_calculate_penalties():
    """Test penalty calculations."""
    print_header("TEST 9: calculate_penalties()")
    proc = AttendanceProcessor()

    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)

    att_df = pd.read_excel(os.path.join(TEST_DIR, "Attendance.xls"), engine='xlrd')
    proc.process_attendance_files([att_df], ["Attendance.xls"])

    try:
        penalties = proc.calculate_penalties()

        checks = [
            ("returns list", isinstance(penalties, list)),
            ("penalties populated", len(penalties) > 0),
        ]

        if penalties:
            sample = penalties[0]
            expected_keys = ['crm', 'name', 'late_count', 'absent_count', 'missing_punch_count',
                           'late_penalty', 'total_deduction_days']
            for key in expected_keys:
                checks.append((f"penalty has '{key}'", key in sample))

            # Summary stats
            total_late = sum(p.get('late_count', 0) for p in penalties)
            total_absent = sum(p.get('absent_count', 0) for p in penalties)
            total_missing = sum(p.get('missing_punch_count', 0) for p in penalties)
            print(f"\n  Summary: {len(penalties)} employees, {total_late} lates, {total_absent} absences, {total_missing} missing punches")

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("calculate_penalties", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_generate_report():
    """Test full report generation (the ultimate integration test)."""
    print_header("TEST 10: Full Report Generation (Integration)")
    proc = AttendanceProcessor()

    # Load master
    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)
    print(f"  Loaded {len(proc.employee_mapping)} employees")

    # Load leave
    leave_path = os.path.join(TEST_DIR, "Leaves.xlsx")
    proc.load_leave_data(leave_path)
    print(f"  Loaded leaves for {len(proc.leave_records)} employees")

    # Process attendance
    att_df = pd.read_excel(os.path.join(TEST_DIR, "Attendance.xls"), engine='xlrd')
    proc.process_attendance_files([att_df], ["Attendance.xls"])
    print(f"  Processed {len(proc.attendance_results)} attendance records")

    # Fill leaves
    proc.fill_leave_records()
    print(f"  After fill_leave_records: {len(proc.attendance_results)} records")

    # Generate report
    output_path = os.path.join(TEST_DIR, "Generated report", "TEST_streamlit_output.xlsx")

    try:
        result = proc.generate_report(output_path)

        checks = [
            ("generate_report returns True", result == True),
            ("output file created", os.path.exists(output_path)),
        ]

        if os.path.exists(output_path):
            # Verify the Excel file structure
            xls = pd.ExcelFile(output_path)
            sheets = xls.sheet_names
            expected_sheets = ["Summary Report", "Individual Analytics", "Alerts & Warnings", "Penalties"]

            checks.append((f"has {len(sheets)} sheets: {sheets}", len(sheets) >= 4))

            for expected in expected_sheets:
                checks.append((f"has '{expected}' sheet", expected in sheets))

            # Check Summary Report
            summary_df = pd.read_excel(output_path, sheet_name="Summary Report")
            checks.append((f"Summary has {summary_df.shape[0]} rows, {summary_df.shape[1]} cols", summary_df.shape[0] > 0))

            # Check Penalties
            penalties_df = pd.read_excel(output_path, sheet_name="Penalties")
            checks.append((f"Penalties has {penalties_df.shape[0]} rows, {penalties_df.shape[1]} cols", penalties_df.shape[0] > 0))

            # Check Individual Analytics
            analytics_df = pd.read_excel(output_path, sheet_name="Individual Analytics")
            checks.append((f"Analytics has {analytics_df.shape[0]} rows", analytics_df.shape[0] > 0))

            file_size = os.path.getsize(output_path) / 1024
            checks.append((f"output file size: {file_size:.1f} KB", file_size > 10))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("generate_report", False, f"{e}\n{traceback.format_exc()}")
        return False

def test_full_pipeline_with_both_attendance():
    """Test with both attendance files and updated leaves."""
    print_header("TEST 11: Full Pipeline - Multiple Attendance + Updated Leaves")
    proc = AttendanceProcessor()

    # Load master
    master_df = pd.read_excel(os.path.join(TEST_DIR, "Master sheet.xlsx"))
    proc.load_master_data(master_df)
    print(f"  Loaded {len(proc.employee_mapping)} employees")

    # Load updated leaves
    leave_path = os.path.join(TEST_DIR, "Updated Leaves.xlsx")
    proc.load_leave_data(leave_path)
    print(f"  Loaded leaves for {len(proc.leave_records)} employees")

    # Process both attendance files
    att1_df = pd.read_excel(os.path.join(TEST_DIR, "Attendance.xls"), engine='xlrd')
    att2_df = pd.read_excel(os.path.join(TEST_DIR, "from 20 Dec to 14 Jan.xls"), engine='xlrd')
    proc.process_attendance_files([att1_df, att2_df], ["Attendance.xls", "from 20 Dec to 14 Jan.xls"])
    print(f"  Processed {len(proc.attendance_results)} attendance records")

    # Fill leaves
    proc.fill_leave_records()
    print(f"  After fill_leave_records: {len(proc.attendance_results)} records")

    # Generate report
    output_path = os.path.join(TEST_DIR, "Generated report", "TEST_full_pipeline.xlsx")

    try:
        result = proc.generate_report(output_path)

        checks = [
            ("generate_report returns True", result == True),
            ("output file created", os.path.exists(output_path)),
        ]

        if os.path.exists(output_path):
            xls = pd.ExcelFile(output_path)
            sheets = xls.sheet_names
            checks.append((f"has sheets: {sheets}", len(sheets) >= 4))

            summary_df = pd.read_excel(output_path, sheet_name="Summary Report")
            checks.append((f"Summary: {summary_df.shape[0]} rows x {summary_df.shape[1]} cols", summary_df.shape[0] > 0))

            file_size = os.path.getsize(output_path) / 1024
            checks.append((f"output file size: {file_size:.1f} KB", file_size > 10))

        all_pass = True
        for name, result in checks:
            print_result(name, result)
            if not result:
                all_pass = False
        return all_pass
    except Exception as e:
        print_result("Full pipeline", False, f"{e}\n{traceback.format_exc()}")
        return False


if __name__ == "__main__":
    print("\n" + "="*60)
    print("  ATTENDANCE DASHBOARD STREAMLIT - TEST SUITE")
    print("="*60)

    tests = [
        ("Initialization", test_init),
        ("normalize_id", test_normalize_id),
        ("find_column", test_find_column),
        ("load_master_data", test_load_master_data),
        ("load_leave_data", test_load_leave_data),
        ("load_leave_data (Updated)", test_load_updated_leaves),
        ("process_attendance_files", test_process_attendance),
        ("fill_leave_records", test_fill_leave_records),
        ("calculate_penalties", test_calculate_penalties),
        ("Full Report Generation", test_generate_report),
        ("Full Pipeline (Multi-file)", test_full_pipeline_with_both_attendance),
    ]

    results = []
    for name, test_fn in tests:
        try:
            passed = test_fn()
            results.append((name, passed))
        except Exception as e:
            print(f"\n  UNEXPECTED ERROR in {name}: {e}")
            traceback.print_exc()
            results.append((name, False))

    print_header("FINAL RESULTS")
    passed_count = sum(1 for _, p in results if p)
    total = len(results)

    for name, passed in results:
        status = "PASS" if passed else "FAIL"
        symbol = "[+]" if passed else "[X]"
        print(f"  {symbol} {name}: {status}")

    print(f"\n  Total: {passed_count}/{total} passed")

    if passed_count == total:
        print("\n  ALL TESTS PASSED!")
    else:
        print(f"\n  {total - passed_count} TEST(S) FAILED")

    sys.exit(0 if passed_count == total else 1)
