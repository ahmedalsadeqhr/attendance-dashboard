"""
Attendance Dashboard Streamlit Web Application v2.2
Web-based attendance processing with Streamlit interface
Based on the desktop CTk application
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime, timedelta
import io
import re
import logging


def read_excel_file(file_obj, filename):
    """Read Excel file with automatic engine detection for .xls and .xlsx files."""
    filename_lower = filename.lower()
    if filename_lower.endswith('.xls') and not filename_lower.endswith('.xlsx'):
        return pd.read_excel(file_obj, engine='xlrd')
    else:
        return pd.read_excel(file_obj)


# Default configuration
DEFAULT_CONFIG = {
    'work_start_time': '12:00 PM',
    'work_end_time': '9:00 PM',
    'late_threshold': '12:00 PM',
    'off_days': [4],  # Friday
    'penalties': {
        'late_1st': 100,
        'late_2nd': 200,
        'late_3rd': 500,
        'late_4th_plus': 500,
        'missing_punch_deduction': 0.5,
        'missing_punch_threshold': 3,
        'missing_punch_warning_threshold': 6,
        'early_departure_deduction': 0.5,
        'absence_deduction': 2,
        'currency': 'EGP'
    }
}


class AttendanceProcessor:
    """Core attendance processing logic (shared with desktop version)"""

    def __init__(self, config=None):
        self.config = config or DEFAULT_CONFIG
        self.employee_mapping = {}
        self.leave_records = []
        self.processed_data = []
        self.conflict_records = []
        self.penalties_data = {}
        self.logs = []

    def log_message(self, message, level='info'):
        """Add log message"""
        self.logs.append({'level': level, 'message': message})

    def normalize_id(self, value):
        """Normalize ID values for consistent matching"""
        if pd.isna(value):
            return ""
        value_str = str(value).strip()
        if value_str.replace('.', '').replace('-', '').isdigit():
            try:
                num = float(value_str)
                if num == int(num):
                    return str(int(num))
            except ValueError:
                pass
        return value_str

    def find_column(self, df, search_terms, exact_matches=None):
        """Find column by search terms or exact matches"""
        columns = df.columns.tolist()

        # First try exact matches
        if exact_matches:
            for exact in exact_matches:
                if exact in columns:
                    return exact
                # Try case-insensitive
                for col in columns:
                    if str(col).lower().strip() == exact.lower().strip():
                        return col

        # Then try partial matches
        for col in columns:
            col_lower = str(col).lower()
            if any(term in col_lower for term in search_terms):
                return col

        return None

    def load_master_data(self, file_obj, filename):
        """Load and process master data file"""
        try:
            self.log_message("Loading master data...")
            df = read_excel_file(file_obj, filename)

            # Find columns
            ac_col = self.find_column(df, ['ac', 'no'], ['AC-No.', 'Ac-No.', 'AC No', 'PS ID', 'PS Id', 'PSID'])
            if not ac_col:
                ac_col = self.find_column(df, ['ps', 'id'], ['PS ID', 'PS Id', 'PSID'])
            crm_col = self.find_column(df, ['crm'], ['CRM'])
            name_col = self.find_column(df, ['name'], ['Name'])
            dept_col = self.find_column(df, ['department', 'dept'], ['Department', 'Dept'])
            pos_col = self.find_column(df, ['position', 'title'], ['Position', 'Title'])
            national_id_col = self.find_column(df, ['identity', 'idnetity', 'national', 'nid'],
                                               ['Idnetity Number', 'Identity Number', 'National ID', 'NID'])
            vendor_col = self.find_column(df, ['vendor'], ['Vendor'])
            ps_id_col = self.find_column(df, ['ps', 'id'], ['PS ID', 'PS Id', 'PSID', 'PS-ID'])

            # Join Date column
            join_date_col = None
            for col in df.columns:
                col_clean = str(col).replace('\n', ' ').lower().strip()
                if 'join date' in col_clean or 'joining date' in col_clean:
                    join_date_col = col
                    break
            if not join_date_col:
                join_date_col = self.find_column(df, ['join', 'joining', 'hired'],
                                                 ['Join Date', 'Joining Date', 'Date of Joining', 'JoinDate',
                                                  'Hire Date', 'HireDate', 'Date Joined', 'DOJ'])

            # Exit Date column - for resigned employees
            exit_date_col = None
            for col in df.columns:
                col_clean = str(col).replace('\n', ' ').lower().strip()
                if any(term in col_clean for term in ['exit date', 'resign', 'end date', 'termination', 'leaving']):
                    exit_date_col = col
                    break
            if not exit_date_col:
                exit_date_col = self.find_column(df, ['exit', 'resign', 'end', 'termination', 'leaving'],
                                                 ['Exit Date', 'Resignation Date', 'End Date', 'Termination Date',
                                                  'Leaving Date', 'Last Day', 'Last Working Day'])

            # Build employee mapping
            self.employee_mapping = {}
            for idx, row in df.iterrows():
                ac_no = self.normalize_id(row[ac_col]) if ac_col else ""
                if ac_no and ac_no != 'nan':
                    # Handle join date
                    join_date_val = self._parse_date_field(row, join_date_col)
                    # Handle exit date
                    exit_date_val = self._parse_date_field(row, exit_date_col)

                    self.employee_mapping[ac_no] = {
                        'crm': str(row[crm_col]).strip() if crm_col and pd.notna(row[crm_col]) else "",
                        'name': str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else "",
                        'department': str(row[dept_col]).strip() if dept_col and pd.notna(row[dept_col]) else "",
                        'position': str(row[pos_col]).strip() if pos_col and pd.notna(row[pos_col]) else "",
                        'national_id': str(row[national_id_col]).strip() if national_id_col and pd.notna(row[national_id_col]) else "",
                        'vendor': str(row[vendor_col]).strip() if vendor_col and pd.notna(row[vendor_col]) else "",
                        'ps_id': self.normalize_id(row[ps_id_col]) if ps_id_col and pd.notna(row[ps_id_col]) else "",
                        'join_date': join_date_val,
                        'exit_date': exit_date_val
                    }

            self.log_message(f"Loaded {len(self.employee_mapping)} employee records", 'success')
            return True

        except Exception as e:
            self.log_message(f"Error loading master data: {str(e)}", 'error')
            return False

    def _parse_date_field(self, row, date_col):
        """Parse a date field from a row"""
        if not date_col or pd.isna(row[date_col]):
            return ""
        try:
            raw_val = row[date_col]
            if isinstance(raw_val, (datetime, pd.Timestamp)):
                return raw_val.strftime('%Y-%m-%d')
            elif hasattr(raw_val, '__int__') and not isinstance(raw_val, str):
                excel_epoch = datetime(1899, 12, 30)
                dt = excel_epoch + timedelta(days=int(raw_val))
                return dt.strftime('%Y-%m-%d')
            else:
                dt = pd.to_datetime(raw_val, dayfirst=True)
                return dt.strftime('%Y-%m-%d')
        except Exception:
            return str(row[date_col]).strip()

    def load_leave_data(self, file_obj, filename):
        """Load and process leave data file"""
        try:
            self.log_message("Loading leave data...")
            self.leave_records = []

            xlsx = pd.ExcelFile(file_obj)
            sheet_names = xlsx.sheet_names
            month_sheets = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

            has_month_sheets = any(month in sheet_names for month in month_sheets)

            if has_month_sheets:
                for sheet_name in sheet_names:
                    if sheet_name in month_sheets:
                        df = pd.read_excel(xlsx, sheet_name=sheet_name)
                        self._process_leave_sheet(df, sheet_name)
            else:
                df = pd.read_excel(xlsx)
                self._process_leave_sheet(df)

            self.log_message(f"Loaded {len(self.leave_records)} leave records", 'success')
            return True

        except Exception as e:
            self.log_message(f"Error loading leave data: {str(e)}", 'error')
            return False

    def _process_leave_sheet(self, df, sheet_name=None):
        """Process a single leave sheet"""
        crm_col = self.find_column(df, ['crm'], ['CRM'])
        date_col = self.find_column(df, ['date', 'leave'], ['Date', 'Leave Date'])
        type_col = self.find_column(df, ['type', 'leave'], ['Type', 'Leave Type'])

        if not crm_col or not date_col:
            return

        for idx, row in df.iterrows():
            crm = str(row[crm_col]).strip() if pd.notna(row[crm_col]) else ""
            if not crm:
                continue

            try:
                leave_date = pd.to_datetime(row[date_col])
                leave_type = str(row[type_col]).strip() if type_col and pd.notna(row[type_col]) else "Leave"

                self.leave_records.append({
                    'crm': crm,
                    'date': leave_date,
                    'leave_type': leave_type
                })
            except Exception:
                continue

    def get_filtered_employee_mapping(self, selected_depts=None, selected_crms=None):
        """Return employee mapping filtered by selected departments and CRMs"""
        if not selected_depts and not selected_crms:
            return self.employee_mapping

        all_depts = set(info['department'] for info in self.employee_mapping.values() if info['department'])
        all_crms = set(info['crm'] for info in self.employee_mapping.values() if info['crm'])

        selected_depts = set(selected_depts) if selected_depts else all_depts
        selected_crms = set(selected_crms) if selected_crms else all_crms

        filtered = {}
        for ac_no, info in self.employee_mapping.items():
            dept = info.get('department', '')
            crm = info.get('crm', '')

            dept_match = not dept or dept in selected_depts
            crm_match = not crm or crm in selected_crms

            if dept_match and crm_match:
                filtered[ac_no] = info

        return filtered

    def extract_date_from_filename(self, filename):
        """Extract date from attendance filename"""
        patterns = [
            r'(\d{4})[-_](\d{2})[-_](\d{2})',
            r'(\d{2})[-_](\d{2})[-_](\d{4})',
            r'(\d{8})',
        ]

        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                groups = match.groups()
                try:
                    if len(groups) == 1:
                        return datetime.strptime(groups[0], '%Y%m%d')
                    elif len(groups[0]) == 4:
                        return datetime(int(groups[0]), int(groups[1]), int(groups[2]))
                    else:
                        return datetime(int(groups[2]), int(groups[1]), int(groups[0]))
                except ValueError:
                    continue

        return datetime.now()

    def get_first_clock_in(self, row):
        """Get first clock in time from row"""
        for i in range(1, 6):
            col_name = f'Clock In {i}'
            if col_name in row.index and pd.notna(row[col_name]):
                return row[col_name]
        return None

    def get_last_clock_out(self, row):
        """Get last clock out time from row"""
        for i in range(5, 0, -1):
            col_name = f'Clock Out {i}'
            if col_name in row.index and pd.notna(row[col_name]):
                return row[col_name]
        return None

    def determine_status(self, clock_in, clock_out, date, crm):
        """Determine attendance status"""
        day_of_week = date.weekday()

        # Check if employee has resigned
        for ac_no, info in self.employee_mapping.items():
            if info.get('crm') == crm:
                exit_date_str = info.get('exit_date', '')
                if exit_date_str:
                    try:
                        exit_date = datetime.strptime(exit_date_str, '%Y-%m-%d').date()
                        if date.date() > exit_date:
                            return "Resigned", "Resigned", "Resigned"
                    except Exception:
                        pass
                break

        # Check if it's an OFF day
        off_days = self.config.get('off_days', [4])
        if day_of_week in off_days:
            if not clock_in and not clock_out:
                return "Weekend", "Weekend", "Weekend"
            else:
                return "Worked on Day Off", "Worked", "Worked"

        # Check for leave
        for leave in self.leave_records:
            if leave['crm'] == crm and leave['date'].date() == date.date():
                return leave['leave_type'], "On Leave", "On Leave"

        # Check for absent
        if not clock_in and not clock_out:
            return "Absent", "No Clock In", "No Clock Out"

        # Check for missing punch
        if not clock_in:
            return "Missing Punch In", "No Clock In", "..."

        if not clock_out:
            try:
                if isinstance(clock_in, str):
                    clock_in_time = datetime.strptime(clock_in, "%I:%M:%S %p").time()
                else:
                    clock_in_time = clock_in

                late_threshold_str = self.config.get('late_threshold', '12:00 PM')
                late_threshold = datetime.strptime(late_threshold_str, "%I:%M %p").time()
                in_status = "Late" if clock_in_time > late_threshold else "On Time"
            except Exception:
                in_status = "..."

            return "Missing Punch Out", in_status, "No Clock Out"

        # Determine if late
        try:
            if isinstance(clock_in, str):
                clock_in_time = datetime.strptime(clock_in, "%I:%M:%S %p").time()
            else:
                clock_in_time = clock_in

            late_threshold_str = self.config.get('late_threshold', '12:00 PM')
            late_threshold = datetime.strptime(late_threshold_str, "%I:%M %p").time()
            is_late = clock_in_time > late_threshold

            if is_late:
                return "Late", "Late", "On Time"
            else:
                return "Normal", "On Time", "On Time"
        except Exception:
            return "Normal", "On Time", "On Time"

    def process_attendance_files(self, files, selected_depts=None, selected_crms=None):
        """Process all attendance files"""
        self.processed_data = []
        filtered_mapping = self.get_filtered_employee_mapping(selected_depts, selected_crms)

        for file_obj, filename in files:
            try:
                self.log_message(f"Processing: {filename}")
                file_obj.seek(0)
                df = read_excel_file(file_obj, filename)
                file_date = self.extract_date_from_filename(filename)

                # Find columns
                ac_col = self.find_column(df, ['ac', 'no'], ['AC-No.', 'Ac-No.', 'AC No'])
                name_col = self.find_column(df, ['name'], ['Name'])
                date_col = self.find_column(df, ['date'], ['Date'])

                for idx, row in df.iterrows():
                    ac_no = self.normalize_id(row[ac_col]) if ac_col and pd.notna(row[ac_col]) else ""
                    if not ac_no or ac_no == 'nan':
                        continue

                    record_date = file_date
                    if date_col and pd.notna(row[date_col]):
                        try:
                            record_date = pd.to_datetime(row[date_col])
                            if hasattr(record_date, 'to_pydatetime'):
                                record_date = record_date.to_pydatetime()
                        except Exception:
                            record_date = file_date

                    clock_in = self.get_first_clock_in(row)
                    clock_out = self.get_last_clock_out(row)

                    emp_info = filtered_mapping.get(ac_no)
                    if not emp_info:
                        # Try name match
                        if name_col and pd.notna(row[name_col]):
                            row_name = str(row[name_col]).strip().lower()
                            for key, info in filtered_mapping.items():
                                if info['name'].lower() == row_name:
                                    emp_info = info
                                    break

                    if emp_info:
                        crm = emp_info['crm']
                        name = emp_info['name']
                        dept = emp_info['department']
                        position = emp_info['position']
                    else:
                        if selected_depts or selected_crms:
                            continue
                        name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else "Unknown"
                        crm = name if name != "Unknown" else f"ID-{ac_no}"
                        dept = ""
                        position = ""

                    status, in_status, out_status = self.determine_status(clock_in, clock_out, record_date, crm)

                    self.processed_data.append({
                        'ac_no': ac_no,
                        'crm': crm,
                        'name': name,
                        'department': dept,
                        'position': position,
                        'date': record_date,
                        'day': record_date.strftime('%A'),
                        'clock_in': clock_in,
                        'clock_out': clock_out,
                        'in_status': in_status,
                        'out_status': out_status,
                        'status': status
                    })

            except Exception as e:
                self.log_message(f"Error processing {filename}: {str(e)}", 'error')

        self.log_message(f"Total records processed: {len(self.processed_data)}", 'success')

    def fill_leave_records(self, selected_depts=None, selected_crms=None):
        """Apply leave records to attendance data"""
        self.log_message("Applying leave records...")
        off_days = self.config.get('off_days', [4])

        if not self.processed_data:
            return

        all_dates = [record['date'] for record in self.processed_data]
        min_date = min(all_dates).date()
        max_date = max(all_dates).date()

        filtered_mapping = self.get_filtered_employee_mapping(selected_depts, selected_crms)
        valid_crms = set(info['crm'] for info in filtered_mapping.values())

        record_index = {}
        for idx, record in enumerate(self.processed_data):
            key = (record['crm'], record['date'].date())
            record_index[key] = idx

        crm_to_emp_info = {}
        for ac_no, info in filtered_mapping.items():
            crm_to_emp_info[info['crm']] = info

        self.conflict_records = []
        updated_count = 0
        added_count = 0

        leave_lookup = {}
        for leave in self.leave_records:
            crm = leave['crm']
            leave_date = leave['date'].date()

            if crm not in valid_crms:
                continue
            if leave_date < min_date or leave_date > max_date:
                continue
            if leave['date'].weekday() in off_days:
                continue

            key = (crm, leave_date)
            leave_lookup[key] = leave

        for key, leave in leave_lookup.items():
            crm, leave_date = key

            if key in record_index:
                idx = record_index[key]
                existing_record = self.processed_data[idx]
                existing_status = existing_record['status']

                if existing_status not in ['Weekend', 'Resigned', 'Absent']:
                    self.conflict_records.append({
                        'crm': crm,
                        'name': existing_record.get('name', ''),
                        'date': leave['date'],
                        'attendance_status': existing_status,
                        'leave_type': leave['leave_type'],
                        'clock_in': existing_record.get('clock_in'),
                        'clock_out': existing_record.get('clock_out')
                    })

                if existing_status in ['Absent', 'Missing Punch In', 'Missing Punch Out', 'Late', 'Normal']:
                    self.processed_data[idx]['status'] = leave['leave_type']
                    self.processed_data[idx]['in_status'] = 'On Leave'
                    self.processed_data[idx]['out_status'] = 'On Leave'
                    updated_count += 1
            else:
                emp_info = crm_to_emp_info.get(crm)
                if emp_info:
                    self.processed_data.append({
                        'ac_no': '',
                        'crm': crm,
                        'name': emp_info['name'],
                        'department': emp_info['department'],
                        'position': emp_info['position'],
                        'date': leave['date'],
                        'day': leave['date'].strftime('%A'),
                        'clock_in': None,
                        'clock_out': None,
                        'in_status': 'On Leave',
                        'out_status': 'On Leave',
                        'status': leave['leave_type']
                    })
                    added_count += 1

        self.log_message(f"Applied {updated_count} leave updates, added {added_count} leave records")
        if self.conflict_records:
            self.log_message(f"Detected {len(self.conflict_records)} leave vs attendance conflicts", 'warning')

    def calculate_penalties(self):
        """Calculate penalties for all employees"""
        penalties_config = self.config.get('penalties', {})
        late_rates = [
            penalties_config.get('late_1st', 100),
            penalties_config.get('late_2nd', 200),
            penalties_config.get('late_3rd', 500),
            penalties_config.get('late_4th_plus', 500)
        ]
        missing_threshold = penalties_config.get('missing_punch_threshold', 3)
        missing_deduction = penalties_config.get('missing_punch_deduction', 0.5)
        absence_deduction = penalties_config.get('absence_deduction', 2)

        employee_stats = {}
        for record in self.processed_data:
            crm = record['crm']
            if crm not in employee_stats:
                for ac_no, info in self.employee_mapping.items():
                    if info['crm'] == crm:
                        employee_stats[crm] = {
                            'name': info['name'],
                            'department': info['department'],
                            'national_id': info.get('national_id', ''),
                            'vendor': info.get('vendor', ''),
                            'ps_id': info.get('ps_id', ''),
                            'join_date': info.get('join_date', ''),
                            'records': [],
                            'working_days': 0,
                            'total_days': 0
                        }
                        break
                else:
                    employee_stats[crm] = {
                        'name': record['name'],
                        'department': record.get('department', ''),
                        'national_id': '',
                        'vendor': '',
                        'ps_id': '',
                        'join_date': '',
                        'records': [],
                        'working_days': 0,
                        'total_days': 0
                    }

            employee_stats[crm]['records'].append(record)
            employee_stats[crm]['total_days'] += 1
            if record['status'] not in ['Weekend', 'Resigned']:
                employee_stats[crm]['working_days'] += 1

        penalties_summary = {}
        for crm, stats in employee_stats.items():
            late_count = sum(1 for r in stats['records'] if r['status'] == 'Late')
            missing_count = sum(1 for r in stats['records'] if 'Missing Punch' in r['status'])
            absence_count = sum(1 for r in stats['records'] if r['status'] == 'Absent')

            late_penalty = 0
            for i in range(late_count):
                if i < 3:
                    late_penalty += late_rates[i]
                else:
                    late_penalty += late_rates[3]

            missing_ded = max(0, (missing_count - missing_threshold) * missing_deduction)
            absence_ded = absence_count * absence_deduction
            total_penalty = late_penalty
            total_deduction = missing_ded + absence_ded
            total_warnings = 1 if late_count >= 3 else 0

            penalties_summary[crm] = {
                'name': stats['name'],
                'department': stats.get('department', ''),
                'national_id': stats.get('national_id', ''),
                'vendor': stats.get('vendor', ''),
                'ps_id': stats.get('ps_id', ''),
                'join_date': stats.get('join_date', ''),
                'late_count': late_count,
                'late_penalty': late_penalty,
                'missing_punch_count': missing_count,
                'missing_deduction': missing_ded,
                'absence_count': absence_count,
                'absence_deduction': absence_ded,
                'total_penalty_egp': total_penalty,
                'total_deduction_days': total_deduction,
                'total_warnings': total_warnings,
                'working_days': stats['working_days'],
                'total_days': stats['total_days']
            }

        self.penalties_data = penalties_summary
        return penalties_summary

    def create_excel_report(self):
        """Create the final Excel report and return as bytes"""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        self.calculate_penalties()
        self.create_summary_sheet(wb)
        self.create_analytics_sheet(wb)
        self.create_alerts_sheet(wb)
        self.create_penalties_sheet(wb)
        self.create_duplicates_sheet(wb)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()

    def _apply_status_color(self, cell, status):
        """Apply color coding to a cell based on status"""
        if '(BD)' in status:
            cell.fill = PatternFill(start_color='E1D5E7', end_color='E1D5E7', fill_type='solid')
        elif status in ['Normal', 'Present', 'Late (Approved)', 'Annual Leave', 'Casual Leave',
                        'Marriage Leave', 'Paternity Leave', 'Maternity Leave', 'Bereavement Leave',
                        'Military Call Leave', 'Early Departure (Approved)']:
            cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
        elif status == 'Late':
            cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
        elif status == 'Absent':
            cell.fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        elif 'Missing Punch' in status:
            cell.fill = PatternFill(start_color='FFD9E6', end_color='FFD9E6', fill_type='solid')
        elif status in ['Sick Leave', 'Unpaid Leave', 'Unpaid leave']:
            cell.fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
        elif status in ['Early Departure', 'Half Day', 'Early Leave (HD)', 'Early Leave (HD) (BD)']:
            cell.fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
        elif status == 'Weekend':
            cell.fill = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
        elif status == 'Resigned':
            cell.fill = PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid')
        elif status == 'Worked on Day Off':
            cell.fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        elif 'Leave' in status:
            cell.fill = PatternFill(start_color='E6F3FF', end_color='E6F3FF', fill_type='solid')

    def create_summary_sheet(self, wb):
        """Create the summary report sheet"""
        ws = wb.create_sheet("Summary Report", 0)

        crms = sorted(set(r['crm'] for r in self.processed_data))
        attendance_dates = [r['date'] for r in self.processed_data]
        if attendance_dates:
            min_date = min(attendance_dates)
            max_date = max(attendance_dates)
            dates = []
            current = min_date
            while current <= max_date:
                dates.append(current)
                current = current + timedelta(days=1)
        else:
            dates = sorted(set(r['date'] for r in self.processed_data))

        matrix = {crm: {date: '' for date in dates} for crm in crms}
        for record in self.processed_data:
            matrix[record['crm']][record['date']] = record['status']

        # Title
        ws.merge_cells('A1:' + get_column_letter(len(dates) + 3) + '1')
        title_cell = ws['A1']
        title_cell.value = "Attendance Summary Report"
        title_cell.font = Font(name='Calibri', size=16, bold=True, color='FFFFFF')
        title_cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 35

        # Headers
        row = 3
        ws.cell(row, 1, "CRM")
        ws.cell(row, 2, "Normal Days")
        ws.cell(row, 3, "Abnormal Days")

        for i, date in enumerate(dates, start=4):
            ws.cell(row, i, date.strftime("%d-%b"))

        for col in range(1, len(dates) + 4):
            cell = ws.cell(row, col)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        off_days = self.config.get('off_days', [4])

        # Justification options
        justification_options = [
            "Normal", "Late (Approved)", "Late", "Absent", "Missing Punch In",
            "Missing Punch In (Justified)", "Missing Punch Out", "Missing Punch Out (Justified)",
            "Early Departure (Approved)", "Early Departure", "Half Day", "Early Leave (HD)",
            "Sick Leave", "Annual Leave", "Casual Leave", "Marriage Leave", "Paternity Leave",
            "Maternity Leave", "Bereavement Leave", "Military Call Leave", "Unpaid Leave",
            "Weekend", "Resigned",
            "Annual Leave (BD)", "Casual Leave (BD)", "Sick Leave (BD)", "Unpaid Leave (BD)",
            "Half Day (BD)", "Early Leave (HD) (BD)", "Early Departure (BD)", "Marriage Leave (BD)",
            "Paternity Leave (BD)", "Maternity Leave (BD)", "Bereavement Leave (BD)"
        ]

        justification_list = ",".join(justification_options)
        dv = DataValidation(
            type="list",
            formula1=f'"{justification_list}"',
            allow_blank=True,
            showDropDown=False
        )
        ws.add_data_validation(dv)

        # Data rows
        data_row = 4
        for crm in crms:
            ws.cell(data_row, 1, crm)

            normal_count = sum(1 for d in dates if matrix[crm].get(d) == 'Normal')
            abnormal_count = sum(1 for d in dates
                                 if matrix[crm].get(d) and matrix[crm].get(d) not in ['Normal', 'Weekend', 'Resigned'])

            ws.cell(data_row, 2, normal_count)
            ws.cell(data_row, 3, abnormal_count)

            for i, date in enumerate(dates, start=4):
                cell = ws.cell(data_row, i)
                status = matrix[crm].get(date, '')

                if not status and date.weekday() in off_days:
                    status = 'Weekend'
                    matrix[crm][date] = status

                cell.value = status
                self._apply_status_color(cell, status)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )

                if status not in ['Weekend', 'Resigned', '']:
                    dv.add(cell)

            data_row += 1

        # Column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 14
        for i in range(4, len(dates) + 4):
            ws.column_dimensions[get_column_letter(i)].width = 10

    def create_analytics_sheet(self, wb):
        """Create individual analytics sheet"""
        ws = wb.create_sheet("Individual Analytics", 1)

        analytics = {}
        for crm in set(r['crm'] for r in self.processed_data):
            records = [r for r in self.processed_data if r['crm'] == crm]
            total = len(records)
            normal = sum(1 for r in records if r['status'] == 'Normal')
            late = sum(1 for r in records if r['status'] == 'Late')
            absent = sum(1 for r in records if r['status'] == 'Absent')
            missing = sum(1 for r in records if 'Missing Punch' in r['status'])

            analytics[crm] = {
                'name': records[0]['name'] if records else '',
                'department': records[0].get('department', '') if records else '',
                'total': total,
                'normal': normal,
                'late': late,
                'absent': absent,
                'missing': missing,
                'attendance_rate': (normal / total * 100) if total > 0 else 0
            }

        # Title
        ws.merge_cells('A1:H1')
        title = ws['A1']
        title.value = "Individual Analytics"
        title.font = Font(size=14, bold=True, color='FFFFFF')
        title.fill = PatternFill(start_color='2E7D32', end_color='2E7D32', fill_type='solid')
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 30

        # Headers
        headers = ['CRM', 'Name', 'Department', 'Total Days', 'Normal', 'Late', 'Absent', 'Attendance Rate']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(3, col, header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='4CAF50', end_color='4CAF50', fill_type='solid')
            cell.alignment = Alignment(horizontal='center')

        # Data
        row = 4
        for crm, data in sorted(analytics.items()):
            ws.cell(row, 1, crm)
            ws.cell(row, 2, data['name'])
            ws.cell(row, 3, data['department'])
            ws.cell(row, 4, data['total'])
            ws.cell(row, 5, data['normal'])
            ws.cell(row, 6, data['late'])
            ws.cell(row, 7, data['absent'])
            ws.cell(row, 8, f"{data['attendance_rate']:.1f}%")
            row += 1

        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
            ws.column_dimensions[col].width = 15

    def create_alerts_sheet(self, wb):
        """Create alerts sheet"""
        ws = wb.create_sheet("Alerts", 2)

        ws.merge_cells('A1:E1')
        title = ws['A1']
        title.value = "Attendance Alerts"
        title.font = Font(size=14, bold=True, color='FFFFFF')
        title.fill = PatternFill(start_color='FF5722', end_color='FF5722', fill_type='solid')
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 30

        headers = ['CRM', 'Name', 'Alert Type', 'Count', 'Severity']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(3, col, header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='E64A19', end_color='E64A19', fill_type='solid')

        row = 4
        for crm, data in self.penalties_data.items():
            if data['late_count'] >= 3:
                ws.cell(row, 1, crm)
                ws.cell(row, 2, data['name'])
                ws.cell(row, 3, "Frequent Late")
                ws.cell(row, 4, data['late_count'])
                ws.cell(row, 5, "HIGH")
                row += 1

            if data['absence_count'] >= 2:
                ws.cell(row, 1, crm)
                ws.cell(row, 2, data['name'])
                ws.cell(row, 3, "Multiple Absences")
                ws.cell(row, 4, data['absence_count'])
                ws.cell(row, 5, "HIGH")
                row += 1

        for col in ['A', 'B', 'C', 'D', 'E']:
            ws.column_dimensions[col].width = 18

    def create_penalties_sheet(self, wb):
        """Create penalties sheet"""
        ws = wb.create_sheet("Penalties", 3)
        currency = self.config.get('penalties', {}).get('currency', 'EGP')

        ws.merge_cells('A1:L1')
        title = ws['A1']
        title.value = f"Attendance Penalties Report ({currency})"
        title.font = Font(size=14, bold=True, color='FFFFFF')
        title.fill = PatternFill(start_color='DC3545', end_color='DC3545', fill_type='solid')
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 30

        headers = ['CRM', 'Name', 'Department', 'Late Count', f'Late Penalty ({currency})',
                   'Missing Punches', 'Punch Ded. (days)', 'Absences', 'Absence Ded. (days)',
                   f'Total Penalty ({currency})', 'Total Ded. (days)', 'Warnings']

        for col, header in enumerate(headers, 1):
            cell = ws.cell(4, col, header)
            cell.font = Font(bold=True, color='FFFFFF', size=9)
            cell.fill = PatternFill(start_color='495057', end_color='495057', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', wrap_text=True)

        row = 5
        for crm, data in sorted(self.penalties_data.items()):
            ws.cell(row, 1, crm)
            ws.cell(row, 2, data['name'])
            ws.cell(row, 3, data.get('department', ''))
            ws.cell(row, 4, data['late_count'])
            ws.cell(row, 5, data['late_penalty'])
            ws.cell(row, 6, data['missing_punch_count'])
            ws.cell(row, 7, data['missing_deduction'])
            ws.cell(row, 8, data['absence_count'])
            ws.cell(row, 9, data['absence_deduction'])
            ws.cell(row, 10, data['total_penalty_egp'])
            ws.cell(row, 11, data['total_deduction_days'])
            ws.cell(row, 12, data['total_warnings'])
            row += 1

        for i, width in enumerate([15, 18, 15, 10, 14, 12, 12, 10, 12, 14, 12, 10], 1):
            ws.column_dimensions[get_column_letter(i)].width = width

    def create_duplicates_sheet(self, wb):
        """Create the Duplicates sheet showing leave vs attendance conflicts"""
        ws = wb.create_sheet("Duplicates", 4)

        ws.merge_cells('A1:G1')
        title = ws['A1']
        title.value = "Leave vs Attendance Conflicts"
        title.font = Font(size=14, bold=True, color='FFFFFF')
        title.fill = PatternFill(start_color='FF6600', end_color='FF6600', fill_type='solid')
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 30

        ws.merge_cells('A2:G2')
        subtitle = ws['A2']
        subtitle.value = "Employees with both attendance and leave on the same date"
        subtitle.font = Font(size=9, italic=True)
        subtitle.alignment = Alignment(horizontal='center')

        headers = ['CRM', 'Name', 'Date', 'Original Attendance Status', 'Leave Type Applied', 'Clock In', 'Clock Out']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(4, col, header)
            cell.font = Font(bold=True, color='FFFFFF', size=10)
            cell.fill = PatternFill(start_color='FF6600', end_color='FF6600', fill_type='solid')
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

        if self.conflict_records:
            row = 5
            for conflict in sorted(self.conflict_records, key=lambda x: (x['crm'], x['date'])):
                clock_in = conflict.get('clock_in')
                clock_out = conflict.get('clock_out')
                clock_in_str = str(clock_in) if clock_in else '-'
                clock_out_str = str(clock_out) if clock_out else '-'

                date_val = conflict['date']
                date_str = date_val.strftime('%d-%b-%Y') if hasattr(date_val, 'strftime') else str(date_val)

                ws.cell(row, 1, conflict['crm'])
                ws.cell(row, 2, conflict['name'])
                ws.cell(row, 3, date_str)
                ws.cell(row, 4, conflict['attendance_status'])
                ws.cell(row, 5, conflict['leave_type'])
                ws.cell(row, 6, clock_in_str)
                ws.cell(row, 7, clock_out_str)

                for col in range(1, 8):
                    cell = ws.cell(row, col)
                    cell.fill = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
                    cell.border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )
                row += 1

            summary_row = row + 1
            ws.cell(summary_row, 1, f"Total Conflicts: {len(self.conflict_records)}")
            ws.cell(summary_row, 1).font = Font(bold=True)
        else:
            ws.merge_cells('A5:G5')
            no_conflict = ws['A5']
            no_conflict.value = "No conflicts detected - All leave records are consistent with attendance data"
            no_conflict.font = Font(size=11, color='28A745')
            no_conflict.alignment = Alignment(horizontal='center')
            ws.row_dimensions[5].height = 30

        for i, width in enumerate([15, 20, 12, 25, 20, 12, 12], 1):
            ws.column_dimensions[get_column_letter(i)].width = width


# Streamlit App
def main():
    st.set_page_config(
        page_title="Attendance Dashboard",
        page_icon="📊",
        layout="wide"
    )

    # Initialize session state
    if 'processor' not in st.session_state:
        st.session_state.processor = AttendanceProcessor()
    if 'master_loaded' not in st.session_state:
        st.session_state.master_loaded = False
    if 'attendance_loaded' not in st.session_state:
        st.session_state.attendance_loaded = False
    if 'leave_loaded' not in st.session_state:
        st.session_state.leave_loaded = False
    if 'available_departments' not in st.session_state:
        st.session_state.available_departments = []
    if 'available_crms' not in st.session_state:
        st.session_state.available_crms = []
    if 'attendance_files' not in st.session_state:
        st.session_state.attendance_files = []

    # Header
    st.title("📊 Attendance Dashboard v2.2")
    st.markdown("*Web-based attendance processing application*")

    # Sidebar: Settings & Filters
    with st.sidebar:
        st.header("⚙️ Settings")

        late_threshold = st.time_input(
            "Late Threshold",
            value=datetime.strptime("12:00", "%H:%M").time()
        )

        off_days = st.multiselect(
            "Off Days",
            options=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
            default=["Friday"]
        )

        # Convert off days to numbers
        day_map = {"Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
                   "Friday": 4, "Saturday": 5, "Sunday": 6}
        off_day_nums = [day_map[d] for d in off_days]

        # Update config
        st.session_state.processor.config['late_threshold'] = late_threshold.strftime("%I:%M %p")
        st.session_state.processor.config['off_days'] = off_day_nums

        st.divider()

        # Filters (only show after master data is loaded)
        if st.session_state.master_loaded:
            st.header("🔍 Filters")

            # Get available departments and CRMs
            depts = sorted(set(info['department'] for info in st.session_state.processor.employee_mapping.values()
                               if info['department']))
            crms = sorted(set(info['crm'] for info in st.session_state.processor.employee_mapping.values()
                              if info['crm']))

            st.session_state.available_departments = depts
            st.session_state.available_crms = crms

            selected_depts = st.multiselect(
                "Departments",
                options=depts,
                default=depts,
                help="Select departments to include"
            )

            selected_crms = st.multiselect(
                "CRMs",
                options=crms,
                default=crms,
                help="Select CRMs to include"
            )

            st.session_state.selected_depts = selected_depts
            st.session_state.selected_crms = selected_crms

            filtered_count = len(st.session_state.processor.get_filtered_employee_mapping(selected_depts, selected_crms))
            total_count = len(st.session_state.processor.employee_mapping)
            st.info(f"📋 {filtered_count} / {total_count} employees selected")

    # Main content
    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("📁 Master Data")
        master_file = st.file_uploader(
            "Upload Master Data",
            type=['xlsx', 'xls'],
            key='master_upload',
            help="Required - Employee master data with CRM, Name, Department, etc."
        )

        if master_file is not None and not st.session_state.master_loaded:
            with st.spinner("Loading master data..."):
                if st.session_state.processor.load_master_data(master_file, master_file.name):
                    st.session_state.master_loaded = True
                    st.success(f"✅ Loaded {len(st.session_state.processor.employee_mapping)} employees")
                else:
                    st.error("❌ Failed to load master data")

        if st.session_state.master_loaded:
            st.success(f"✅ {len(st.session_state.processor.employee_mapping)} employees loaded")

    with col2:
        st.subheader("📅 Attendance Files")
        attendance_files = st.file_uploader(
            "Upload Attendance Files",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            key='attendance_upload',
            help="Required - Daily attendance files"
        )

        if attendance_files:
            st.session_state.attendance_files = [(f, f.name) for f in attendance_files]
            st.session_state.attendance_loaded = True
            st.success(f"✅ {len(attendance_files)} file(s) uploaded")

    with col3:
        st.subheader("🏖️ Leave Sheet")
        leave_file = st.file_uploader(
            "Upload Leave Sheet",
            type=['xlsx', 'xls'],
            key='leave_upload',
            help="Optional - Leave records"
        )

        if leave_file is not None and not st.session_state.leave_loaded:
            with st.spinner("Loading leave data..."):
                if st.session_state.processor.load_leave_data(leave_file, leave_file.name):
                    st.session_state.leave_loaded = True
                    st.success(f"✅ Loaded {len(st.session_state.processor.leave_records)} leave records")
                else:
                    st.error("❌ Failed to load leave data")

        if st.session_state.leave_loaded:
            st.success(f"✅ {len(st.session_state.processor.leave_records)} leave records loaded")

    st.divider()

    # Generate Report Button
    ready = st.session_state.master_loaded and st.session_state.attendance_loaded

    if st.button("🚀 GENERATE REPORT", type="primary", disabled=not ready, use_container_width=True):
        with st.spinner("Processing attendance data..."):
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Get filter selections
            selected_depts = getattr(st.session_state, 'selected_depts', None)
            selected_crms = getattr(st.session_state, 'selected_crms', None)

            # Process attendance files
            status_text.text("Processing attendance files...")
            progress_bar.progress(20)
            st.session_state.processor.process_attendance_files(
                st.session_state.attendance_files,
                selected_depts,
                selected_crms
            )

            # Apply leave records
            if st.session_state.processor.leave_records:
                status_text.text("Applying leave records...")
                progress_bar.progress(50)
                st.session_state.processor.fill_leave_records(selected_depts, selected_crms)

            # Generate Excel report
            status_text.text("Generating Excel report...")
            progress_bar.progress(80)
            excel_data = st.session_state.processor.create_excel_report()

            progress_bar.progress(100)
            status_text.text("✅ Report generated successfully!")

            # Download button
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="📥 Download Report",
                data=excel_data,
                file_name=f"Attendance_Report_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            # Show logs
            with st.expander("📋 Processing Log"):
                for log in st.session_state.processor.logs:
                    if log['level'] == 'error':
                        st.error(log['message'])
                    elif log['level'] == 'warning':
                        st.warning(log['message'])
                    elif log['level'] == 'success':
                        st.success(log['message'])
                    else:
                        st.info(log['message'])

            # Show conflicts if any
            if st.session_state.processor.conflict_records:
                st.warning(f"⚠️ Detected {len(st.session_state.processor.conflict_records)} leave vs attendance conflicts")
                with st.expander("View Conflicts"):
                    conflict_df = pd.DataFrame(st.session_state.processor.conflict_records)
                    st.dataframe(conflict_df)

    if not ready:
        st.info("📌 Please upload Master Data and at least one Attendance file to generate a report.")

    # Reset button
    if st.button("🔄 Reset All"):
        st.session_state.processor = AttendanceProcessor()
        st.session_state.master_loaded = False
        st.session_state.attendance_loaded = False
        st.session_state.leave_loaded = False
        st.session_state.attendance_files = []
        st.rerun()

    # Footer
    st.divider()
    st.caption("Attendance Dashboard v2.2 - Web Edition | Built with Streamlit")


if __name__ == "__main__":
    main()
