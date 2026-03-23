import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
from dotenv import load_dotenv

load_dotenv()

class GoogleSheetsStorage:
    def __init__(self, credentials_file, spreadsheet_id):
        self.scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, self.scope)
        self.client = gspread.authorize(self.creds)
        self.spreadsheet = self.client.open_by_key(spreadsheet_id)
        self._initialize_sheets()

    def _initialize_sheets(self):
        """
        Creates worksheets if they don't exist.
        """
        try:
            self.attendance_ws = self.spreadsheet.worksheet("Attendance")
        except gspread.exceptions.WorksheetNotFound:
            self.attendance_ws = self.spreadsheet.add_worksheet(title="Attendance", rows="1000", cols="5")
            self.attendance_ws.append_row(["User ID", "Фамилия", "Action", "Date", "Time"])

        try:
            self.employees_ws = self.spreadsheet.worksheet("Employees")
        except gspread.exceptions.WorksheetNotFound:
            self.employees_ws = self.spreadsheet.add_worksheet(title="Employees", rows="500", cols="2")
            self.employees_ws.append_row(["user_id", "Фамилия"])

        try:
            self.summary_ws = self.spreadsheet.worksheet("Summary")
        except gspread.exceptions.WorksheetNotFound:
            self.summary_ws = self.spreadsheet.add_worksheet(title="Summary", rows="500", cols="31")
            self.summary_ws.append_row(["Фамилия"])

    def get_employee_name(self, user_id):
        """
        Returns the registered full name for a user_id.
        """
        records = self.employees_ws.get_all_records()
        for rec in records:
            if str(rec['user_id']) == str(user_id):
                return rec['Фамилия']
        return None

    def register_employee(self, user_id, full_name):
        """
        Registers a new employee.
        """
        # Check if already registered
        records = self.employees_ws.get_all_records()
        for rec in records:
            if str(rec['user_id']) == str(user_id):
                return False
        
        self.employees_ws.append_row([user_id, full_name])
        
        # Also ensure they have a row in Summary
        summary_names = self.summary_ws.col_values(1)
        if full_name not in summary_names:
            self.summary_ws.append_row([full_name])
            
        return True

    def _update_summary(self, full_name, date_str, value="+"):
        """
        Updates the Summary sheet with a value (plus or hours).
        """
        # Get headers (dates)
        headers = self.summary_ws.row_values(1)
        
        # Find or create column for date
        try:
            date_col = headers.index(date_str) + 1
        except ValueError:
            date_col = len(headers) + 1
            self.summary_ws.update_cell(1, date_col, date_str)
        
        # Find row for employee
        names = self.summary_ws.col_values(1)
        try:
            # Case-insensitive search for safety
            emp_row = next(i for i, name in enumerate(names) if name.strip().lower() == full_name.strip().lower()) + 1
        except StopIteration:
            emp_row = len(names) + 1
            self.summary_ws.update_cell(emp_row, 1, full_name)
        
        # Update the cell
        self.summary_ws.update_cell(emp_row, date_col, value)

    def add_attendance(self, user_id, full_name, action, value="+"):
        """
        Adds attendance record.
        """
        now = datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        time_str = now.strftime("%H:%M:%S")
        
        self.attendance_ws.append_row([user_id, full_name, action, date_str, time_str])
        
        if action in ["Пришел", "Ушел раньше"]:
            self._update_summary(full_name, date_str, value)
            
        return True
