"""
Format D handler for month 11.2568.xlsx

Structure:
- 36 columns
- Header at row 4 (skiprows=4) - DIFFERENT from other formats!
- Has ลำดับ column at position 0, so ID is at column 1, Name at column 2
- Same column structure as Format C (month 10)
- Multi-machine needs to sum columns 28 + 29
"""

import pandas as pd
from typing import List, Dict, Any

from .base_format import BaseFormatHandler
from config.absence_mapping import FORMAT_CONFIGS
from file_io.excel_reader import load_excel_file, parse_value
from models.employee import extract_name_key_and_notes, extract_nickname


class AbsenceFormat11(BaseFormatHandler):
    """Handler for Format D (month 11)."""

    def __init__(self):
        self.config = FORMAT_CONFIGS['D']

    def get_format_config(self) -> Dict[str, Any]:
        return self.config

    @property
    def format_name(self) -> str:
        return "Format D (11)"

    def extract_employees(self, filepath: str) -> List[Dict[str, Any]]:
        """
        Extract employee data from Format D Excel file.

        Same column mapping as Format C, but header row is 4 instead of 3:
            [0] Work Days = col 7
            [1] Absent = col 17
            [2] Personal Leave = col 15
            [3] Sick w/Cert = col 13
            [4] Sick w/o Cert = col 14
            [5] Maternity = col 16
            [6] Late Grace = col 21
            [7] Late Penalty = col 22
            [8] OT Leave = col 20
            [9] Suspension = col 18
            [10] Annual Leave = col 8
            [11] OT 2.5hr = col 9
            [12] OT >2.5hr = col 10
            [13] Holiday Work = col 11
            [14] Holiday OT = col 12
            [15] Night Shift = col 27
            [16] Multi-Machine = col 28 + col 29
        """
        df = load_excel_file(filepath, skiprows=self.config['header_row'])
        employees = []

        absence_cols = self.config['absence_cols']
        multi_machine_cols = self.config['multi_machine_cols']

        for _, row in df.iterrows():
            emp_id = row.iloc[self.config['id_col']]
            full_name = row.iloc[self.config['name_col']]
            position = row.iloc[self.config['position_col']]
            department = row.iloc[self.config['department_col']]
            pay_type = row.iloc[self.config['paytype_col']]

            # Skip if no name
            if pd.isna(full_name) or str(full_name).strip() == '':
                continue

            # Extract key, display name, and notes
            key, display_name, note = extract_name_key_and_notes(full_name)
            if not key:
                continue

            # Clean employee ID
            emp_id_str = str(emp_id).strip() if pd.notna(emp_id) else ''

            # Extract nickname
            nickname = extract_nickname(full_name)

            # Extract absence totals with column remapping
            monthly_totals = []
            for i, col_idx in enumerate(absence_cols):
                if col_idx is None:
                    # Multi-machine: sum columns 28 + 29
                    val = parse_value(row.iloc[multi_machine_cols[0]]) + \
                          parse_value(row.iloc[multi_machine_cols[1]])
                else:
                    val = parse_value(row.iloc[col_idx])
                monthly_totals.append(val)

            employees.append({
                'primary_key': key,
                'name_key': key,
                'emp_id': emp_id_str,
                'nickname': nickname,
                'display_name': display_name,
                'note': note,
                'position': position,
                'department': department,
                'payType': pay_type,
                'totals': monthly_totals
            })

        return employees
