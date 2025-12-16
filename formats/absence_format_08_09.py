"""
Format B handler for months 08-09.2568.xlsx

Structure:
- 58 columns
- Header at row 3 (skiprows=3)
- ID at column 0, Name at column 1
- Monthly totals available at columns 5-23 (use directly, no summing needed)
"""

import pandas as pd
from typing import List, Dict, Any

from .base_format import BaseFormatHandler
from config.absence_mapping import FORMAT_CONFIGS
from file_io.excel_reader import load_excel_file, parse_value
from models.employee import extract_name_key_and_notes, extract_nickname


class AbsenceFormat0809(BaseFormatHandler):
    """Handler for Format B (months 08-09)."""

    def __init__(self):
        self.config = FORMAT_CONFIGS['B']

    def get_format_config(self) -> Dict[str, Any]:
        return self.config

    @property
    def format_name(self) -> str:
        return "Format B (08-09)"

    def extract_employees(self, filepath: str) -> List[Dict[str, Any]]:
        """
        Extract employee data from Format B Excel file.

        Uses monthly totals directly from columns 5-23.
        Column mapping:
            [0] Work Days = col 5
            [1] Absent = col 8
            [2] Personal Leave = col 9
            [3] Sick w/Cert = col 10
            [4] Sick w/o Cert = col 11
            [5] Maternity = col 12
            [6] Late Grace = col 13
            [7] Late Penalty = col 14
            [8] OT Leave = col 15
            [9] Suspension = col 16
            [10] Annual Leave = col 17
            [11] OT 2.5hr = col 18
            [12] OT >2.5hr = col 19
            [13] Holiday Work = col 20
            [14] Holiday OT = col 21
            [15] Night Shift = col 22
            [16] Multi-Machine = col 23
        """
        df = load_excel_file(filepath, skiprows=self.config['header_row'])
        employees = []

        absence_cols = self.config['absence_cols']

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

            # Extract absence totals from mapped columns
            monthly_totals = []
            for col_idx in absence_cols:
                monthly_totals.append(parse_value(row.iloc[col_idx]))

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
