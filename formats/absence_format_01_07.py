"""
Format A handler for months 01-07.2568.xlsx

Structure:
- 42 columns
- Header at row 3 (skiprows=3)
- ID at column 0, Name at column 1
- Two halves: columns 5-21 (first half) + columns 22-38 (second half)
- Sum both halves for monthly total
"""

import pandas as pd
from typing import List, Dict, Any

from .base_format import BaseFormatHandler
from config.absence_mapping import FORMAT_CONFIGS
from file_io.excel_reader import load_excel_file, parse_value
from models.employee import extract_name_key_and_notes, extract_nickname


class AbsenceFormat0107(BaseFormatHandler):
    """Handler for Format A (months 01-07)."""

    def __init__(self):
        self.config = FORMAT_CONFIGS['A']

    def get_format_config(self) -> Dict[str, Any]:
        return self.config

    @property
    def format_name(self) -> str:
        return "Format A (01-07)"

    def extract_employees(self, filepath: str) -> List[Dict[str, Any]]:
        """
        Extract employee data from Format A Excel file.

        Sums first half (cols 5-21) and second half (cols 22-38) for each absence type.
        """
        df = load_excel_file(filepath, skiprows=self.config['header_row'])
        employees = []

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

            # Sum first half + second half for each absence type
            cols_half1 = self.config['absence_cols_half1']
            cols_half2 = self.config['absence_cols_half2']

            monthly_totals = []
            for i in range(17):
                first_half = parse_value(row.iloc[cols_half1[i]])
                second_half = parse_value(row.iloc[cols_half2[i]])
                monthly_totals.append(first_half + second_half)

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
