"""
Absence Type Definitions and Column Mappings for Each Excel Format

This file defines the 17 standard absence types and maps them to column positions
for each of the 4 different Excel formats used across months 01-11.2568.
"""

# The 17 standard absence types (order matters - index used for totals array)
ABSENCE_TYPES = [
    'วันทำงาน',           # 0: Work Days
    'ขาดงาน',             # 1: Absent
    'ลากิจ',              # 2: Personal Leave
    'ป่วยมีใบรพ.',         # 3: Sick w/Cert
    'ป่วยไม่มีรพ.',        # 4: Sick w/o Cert
    'ลาคลอด',             # 5: Maternity
    'ลืมสแกน/มาสาย',       # 6: Late Grace
    'มาสายเกิน',          # 7: Late Penalty
    'ลาOT',              # 8: OT Leave
    'พักงาน',             # 9: Suspension
    'พักร้อน',            # 10: Annual Leave
    'OT 2.5 ชม',         # 11: OT 2.5hr
    'OT >2.5 ชม',        # 12: OT >2.5hr
    'ทำงานวันหยุด',        # 13: Holiday Work
    'OT วันหยุด',         # 14: Holiday OT
    'กะดึก',              # 15: Night Shift
    'ควบคุม 2 เครื่อง',    # 16: Multi-Machine
]

# Column headers for export (Thai + English)
ABSENCE_COLUMN_HEADERS = [
    'วันทำงาน (Work Days)',
    'ขาดงาน (Absent)',
    'ลากิจ (Personal Leave)',
    'ป่วยมีใบรพ. (Sick w/Cert)',
    'ป่วยไม่มีรพ. (Sick w/o Cert)',
    'ลาคลอด (Maternity)',
    'ลืมสแกนนิ้ว/มาสาย (Late Grace)',
    'มาสายเกิน (Late Penalty)',
    'ลาOT (OT Leave)',
    'ให้หยุด/พักงาน (Suspension)',
    'พักร้อน (Annual Leave)',
    'OT 2.5 ชม',
    'OT >2.5 ชม',
    'ทำงานวันหยุด (Holiday Work)',
    'OT วันหยุด (Holiday OT)',
    'กะดึก (Night Shift)',
    'ควบคุม 2 เครื่อง (Multi-Machine)'
]

# Format configurations
# Each format defines how to extract the 17 absence types from the Excel columns
FORMAT_CONFIGS = {
    # Format A: Months 01-07
    # 42 columns, header at row 3, has two halves that need to be summed
    # Columns 5-21 = first half, Columns 22-38 = second half
    'A': {
        'files': ['01', '02', '03', '04', '05', '06', '07'],
        'header_row': 3,  # skiprows=3 in pandas
        'id_col': 0,
        'name_col': 1,
        'position_col': 2,
        'department_col': 3,
        'paytype_col': 4,
        'sum_halves': True,
        # First half columns (add 17 to get second half)
        # Index in array -> column position
        # [0] Work Days: col 5 + col 22
        # [1] Absent: col 6 + col 23
        # etc.
        'absence_cols_half1': [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21],
        'absence_cols_half2': [22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38],
    },

    # Format B: Months 08-09
    # 58 columns, header at row 3, has monthly totals in columns 5-23
    # Also has two halves but we use the monthly totals directly
    'B': {
        'files': ['08', '09'],
        'header_row': 3,  # skiprows=3 in pandas (data header at row index 3)
        'id_col': 0,
        'name_col': 1,
        'position_col': 2,
        'department_col': 3,
        'paytype_col': 4,
        'sum_halves': False,  # Use monthly totals directly
        # Mapping: Standard type index -> column position
        # Based on actual headers from 08.2568.xlsx row 3:
        # col 5: วันทำงานในเดือนนี้ -> [0] Work Days
        # col 8: ขาดงานในเดือนนี้ -> [1] Absent
        # col 9: ลากิจ-เดือนนี้ -> [2] Personal Leave
        # col 10: ป่วยมีใบรพ.ในเดือนนี้ -> [3] Sick w/Cert
        # col 11: ป่วยไม่มีรพ.ในเดือนนี้ -> [4] Sick w/o Cert
        # col 12: ลาคลอด -> [5] Maternity
        # col 13: ลืมสแกนนิ้ว -> [6] Late Grace
        # col 14: มาสายเกิน -> [7] Late Penalty
        # col 15: ลาOTในเดือนนี้ -> [8] OT Leave
        # col 16: พักงาน -> [9] Suspension
        # col 17: พักร้อนเดือนนี้ -> [10] Annual Leave
        # col 18: OT 2.5 ชม -> [11] OT 2.5hr
        # col 19: OT >2.5 ชม -> [12] OT >2.5hr
        # col 20: ทำงานในวันหยุด -> [13] Holiday Work
        # col 21: OT ในวันหยุด -> [14] Holiday OT
        # col 22: กะดึก -> [15] Night Shift
        # col 23: ควบคุม 2 เครื่อง -> [16] Multi-Machine
        'absence_cols': [5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],
    },

    # Format C: Month 10
    # 41 columns, header at row 3, has ลำดับ column at position 0
    # ID and Name columns shifted by 1
    'C': {
        'files': ['10'],
        'header_row': 3,  # skiprows=3 in pandas
        'id_col': 1,  # Shifted due to ลำดับ column
        'name_col': 2,
        'position_col': 3,
        'department_col': 4,
        'paytype_col': 5,
        'status_col': 6,  # New: สถานะ column
        'sum_halves': False,
        # Mapping: Standard type index -> column position
        # IMPORTANT: Columns are in different order than Format A/B!
        # Based on actual headers from 10.2568.xlsx row 3:
        # col 7: วันที่มาทำงาน -> [0] Work Days
        # col 17: ขาดงาน -> [1] Absent
        # col 15: ลากิจ -> [2] Personal Leave
        # col 13: ลาป่วย+ใบรพ -> [3] Sick w/Cert
        # col 14: ลาป่วยไม่มีใบรพ -> [4] Sick w/o Cert
        # col 16: ลาคลอด -> [5] Maternity
        # col 21: ลืมสแกนนิ้ว -> [6] Late Grace
        # col 22: เข้างานช้า -> [7] Late Penalty
        # col 20: ลาโอที -> [8] OT Leave
        # col 18: พักงาน -> [9] Suspension
        # col 8: ลาพักร้อน -> [10] Annual Leave
        # col 9: ทำโอที -> [11] OT 2.5hr
        # col 10: โอทีที่เกินเวลาปกติ -> [12] OT >2.5hr
        # col 11: ทำงานวันหยุด -> [13] Holiday Work
        # col 12: โอทีวันหยุด -> [14] Holiday OT
        # col 27: วันที่เข้ากะดึก -> [15] Night Shift
        # col 28+29: เครื่องคู่ x40 + x60 -> [16] Multi-Machine (SUM both)
        'absence_cols': [7, 17, 15, 13, 14, 16, 21, 22, 20, 18, 8, 9, 10, 11, 12, 27, None],
        'multi_machine_cols': [28, 29],  # Need to sum these for index 16
    },

    # Format D: Month 11
    # 36 columns, header at row 4 (different!), same structure as Format C
    'D': {
        'files': ['11'],
        'header_row': 4,  # skiprows=4 in pandas (header at row index 4!)
        'id_col': 1,
        'name_col': 2,
        'position_col': 3,
        'department_col': 4,
        'paytype_col': 5,
        'status_col': 6,
        'sum_halves': False,
        # Same column mapping as Format C
        'absence_cols': [7, 17, 15, 13, 14, 16, 21, 22, 20, 18, 8, 9, 10, 11, 12, 27, None],
        'multi_machine_cols': [28, 29],
    },
}


def get_format_for_file(filename: str) -> str:
    """
    Determine which format to use based on filename.

    Args:
        filename: The Excel filename (e.g., '01.2568.xlsx')

    Returns:
        Format key ('A', 'B', 'C', or 'D')

    Raises:
        ValueError: If no format found for the file
    """
    # Extract month number from filename
    import re
    match = re.match(r'^(\d{2})\.2568\.xlsx$', filename)
    if not match:
        raise ValueError(f"Invalid filename format: {filename}")

    month = match.group(1)

    for format_key, config in FORMAT_CONFIGS.items():
        if month in config['files']:
            return format_key

    raise ValueError(f"No format configuration found for month {month}")
