"""
Excel file reading utilities.
"""

import pandas as pd
from typing import Optional


def load_excel_file(filepath: str, skiprows: int = 3) -> pd.DataFrame:
    """
    Load an Excel file with specified header row skip.

    Args:
        filepath: Path to Excel file
        skiprows: Number of rows to skip (default 3 for most formats)

    Returns:
        pandas DataFrame with the data
    """
    return pd.read_excel(filepath, skiprows=skiprows)


def parse_value(val) -> float:
    """
    Parse a cell value to numeric, handling various empty/missing formats.

    Args:
        val: Cell value (could be number, string, NaN, etc.)

    Returns:
        Float value, or 0 if empty/invalid
    """
    if pd.isna(val) or val == '' or str(val).strip() in ['-', ' - ', '--']:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
