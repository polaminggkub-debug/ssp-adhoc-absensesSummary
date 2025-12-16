#!/usr/bin/env python3
"""
Multi-Format Absence Data Aggregation Tool

Aggregates monthly absence Excel files from different formats (01-11.2568.xlsx)
into a comprehensive yearly summary with 6 sheets:
1. Executive Summary - CEO view (30 seconds)
2. Suspicious - Flagged employees for HR review
3. Master Match - Employee master matching audit trail
4. Merged Names - Audit trail for merged employees
5. Data Traceback - File-by-file breakdown
6. Employees - Complete detailed data

Usage:
    python3 main.py

Input: 01.2568.xlsx through 12.2568.xlsx (in current directory)
       employee_master.xlsx (optional - for ID/name standardization)
Output: absence-summary-2568.xlsx
"""

import glob
import os
import sys
from typing import List, Tuple, Dict, Any, Optional

import pandas as pd

from config.absence_mapping import get_format_for_file, FORMAT_CONFIGS
from formats.absence_format_01_07 import AbsenceFormat0107
from formats.absence_format_08_09 import AbsenceFormat0809
from formats.absence_format_10 import AbsenceFormat10
from formats.absence_format_11 import AbsenceFormat11
from file_io.excel_reader import load_excel_file, parse_value
from services.aggregator import aggregate_yearly_totals
from services.master_matcher import apply_master_data
from services.excel_exporter import (
    create_output_dataframe,
    calculate_summary_stats,
    export_to_excel
)


def find_monthly_files() -> List[str]:
    """
    Find all XX.2568.xlsx files in current directory.

    Returns:
        Sorted list of file paths
    """
    files = glob.glob('[0-1][0-9].2568.xlsx')
    return sorted(files)


def get_format_handler(filename: str):
    """
    Get the appropriate format handler for a file.

    Args:
        filename: Excel filename (e.g., '01.2568.xlsx')

    Returns:
        Format handler instance
    """
    format_key = get_format_for_file(filename)

    handlers = {
        'A': AbsenceFormat0107,
        'B': AbsenceFormat0809,
        'C': AbsenceFormat10,
        'D': AbsenceFormat11,
    }

    handler_class = handlers.get(format_key)
    if not handler_class:
        raise ValueError(f"No handler for format {format_key}")

    return handler_class()


def process_file(filepath: str) -> Tuple[List[dict], str]:
    """
    Process a single Excel file.

    Args:
        filepath: Path to Excel file

    Returns:
        Tuple of (list of employee dicts, format name)
    """
    handler = get_format_handler(filepath)
    employees = handler.extract_employees(filepath)
    return employees, handler.format_name


def extract_section_data(filepath: str) -> Optional[Dict[str, Any]]:
    """
    Extract section-level totals from a file for traceback.

    Args:
        filepath: Path to Excel file

    Returns:
        Dict with section breakdowns, or None if file has no sections.
        Format A (01-07): {'sections': ['First Half', 'Second Half'], 'section0': [17 totals], 'section1': [17 totals]}
        Format B (08-09): {'sections': ['Monthly Total', 'First Half', 'Second Half'], ...}
        Format C/D: None (no sections)
    """
    format_key = get_format_for_file(filepath)
    config = FORMAT_CONFIGS[format_key]

    if format_key == 'A':
        # Format A: Has two halves
        df = load_excel_file(filepath, skiprows=config['header_row'])
        cols_half1 = config['absence_cols_half1']
        cols_half2 = config['absence_cols_half2']

        half1_totals = [0.0] * 17
        half2_totals = [0.0] * 17

        for _, row in df.iterrows():
            # Skip empty rows
            if pd.isna(row.iloc[config['name_col']]):
                continue

            for i in range(17):
                half1_totals[i] += parse_value(row.iloc[cols_half1[i]])
                half2_totals[i] += parse_value(row.iloc[cols_half2[i]])

        return {
            'sections': ['First Half', 'Second Half'],
            'section0': half1_totals,
            'section1': half2_totals
        }

    elif format_key == 'B':
        # Format B: Has monthly totals + two halves
        df = load_excel_file(filepath, skiprows=config['header_row'])
        absence_cols = config['absence_cols']

        # Columns 24-40 are first half, 41-57 are second half
        # Based on actual Format B structure
        cols_half1 = [24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40]
        cols_half2 = [41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57]

        half1_totals = [0.0] * 17
        half2_totals = [0.0] * 17

        for _, row in df.iterrows():
            if pd.isna(row.iloc[config['name_col']]):
                continue

            for i in range(17):
                half1_totals[i] += parse_value(row.iloc[cols_half1[i]])
                half2_totals[i] += parse_value(row.iloc[cols_half2[i]])

        return {
            'sections': ['First Half', 'Second Half'],
            'section0': half1_totals,
            'section1': half2_totals
        }

    # Format C/D: No sections
    return None


def main():
    """Main execution function."""
    print('=' * 60)
    print('Multi-Format Absence Data Aggregation Tool')
    print('=' * 60)

    # Find files
    print('\nFinding monthly files...')
    files = find_monthly_files()

    if not files:
        print('ERROR: No monthly files found (XX.2568.xlsx)')
        print('Please ensure files are in the current directory.')
        sys.exit(1)

    print(f'Found {len(files)} files: {", ".join(files)}')

    # Process all months
    all_months_data = []
    processed_files = []  # Track successfully processed files
    section_data_list = []  # Track section breakdowns for traceback
    format_summary = {}

    print('\nProcessing files...')
    for filepath in files:
        try:
            employees, format_name = process_file(filepath)
            all_months_data.append(employees)
            processed_files.append(filepath)

            # Extract section data for traceback
            section_data = extract_section_data(filepath)
            section_data_list.append(section_data)

            # Track format usage
            if format_name not in format_summary:
                format_summary[format_name] = []
            format_summary[format_name].append(filepath)

            print(f'  {filepath}: {len(employees)} employees ({format_name})')
        except Exception as e:
            print(f'  ERROR processing {filepath}: {e}')
            import traceback
            traceback.print_exc()
            continue

    if not all_months_data:
        print('ERROR: No data could be processed')
        sys.exit(1)

    # Show format summary
    print('\nFormat Summary:')
    for fmt, files in format_summary.items():
        print(f'  {fmt}: {", ".join(files)}')

    # Aggregate
    print(f'\nAggregating {len(all_months_data)} months...')
    aggregated = aggregate_yearly_totals(all_months_data)
    print(f'-> {len(aggregated)} unique employees')

    # Calculate total raw records
    total_raw = sum(len(month) for month in all_months_data)
    merged = total_raw - len(aggregated)
    print(f'-> {merged} duplicate records merged')

    # Apply master employee data if available
    match_audit = None
    master_file = 'employee_master.xlsx'
    if os.path.exists(master_file):
        print(f'\nApplying master employee data from {master_file}...')
        aggregated, match_audit = apply_master_data(aggregated, master_file)
    else:
        print(f'\nNote: {master_file} not found - skipping master matching')

    # Calculate summary statistics with file and section traceback
    print('\nCalculating data traceback...')
    summary_stats = calculate_summary_stats(aggregated, all_months_data, processed_files, section_data_list)
    summary_df = pd.DataFrame(summary_stats)

    # Validate: Check for duplicate IDs before export
    emp_ids = [emp.get('emp_id', '') for emp in aggregated if emp.get('emp_id', '')]
    duplicate_ids = [emp_id for emp_id in set(emp_ids) if emp_ids.count(emp_id) > 1]
    if duplicate_ids:
        print(f'\n⚠ WARNING: {len(duplicate_ids)} duplicate IDs found!')
        for dup_id in sorted(duplicate_ids)[:10]:
            print(f'  - {dup_id}')
        if len(duplicate_ids) > 10:
            print(f'  ... and {len(duplicate_ids) - 10} more')
    else:
        print('\n✓ No duplicate IDs - data is clean')

    # Export
    print('\nExporting to Excel...')
    df_output = create_output_dataframe(aggregated)
    export_to_excel(df_output, summary_df, aggregated, all_months_data, match_audit)

    # Final summary
    print('\n' + '=' * 60)
    print('COMPLETE!')
    print('=' * 60)
    print(f'Files processed: {len(processed_files)}')
    print(f'Total records: {total_raw}')
    print(f'Unique employees: {len(aggregated)}')
    print(f'Records merged: {merged}')
    if match_audit:
        matched = sum(1 for m in match_audit if m['match_type'] != 'UNMATCHED')
        unmatched = len(match_audit) - matched
        print(f'Master matched: {matched}, Unmatched: {unmatched}')
    print(f'\nOutput: absence-summary-2568.xlsx')
    print('\nSheets:')
    print('  1. Executive Summary - CEO overview')
    print('  2. Suspicious - Records for HR review')
    if match_audit:
        print('  3. Master Match - Master data matching audit')
        print('  4. Merged Names - Merge audit trail')
        print('  5. Data Traceback - File-by-file breakdown')
        print('  6. Employees - Complete data')
    else:
        print('  3. Merged Names - Merge audit trail')
        print('  4. Data Traceback - File-by-file breakdown')
        print('  5. Employees - Complete data')


if __name__ == '__main__':
    main()
