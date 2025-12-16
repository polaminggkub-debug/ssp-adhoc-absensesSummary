"""
Master Employee Matching Service

Matches aggregated employee data against the official employee_master.xlsx
to ensure consistent IDs and names in the output.
"""

import pandas as pd
from typing import Dict, List, Tuple, Optional
from difflib import SequenceMatcher

from models.employee import extract_name_key_and_notes


def load_employee_master(filepath: str = 'employee_master.xlsx') -> pd.DataFrame:
    """
    Load the employee master file.

    Returns:
        DataFrame with columns: emp_id, full_name, name_key
    """
    df = pd.read_excel(filepath, skiprows=1)
    df.columns = ['ลำดับ', 'รหัส', 'ชื่อ-นามสกุล', 'จำนวนเงิน', 'ลงชื่อ']

    # Clean and extract key fields
    master_records = []
    for _, row in df.iterrows():
        emp_id = str(row['รหัส']).strip() if pd.notna(row['รหัส']) else ''
        full_name = str(row['ชื่อ-นามสกุล']).strip() if pd.notna(row['ชื่อ-นามสกุล']) else ''

        if not emp_id or not full_name:
            continue

        # Extract name key for matching
        name_key, display_name, _ = extract_name_key_and_notes(full_name)

        master_records.append({
            'master_id': emp_id,
            'master_name': full_name,
            'master_display': display_name,
            'name_key': name_key
        })

    return pd.DataFrame(master_records)


def similarity_ratio(a: str, b: str) -> float:
    """Calculate similarity ratio between two strings."""
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def find_best_match(
    name_key: str,
    emp_id: str,
    master_df: pd.DataFrame,
    threshold: float = 0.75
) -> Optional[Dict]:
    """
    Find the best matching employee in master data.

    IMPORTANT: Employee IDs can be REUSED for different people over time!
    We must verify that both ID AND name match to avoid merging different employees.

    Priority:
    1. ID match WITH name verification (ID matches AND name similarity >= 85%)
    2. Exact name_key match (for employees whose IDs changed)

    Returns:
        Dict with master_id, master_name, master_display, match_type, confidence
        or None if no match found
    """
    # Try ID match WITH name verification
    # IDs can be reused for different people, so we must check name too
    if emp_id and name_key:
        id_parts = [id_str.strip() for id_str in emp_id.split('|')]
        for single_id in id_parts:
            if not single_id:
                continue
            id_matches = master_df[master_df['master_id'] == single_id]
            if len(id_matches) == 1:
                row = id_matches.iloc[0]
                # Verify name similarity to prevent wrong merges
                name_sim = similarity_ratio(name_key, row['name_key'])
                if name_sim >= 0.85:  # Name must be at least 85% similar
                    return {
                        'master_id': row['master_id'],
                        'master_name': row['master_name'],
                        'master_display': row['master_display'],
                        'match_type': 'ID+Name',
                        'confidence': name_sim
                    }
                # ID matched but name didn't - this is a REUSED ID, skip it

    # Try exact name_key match (for employees whose IDs may have changed)
    if name_key:
        name_matches = master_df[master_df['name_key'] == name_key]
        if len(name_matches) == 1:
            row = name_matches.iloc[0]
            return {
                'master_id': row['master_id'],
                'master_name': row['master_name'],
                'master_display': row['master_display'],
                'match_type': 'Name',
                'confidence': 1.0
            }

    return None


def apply_master_data(
    aggregated_data: List[Dict],
    master_filepath: str = 'employee_master.xlsx',
    threshold: float = 0.75
) -> Tuple[List[Dict], List[Dict]]:
    """
    Apply master employee data to aggregated records.

    This function:
    1. Matches each record to master data
    2. Uses master ID as the final employee ID (when matched)
    3. Merges records that match to the same master ID

    Args:
        aggregated_data: List of employee dicts from aggregation
        master_filepath: Path to employee master file
        threshold: Minimum similarity for fuzzy matching

    Returns:
        Tuple of (updated_data, match_audit)
        - updated_data: Merged records with master IDs/names applied
        - match_audit: List of match details for traceability
    """
    # Load master data
    master_df = load_employee_master(master_filepath)
    print(f'  Loaded {len(master_df)} employees from master file')

    # First pass: match each record to master
    match_audit = []
    matched_records = {}  # master_id -> list of matched employee records
    unmatched_records = []

    for emp in aggregated_data:
        original_id = emp.get('emp_id', '')
        original_name = emp.get('name', '') or emp.get('display_name', '')
        name_key = emp.get('name_key', '')
        notes = emp.get('notes', '') or ''

        # Find best match
        match = find_best_match(name_key, original_id, master_df, threshold)

        if match:
            master_id = match['master_id']

            # Add to audit
            match_audit.append({
                'master_id': master_id,
                'master_name': match['master_name'],
                'original_id': original_id,
                'original_name': original_name,
                'original_notes': notes,
                'match_type': match['match_type'],
                'confidence': match['confidence']
            })

            # Group by master_id for merging
            if master_id not in matched_records:
                matched_records[master_id] = {
                    'master_id': master_id,
                    'master_name': match['master_name'],
                    'master_display': match['master_display'],
                    'records': []
                }
            matched_records[master_id]['records'].append(emp)
        else:
            # Unmatched - keep original
            match_audit.append({
                'master_id': '',
                'master_name': '',
                'original_id': original_id,
                'original_name': original_name,
                'original_notes': notes,
                'match_type': 'UNMATCHED',
                'confidence': 0.0
            })
            unmatched_records.append(emp)

    # Second pass: merge records with same master_id
    updated_data = []

    for master_id, group in matched_records.items():
        records = group['records']

        if len(records) == 1:
            # Single record - just update with master info
            merged = records[0].copy()
        else:
            # Multiple records - merge them
            merged = records[0].copy()
            merged['original_names'] = set()
            merged['notes'] = set()
            merged['merge_reasons'] = set()
            merged['totals'] = [0.0] * 17

            for rec in records:
                # Collect original names
                orig_names = rec.get('original_names', '')
                if orig_names:
                    if isinstance(orig_names, set):
                        merged['original_names'].update(orig_names)
                    else:
                        for n in orig_names.split(' | '):
                            if n.strip():
                                merged['original_names'].add(n.strip())
                else:
                    merged['original_names'].add(rec.get('name', ''))

                # Collect notes
                notes = rec.get('notes', '')
                if notes:
                    if isinstance(notes, set):
                        merged['notes'].update(notes)
                    else:
                        for n in notes.split(' | '):
                            if n.strip():
                                merged['notes'].add(n.strip())

                # Track merge reason
                orig_id = rec.get('emp_id', '')
                orig_name = rec.get('name', '')
                merged['merge_reasons'].add(f"Master Merge: {orig_id} ({orig_name})")

                # Sum totals
                for i in range(17):
                    merged['totals'][i] += rec['totals'][i]

            # Convert sets to strings
            merged['original_names'] = ' | '.join(sorted(merged['original_names']))
            merged['notes'] = ' | '.join(sorted(merged['notes']))
            merged['merge_reasons'] = ' | '.join(sorted(merged['merge_reasons']))

        # Apply master data
        merged['emp_id'] = master_id
        merged['name'] = group['master_display']
        merged['master_full_name'] = group['master_name']

        updated_data.append(merged)

    # Add unmatched records (keep original IDs)
    for emp in unmatched_records:
        updated_emp = emp.copy()
        updated_emp['master_full_name'] = ''
        updated_data.append(updated_emp)

    # Fix duplicate IDs: add suffix for unmatched employees with same ID
    # This happens when the same ID was reused for different people
    id_counts = {}
    for emp in updated_data:
        emp_id = emp.get('emp_id', '')
        if emp_id:
            id_counts[emp_id] = id_counts.get(emp_id, 0) + 1

    # Find IDs that appear more than once
    duplicate_ids = {emp_id for emp_id, count in id_counts.items() if count > 1}

    if duplicate_ids:
        # Add suffix to duplicates (keep first occurrence as-is, add -A, -B, etc. to others)
        id_suffix_counters = {emp_id: 0 for emp_id in duplicate_ids}

        for emp in updated_data:
            emp_id = emp.get('emp_id', '')
            if emp_id in duplicate_ids:
                count = id_suffix_counters[emp_id]
                if count > 0:
                    # Add suffix: -A, -B, -C, etc.
                    suffix = chr(ord('A') + count - 1)
                    emp['emp_id'] = f"{emp_id}-{suffix}"
                id_suffix_counters[emp_id] += 1

        print(f'  Fixed {len(duplicate_ids)} duplicate IDs with suffixes')

    # Sort by emp_id
    updated_data.sort(key=lambda x: (x['emp_id'] == '', x['emp_id'], x.get('name', '')))

    matched_count = len(matched_records)
    unmatched_count = len(unmatched_records)
    merged_count = sum(1 for g in matched_records.values() if len(g['records']) > 1)

    print(f'  Matched: {matched_count}, Unmatched: {unmatched_count}')
    if merged_count > 0:
        print(f'  Merged {merged_count} groups with same master ID')

    return updated_data, match_audit


def get_unmatched_employees(match_audit: List[Dict]) -> List[Dict]:
    """Get list of employees that couldn't be matched to master."""
    return [m for m in match_audit if m['match_type'] == 'UNMATCHED']


def get_fuzzy_matches(match_audit: List[Dict]) -> List[Dict]:
    """Get list of fuzzy matches for review."""
    return [m for m in match_audit if m['match_type'] == 'Fuzzy']
