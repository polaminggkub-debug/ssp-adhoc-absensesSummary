"""
Employee aggregation service with fuzzy name matching.

This module handles deduplication and merging of employee records across months.
"""

from difflib import SequenceMatcher
from typing import List, Dict, Any, Optional, Tuple
import re
import pandas as pd


def extract_nickname(display_name: str) -> str:
    """Extract Thai nickname from display name."""
    if not display_name:
        return ''
    match = re.search(r'\(([ก-๙]+)\)', display_name)
    return match.group(1) if match else ''


def extract_thai_only_name(display_name: str) -> str:
    """
    Extract Thai-only name part (for short Thai names like 'นาย เสร็จ').
    Returns the Thai word after prefix if it's a Thai-only name.
    """
    if not display_name:
        return ''
    # Check if this is a short Thai name (prefix + Thai word only)
    match = re.match(r'^(นาย|นาง|นางสาว)\s+([ก-๙]+)$', display_name.strip())
    if match:
        return match.group(2)  # Return the Thai name part
    return ''


def nicknames_match(name1: str, name2: str) -> bool:
    """
    Check if two names refer to same person via nickname matching.
    Handles cases like:
    - "นาย PISET SAY (เสร็จ)" vs "นาย เสร็จ"
    - Both have same nickname in parentheses
    """
    nick1 = extract_nickname(name1)
    nick2 = extract_nickname(name2)
    thai1 = extract_thai_only_name(name1)
    thai2 = extract_thai_only_name(name2)

    # Case 1: Both have nicknames and they match
    if nick1 and nick2 and nick1 == nick2:
        return True

    # Case 2: One is Thai short name, other has matching nickname
    if thai1 and nick2 and thai1 == nick2:
        return True
    if thai2 and nick1 and thai2 == nick1:
        return True

    # Case 3: Both are Thai short names and they match
    if thai1 and thai2 and thai1 == thai2:
        return True

    return False


def similarity_ratio(s1: str, s2: str) -> float:
    """
    Calculate string similarity (0-1).

    Args:
        s1: First string
        s2: Second string

    Returns:
        Similarity ratio between 0.0 and 1.0
    """
    return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()


def normalize_name_parts(name_str: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Extract and normalize name parts (prefix, firstname, lastname).

    Args:
        name_str: Full name string

    Returns:
        Tuple of (prefix, firstname, lastname)
    """
    if not name_str or pd.isna(name_str):
        return None, None, None

    name_str = str(name_str).split('/')[0].strip()  # Remove notes
    name_str = re.sub(r'\s*\([ก-๙]+\)\s*', '', name_str).strip()  # Remove nickname

    parts = name_str.split()
    if not parts:
        return None, None, None

    first_part = parts[0]
    if first_part in ['นาย', 'นาง', 'นางสาว']:
        prefix = first_part
        firstname = parts[1] if len(parts) > 1 else ''
        lastname = parts[2] if len(parts) > 2 else ''
    else:
        # Thai name with prefix merged
        if first_part.startswith('นางสาว'):
            prefix, firstname = 'นางสาว', first_part[6:]
        elif first_part.startswith('นาง'):
            prefix, firstname = 'นาง', first_part[3:]
        elif first_part.startswith('นาย'):
            prefix, firstname = 'นาย', first_part[3:]
        else:
            prefix, firstname = '', first_part
        lastname = parts[1] if len(parts) > 1 else ''

    return prefix.strip(), firstname.strip(), lastname.strip()


def find_fuzzy_match(
    prefix: str,
    firstname: str,
    lastname: str,
    employee_map: Dict[str, Dict],
    threshold: float = 0.85
) -> Optional[str]:
    """
    Find fuzzy match in employee_map for potential typos/variations.

    Args:
        prefix: Name prefix (นาย/นาง/นางสาว)
        firstname: First name
        lastname: Last name
        employee_map: Map of existing employees
        threshold: Minimum similarity for lastname match

    Returns:
        Matching key if found, None otherwise
    """
    for key, emp in employee_map.items():
        parts = key.split('|')
        if len(parts) >= 3:
            existing_prefix, existing_firstname, existing_lastname = parts[0], parts[1], parts[2]

            # Prefix must match exactly
            if prefix != existing_prefix:
                continue

            # Firstname must match exactly or very closely
            fname_sim = similarity_ratio(firstname, existing_firstname)
            if fname_sim < 0.95:  # Firstname must be nearly identical
                continue

            # Lastname can have typos - check similarity
            lname_sim = similarity_ratio(lastname, existing_lastname)
            if lname_sim >= threshold:
                return key

    return None


def aggregate_yearly_totals(all_months_data: List[List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """
    Aggregate absence data using employee ID as primary key.

    Strategy:
    - If employee has an ID: aggregate by ID (all records with same ID become one employee)
    - If employee has no ID: aggregate by name key with fuzzy matching

    This ensures that when an employee code is reused for different people,
    all their data gets combined under that single ID (as per master file).

    Args:
        all_months_data: List of monthly employee data lists

    Returns:
        List of aggregated employee dictionaries
    """
    # Separate maps for ID-based and name-based aggregation
    id_employee_map = {}    # emp_id -> employee data (for employees with IDs)
    name_employee_map = {}  # name_key -> employee data (for employees without IDs)

    for month_data in all_months_data:
        for emp in month_data:
            name_key = emp['primary_key']
            display_name = emp['display_name']
            emp_id = emp['emp_id'].strip() if emp['emp_id'] else ''

            if emp_id:
                # Employee has ID - but IDs can be REUSED for different people!
                # We must verify name or nickname similarity before merging
                if emp_id in id_employee_map:
                    existing = id_employee_map[emp_id]
                    # Check if names are similar enough (85%+ similarity)
                    name_sim = similarity_ratio(name_key, existing['name_key'])

                    # Also check nickname match - same nickname = same person
                    # This handles cases like "นาย เสร็จ" vs "นาย PISET SAY (เสร็จ)"
                    nickname_match = nicknames_match(display_name, existing['name'])

                    if name_sim >= 0.85 or nickname_match:
                        # Same person - merge into existing record
                        target = existing
                        if display_name != target['name'] and display_name not in target['original_names']:
                            target['merge_reasons'].add(f"ID Merge: {display_name}")
                    else:
                        # Different person with reused ID - treat as separate
                        # Use compound key: emp_id + name_key
                        compound_key = f"{emp_id}|{name_key}"
                        if compound_key not in id_employee_map:
                            id_employee_map[compound_key] = {
                                'name': display_name,
                                'name_key': name_key,
                                'emp_id': emp_id,  # Keep original ID for traceability
                                'notes': set(),
                                'original_names': set(),
                                'merge_reasons': set(),
                                'position': emp['position'],
                                'department': emp['department'],
                                'payType': emp['payType'],
                                'totals': [0] * 17
                            }
                        target = id_employee_map[compound_key]
                else:
                    # First time seeing this ID
                    id_employee_map[emp_id] = {
                        'name': display_name,
                        'name_key': name_key,
                        'emp_id': emp_id,
                        'notes': set(),
                        'original_names': set(),
                        'merge_reasons': set(),
                        'position': emp['position'],
                        'department': emp['department'],
                        'payType': emp['payType'],
                        'totals': [0] * 17
                    }
                    target = id_employee_map[emp_id]

            else:
                # No ID - aggregate by exact name only
                # DISABLED fuzzy matching - it was merging completely different employees
                # like นาง CHO ZIN with นาย WIN TUN and นาย YE KYAW
                matched_key = name_key

                if matched_key not in name_employee_map:
                    name_employee_map[matched_key] = {
                        'name': display_name,
                        'name_key': matched_key,
                        'emp_id': '',
                        'notes': set(),
                        'original_names': set(),
                        'merge_reasons': set(),
                        'position': emp['position'],
                        'department': emp['department'],
                        'payType': emp['payType'],
                        'totals': [0] * 17
                    }

                target = name_employee_map[matched_key]

            # Add data to target
            target['original_names'].add(display_name)
            if emp['note']:
                target['notes'].add(emp['note'])
            for i in range(17):
                target['totals'][i] += emp['totals'][i]

    # Combine both maps
    all_employees = list(id_employee_map.values()) + list(name_employee_map.values())

    # SECOND PASS: Merge employees with same name but different IDs
    # This handles cases like HTET WINT having both R88019 and SBI2101
    name_key_map = {}  # name_key -> merged employee record
    final_employees = []

    for emp in all_employees:
        name_key = emp['name_key']

        if name_key in name_key_map:
            # Same person with different ID - merge
            existing = name_key_map[name_key]

            # Combine IDs (pipe-separated)
            existing_ids = set(existing['emp_id'].split(' | ')) if existing['emp_id'] else set()
            new_ids = set(emp['emp_id'].split(' | ')) if emp['emp_id'] else set()
            combined_ids = existing_ids | new_ids
            combined_ids.discard('')  # Remove empty strings
            existing['emp_id'] = ' | '.join(sorted(combined_ids))

            # Track the merge
            existing['merge_reasons'].add(f"Same Name: {emp['emp_id']} ({emp['name']})")

            # Combine original names
            if isinstance(emp['original_names'], set):
                existing['original_names'].update(emp['original_names'])
            elif emp['original_names']:
                for n in emp['original_names'].split(' | '):
                    if n.strip():
                        existing['original_names'].add(n.strip())

            # Combine notes
            if isinstance(emp['notes'], set):
                existing['notes'].update(emp['notes'])
            elif emp['notes']:
                for n in emp['notes'].split(' | '):
                    if n.strip():
                        existing['notes'].add(n.strip())

            # Sum totals
            for i in range(17):
                existing['totals'][i] += emp['totals'][i]
        else:
            # First time seeing this name_key
            # Ensure sets are preserved for merging
            if not isinstance(emp['original_names'], set):
                names_set = set()
                if emp['original_names']:
                    for n in emp['original_names'].split(' | '):
                        if n.strip():
                            names_set.add(n.strip())
                emp['original_names'] = names_set

            if not isinstance(emp['notes'], set):
                notes_set = set()
                if emp['notes']:
                    for n in emp['notes'].split(' | '):
                        if n.strip():
                            notes_set.add(n.strip())
                emp['notes'] = notes_set

            if not isinstance(emp['merge_reasons'], set):
                reasons_set = set()
                if emp['merge_reasons']:
                    for r in emp['merge_reasons'].split(' | '):
                        if r.strip():
                            reasons_set.add(r.strip())
                emp['merge_reasons'] = reasons_set

            name_key_map[name_key] = emp

    all_employees = list(name_key_map.values())

    # Convert sets to strings
    for emp in all_employees:
        emp['notes'] = ' | '.join(sorted(emp['notes'])) if emp['notes'] else ''
        emp['original_names'] = ' | '.join(sorted(emp['original_names'])) if emp['original_names'] else ''
        emp['merge_reasons'] = ' | '.join(sorted(emp['merge_reasons'])) if emp['merge_reasons'] else ''

    # Sort by employee ID first, then by name
    return sorted(all_employees, key=lambda x: (x['emp_id'] == '', x['emp_id'], x['name']))


def extract_name_key_and_notes(full_name: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Extract matching key and notes from full name string.
    Re-exported from models.employee for backwards compatibility.
    """
    from models.employee import extract_name_key_and_notes as _extract
    return _extract(full_name)
