"""
Employee data model and name parsing utilities.
"""

import re
import pandas as pd
from typing import Optional, Tuple, Dict, Any
from dataclasses import dataclass, field
from typing import List, Set


@dataclass
class Employee:
    """
    Represents an employee with absence data.

    Attributes:
        primary_key: Key used for matching/deduplication
        name_key: Name-based key for fuzzy matching
        emp_id: Employee ID (may be pipe-separated if multiple)
        nickname: Thai nickname in parentheses
        display_name: Formatted name for display
        note: Notes from name field (quit dates, transfers, etc.)
        position: Job position
        department: Department name
        pay_type: Payment type (daily/monthly)
        totals: List of 17 absence type totals
    """
    primary_key: str
    name_key: str
    emp_id: str
    nickname: str
    display_name: str
    note: Optional[str]
    position: str
    department: str
    pay_type: str
    totals: List[float] = field(default_factory=lambda: [0.0] * 17)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for compatibility with existing code."""
        return {
            'primary_key': self.primary_key,
            'name_key': self.name_key,
            'emp_id': self.emp_id,
            'nickname': self.nickname,
            'display_name': self.display_name,
            'note': self.note,
            'position': self.position,
            'department': self.department,
            'payType': self.pay_type,
            'totals': self.totals,
        }


def normalize_prefix(prefix: str) -> str:
    """
    Normalize Thai prefix abbreviations to full form.

    This ensures that the same person with different prefix formats
    (e.g., 'น.ส.' vs 'นางสาว') will be matched as the same person.
    """
    PREFIX_ABBREVIATIONS = {
        'น.ส.': 'นางสาว',
        'นส.': 'นางสาว',
        'น.ส': 'นางสาว',
        'นส': 'นางสาว',
        'น.': 'นาย',  # rare but possible
    }
    return PREFIX_ABBREVIATIONS.get(prefix, prefix)


def extract_name_key_and_notes(full_name: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Extract matching key and notes from full name string.

    Key = prefix|firstname|lastname (ignores nickname for fuzzy matching)

    Example: "นาง CHHUN ORNG LY (รี)/ลาออก 27/03"
    -> key: "นาง|CHHUN|ORNG", display_name: "นาง CHHUN ORNG (รี)", note: "ลาออก 27/03"

    Args:
        full_name: Raw name string from Excel

    Returns:
        Tuple of (matching_key, display_name, note)
    """
    if not full_name or pd.isna(full_name):
        return None, None, None

    full_name = str(full_name).strip()

    # Extract notes after /
    note = None
    match = re.search(r'/(?=[ก-๙a-zA-Z])', full_name)
    if match:
        note = full_name[match.start()+1:].strip()
        name_part = full_name[:match.start()].strip()
    else:
        name_part = full_name

    # Extract nickname (Thai in parentheses)
    nick_match = re.search(r'\(([ก-๙]+)\)', name_part)
    nickname = nick_match.group(1) if nick_match else ''

    # Remove nickname from name
    name_clean = re.sub(r'\s*\([ก-๙]+\)\s*', '', name_part).strip()
    parts = name_clean.split()

    if not parts:
        return None, None, None

    first_part = parts[0]

    # Detect if Thai or Foreign name pattern
    # Check for full prefixes first (with or without space)
    if first_part in ['นาย', 'นาง', 'นางสาว']:
        # Foreign name: นาง FIRSTNAME LASTNAME [EXTRA...]
        prefix = first_part
        firstname = parts[1] if len(parts) > 1 else ''
        lastname = parts[2] if len(parts) > 2 else ''
    # Check for abbreviated prefixes (e.g., น.ส., นส.)
    elif first_part in ['น.ส.', 'นส.', 'น.ส', 'นส']:
        prefix = normalize_prefix(first_part)
        firstname = parts[1] if len(parts) > 1 else ''
        lastname = parts[2] if len(parts) > 2 else ''
    else:
        # Thai name: นายFirstname Lastname or น.ส.Firstname
        if first_part.startswith('นางสาว'):
            prefix, firstname = 'นางสาว', first_part[6:]
        elif first_part.startswith('น.ส.'):
            prefix, firstname = 'นางสาว', first_part[4:]  # Normalize to full form
        elif first_part.startswith('นส.'):
            prefix, firstname = 'นางสาว', first_part[3:]  # Normalize to full form
        elif first_part.startswith('นาง'):
            prefix, firstname = 'นาง', first_part[3:]
        elif first_part.startswith('นาย'):
            prefix, firstname = 'นาย', first_part[3:]
        else:
            prefix, firstname = '', first_part
        lastname = parts[1] if len(parts) > 1 else ''

    # Create matching key (ignore nickname - same person might have different nicknames)
    key = f'{prefix}|{firstname}|{lastname}'

    # Create display name (always add space after prefix, include nickname if available)
    if nickname:
        display_name = f'{prefix} {firstname} {lastname} ({nickname})'
    else:
        display_name = f'{prefix} {firstname} {lastname}'

    # Clean up extra spaces
    display_name = ' '.join(display_name.split())

    return key, display_name, note


def extract_nickname(full_name: str) -> str:
    """
    Extract Thai nickname from name string.

    Args:
        full_name: Full name string

    Returns:
        Nickname string or empty string if not found
    """
    if not full_name or pd.isna(full_name):
        return ''

    full_name_clean = str(full_name).split('/')[0].strip()
    nick_match = re.search(r'\(([ก-๙]+)\)', full_name_clean)
    return nick_match.group(1) if nick_match else ''
