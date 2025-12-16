# SKILL.md - Absence Aggregation Tool Skills & Techniques

This document outlines the key technical skills and techniques used in the absence aggregation tool.

## 1. String Processing & Fuzzy Matching

### Skill: Name Normalization & Extraction
**File**: `aggregate_absence.py` → `extract_name_key_and_notes()`

**Technique**: Parse complex Thai/English names with nested information
```python
Input: "นาง CHHUN ORNG (รี)/ลาออก 27/03"
Extract:
  - Prefix: "นาง" (Mrs)
  - First Name: "CHHUN"
  - Last Name: "ORNG"
  - Nickname: "รี"
  - Notes: "ลาออก 27/03" (quit on 27/03)
```

**Skills Demonstrated**:
- Regular expressions (regex) for pattern matching
- Thai character handling (`[ก-๙]+` pattern)
- String manipulation and parsing
- Conditional logic for name structure variations

**Code Pattern**:
```python
# Extract notes after /
match = re.search(r'/(?=[ก-๙a-zA-Z])', full_name)
note = full_name[match.start()+1:].strip()

# Extract nickname in Thai parentheses
nick_match = re.search(r'\(([ก-๙]+)\)', name_part)
nickname = nick_match.group(1) if nick_match else ''
```

### Skill: Fuzzy String Matching
**File**: `aggregate_absence.py` → `similarity_ratio()` + `find_fuzzy_match()`

**Technique**: Detect similar names despite typos/variations
```python
"ORNG" vs "ORNG LY" → 80% similarity → MATCHED
"CHHUN" vs "CHHUN" → 100% similarity → MATCH REQUIRED
```

**Skills Demonstrated**:
- Sequence matching algorithm (difflib.SequenceMatcher)
- Threshold-based decision making
- Multi-criteria matching (exact + fuzzy)

**Code Pattern**:
```python
from difflib import SequenceMatcher

def similarity_ratio(s1, s2):
    return SequenceMatcher(None, s1.lower(), s2.lower()).ratio()

# Result: 0.0 (no match) to 1.0 (perfect match)
# Threshold: 0.80 = require 80% similarity
```

## 2. Data Aggregation & Merging

### Skill: Grouping with Deduplication
**File**: `aggregate_absence.py` → `aggregate_yearly_totals()`

**Technique**: Merge 1,214 duplicate records into 240 unique employees

**Skills Demonstrated**:
- Dictionary-based grouping (employee_map)
- Fuzzy matching across multiple months
- Set data structures for tracking unique IDs
- Data consolidation logic

**Algorithm**:
```
For each monthly data:
  For each employee record:
    1. Generate matching key: prefix|firstname|lastname
    2. Try exact match in employee_map
    3. If not found, try fuzzy match (80% lastname similarity)
    4. If still not found, create new entry
    5. If found, merge:
       - Add ID to employee's ID set (if new)
       - Add notes to employee's notes set (if new)
       - Sum absence totals: totals[i] += emp['totals'][i]
```

### Skill: Multi-source Data Consolidation
**Technique**: Combine notes from multiple records, handle ID changes

**Code Pattern**:
```python
if matched_key not in employee_map:
    employee_map[matched_key] = {
        'name': display_name,
        'emp_ids': set(),      # Track ALL IDs
        'notes': set(),        # Track ALL notes
        'totals': [0] * 17
    }

# Add data from this month
employee_map[matched_key]['emp_ids'].add(emp['emp_id'])
employee_map[matched_key]['notes'].add(emp['note'])

# Later: Convert sets to strings
emp['notes'] = ' | '.join(sorted(emp['notes']))
emp['emp_id'] = ' | '.join(sorted(emp['emp_ids']))
```

## 3. Data Quality Detection

### Skill: Pattern-Based Anomaly Detection
**File**: `aggregate_absence.py` → `create_suspicious_sheet()`

**Technique**: Flag employees with data quality issues (5 criteria)

**Skills Demonstrated**:
- Pattern detection (string contains checks)
- Multi-criteria flagging logic
- Rule-based filtering

**Patterns Detected**:
```python
# Flag 1: Multiple IDs
if '|' in str(emp_id):  # "SBI1999 | SBI2041"
    flag_multiple_ids = '⚠ YES'

# Flag 2: Merged names
if '/' in str(name):    # "NAME1 / NAME2"
    flag_merged_name = '⚠ YES'

# Flag 3-5: Notes-based flags
if 'ลาออก' in str(notes):      # Quit
if 'เริ่มใหม่' in str(notes):  # Restart
if 'ย้ายมา' in str(notes):     # Transfer
```

**Result**: 39 problematic records identified automatically (no manual review needed)

## 4. Excel Data Processing

### Skill: Multi-sheet Excel Export with Formatting
**File**: `aggregate_absence.py` → `export_to_excel()`

**Technique**: Write 4 sheets with different formatting per sheet

**Skills Demonstrated**:
- pandas ExcelWriter API
- openpyxl styling (Font, PatternFill, Alignment)
- Per-sheet conditional formatting
- Column width calculation and auto-fit

**Code Pattern**:
```python
from openpyxl.styles import Alignment, PatternFill, Font

with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    # Write sheets
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Post-process with openpyxl
    ws = writer.sheets['Sheet1']

    # Auto-fit columns
    for column in ws.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

    # Conditional formatting
    for cell in ws[1]:
        if cell.value:
            cell.font = Font(bold=True, size=12)
```

### Skill: Avoiding Excel Formula Errors
**Issue**: `===` and `==` are interpreted as Excel formulas → Err:510/Err:509

**Solution**: Use bracket notation `[TEXT]` instead
```python
# ❌ BAD: Causes Excel formula error
summary.append({'Metric': '=== SECTION HEADER ===', 'Value': ''})

# ✅ GOOD: Just text, no formula interpretation
summary.append({'Metric': '[SECTION HEADER]', 'Value': ''})
```

**Formatting Detection**:
```python
if cell.value and '[' in str(cell.value) and ']' in str(cell.value):
    cell.font = Font(bold=True, size=12)
    cell.fill = PatternFill(start_color='D3D3D3', ...)
```

## 5. Merge Audit Trail

### Skill: Tracking Merge Reasons
**File**: `aggregate_absence.py` → `aggregate_yearly_totals()` + `create_merged_names_sheet()`

**Technique**: Record why each employee merge happened and show audit trail

**Skills Demonstrated**:
- Merge reason tracking during aggregation
- Original name preservation
- Monthly ID progression tracking
- Multi-column audit trail output

**Merge Types Detected**:
```python
# Type 1: ID Change - Same name, different IDs across months
if has_multiple_ids:
    merge_type = 'ID Change'
    # Example: R88066 → SBI2078 (transfer to new system)

# Type 2: Fuzzy 80% - Names matched via similarity algorithm
if has_fuzzy_merge:
    merge_type = 'Fuzzy 80%'
    # Example: อ่ำสาริกา → อ่ำสาริต (typo in lastname)

# Type 3: Name Variation - Same ID, different name spellings
if has_multiple_names:
    merge_type = 'Name Variation'
    # Example: PANHA YON (ยอน) | PANHA YON (ยา) (nickname changed)
```

**Monthly ID Tracking**:
```python
# Track which ID appeared in each month
month_ids = {}
for m_idx, month_data in enumerate(all_months_data):
    found_ids = []
    for monthly_emp in month_data:
        if monthly_emp['emp_id'] in ids_list:
            found_ids.append(monthly_emp['emp_id'])
    month_ids[month_label] = ' | '.join(sorted(found_ids)) if found_ids else '-'
```

**Output**: Sheet with columns: Final Name, Original Names, Merge Type, Jan, Feb, Mar, Apr, May, Jun, Jul

## 6. Summary Statistics & Reporting

### Skill: Executive Summary Generation
**File**: `aggregate_absence.py` → `create_executive_summary()`

**Technique**: Organize complex data into 5 sections for 30-second read

**Skills Demonstrated**:
- Hierarchical data organization
- Percentage calculations
- Comparative analysis
- Business insight derivation

**Structure**:
```
[PERIOD & SCOPE]
  - Data period
  - Employee count
  - Records merged

[WORKFORCE OVERVIEW]
  - Work days total
  - Employees requiring review (with breakdown)

[TOP ABSENCE CATEGORIES]
  - Top 7 absence types
  - Each with total and percentage

[DEPARTMENT CONCENTRATION]
  - Top 5 departments
  - Each with count and percentage

[KEY INSIGHTS]
  - Compliance alerts
  - Risk flags
  - Data quality notes
```

### Skill: Validation & Verification
**File**: `aggregate_absence.py` → `calculate_summary_stats()`

**Technique**: Prove calculations are correct (raw vs summary comparison)

**Code Pattern**:
```python
# Calculate raw totals (before merging)
raw_totals = [0] * 17
for month_data in all_months_data:
    for emp in month_data:
        for i in range(17):
            raw_totals[i] += emp['totals'][i]

# Calculate summary totals (after merging)
aggregated_totals = [0] * 17
for emp in aggregated_data:
    for i in range(17):
        aggregated_totals[i] += emp['totals'][i]

# Verify they match
if raw_total == summary_total:
    match = '✓'  # Data integrity confirmed
else:
    match = '❌ DIFF!'  # Problem detected
```

## 6. Data Structure Design

### Skill: Dictionary-Based Data Modeling
**Technique**: Use Python dicts to represent complex entities

**Employee Record Structure**:
```python
{
    'emp_id': 'SBI2068',                          # Single or pipe-separated
    'name': 'นาง CHHUN ORNG (รี)',               # Display name
    'notes': 'ลาออก 27/03 | เข้า 17/03',        # Combined from all months
    'position': 'พนักงานฝ่ายผลิต:Operator',
    'department': 'PRODUCTION-F',
    'payType': 'รายวัน',
    'totals': [28341, 9, 675, 102.5, 0, ...]    # 17 absence type totals
}
```

**Benefits**:
- Flexible schema (easy to add fields)
- Named access (readable code)
- Serializable to JSON/Excel

## 7. File I/O & Pattern Matching

### Skill: File Discovery
**File**: `aggregate_absence.py` → `find_monthly_files()`

**Technique**: Automatically find files matching pattern

```python
import glob

files = glob.glob('[0-1][0-9].2568.xlsx')  # Matches 01-12.2568.xlsx
return sorted(files)  # Returns: ['01.2568.xlsx', '02.2568.xlsx', ...]
```

**Benefits**:
- No hardcoded filenames
- Handles variable month counts (1-12 files)
- Sorted output for predictable processing

## 8. Error Handling & Data Validation

### Skill: Graceful Missing Data Handling
**File**: `aggregate_absence.py` → `parse_value()`

**Technique**: Convert Excel cell values safely

```python
def parse_value(val):
    if pd.isna(val) or val == '' or str(val).strip() in ['-', ' - ']:
        return 0  # Treat missing as zero
    try:
        return float(val)
    except:
        return 0  # Treat unparseable as zero
```

**Handles**:
- Missing cells (NaN)
- Empty strings
- Dash characters ("-")
- Non-numeric values
- Multi-format inputs

## Summary: Key Technical Competencies

| Skill | Technique | File Location |
|-------|-----------|---------------|
| **Regex & String Parsing** | Extract components from complex names | `extract_name_key_and_notes()` |
| **Fuzzy Matching** | Detect similar names despite typos | `similarity_ratio()` + `find_fuzzy_match()` |
| **Data Aggregation** | Merge 1,214 records → 240 unique | `aggregate_yearly_totals()` |
| **Merge Audit Trail** | Track merge reasons + monthly ID progression | `create_merged_names_sheet()` |
| **Anomaly Detection** | Flag 39 problematic records | `create_suspicious_sheet()` |
| **Excel Processing** | Multi-sheet export with formatting | `export_to_excel()` |
| **Error Prevention** | Use brackets not equals for headers | Formula error avoidance |
| **Data Validation** | Raw vs summary totals verification | `calculate_summary_stats()` |
| **Executive Reporting** | 5-section CEO summary in 30 seconds | `create_executive_summary()` |

---

**Difficulty Level**: Intermediate → Advanced
**Data Volume**: 1,214 records → 240 clean records
**Quality Gate**: 39 problematic records flagged
**Validation**: 100% total verification
