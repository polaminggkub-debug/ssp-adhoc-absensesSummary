# Absence Data Aggregation Tool

Aggregates monthly absence Excel files (01-11.2568.xlsx) into a comprehensive yearly summary report with master data matching and full audit trails.

## Status: Production Ready

**Version**: 2.1
**Last Updated**: December 17, 2025
**Data Period**: January - November 2568 (Thai Buddhist Year)

## Quick Start

```bash
pip install pandas openpyxl
python3 main.py
```

**Input Files:**
- `01.2568.xlsx` through `11.2568.xlsx` - Monthly absence data (4 different formats)
- `employee_master.xlsx` - Official employee list (optional but recommended)

**Output:** `absence-summary-2568.xlsx` with 6 sheets

## Key Features

### Multi-Format Support
Handles 4 different Excel column structures:
- **Format A** (months 01-07): 42 columns, two halves to sum
- **Format B** (months 08-09): 58 columns, monthly totals
- **Format C** (month 10): 41 columns, different column order
- **Format D** (month 11): 36 columns, header at row 4

### Smart Matching Logic

**Three-layer matching to handle inconsistent data entry:**

1. **ID + Name similarity (85%+)**: Same ID and similar name = same person
2. **ID + Thai nickname match**: Same ID and same Thai nickname (even if romanized name differs)
3. **Same name, different IDs**: Merges employees who changed IDs over time

**Disabled fuzzy matching** - caused incorrect merges of different employees.

### Thai Nickname Matching
Handles inconsistent name formats across months:
- `นาย เสร็จ` (Thai short) matches `นาย PISET SAY (เสร็จ)` (full with nickname)
- `นาง สนดิน` matches `นาง SAN DIN (สนดิน)`
- Same ID + same Thai nickname = same person

### Prefix Normalization
Thai prefix abbreviations are normalized:
- `น.ส.`, `นส.` → `นางสาว`
- Ensures same person with different prefix formats is matched

### Master Data Integration
- Matches employees to `employee_master.xlsx` for standardized IDs/names
- Handles ID reuse (same ID assigned to different people over time)
- Full audit trail in Master Match sheet

### 17 Absence Types Tracked
1. วันทำงาน (Work Days)
2. ขาดงาน (Absent)
3. ลากิจ (Personal Leave)
4. ป่วยมีใบรพ. (Sick w/Cert)
5. ป่วยไม่มีรพ. (Sick w/o Cert)
6. ลาคลอด (Maternity)
7. ลืมสแกนนิ้ว/มาสาย (Late Grace)
8. มาสายเกิน (Late Penalty)
9. ลาOT (OT Leave)
10. ให้หยุด/พักงาน (Suspension)
11. พักร้อน (Annual Leave)
12. OT 2.5 ชม
13. OT >2.5 ชม
14. ทำงานวันหยุด (Holiday Work)
15. OT วันหยุด (Holiday OT)
16. กะดึก (Night Shift)
17. ควบคุม 2 เครื่อง (Multi-Machine)

## Output: 6 Excel Sheets

### 1. Executive Summary
CEO-level overview (30-second read):
- Period & scope
- Workforce overview
- Top absence categories
- Department concentration (Top 5)
- Key insights

### 2. Suspicious
Flagged records for HR review:
- Multiple IDs (job changes)
- Quit/restart/transfer flags
- Incomplete merges

### 3. Master Match
Employee matching audit trail:
- Match types: ID+Name, Name, UNMATCHED
- For unmatched employees:
  - `ลาออก (Resigned)` - if source data contains resignation keywords
  - `สุดท้าย: เดือน XX` - last month they appeared
  - Any notes from source files

### 4. Merged Names
Audit trail for merged records:
- Shows which records were combined
- Merge types: Same ID, ID Change, Name Variation, Other

### 5. Data Traceback
File-by-file breakdown:
- Totals per file and section (First Half / Second Half)
- Thai column headers matching Employees sheet
- Verifiable against source files

### 6. Employees
Complete detailed data:
- All 17 absence columns (Thai + English headers)
- Master name column for matched employees
- Notes with resignation/transfer info

## Sample Results

```
Files processed: 11
Total records: 1956
Unique employees: 289
Records merged: 1667
Master matched: 172, Unmatched: 117
```

## Manual Verification Checklist

After running, verify these merges are correct:

### Nickname-Based Merges
| ID | Name 1 | Name 2 |
|---|---|---|
| SBI729 | นาย เสร็จ | นาย PISET SAY (เสร็จ) |
| SBI705 | นาง สนดิน | นาง SAN DIN (สนดิน) |
| SBI844 | นายอิน | นายDOEUR IN (อิน) |
| SBI2081 | นาย RAN ROM (รม) | นาย SAROM TONH (รม) |
| SBI2082 | นาง THOO (ทู) | นาง THOU HUM (ทู) |
| SBI1953 | นาย MEAN CHIEV (เทียว) | นาย CHIEV MEAN (เทียว) |

### ID Change Merges
| Name | IDs |
|---|---|
| นาย MIN MIN (เม) | R88006 → SBI2107 |
| นาย HTET WINT (วิว) | R88019 → SBI2101 |
| นาย KYAW NAING (บี) | R88037 → SBI2062 |
| นาย HTET ZAW (ลี) | SBI2148/2149/2151 |
| นางสาว NWAY NWAY (นวย) | SBI2188/2190 |

## Project Structure

```
ssp-adhoc-absensesSummary/
├── main.py                      # Entry point
├── config/
│   └── absence_mapping.py       # 17 absence types & 4 format configs
├── formats/
│   ├── absence_format_01_07.py  # Format A
│   ├── absence_format_08_09.py  # Format B
│   ├── absence_format_10.py     # Format C
│   └── absence_format_11.py     # Format D
├── file_io/
│   └── excel_reader.py          # Excel utilities
├── models/
│   └── employee.py              # Employee model & name parsing
├── services/
│   ├── aggregator.py            # Deduplication & merging
│   ├── master_matcher.py        # Master data matching
│   └── excel_exporter.py        # Excel output (6 sheets)
└── README.md
```

## Key Design Decisions

### Why no fuzzy matching?
Fuzzy matching (75-85% similarity) caused incorrect merges:
- นาง CHO ZIN was merged with นาย WIN TUN and นาย YE KYAW
- Different genders (นาย/นาง) being combined

**Solution**: Only exact ID+Name or nickname-based matches are used.

### Why ID + Name verification?
Employee IDs can be **reused** for different people:
- SBI2183 was นาง CHO ZIN in some months
- Same SBI2183 was นาย YE KYAW PAING in other months

**Solution**: When matching by ID, also verify name similarity >= 85% OR Thai nickname matches.

### Why nickname matching?
Data entry was inconsistent across months:
- Some months: `นาย เสร็จ` (Thai short name)
- Other months: `นาย PISET SAY (เสร็จ)` (full romanized with nickname)

**Solution**: Same ID + same Thai nickname = same person, even if romanized names differ.

### Unmatched Employee Notes
For employees not found in master, Note shows only **facts from source data**:
- `ลาออก (Resigned)` - only if source contains this keyword
- `สุดท้าย: เดือน XX` - last month they appeared
- Source file notes (e.g., salary adjustments)

No assumptions made - just data from the files.

## Dependencies

- Python 3.x
- pandas
- openpyxl

## Troubleshooting

**Duplicate IDs in output?**
- This happens when same ID was assigned to different people
- Check Master Match sheet for match details
- Review Merged Names sheet for merge audit

**Employee not matched to master?**
- Check Master Match sheet Note column for last appearance
- `สุดท้าย: เดือน 07` means last seen in July
- May indicate resignation even without explicit note

**Numbers don't match source files?**
- Check Data Traceback sheet for file-by-file breakdown
- First Half / Second Half sections for Format A/B files
- Compare against source file totals
