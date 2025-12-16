# Absence Data Aggregation Tool

Automated Python script to consolidate monthly absence/attendance records from 12 Excel files into a comprehensive yearly summary with CEO-level insights, quality control flags, and detailed employee data.

## Quick Start

### Prerequisites
```bash
pip install pandas openpyxl
```

### Usage
```bash
python3 aggregate_absence.py
```

**Input:** 12 monthly Excel files (01.2568.xlsx through 12.2568.xlsx)
**Output:** absence-summary-2568.xlsx with 5 sheets

## Output File Structure

The generated Excel file contains **5 sheets** (in this order):

### 1. Executive Summary (First Page - CEO View)
**30-second overview of the complete situation**

Shows:
- **Period & Scope**: Data range, employee count, record merges
- **Workforce Overview**: Total work days, employees requiring review (broken down by reason)
- **Top Absence Categories**: The 7 most common absence types with percentages
- **Department Concentration**: Top 5 departments by headcount
- **Key Insights**: Business-level alerts and compliance indicators

**Purpose**: CEO/Manager opens file → reads this sheet → understands the situation immediately

**Design**: One screen, no scrolling, clean gray section headers, left-aligned text for readability

### 2. Suspicious
**39 employee records flagged for manual review**

Shows individual employees with these issues:
- **Multiple IDs?** → Employee had job changes (verify all months counted correctly)
- **Merged Name?** → System detected incomplete name merges (verify it's one person)
- **Quit (ลาออก)?** → Employee quit mid-year (data only includes working months)
- **Restart (เริ่มใหม่)?** → Employee restarted after break (partial-year data)
- **Transfer (ย้ายมา)?** → Transferred from old system (ID may have changed)

Plus: Employee ID, Name, and full Notes

**Purpose**: HR/Payroll reviews the 39 flagged cases (reduces manual review from 240 to 39)

**Design**: Bold red text for "⚠ YES" flags, easy to spot issues, large readable font

### 3. Merged Names (Merge Audit Trail)
**21 employees with merged records - shows exactly what was combined**

Shows one row per merged employee with:
- **Final Name**: The name we settled on after merging
- **Original Names**: All name variations that were merged (e.g., "นาย ZAW MYO (ซอ) | นาย ZAW MYO (ยง)")
- **Merge Type**: Why the merge happened
  - `ID Change` - Same name, different employee IDs across months
  - `Fuzzy 80%` - Names matched via 80% similarity algorithm
  - `Name Variation` - Same ID, different name spellings
- **Jan-Jul columns**: Which employee ID appeared in each month (shows ID progression)

**Purpose**: Audit trail for HR to verify merges are correct. Helps catch incorrect merges (e.g., two different people merged as one)

**Example row**:
| Final Name | Original Names | Type | Jan | Feb | Mar | Apr | May | Jun | Jul |
|---|---|---|---|---|---|---|---|---|---|
| นาย ZAW MYO (ซอ) | นาย ZAW MYO (ซอ) \| นาย ZAW MYO (ยง) | ID Change | R88016 \| R88041 | R88016 \| R88041 | R88041 \| SBI2061 | SBI2061 \| SBI2077 | ... | ... | ... |

**Design**: One row per merged person, monthly columns show ID changes over time

### 4. Summary
**Data validation sheet - proves calculation correctness**

Shows:
- Metadata: Record counts before/after merge, data period
- Comparison table: Raw file totals vs. Summary totals for all 17 absence types
- Verification: ✓ mark if numbers match (confirms no data loss during merging)

**Purpose**: Technical validation that aggregation is mathematically correct

**Design**: Two-column comparison for easy verification

### 5. Employees
**Complete detailed dataset - 240 rows × 23 columns**

Shows every employee with:
- Employee ID (may have multiple if they had job changes)
- Name and Notes (includes quit dates, transfer info, etc.)
- Position, Department, Pay Type
- Yearly totals for all 17 absence types

**Purpose**: Detailed lookup - find a specific person's data

**Design**: Center-aligned, auto-fit columns, sorted by name

## Data Processing Overview

### Input Data Structure
12 monthly Excel files (01.2568.xlsx - 12.2568.xlsx) with:
- 161+ employees per month
- 2 halves per month (days 1-15 and 16-31)
- 17 absence/work types per employee
- Employee info: ID, Name (with nickname), Position, Department, Pay Type

### Processing Steps

**1. Load & Extract** (per monthly file)
- Read Excel file, skip header rows
- Extract: Employee ID, Name, Position, Department, Pay Type
- Parse name into components (prefix, firstname, lastname, nickname)
- Extract notes (quit dates, transfers, etc.)
- Combine first-half and second-half month data into monthly totals

**2. Aggregate** (across all months)
- Use fuzzy name matching to detect duplicate employees
  - Match by: prefix + firstname + lastname (ignore nickname)
  - Allow lastname typos at 80% similarity threshold
- Merge duplicate records:
  - Keep track of all employee IDs (if job changes occurred)
  - Combine all notes from different months
  - Sum absence totals across all months
- Result: 1,214 raw monthly records → 240 unique employees

**3. Validate**
- Calculate raw file totals (before merging)
- Calculate summary totals (after merging)
- Verify they match (proves no data loss)

**4. Export**
- Create 5 sheets with different purposes
- Apply formatting for readability
- Auto-fit columns, adjust fonts and colors

## Technical Details

### Fuzzy Matching Logic
When an employee's name appears with variations (typos, spelling differences):

```
Input: "นาง CHHUN ORNG (รี)" and "นาง CHHUN ORNG LY (รี)"
Extract: prefix|firstname|lastname = นาง|CHHUN|ORNG vs นาง|CHHUN|ORNG LY
Match?: ORNG vs ORNG LY = 80% similarity → MATCHED
Result: Merged into one employee record
```

Threshold: 80% similarity on lastname (configured in code)

### 17 Absence Types Tracked
1. วันทำงาน (Work Days)
2. ขาดงาน (Absent/Unexcused)
3. ลากิจ (Personal Leave)
4. ป่วยมีใบรพ. (Sick w/Certificate)
5. ป่วยไม่มีรพ. (Sick w/o Certificate)
6. ลาคลอด (Maternity Leave)
7. ลืมสแกนนิ้ว/มาสาย (Late - Grace Period)
8. มาสายเกิน (Late - Penalty)
9. ลาOT (OT Leave)
10. ให้หยุด/พักงาน (Suspension)
11. พักร้อน (Annual Leave)
12. OT 2.5 ชม (OT 2.5 hours)
13. OT >2.5 ชม (OT >2.5 hours)
14. ทำงานวันหยุด (Holiday Work)
15. OT วันหยุด (Holiday OT)
16. กะดึก (Night Shift)
17. ควบคุม 2 เครื่อง (Control 2 Machines/Meeting)

## Files

**Main Script:**
- `aggregate_absence.py` - Complete aggregation tool (~750 lines)

**Dependencies:**
- pandas - Data processing
- openpyxl - Excel file handling

**Input Files:**
- 01.2568.xlsx through 12.2568.xlsx (monthly data)

**Output:**
- absence-summary-2568.xlsx (generated Excel report)

## Code Functions

### Core Functions

**`find_monthly_files()`**
- Finds all XX.2568.xlsx files in current directory
- Returns sorted list (01 → 12)

**`load_monthly_file(filepath)`**
- Loads Excel file, skips header rows
- Returns pandas DataFrame

**`extract_absence_data(df)`**
- Processes one monthly file
- Extracts employee info and combined absence totals
- Returns list of employee dicts

**`aggregate_yearly_totals(all_months_data)`**
- Applies fuzzy matching across all months
- Merges duplicate employee records
- Sums totals across all 12 months
- Returns list of 240 unique employees with yearly totals

**`calculate_summary_stats(aggregated_data, all_months_data)`**
- Calculates raw vs. summary totals for validation
- Returns list of comparison rows

**`create_suspicious_sheet(df)`**
- Scans 240 employees for problematic patterns
- Flags: multiple IDs, quit records, transfers, restarts, incomplete merges
- Returns DataFrame with 39 flagged records

**`create_executive_summary(aggregated_data, suspicious_df, all_months_data)`**
- Creates CEO overview with 5 sections
- Shows period, workforce overview, top absences, departments, insights
- Returns DataFrame formatted for one-page view

**`create_merged_names_sheet(df, aggregated_data, all_months_data)`**
- Shows all merged employees with audit trail
- Displays original names, merge type, and monthly ID progression
- Returns DataFrame with 21 merged employee records

**`create_output_dataframe(aggregated_data)`**
- Formats employee data for final export
- 23 columns: ID, Name, Notes, Position, Dept, PayType + 17 absence types
- Returns DataFrame with all 240 employees

**`export_to_excel(df, summary_df, aggregated_data, all_months_data, filename)`**
- Writes all 5 sheets to Excel file
- Applies formatting: fonts, colors, column widths, alignment
- Organizes sheets: Executive Summary → Suspicious → Merged Names → Summary → Employees

### Helper Functions

**`parse_value(val)`**
- Converts Excel cell values to numbers
- Handles missing/empty values as 0

**`similarity_ratio(s1, s2)`**
- Calculates string similarity 0-1 (using SequenceMatcher)
- Used for fuzzy matching lastname variations

**`extract_name_key_and_notes(full_name)`**
- Parses Thai/English names with notes
- Example: "นาง CHHUN ORNG (รี)/ลาออก 27/03"
- Returns: matching_key, display_name, notes

**`normalize_name_parts(name_str)`, `find_fuzzy_match(...)`**
- Support fuzzy matching logic

## Common Use Cases

### Case 1: CEO Wants Overview
→ Open absence-summary-2568.xlsx
→ Look at **Executive Summary** sheet
→ Takes 30 seconds

### Case 2: HR Needs to Verify Data
→ Open absence-summary-2568.xlsx
→ Go to **Suspicious** sheet
→ Review 39 flagged records
→ Takes 15-30 minutes

### Case 3: HR Needs to Audit Merges
→ Open absence-summary-2568.xlsx
→ Go to **Merged Names** sheet
→ Check Original Names column for incorrect merges (e.g., two different people merged)
→ Review monthly ID progression to understand when IDs changed
→ Takes 10-20 minutes

### Case 4: Finance Needs Detailed Breakdown
→ Open absence-summary-2568.xlsx
→ Go to **Employees** sheet
→ Filter/sort as needed
→ Takes 10+ minutes

### Case 5: Payroll Needs to Verify Calculations
→ Open absence-summary-2568.xlsx
→ Go to **Summary** sheet
→ Check ✓ marks in "Match?" column
→ Confirms no data loss during aggregation
→ Takes 5 minutes

## Key Features

✅ **Deduplication**: Merges 1,214 monthly records into 240 unique employees
✅ **Fuzzy Matching**: Handles name typos and variations across months
✅ **Merge Audit Trail**: Shows exactly what was merged and why (Merged Names sheet)
✅ **Data Validation**: Raw vs. summary totals comparison proves correctness
✅ **Quality Flags**: Auto-detects 39 problematic cases for manual review
✅ **CEO View**: One-sheet executive summary with business insights
✅ **Readable Formatting**: Large fonts, colors, auto-fit columns for clarity
✅ **Thai/English**: Bilingual headers and content
✅ **Flexible**: Easily processes 01-12 months, extends to more files

## Limitations & Known Issues

- **Name merging**: Two employees might be detected as same person if:
  - Same prefix + firstname + lastname, but very different nicknames
  - Common name variations exceed 80% similarity threshold

- **Multiple months partial data**: Employees with quit/restart dates will have correct totals but only count the months worked (expected behavior, not a bug)

- **Department variations**: If department names change between months, they may appear as two departments (use Notes column to identify)

## Troubleshooting

**No output file generated?**
- Ensure all 12 monthly files (01.2568.xlsx - 12.2568.xlsx) are in current directory
- Check if filenames match pattern exactly

**Numbers don't match expected totals?**
- Check **Summary** sheet for ✓ or ❌ marks
- Raw totals = numbers before any deduplication
- Summary totals = final numbers with merged employees
- These will differ by about 3-5% due to duplicate record elimination

**Employee appears multiple times?**
- Check **Suspicious** sheet for "Multiple IDs?" flag
- This indicates job changes (different pay periods)
- Review Notes column for quit/transfer information

**Names look wrong?**
- Check **Suspicious** sheet for "Merged Name?" flag
- Look for "/" character indicating incomplete merge
- Review Notes column for context

## Future Enhancements

- [ ] Add charts/visualizations to Executive Summary
- [ ] Support for 12+ months of data
- [ ] Department-level summaries
- [ ] Year-over-year comparison (2567 vs 2568)
- [ ] Email export option (send sheet to CEO automatically)
- [ ] Dashboard integration (Power BI, Google Sheets)

## Support

For issues or feature requests, check:
1. Ensure input files match format (XX.2568.xlsx)
2. Review **Suspicious** sheet for data quality issues
3. Check **Summary** sheet for calculation validation
4. Verify matching logic in `find_fuzzy_match()` function

---

**Version**: 1.0
**Last Updated**: December 16, 2025
**Data Period**: January - July 2568 (Thai Buddhist Year)
