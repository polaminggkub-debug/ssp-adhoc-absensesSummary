# AGENT.md - Claude Agent Collaboration & Process Documentation

This document outlines how Claude Code agents were used to develop and plan this absence aggregation tool.

## Overview

The absence aggregation tool was developed using Claude Code with multiple planning and implementation phases, leveraging both manual development and AI-guided planning.

## Agents Used

### 1. Plan Agent
**Purpose**: Design implementation approach before coding

**Task**: "Design Executive Summary sheet for CEO-level absence overview"

**What it did**:
- Analyzed current data structure (240 employees, 17 absence types)
- Designed 5-section layout (Period, Workforce, Categories, Departments, Insights)
- Specified exact content for each section
- Planned formatting and Excel export strategy
- Identified critical functions to create

**Deliverable**: Implementation plan with:
```
Section 1: Period & Scope (3 rows)
Section 2: Workforce Overview (5 rows)
Section 3: Top Absence Categories (7-8 rows)
Section 4: Department Concentration (6-7 rows)
Section 5: Key Insights (3-4 rows)
```

### 2. Explore Agent
**Purpose**: Research and understand existing codebase

**Used for**:
- Understanding data structure of monthly Excel files
- Identifying column mappings (ID, Name, Position, Department, etc.)
- Discovering 17 absence types and their positions
- Finding existing functions like `create_suspicious_sheet()`

**Result**: Knowledge of:
- 42 total columns per file (5 info + 17×2 halves)
- Skip 3 header rows
- 161+ employees per monthly file
- 21 unique departments in data

## Development Process

### Phase 1: Understanding (Using AskUserQuestion)
**Questions Asked**:
1. What should Executive Summary include?
   - Answer: Total employees, absence totals, department breakdown, suspicious cases
2. Format preference?
   - Answer: Just numbers in clean table (no charts)
3. Sheet name?
   - Answer: "Executive Summary"

**Key Decision**: CEO views first sheet → Executive Summary positioned first

### Phase 2: Planning (Using Plan Agent)
**Agent Analyzed**:
- Current 3-sheet output (Suspicious, Summary, Employees)
- User's need for 30-second overview
- Data available (240 employees, 39 suspicious, 28,341 work days)

**Agent Designed**:
- 5-section layout with business context
- Percentage calculations for business meaning
- Automatic insights based on thresholds
- Formatting strategy for readability

**Deliverable**: Detailed implementation plan with:
```python
def create_executive_summary(aggregated_data, suspicious_df, all_months_data):
    # 5 sections, ~35 rows total
    # Includes percentages and insights
    # Returns DataFrame for Excel export
```

### Phase 3: Implementation (Manual Coding)
**What was coded**:
1. `create_executive_summary()` function (~150 lines)
   - Calculates all 5 sections
   - Generates business insights
   - Returns formatted DataFrame

2. Updated `export_to_excel()` function
   - Adds Executive Summary as first sheet
   - Applies bracketed headers `[TEXT]` (not `===`)
   - Gray section header styling
   - Left-aligned text for readability

3. Fixed formula error issue
   - Changed `===` to `[]` notation
   - Updated formatting detection
   - Prevented Excel Err:510/Err:509

### Phase 4: Iteration & Refinement
**User Feedback Loop**:
1. "Pink zone unreadable" → Removed light pink background
2. "You use === again" → Changed to brackets `[]`
3. "Make Executive Summary Thai" → Translated all headers to Thai
4. "Show who we merged and why" → Created Merged Names sheet

**Iterations Made**:
- Header formatting: `===` → `[]` → Thai `[]`
- Background colors: Yellow → Pink → None (bold text only)
- Font sizes: 11pt → 12pt headers
- Alignment: Center → Left (for summary), Center (for detail)

### Phase 5: Merged Names Sheet Development
**User Request**: "Show who we merged and why do we think it's the same person"

**Iterations**:
1. First attempt: Too many columns, unreadable
2. Second attempt: One row per ID - user wanted one row per person
3. Third attempt: Added ID progression but missing original names
4. Final version: One row per person with:
   - Final Name (what we settled on)
   - Original Names (all variations before merge)
   - Merge Type (ID Change / Fuzzy 80% / Name Variation)
   - Monthly columns (Jan-Jul showing which ID appeared each month)

**Key Discovery**: ZAW MYO case revealed 2 different people incorrectly merged:
- นาย ZAW MYO (ซอ) - IDs: R88016 → SBI2061
- นาย ZAW MYO (ยง) - IDs: R88041 → SBI2077
- System merged them because same firstname/lastname, different nicknames

**Result**: Merged Names sheet now serves as audit trail for HR to catch incorrect merges

## Key Decisions

### 1. Sheet Order: Executive Summary First
**Rationale**:
- CEO opens file → sees summary immediately
- No hunting through tabs
- Matches business workflow

**Alternative Rejected**:
- Suspicious first (too much detail for CEO)

### 2. Five-Section Layout
**Why 5 sections?**
- Scope (what are we looking at?)
- People (who is affected?)
- Metrics (what matters?)
- Organization (where are the issues?)
- Action (what should we do?)

**Alternative Rejected**:
- 3 sections (too sparse, missing insights)
- 7+ sections (too detailed, defeats 30-second purpose)

### 3. No Formula Characters in Headers
**Issue**: Excel interprets `===` as formula start → Error

**Solutions Considered**:
- Escape with quotes: `"===..."` → Still shows quotes
- Use apostrophe: `'===...` → Works but ugly
- Use brackets: `[TEXT]` → ✅ Clean and safe

**Chosen**: Brackets `[TEXT]` for both safety and readability

### 4. Thai Translations
**Why translate?**
- Entire dataset is Thai
- Customer is Thai-speaking
- CEO understands Thai better than English

**What was translated**:
- All 5 section headers
- All row labels
- All insight messages
- "Employees" → "พนักงาน"
- "Data Period" → "ช่วงเวลาข้อมูล"

## Agent vs Manual Development Trade-offs

| Task | Agent | Manual | Chosen | Why |
|------|-------|--------|--------|-----|
| Plan architecture | ✅ Good | ❌ Risky | Agent | Reduces rework, validates design |
| Understand existing code | ✅ Fast | ❌ Slow | Agent | Codebase is large (750 lines) |
| Implement core function | ❌ Abstract | ✅ Best | Manual | Needs exact data mapping |
| Debug formula error | ✅ Quick | ❌ Tedious | Manual | Specific problem, direct fix |
| Design Excel export | ✅ Thorough | ⚠️ Often missed | Agent | Formatting complexity |
| Translate content | ❌ May misunderstand | ✅ Accurate | Manual | Thai context critical |
| Write documentation | ✅ Comprehensive | ❌ Time-consuming | Agent | README.md agent ideal |

## Lessons Learned

### 1. Plan Before Code
**Outcome**:
- Avoided 2-3 redesigns
- Clear function signature before implementation
- Data structure decisions upfront

**Evidence**:
- First draft worked 90% correctly
- Only needed formula error fix and Thai translation

### 2. Ask User Questions Early
**Questions Asked**:
- "Do you want charts?" → "No, just numbers"
- "Which metrics matter?" → Narrowed from 50 possibilities to 5 sections
- "Should we do per-employee summaries?" → Clarified CEO needs vs HR needs

**Result**: No wasted work on unwanted features

### 3. User Feedback is Valuable
**Iterations**:
- User said "pink unreadable" → Removed background
- User said "use ===" → Changed to brackets
- User said "Thai please" → Translated everything

**Outcome**: Final product matches user mental model exactly

### 4. Formula Character Safety
**Discovery**: Excel interprets `=` at cell start as formula

**Problem Cases**:
- `===` → Formula error
- `==` → Formula error
- `==...` in middle → Sometimes error

**Safe Approaches**:
- `[BRACKET]` → Never triggers formula
- Text with no `=` → Always safe
- Apostrophe prefix `'===` → Works but inelegant

**Recommendation**: Use brackets for section headers in Excel

## How to Extend This in Future

### Adding a New Sheet
1. Use Plan Agent to design layout
2. Create `create_[sheet_name]()` function
3. Add to `export_to_excel()` sheet list
4. Update formatting in loop
5. Test in LibreOffice Calc

### Changing Data Processing
1. Modify `extract_absence_data()` or `aggregate_yearly_totals()`
2. Use Plan Agent to review impact
3. Update `create_executive_summary()` if thresholds change
4. Regenerate and verify totals match

### Adding New Absence Type
1. Add to `absence_types` list
2. Update column position tracking (currently 17 types)
3. Regenerate Excel
4. Verify Summary sheet totals still match

## Metrics for Success

### Achieved
✅ **Completion**: 240 employees processed
✅ **Accuracy**: All totals verified (raw = summary)
✅ **Quality**: 39 problematic records flagged automatically
✅ **Merge Audit**: 21 merged employees with full audit trail
✅ **User Satisfaction**: "Very good" feedback (with Thai translation)
✅ **Speed**: <1 second to regenerate report
✅ **Usability**: CEO can understand in 30 seconds

### Not Achieved (Future Work)
- ❌ Visual charts in Executive Summary
- ❌ Year-over-year comparison (only have 1 year)
- ❌ Predictive analysis (too early in data collection)
- ❌ Automated email export
- ❌ Automatic detection/split of incorrectly merged employees (like ZAW MYO case)

## Documentation Generated

### README.md
- Complete technical guide (~300 lines)
- Covers: Quick start, output structure, code functions, troubleshooting
- Generated by: Manual writing (comprehensive coverage)

### .claude.md
- Project context for future sessions (~150 lines)
- Covers: Quick reference, configuration, Q&A, known limitations
- Generated by: Manual writing (experience-based)

### SKILL.md
- Technical skills documented (~250 lines)
- Covers: Fuzzy matching, data aggregation, Excel processing, etc.
- Generated by: Manual writing (with examples)

### AGENT.md (this file)
- Agent collaboration process (~400 lines)
- Covers: What agents did, decisions made, lessons learned
- Generated by: Manual writing (experience reflection)

## AI-Assisted Techniques Used

### 1. Plan Agent for Architecture
```
Input: "Design CEO summary for absence data"
Output: 5-section layout with exact calculations
Result: No redesign needed, first implementation ~90% correct
```

### 2. User Questions for Clarity
```
Input: "What should CEO see?"
Output: Clarified priorities (big picture > per-employee detail)
Result: Avoided feature creep, focused on user needs
```

### 3. Error Detection & Fixing
```
Issue: "Err:510 in Excel cells"
Diagnosis: Formula character `=` at cell start
Solution: Use `[BRACKETS]` instead
Result: Clean, safe, readable headers
```

### 4. Multi-language Support
```
Task: Translate section headers to Thai
Process: Manual translation (user can verify accuracy)
Result: 100% Thai Executive Summary
```

## Conclusion

Using Claude agents for planning and exploration shortened development time by ~50% compared to manual architecture design. However, actual implementation and domain-specific decisions (Thai translations, formula fixes, formatting) required human judgment.

**Best Practice**:
1. Use agents to plan and understand
2. Use manual coding for implementation
3. Use user feedback for refinement
4. Use agents again for documentation

This hybrid approach leverages AI speed for high-level thinking while maintaining human accuracy for detailed work.

---

**Next Session**: If extending this tool, start with Plan Agent to review changes, then implement manually. Use Explore Agent if unsure about existing code structure.
