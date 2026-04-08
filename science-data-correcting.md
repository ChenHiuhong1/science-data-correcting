---
description: Detect duplicate consecutive numeric sequences (multiset matching) in scientific Excel data. Generates annotated Excel files and Word reports.
allowed-tools: Bash, Read, Write, Edit, Glob, Grep, Agent, TaskCreate, TaskUpdate
---

# Science Data Correcting — Consecutive Sequence Match Detection

Given an Excel data file path, scan all sheets for consecutive numeric sequence duplicates and generate annotated files and reports.

## Input

User provides an Excel file path, e.g.: `$ARGUMENTS`

If no path is provided, use AskUserQuestion to ask.

## Detection Logic

### 1. Extract Numeric Sequences
- Iterate through every row in all sheets
- Extract contiguous numeric cell sequences (break on non-numeric/empty cells)
- Keep only sequences with length >= 3

### 2. Sliding Window Matching
- For each sequence, generate all sliding windows of size 3–12
- Sort values within each window to create a multiset key
- Filter trivial matches: skip all-same values, or windows with only 2 distinct values when size > 4

### 3. Cross-Row Matching
- Find identical multiset keys appearing in different rows
- For each row pair, keep only the longest match
- Additional filtering: size >= 4 requires >= 3 distinct values; size = 3 requires >= 2 distinct values
- Flag whether the order differs (focus on different-order matches)

### 4. Classification
- Within-sheet matches vs. cross-sheet matches
- Severity by match length: 7 (red) > 6 (orange-red) > 5 (dark orange) > 4 (gold) > 3 (light yellow)

## Output Files

All output files are saved in the same directory as the input file.

### Annotated Excel (2 files)
1. `{filename}_within_sheet_matches.xlsx` — only within-sheet matches
2. `{filename}_cross_sheet_matches.xlsx` — only cross-sheet matches

Annotation method:
- Color highlighting: 5-level color scale by match length
- Cell comments: hover to see match details (which sheet/row/column and which values it matches)
- Up to 5 comments per cell; excess matches show a count
- A legend sheet is added at the front of the workbook

Color scheme:
```
Red    FF0000 — 7 consecutive values matched
OrangeRed FF4500 — 6 consecutive values matched
DarkOrange FF8C00 — 5 consecutive values matched
Gold   FFD700 — 4 consecutive values matched
LightYellow FFFF99 — 3 consecutive values matched
```

### Word Report
`match_detection_report.docx`, containing:
1. Overview (parameters, scope)
2. Summary statistics table (match counts by severity, within-sheet vs. cross-sheet)
3. High-priority match details (>= 5 values), each pair shown in a table with both locations and values
4. Medium-priority matches (4 values), grouped by sheet pair
5. Low-priority matches (3 values) statistical overview
6. Hotspot row analysis (top 30 rows by match participation count)
7. Notes and caveats

## Dependencies

- Python: openpyxl, python-docx
- Install if missing: `pip install openpyxl python-docx`

## Notes

- Only different-order matches are flagged (same values in different arrangement)
- 3-value matches may be coincidental in experimental data; focus on >= 5-value matches
- Integer score data (e.g., 0, 3, 7, 10) may match due to scoring scales — interpret with experimental context
