# science-data-correcting

A [Claude Code](https://claude.ai/claude-code) skill that detects suspicious consecutive numeric sequence matches across Excel spreadsheets — useful for scientific data integrity checks.

## What It Does

Given an Excel file with multiple sheets of experimental data, this skill:

1. Scans every row across all sheets for contiguous numeric sequences
2. Detects **multiset matches** — cases where the same set of values appears in different rows but in a different order (e.g., `[0.9, 1.09, 1.01]` in one row and `[1.01, 1.09, 0.9]` in another)
3. Generates annotated Excel files with color-coded highlights and cell comments pointing to each match
4. Produces a Word report summarizing all findings

## Output Files

| File | Description |
|------|-------------|
| `*_同sheet匹配.xlsx` | Matches found within the same sheet |
| `*_跨sheet匹配.xlsx` | Matches found across different sheets |
| `连续匹配检测报告.docx` | Full Word report with statistics, details, and hotspot analysis |

### Color Scheme

| Color | Match Size |
|-------|-----------|
| Red | 7 consecutive values |
| OrangeRed | 6 consecutive values |
| DarkOrange | 5 consecutive values |
| Gold | 4 consecutive values |
| LightYellow | 3 consecutive values |

Each annotated cell includes a hover comment describing exactly which sheet/row/column it matches with.

## Usage

In Claude Code, run:

```
/science-data-correcting path/to/your/data.xlsx
```

## Requirements

- Python 3
- `openpyxl` — Excel read/write
- `python-docx` — Word report generation

```bash
pip install openpyxl python-docx
```

## Detection Logic

- Extracts contiguous numeric runs from each row (breaks on empty/non-numeric cells)
- Generates sliding windows of size 3-12
- Filters trivial matches (all-same values, low-diversity sequences)
- Groups by sorted value tuple (multiset key) to find cross-row matches
- For each row pair, keeps only the longest match
- Focuses on **different-order** matches — same values arranged differently across rows

## License

MIT
