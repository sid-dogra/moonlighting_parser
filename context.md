# Moonlighting Schedule Processor - Context

## Overview
This system processes monthly moonlighting schedules for medical residents and fellows. It parses Amion HTML exports and generates trainee moonlighting summaries.

## Web-Based Tool (GitHub Pages)

A browser-based version of the parser is hosted at:
**https://sid-dogra.github.io/moonlighting_parser/**

- Source code: `docs/index.html`
- GitHub repo: https://github.com/sid-dogra/moonlighting_parser
- All processing happens client-side (no server, files stay local)
- Uses SheetJS library to parse Trainees.xlsx in browser
- Same shift mappings as Python version

### Updating the Web Tool
When new shifts are added, update the `RESIDENCY_SHIFT_MAPPINGS` and `CONTRAST_SHIFT_MAPPINGS` objects in `docs/index.html`, then push to GitHub.

### New Academic Year (July)
When fellows rotate (new fellows start, old fellows graduate):
1. Update `Trainees.xlsx` - add new fellows to the Fellows column, remove graduated ones
2. No code changes needed - abbreviation map is built dynamically from the HTML each time
3. New fellows will be automatically recognized as long as they appear in the HTML legend

## Command-Line Workflow: Amion HTML Parser (`amion_parser.py`)

### Input Files
Place two HTML files exported from Amion in a month folder (e.g., `May/`):
- `Radiology Residency Schedule, {M}_1 to {M}_31, {YYYY}.html`
- `Radiology - Contrast Coverage Schedule, {M}_1 to {M}_31, {YYYY}.html`

### Usage
```bash
python amion_parser.py May
```

### Output
- `{Month}/{Month}_Trainee_Summary.txt` - Combined summary of all trainee shifts

## HTML Parsing Details

### Residency Schedule (shifts as columns, dates as rows)
The parser reads the main schedule table where:
- Row 1: Shift headers (column names)
- Subsequent rows: Date in first column, person names in shift columns

### Contrast Coverage Schedule (shifts as rows, dates as columns)
- Header row contains dates
- Each row is a shift type with person abbreviations in date columns
- Abbreviation map built dynamically from HTML legend (no manual updates needed for new people)

### Abbreviation Map (Contrast Schedule)
The HTML contains a legend mapping names to abbreviations (e.g., "Emmy Hu" → "EH").
- **Python (BeautifulSoup)**: Sees combined text "Emmy HuEH" due to malformed HTML
- **Web (DOMParser)**: Sees separate cells, so looks at adjacent cell pairs: `<td>Emmy Hu</td><td>EH</td>`

Both approaches build the same map; they just parse the HTML differently.

## Shift Mappings

### Residency Schedule Shifts (`RESIDENCY_SHIFT_MAPPINGS`)
| Amion Shift Name | Display Name | Hours |
|------------------|--------------|-------|
| ECEP Weekday 5p-11p | ECEP | 6 |
| ECEP Holiday-Weekend 5p-9p | ECEP | 5 |
| Supershift Coverage 5p-9p | Bellevue Supershift | 4 |
| Scanner Coverage - Gram Weekday 5p-9p | Gramercy weekday | 4 |
| Scanner Coverage - HCC Weekday 5p-7p | HCC weekday | 2 |
| Scanner Coverage - 41st St Weekday 7p-9p | 41st St | 2 |
| Scanner Coverage - 53rd St Weekday 5p-9p | 53rd St weekday | 4 |
| Scanner Coverage - CBI Weekday 5p-9p | CBI | 4 |
| Scanner Coverage - Bay Ridge Weekday 6p-9p | Bay Ridge | 4 |
| Scanner Coverage - CC Weekday 5p-8p | Cancer Center weekday | 3 |
| Scanner Coverage - Gram Weekend 9a-5p | Gramercy weekend | 8 |
| Scanner Coverage - Gram Weekend 8a-4p | Gramercy weekend | 8 |
| Scanner Coverage - CBI Weekend 8a-4p | CBI weekend | 8 |
| Scanner Coverage - Bay Ridge Weekend 8a-4p | Bay Ridge | 8 |
| Scanner Coverage - Forest Hills Weekend 8a-4p | Forest Hills | 8 |
| Scanner Coverage - 53rd Weekend 8a-4p | 53rd St weekend | 8 |
| Scanner Coverage - Garden City Weekend 8a-4p | Garden City | 8 |
| Scanner Coverage - Lake Success Weekend 8a-4p | Lake Success | 8 |
| Scanner Coverage - 32nd St Weekend 8a-4p | 32nd St weekend | 8 |
| Scanner Coverage - 41st St Weekend 8a-4p | 41st St weekend | 8 |
| Scanner Coverage - 41st St Weekend 4p-8p | 41st St weekend PM | 4 |
| Scanner Coverage - Greenpoint Weekend 8a-4p | Greenpoint | 8 |
| Scanner Coverage - CC Weekend 8a-5:30p | Cancer Center weekend 8A-5:30PM | 9.5 |
| Scanner Coverage - CC Weekend 8a-4p | Cancer Center weekend 8A-4PM | 8 |
| Scanner Coverage - CC Weekend 8:30a-4p | Cancer Center weekend | 7.5 |

### Contrast Coverage Shifts (`CONTRAST_SHIFT_MAPPINGS`)
| Amion Shift Name | Display Name | Hours |
|------------------|--------------|-------|
| 32nd Early (7a-8a) | 7AM 32nd | 1 |
| CBI Early (7a-8a) | 7AM CBI | 1 |
| 32nd Late (5p-9p) | 32nd Late | 4 |
| 41st Late (5p-7p) | 41st St late | 2 |
| 32nd Sat (8a-4p) | 32nd St weekend | 8 |
| 41st Sat (8a-4p) | 41st St weekend | 8 |
| CBI Sat (8a-4p) | CBI weekend | 8 |
| HCC Sat (9a-5p) | HCC weekend | 8 |
| 53rd (8a-5p) | 53rd St weekend | 9 |
| Gramercy (8a-5p) | Gramercy weekend | 9 |
| Premier Fellow | Premier | 1 |
| FPO Weekend | FPO Weekend | 8 |

## Trainee List (`Trainees.xlsx`)
Format:
- **Column A**: "Residents" header in A1, resident names below (one per row)
- **Column C**: "Fellows" header in C1, fellow names below (one per row)

Only people in this list are included in the summary. Premier Fellow shifts only count for fellows.

## Name Handling
- Names normalized from "Last, First" to "First Last" format
- Name aliases defined for people listed differently in Amion (e.g., "Young Joon Kwon" → "Fred Kwon")

## Data Flow
```
May/
├── Radiology Residency Schedule, 5_1 to 5_31, 2026.html
├── Radiology - Contrast Coverage Schedule, 5_1 to 5_31, 2026.html
        │
        ▼
┌───────────────────────┐
│   amion_parser.py     │
│ - Load trainees       │
│ - Parse residency HTML│
│ - Parse contrast HTML │
│ - Deduplicate shifts  │
│ - Generate summary    │
└───────────────────────┘
        │
        ▼
May/May_Trainee_Summary.txt
```

## Troubleshooting

### Missing Shifts
If a shift appears in Amion but not in the summary:
1. Check if the shift name is in `RESIDENCY_SHIFT_MAPPINGS` or `CONTRAST_SHIFT_MAPPINGS`
2. New shift variants may need to be added (e.g., "Gram Weekend 8a-4p" vs "Gram Weekend 9a-5p")
3. Check if the person is in `Trainees.xlsx`

### Recent Fixes
- **June 2026**: Fixed web parser for Contrast Coverage Schedule
  - Table selection: Web parser now uses same multi-criteria selection as Python (data rows, clean headers, row count)
  - Abbreviation map: Web parser now looks at adjacent cells for name/abbreviation pairs (browser parses HTML differently than Python's BeautifulSoup)
- **May 2026**: Added `Scanner Coverage - Gram Weekend 8a-4p` mapping (was missing, affected Mira Yousef and Yvan James)

## Legacy System (deprecated)
The old `site_excel_generator.py` processed Excel files directly. This is no longer used - HTML parsing via `amion_parser.py` is the current method.
