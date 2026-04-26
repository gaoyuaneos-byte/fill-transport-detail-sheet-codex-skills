---
name: fill-transport-detail-sheet
description: Fill Chinese reimbursement transport detail Excel sheets from Didi/Hellobike ride itinerary PDFs or pasted ride tables. Use when the user asks to put taxi, Didi, ride-hailing, 顺风车, 哈啰出行, or 打车行程单 records into a 交通明细表.xlsx, preserve the sheet format, set project number/name, add route/work notes, handle toll/highway-fee rows, and verify row counts and totals.
---

# Fill Transport Detail Sheet

## Goal

Turn ride itinerary sources into a completed `交通明细表.xlsx`-style workbook.

Use this workflow for:
- Didi PDFs named like `滴滴出行行程报销单...pdf` or `滴滴顺风车行程单...pdf`
- Hellobike/哈啰 pasted tables or PDFs with missing text
- Excel templates whose data columns are: `序号, 项目号, 项目名称, 日期, 城市, 起点, 终点, 金额(RMB), 说明（工作内容）`

## Quick Start

```powershell
node scripts/fill_transport_sheet.js `
  --source "交通明细表.xlsx" `
  --target "交通明细表_已填.xlsx" `
  --project-no "PAEEXXXX" `
  --project-name "XX项目" `
  --clear-existing
```

## Workflow

1. List files in the working directory. Identify the active `交通明细表.xlsx` and all current itinerary PDFs/tables.
2. Read the workbook with `xlsx`; the script auto-detects the header row and data start position.
3. Extract PDF text with `pdf-parse`. The script auto-detects Didi vs Hellobike format by filename and content.
4. If a PDF table is blank or totals only, the script warns and continues. Tell the user exactly what is missing and ask for/paste the ride table. Do not invent rows.
5. Fill rows using the template's existing column widths/styles from the detected data row.
6. For project metadata, prefer explicit user input; otherwise infer conservatively from existing paths, filenames, destinations, or prior workbook rows and mention the assumption.
7. Sort rows by departure time unless the user asks to preserve source order.
8. Put toll/highway fee rows (`高速费`) in their own amount rows. If the user says no work content for these rows, use `--blank-highway-notes`.
9. If the user provides date-to-work-content mapping, set the note column by trip date. For dates not provided, keep route notes unless asked to blank them.
10. Validate before final response: data row count, continuous serial numbers, amount total against PDF/pasted totals, and representative row samples.

## CLI Options

| Flag | Description | Default |
|------|-------------|---------|
| `--source` | Template workbook path | `交通明细表.xlsx` |
| `--target` | Output workbook path | `交通明细表_已填.xlsx` |
| `--project-no` | Project number **(required)** | — |
| `--project-name` | Project name **(required)** | — |
| `--pdf-dir` | Directory containing itinerary PDFs | current directory |
| `--source-format` | Force parser: `didi`, `hellobike`, or `auto` | `auto` |
| `--trips-json` | Use manually prepared trip records instead of PDFs | — |
| `--work-json` | Date-to-work-content map: `{"2026-01-20":"..."}` | — |
| `--clear-existing` | Clear old data rows before writing | `false` |
| `--blank-highway-notes` | Leave notes blank for 高速费 rows | `false` |

## Trip JSON Format

```json
[
  {
    "dateTime": "2026-01-21 10:27",
    "city": "XX市",
    "start": "XX区|XX地名",
    "end": "XX县|XX站-进站口",
    "amount": 39.10,
    "feeType": "行程费"
  }
]
```

Use `"feeType": "高速费"` for toll/highway rows.

## Parsers

### Didi (滴滴) Parser

Handles PDFs containing: 快车, 特惠快车, 专车, 优享, 出租车, 顺风车, 独享.

Address splitting uses `区|`/`县|`/`路|` pipe-delimited district markers to separate start/end locations. Falls back to midpoint splitting when markers are absent.

### Hellobike (哈啰) Parser

Handles tab-separated or multi-space-separated pasted tables with `YYYY-MM-DD HH:mm` date-time format. Header lines containing 日期/起点/终点 are auto-skipped.

### Auto-Detection

Format is detected by filename keywords (`滴滴`/`didi` → Didi, `哈啰`/`hellobike`/`顺风车` → Hellobike) and content patterns. Override with `--source-format`.

## Notes And Edge Cases

- Didi PDFs often split words and city names across lines (`特惠快\n车`, `XX\n市`). The script normalizes known service types plus any `XX 市` → `XX市` pattern generically.
- Some Hellobike PDFs show a grid but no row text. The script warns on zero-result extractions. Confirm by screenshot/table extraction, then paste details via `--trips-json`.
- Header row is auto-detected by scanning for `序号, 项目号, 日期, 起点, 终点, 金额` column headers. Falls back to row 4 if detection is ambiguous.
- If the workbook currently contains a prior project's data, the script clears rows below the header unless you omit `--clear-existing`.
- Preserve Chinese place names as provided in start/end columns. For generated route notes, strip `区|`/`县|` prefixes and use `从A去B`.
- Excel serial dates use the workbook's date style from the template row, commonly `m/d/yy h:mm`.
- Validate sums with cents using numeric totals, not formatted text.

## Dependencies

```bash
npm install pdf-parse xlsx
```
