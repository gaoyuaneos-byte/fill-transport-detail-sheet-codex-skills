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

## Workflow

1. List files in the working directory. Identify the active `交通明细表.xlsx` and all current itinerary PDFs/tables.
2. Read the workbook with `xlsx`; inspect sheet name, `!ref`, header row, existing project values, and whether old project rows need clearing.
3. Extract PDF text with `pdf-parse`. Render or table-extract suspicious PDFs when text is incomplete.
4. If a PDF table is blank or totals only, tell the user exactly what is missing and ask for/paste the ride table. Do not invent rows.
5. Fill rows using the template's existing column widths/styles. Usually clear old data rows and keep rows 1-4.
6. For project metadata, prefer explicit user input; otherwise infer conservatively from existing paths, filenames, destinations, or prior workbook rows and mention the assumption.
7. Sort rows by departure time unless the user asks to preserve source order.
8. Put toll/highway fee rows (`高速费`) in their own amount rows. If the user says no work content for these rows, leave the note blank.
9. If the user provides date-to-work-content mapping, set the note column by trip date. For dates not provided, keep route notes unless asked to blank them.
10. Validate before final response: data row count, continuous serial numbers, amount total against PDF/pasted totals, and representative row samples.

## Reusable Script

Use `scripts/fill_transport_sheet.js` as a helper when practical.

Typical command:

```powershell
node C:\Users\Administrator\.codex\skills\fill-transport-detail-sheet\scripts\fill_transport_sheet.js `
  --source "交通明细表.xlsx" `
  --target "交通明细表_已填.xlsx" `
  --project-no "PAEE2512025" `
  --project-name "杭州倍特" `
  --clear-existing
```

Useful options:
- `--pdf-dir <dir>`: directory containing PDFs; defaults to current working directory
- `--trips-json <file>`: use manually prepared trip records instead of PDF extraction
- `--work-json <file>`: date-to-work-content map, e.g. `{ "2026-01-20": "..." }`
- `--blank-highway-notes`: blank notes for rows whose `feeType` is `高速费`
- `--target <file>`: always write a new output workbook unless the user explicitly asks to overwrite

Trip JSON format:

```json
[
  {
    "dateTime": "2026-01-21 10:27",
    "city": "合肥市",
    "start": "肥东县|畅和家园-西南门",
    "end": "包河区|合肥南站-西进站口",
    "amount": 39.1,
    "feeType": "行程费"
  }
]
```

## Notes And Edge Cases

- Didi PDFs often split words and city names across lines (`特惠快\n车`, `杭州\n市`). Normalize whitespace before parsing.
- Some Hellobike PDFs show a grid but no row text. Confirm by text extraction plus screenshot/table extraction, then request pasted details.
- If the workbook currently contains a prior project's data, clear rows below the header unless the user explicitly asks to append.
- Preserve Chinese place names as provided in start/end columns. For generated route notes, strip prefixes before `|` and use `从A去B`.
- Excel serial dates should use the workbook's date style, commonly `m/d/yy h:mm`.
- Validate sums with cents using numeric totals, not formatted text.
