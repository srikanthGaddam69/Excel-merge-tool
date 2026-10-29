# Excel Merge Tool

A single-file, fully offline browser tool that copies data from one Excel workbook into another by matching a **key column** (for example `Reqt_ID`), then either **replaces** or **appends to** the target cells — while keeping the original workbook's formatting exactly as it was.

Typical use: you own a master tracker (a requirements matrix, an asset register, a test log) and colleagues send back their own spreadsheets with updates against the same IDs. Instead of copy-pasting row by row, load both files, map the columns once, and download the updated master.

![Screenshot](docs/screenshot.png)

## Features

- **Match by key column** — any column in either file; spaces trimmed, optional ignore-case.
- **Replace or Concat per column** — append mode uses a separator you choose (default ` | `).
- **Blank protection** — an empty cell in the shared file never wipes existing data.
- **Combine duplicate IDs** — if the shared file lists the same ID on several rows, their values are joined (blanks and repeats skipped) instead of keeping only the last row. On by default, with its own separator; works with Replace and Concat rules.
- **Formatting preserved** — only the changed cells are rewritten inside the `.xlsx` package. Fills, fonts, borders, column widths, freeze panes, filters, data validation, other sheets, charts and pivots are untouched.
- **Formulas recalculate on open** — dependent status formulas pick up the new values automatically.
- **Excel's 32,767-character cell limit** handled — over-long text is trimmed and listed.
- **Change log** — download a CSV of every cell changed (key, column, cell, old value, new value) for your audit trail.
- **Reusable rule sets** — save your column mappings to `.json` and load them next time.
- **Preview before writing** — see matched/unmatched rows, keys only in the shared file, and how many IDs are duplicated in the shared file.
- **Private and offline** — no install, no server, no internet; data never leaves your computer.

## Quick start

1. Download [`dist/excel-merge-tool.html`](dist/excel-merge-tool.html).
2. Double-click it to open in Chrome, Edge or Firefox.
3. Load your original workbook and the shared workbook → choose sheets and key columns → **Preview matches** → add column rules → **Process & download**.

A step-by-step guide is in [`docs/User-Guide.pdf`](docs/User-Guide.pdf). Try it first with the demo files in [`samples/`](samples/).

## How it decides what to change

| Situation | Result |
|---|---|
| Key found, shared cell has a value, **Replace** | Target cell gets the shared value |
| Key found, shared cell has a value, **Concat** | `existing + separator + shared` (or just the shared value if the target was empty) |
| Key found, shared cell is blank | Original value kept |
| Key not in shared file | Row left unchanged |
| Key only in shared file | Ignored (listed in the preview) |
| Same key twice in the original | Every copy is updated |
| Same key on several rows of the shared file, **Combine on** (default) | Non-blank values joined in sheet order with the combine separator; repeated values dropped. With Concat the joined text is appended after the existing value |
| Same key on several rows of the shared file, **Combine off** | Last row wins |
| Plain number in the shared file (Replace) | Written as a number |

## Supported files

| Original workbook | Result |
|---|---|
| `.xlsx`, `.xlsm` | Values merged, **formatting and macros kept** |
| `.xls`, `.csv` | Values merged and saved as `.xlsx`; formatting is **not** kept — save as `.xlsx` first |

The shared workbook can be `.xlsx`, `.xlsm`, `.xls` or `.csv`.

## Known limitations

- Static cell colours do not change with the value (a red "Open" cell stays red after it becomes "Closed"). Use Excel conditional formatting if colours should follow the status.
- If a target cell contained a formula, the formula is replaced by the merged value (reported in the summary).
- A target cell that anchors a *shared* formula is skipped and reported, to avoid corrupting neighbouring formulas.
- Pivot tables keep their last values until you click **Refresh All** in Excel.
- Header must be a single row; the header row number is configurable.

## Repository layout

```
dist/excel-merge-tool.html   ← the tool (single offline file, ready to use)
src/                         ← source: index.html, styles.css, app.js, merge-core.js
vendor/                      ← SheetJS and JSZip libraries + their licences
samples/                     ← demo workbooks and the script that generates them
tests/run-tests.js           ← automated tests of the merge engine
docs/                        ← PDF user guide and screenshot
build.py                     ← inlines src + vendor into dist/
```

## Building and testing

Requires Python 3 (build) and Node.js 18+ (tests).

```bash
python3 build.py      # rebuilds dist/excel-merge-tool.html
npm install           # installs jszip and @xmldom/xmldom for the tests
npm test              # runs 19 tests of the merge engine
```

Tests cover replace/concat/blank retention, formatting preservation (styles, freeze panes, filters, validation, other sheets), the 32,767-character limit, duplicate keys, case-insensitive matching, number types, non-first header rows, illegal XML characters, formula/calcChain handling, sheet names with special characters, a 5,000-row workbook, the change log, and combining duplicate IDs (Replace and Concat, blanks, repeats, numbers, separator, ignore-case).

## Security note

The bundled SheetJS is the last version on npm (**0.18.5**), which has two published advisories (CVE-2023-30533 prototype pollution, CVE-2024-22363 ReDoS) triggered by **specially crafted** files. Risk is low for workbooks you and your colleagues create, but to remove it:

1. Download `xlsx.full.min.js` version 0.20.3 or later from <https://cdn.sheetjs.com/>.
2. Replace `vendor/xlsx.full.min.js` with it.
3. Run `python3 build.py` and `npm test`.

## Licence

MIT — see [LICENSE](LICENSE). Bundled third-party libraries keep their own licences — see [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
