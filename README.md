# B3 Jira Excel Reconciliation Tool

A Python utility that reconciles a **Jira issue export** against a **master tracking sheet** (the "B3 sheet"), automatically detecting issues that exist in Jira but are missing from the master, and copying them across into a dated output workbook.

## What it does

`finished_b3_excel_code.py` performs a one-way sync from a Jira export into a master sheet:

1. **Loads three Excel files** — the Jira export, the master (B3) sheet, and an output workbook (created automatically if it doesn't exist).
2. **Creates a date-stamped worksheet** (e.g. `19-07-2026 21.15`) at the front of the output workbook so each run is timestamped and preserved.
3. **Maps columns by header name** rather than fixed positions, so the script keeps working even if column order changes between exports.
4. **Diffs by `Issue key`** — builds the list of Jira issue keys not already present in the master sheet.
5. **Copies the missing rows** (all mapped fields) into both the dated output sheet and the master sheet, tagging each new row with the run date in column 1.
6. **Saves** the output workbook.

## Fields tracked

The script maps and transfers the following Jira columns:

- `Issue key`, `Issue id`, `Issue Type`
- `Updated`, `Status`, `Priority`, `Resolution`
- `Summary`
- `Custom field (Responsible Team)`
- `Custom field (Preventive Action Category)`
- `Project key`, `Project name`, `Project type`, `Project lead`, `Project description`, `Project url`
- `Assignee`, `Reporter`, `Creator`

## Requirements

- Python 3.x
- [`openpyxl`](https://pypi.org/project/openpyxl/)

```bash
pip install openpyxl
```

## Configuration

Edit the three file paths near the top of the script before running:

```python
jira_file_path   = r'C:\path\to\Jira.xlsx'                # Jira export (source)
master_file_path = r'C:\path\to\b3_sheet.xlsx'            # Master tracking sheet
output_file_path = r'C:\path\to\summa_output_file.xlsx'   # Output (auto-created)
```

> The Jira export and master sheet must contain a header row (row 1) with the field names listed above. Matching is done on `Issue key`.

## Usage

```bash
python finished_b3_excel_code.py
```

On each run the script prints the detected header indexes and the list of new issues found, then writes them into a new dated sheet in the output workbook.

## How it works (technical notes)

- **Header-index resolution:** two dictionaries (`Jira_header_index`, `master_header_index`) are populated by scanning row 1 of each sheet and recording the 1-based column index of each expected header.
- **Set difference:** issue keys from the Jira column are compared against the master column; any key in Jira but not in the master is queued for copying (`not_in_list_b`).
- **Timestamping:** every appended row is stamped with the current date/time in column 1, and each run gets its own worksheet, giving a simple audit trail of what was added and when.

## Notes / limitations

- Paths are currently hard-coded for a Windows environment — adjust for your OS.
- The sync is one-directional (Jira → master/output); it adds new issues but does not update or remove existing rows.
- Duplicate detection is based solely on `Issue key`.

---

*Utility script for automating preventive-action / issue tracking reconciliation between Jira exports and a master spreadsheet.*
