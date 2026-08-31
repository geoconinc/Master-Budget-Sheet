# Master Budget Sheet

Python script that turns a BST **Project Summary** Excel export into Geocon’s master project budget workbook (`all_geocon_budgets.xlsx`) and syncs it to SharePoint.

## What it does

1. Downloads the existing master file from SharePoint
2. Reads the local BST export: `C:\GeoconBST_Summary\Project Summary.xlsx` (first sheet)
3. Cleans and filters project rows
4. Rebuilds `all_geocon_budgets.xlsx` with formatted columns and % Complete color coding
5. Uploads the updated master back to SharePoint
6. Deletes the local temporary master file

## Cleaning rules

- Fills missing numeric fields (`Revenue % Complete`, `Backlog`, `TD Budget Effort`, etc.) with `0`
- Drops rows where `Project Code` starts with `*`
- Drops blank / null project codes
- Keeps only the fields needed for the master sheet
- Sorts remaining rows by Project Code

## Output layout (`all_geocon_budgets.xlsx`)

| Row | Content |
|-----|---------|
| 1 | Color legend for % Complete |
| 2 | Column headers |
| 3+ | Project data |

| Output column | Source field |
|---------------|--------------|
| Project Number | Project Code |
| % Complete | Revenue % Complete (rounded, color-coded) |
| $ Left | Backlog |
| Total Budget | TD Budget Effort |
| Project Name | Project Name |
| Client | Project Client Name |
| Project Manager | Project Manager Name |
| Project Director | Project Director Name |
| Organization | Project Organization Name |

**% Complete colors**

| Range | Color |
|-------|--------|
| 50%–74% | Light yellow |
| 75%–89% | Moccasin |
| 90%–100% | Light coral |
| >100% | Thistle |

## SharePoint location

- Site: `https://geoconmail.sharepoint.com/sites/GeoconCentral`
- Folder: `Shared Documents/GeoconDocuments/ProjectNumbers`
- File: `all_geocon_budgets.xlsx`

## Setup

### Dependencies

```bash
pip install pandas openpyxl Office365-REST-Python-Client
```

### Credentials

Set these environment variables before running (do not hardcode passwords in the script):

```bash
# Windows (PowerShell)
$env:SHAREPOINT_EMAIL = "you@geoconinc.com"
$env:SHAREPOINT_PASSWORD = "your-password"

# Windows (Command Prompt)
set SHAREPOINT_EMAIL=you@geoconinc.com
set SHAREPOINT_PASSWORD=your-password
```

### Input file

Place the BST export at:

```text
C:\GeoconBST_Summary\Project Summary.xlsx
```

## Run

```bash
python update_budget_sheet.py
```

Or call `run_update_process()` from another script.

## Notes for GMS

The master workbook on SharePoint is the all-projects budget sheet. GMS can surface that file (or a refreshed copy of it) so everyone can view current project budgets in one place.
