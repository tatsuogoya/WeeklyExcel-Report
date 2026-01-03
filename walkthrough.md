# User Guide: Weekly Report Automation

## How to Run
1.  Navigate to your project folder: `C:\Users\tatsu\AI\MyProjects\MNA\Weekly_Report`
2.  Double-click the **`run_report.bat`** file.
3.  A black terminal window will open.

## Inputs Explained

### 1. "Enter Source Tab Name"
- **Default**: `2026`
- **What to type**:
    -   Press **ENTER** to use the default (`2026`).
    -   Type `2025` to run the report using the "2025" tab from your source Excel file.
    -   Type any other tab name if your source file changes.

### 2. "Enter Reference Date"
- **Default**: `Today`
- **What to type**:
    -   Press **ENTER** to generate the report for the **Previous Week** (relative to today).
    -   Type a date (e.g., `2025-12-29`) to pretend "Today is Dec 29th". The tool will then generate the report for the week *before* that date.
    -   *Tip: Use this to regenerate old reports from the past.*

## Output
- The tool creates a file named `Weekly_Report_YYYY-MM-DD.xlsx`.
- It uses the template format perfectly.
- It filters for:
    -   **Tickets**: Created in Last Week **OR** (Status is "Open" AND Created Before/During Week).
    -   **New Users**: Created in Last Week **AND** Category is "Create Account".

## Validation Scenarios (Why is a ticket missing/present?)
- **Case 1: Ticket is CLOSED** (`SCTASK0961884`)
    -   If date is in Last Week, it is INCLUDED.
    -   If date is OLD and Status is Closed, it is EXCLUDED.
- **Case 2: Date is OLD but OPEN** (`SCTASK0938232`)
    -   **INCLUDED**. Since it is still "Open" as of the report date, it appears in the report (Backlog logic).
- **Case 3: Future Open Ticket**
    -   **EXCLUDED**. The tool filters out tickets created *after* the report week end date.

## Troubleshooting
- **"PermissionError" / "File is open"**:
    -   This means the report file is currently open in Excel.
    -   **Fix**: Close output Excel file. Click back on the black terminal window. Press **Enter**. It will retry and succeed!
- **Data Mismatch?**:
    -   The tool now **Cleans the Template** before writing. This removes old "Future Tickets".
    -   It acts like a fresh report every time.
- **Reference Date Tip**:
    -   If today is Jan 2nd and you want the report for Dec 25th -> Type `2026-01-02` (or just Enter for Today).
    -   If you want to re-run an old report for Sept 1st -> Type `2025-09-08` (Last week relative to Sept 8 is Sept 1-7).
    -   **Wait, logic check**: `get_last_week` calculates previous Monday.
    -   If you type `2025-09-01` (Monday), "Last Week" is Aug 25 - Aug 31.
    -   If you want the report FOR the week of Sept 1st, you should enter a date in the *following* week (e.g., Sept 8th).
