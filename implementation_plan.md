# Technical Specification & Implementation Guide

> **Status**: Completed & Verified (2026-01-02)
> **Context**: This document describes the architecture of the **Weekly Report Automation** tool. It is designed to be read by AI Assistants (Cursor/Copilot) or Developers to understand how to maintain or extend the logic.

## 1. Project Overview
**Goal**: Automate the generation of a Weekly Status Report for IFS/MNA projects.
**Input**: `NA Daily work.xlsx` (Source Data)
**Template**: `alphast_SNOW_report_12_19_25.xlsx` (Style/Layout Source)
**Output**: `Weekly_Report_YYYY-MM-DD.xlsx` (Final Artifact)

## 2. Core Architecture
The solution is a single-script Python automation (`generate_report.py`) wrapped in a batch file (`run_report.bat`).

### Data Flow
1.  **Clone Template**: Copy `alphast_SNOW_report...xlsx` to `Weekly_Report_[EndDate].xlsx`.
2.  **Clean Output**:
    *   **Partial Cleaning**: Wipes Columns H-Z (Data Area) from Row 3+.
    *   **Dashboard Protection**: Wipes Columns A-G only from Row 9+ (preserving Header/Stats at Top).
    *   **Object Removal**: Deletes Excel `ListObjects` (Tables) to prevent XML corruption during row shifting.
3.  **Read Source**: Loads `NA Daily work.xlsx` (Tabs: `2026` or `2025`, and `New Users`).
4.  **Filter & Transform**: Applies business logic to select relevant tickets.
5.  **Write Data**: Appends data to the output file, verifying column mappings.
6.  **Calculate Stats**: Counts "Closed/Open" for the dashboard header.
7.  **Final Polish**: Applies visual styles (Borders/Fonts) and adds a standard AutoFilter.

## 3. Critical Business Logic (The "Why")

### A. The "Backlog Filter" (Tickets)
*User Requirement*: "Show tickets from Last Week OR any Open tickets."
**Logic Implementation**:
```python
# 1. Created Last Week
mask_week = (df['Date Created'] >= start_date) & (df['Date Created'] <= end_date)

# 2. Status is Open (Regardless of creation date, excluding Future)
mask_open = (df['Status'] == "open") & (df['Date Created'] <= end_date)

# Final Filter
filtered_tickets = df[mask_week | mask_open]
```

### B. The "Anti-Corruption" Strategy (Excel)
*Problem*: Manipulating rows inside an existing Excel Table (`ListObject`) caused "Repaired Records" errors.
*Solution*:
*   **Delete Tables**: The script explicitly deletes `ws.tables` from the output.
*   **Mimic Styles**: It uses `clean_and_get_styles()` to memorize the font/border of `Row 3` and applies it to all new rows.
*   **Standard Filter**: It applies `ws.auto_filter.ref` instead of a Table Table.

## 4. Configuration & Mappings

### Column Mapping (`generate_report.py`)
To change which source column maps to which target column, edit `map_tickets`:
```python
map_tickets = {
    8: 'Ticket No.',     # Target Col 8 (H) <- Source 'Ticket No.' (Col I)
    9: 'REQ No.',        # Target Col 9 (I)
    10: 'Category',      # Target Col 10 (J)
    11: 'Description',   # Target Col 11 (K)
    15: 'Contact',       # Target Col 15 (O)
    16: 'Team member',   # Target Col 16 (P)
}
```

### Dashboard Headers
Static headers are force-written to ensure integrity:
*   `C3`: "IFS AMS Workload Summary"
*   `D5`: "For NAFTA Marelli USA"
*   `D7`: Dynamic Period String + Stats (e.g., "Period: ... 37 closed, 13 open")

## 5. Maintenance Guide (For Cursor/AI)

### How to Modify...
*   **Add a Column**:
    1.  Update `map_tickets` dictionary.
    2.  Ensure the Template has space (or update `clean_and_get_styles` range `1, 27`).
*   **Change Filter Logic**:
    1.  Go to `generate_report.py` -> `Processing Tickets` section.
    2.  Modify `mask_week`, `mask_open`, or their combination.
*   **Fix "White Background"**:
    1.  Check `clean_and_get_styles`. Ensure `clean_col_start` is set to `8` (H) for Tickets sheet to protect Dashboard (A-G).

## 6. Files
*   `generate_report.py`: Main Logic.
*   `run_report.bat`: Execution Wrapper.
*   `inspect_*.py`: Debugging utilities (safe to delete).
