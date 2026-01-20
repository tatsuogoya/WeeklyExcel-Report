# IFS AMS Workload Summary

## Overview

This project generates a **Weekly ServiceNow Report** from daily ticket data
stored in **NA Daily work.xlsx**.

Users upload an Excel file, select a report period,
view the report on a web page, and download it as a **PDF**.

---

## Features (MVP)

- Web-based UI
- Excel upload (`.xlsx`)
- Begin Date / End Date selection
- Weekly report displayed on screen
- PDF download
- Excel export (Legacy support)
- Clear validation and error messages

---

## Tech Stack

- Backend: Python + FastAPI
- Excel Processing: pandas, openpyxl
- PDF: reportlab
- Testing: pytest

---

## Input File

### NA Daily work.xlsx

- Year-based sheets (e.g. `2025`, `2026`)
- Special sheet: `New Users` (optional, shows all records regardless of date filter)
- Required columns (for year sheets):

| Column |
|------|
| Date |
| Ticket No. |
| REQ No. |
| Type |
| Requested for |
| Assign To |
| Request Detail |
| Time - Arrive |
| Time - Close |
| Status |
| Remarks |

---

## Report Logic

### Summary (Header)
- Date range: begin_date ～ end_date
- open_count: All OPEN tickets (any date)
- closed_count: CLOSE tickets with Time - Close within date range
- Processes current year + previous year sheets only

---

### Weekly Activity (Left Section)
- All OPEN tickets (no date filtering)
- Columns: Ticket #, Description
- Sorted by Received → Ticket No.

---

### Weekly Workload Details (Right Section)
- All OPEN tickets (any date) + CLOSE tickets with Time - Close in range
- Columns: Ticket #, Status, REQ No., Type, Description, Requested for, PIC, Received, Resolved
- Sorted by Status DESC (OPEN before CLOSE), then Received ASC

---

### New Users Section (Optional)
- All records from the `New Users` sheet (no date filtering)
- Displays: Date, Ticket #, User, Status, Department
- Appears at the bottom of web view and PDF
- Only shown if `New Users` sheet exists

---

## Output

### Web Screen
- Summary
- Weekly Activity table (Date Range Based)
- Weekly Workload Details table

### PDF
- Same content as web screen
- Business-style layout

---

## Error Handling

- Invalid input → 400
- Excel schema issues → 400
- No matching data → 400
- Unexpected error → 500

---

## Non-Goals

- Monthly reports
- Authentication
- Background processing

---

## Status

This repository is under active development.
The current version targets **Professional Web Dashboard + PDF/Excel output (V2)**.
