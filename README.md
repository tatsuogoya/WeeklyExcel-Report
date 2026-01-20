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
- Required columns:

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
- open_count: Status = OPEN
- closed_count: Status = CLOSE
- All sheets included

---

### Weekly Activity (Left Section)
- Date is within period
- Status = OPEN only
- Columns: Ticket #, Description, Remarks
- Sorted by Received → Ticket No.

---

### Weekly Workload Details (Right Section)
- Date is within period
- Status = OPEN or CLOSED
- Columns: Ticket #, REQ No., Type, Description, Requested for, PIC, Received, Resolved

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
