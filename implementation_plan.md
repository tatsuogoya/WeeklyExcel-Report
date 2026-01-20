# IFS AMS Workload Summary - Implementation Plan

## 1. Purpose

This application generates a **Weekly Report** from daily ServiceNow ticket data
stored in Excel (**NA Daily work.xlsx**).

The output is:
- Displayed on a **web screen**
- Downloadable as a **PDF**

The output is:
1. Displayed on a **professional web dashboard** (with real-time sorting and branding)
2. Downloadable as a **consistent business PDF**
3. Exportable to **Excel** (Legacy template support)

---

## 2. Scope (MVP)

### In Scope
- Upload Excel file (`.xlsx`)
- Input report period (Begin Date / End Date)
- Display weekly report on web page
- Download the same report as PDF
- Validation and clear error handling

### Out of Scope
- Monthly reports
- Authentication / authorization
- Background jobs
- Excel output / macros

---

## 3. High-Level Architecture

---

## 4. API Design

### GET /
- Returns HTML page
- Upload form + date inputs
- Result display area

### POST /generate
- Input:
  - Excel file
  - begin_date
  - end_date
- Output:
  - JSON representing the weekly report

### POST /pdf
- Input:
  - Excel file
  - begin_date
  - end_date
- Output:
  - PDF (`application/pdf`)
- Uses the same business logic as `/generate`

---

## 5. Core Business Logic (Service Layer)

### 5.1 Input Data
- Excel file: `NA Daily work.xlsx`
- Sheets:
  - Year-based (e.g. `2025`, `2026`)

### Required Columns
- Date
- Ticket No.
- REQ No.
- Type
- Requested for
- Assign To
- Request Detail
- Time - Arrive
- Time - Close
- Status
- Remarks

---

## 6. Report Structure

The weekly report contains **three logical parts**:

---

### 6.1 Summary (Period Header)

**Date Range**

**Counts**
- `open_count`: Status = "OPEN"
- `closed_count`: Status = "CLOSE"

Rules:
- All sheets are included
- Status values are uppercase

---

### 6.2 Weekly Workload Details (Activity)

**Purpose**
- Show what happened during the selected period

**Selection Rules**
- `begin_date <= Date <= end_date`

### 6.3 Right Section – Weekly Workload Details

**Purpose**
- Show current unresolved tickets

**Year Context**
- `current_year = end_date.year`
- `previous_year = current_year - 1`

**Data Source**
- Sheets: current_year and previous_year (if exists)

**Selection Rules**
- `begin_date <= Date <= end_date`
- Status = "OPEN" or "CLOSED" (Filtered by Date Range and Year Context)

**Sorting**
- Received (Time - Arrive) ascending
- Ticket No. ascending

---

## 7. Data Mapping (Unified Model)

All outputs (Web + PDF) use the same field names:

- ServiceNow Ticket #
- REQ No.
- Type
- Description
- Requested for
- PIC
- Received
- Resolved
- Remarks

Optional (display/debug):
- Status
- Date

---

## 8. PDF Generation

### Strategy (MVP)
- Use `reportlab`
- Programmatic table rendering
- Fixed, business-report-style layout

### Content
- Summary section
- Weekly Activity table
- Weekly Workload Details table

---

## 9. Validation & Error Handling

### Validation
- File is `.xlsx`
- begin_date <= end_date
- Required columns exist
- Date fields are parseable

### Errors
| Type | HTTP | Meaning |
|----|----|----|
| InputPayloadError | 400 | Invalid request |
| ExcelSchemaError | 400 | Missing sheet / column |
| NoDataWarning | 400 | No matching data |
| SystemError | 500 | Unexpected error |

---

## 10. Testing Strategy

### Unit Tests (pytest)
- Date filtering
- Section separation logic
- Summary count accuracy
- Sorting order

### Integration Tests
- `/generate` JSON structure
- `/pdf` returns valid PDF

PDF tests are lightweight:
- File is generated
- Key text exists (period, ticket no.)

---

---

## 11. Visual Identity (Professional UI)

- **Typography**: Inter (Google Fonts) for modern readability.
- **Palette**: Dodger Blue, Amber-Gold, and Emerald Green highlights.
- **Branding**: Professional logo integration with balanced typography.
- **Interactions**: Smooth CSS transitions and fade-in animations.

## 12. Repository Structure (Current)
/app
/api
/services
/infra
/excel
/tests
/docs
implementation_plan.md
README.md


---

## 12. Rules for AI Tools

- Design must be finalized before coding
- README.md + implementation_plan.md are the **authoritative specs**
- No AI-invented behavior allowed


