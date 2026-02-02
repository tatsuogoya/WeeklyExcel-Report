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
- Show current unresolved tickets and recently closed ones

**Year Context**
- `current_year = datetime.now().year`
- `previous_year = current_year - 1`

**Data Source**
- Sheets: current_year and previous_year only

**Selection Rules**
- All OPEN tickets (any date)
- CLOSE tickets where Time - Close falls within begin_date to end_date

**Sorting**
- Status DESC (OPEN before CLOSE)
- Received (Time - Arrive) ascending
- Ticket No. ascending

---

### 6.4 New Users Section (Optional)

**Purpose**
- Display newly added users/onboarded staff

**Data Source**
- Sheet: "New Users" (if exists, optional)

**Selection Rules**
- All records in the sheet are included.

**Sorting Rules**
- Sorted by **Date Created** descending (newest first).
- Handled exclusively in the Service Layer.

**Display**
- Appears only if "New Users" sheet exists.
- Shows at bottom of web view and PDF.
- **Dynamic Columns**: All columns from the Excel sheet are displayed (no hardcoded list).
- **Highlighting**: Rows created within the selected `begin_date` to `end_date` are displayed in **bold blue text**.
- Columns mapped if found: Ticket No, Date Created, User Name, Function / Department, Email address. Scatter columns are preserved.

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
- Summary section (period, open/closed counts)
- Weekly Activity table (all OPEN tickets)
- Weekly Workload Details table (OPEN + CLOSE tickets in range)
- New Users section (if exists)

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

## 13. Refactoring and Maintainability Cleanup

To ensure long-term maintenance, the following structural improvements are being implemented:

### Service Layer Modularization
- `report_parser.process_report_data` is being split into smaller, focused helpers:
    - `_normalize_activity_data`: Handles standard year-based sheets.
    - `_process_new_users`: Handles the "New Users" sheet logic.
    - `_detect_header_row`: Logic for scanning messy Excel headers.

### Logic Consolidation
- PDF row highlighting logic is being consolidated to use backend flags (`is_new_user`) rather than re-calculating dates in the PDF service.
- Table preparation logic in `pdf_service` is being unified into a single robust helper.

### Frontend Cleanup
- Removal of debug `console.log` statements.
- Strict reliance on backend flags for conditional styling.

### Repository Hygiene
- Archived legacy scripts (`generate_report.py`, `inspect_excel.py`, `test_log.txt`) for a cleaner root directory.

---

## 14. Phase 3: Monthly Report (Strategic Implementation)

### 📊 Vision
The monthly report transforms raw ticket data into a strategic KPI summary (Marelli IFS Monthly KPI) for management review. It focuses on annual trends, category-specific performance, and service/incident categorization.

### 🖼️ Report Structure (Based on Reference Images)

#### 1. Title & TOC
- **Branding**: Alpash / Marelli Logo.
- **Title**: Marelli IFS Monthly KPI - Monthly Report for [Month-Year].
- **Sections**: 
  1. Total incidents and SRs.
  2. Development hours, how much is planned and utilized (Placeholder for now).

#### 2. Annual Summary (Stacked Bar Chart)
- **Purpose**: Monthly ticket volume breakdown for the entire year.
- **X-Axis**: Jan to Dec.
- **Y-Axis**: Ticket Count (0-200 range).
- **Categories (Stacks)**:
  - `Miscellaneous` (Navy)
  - `Development` (Orange)
  - `Transfer to another group` (Teal)
  - `Permissions control` (Purple)
  - `Create account` (Olive/Green)
  - `Reset password` (Red/Brown)
  - `Deactivate account` (Blue)
- **Data Table**: A corresponding table below the chart showing raw monthly counts for each category.

#### 3. Total Incidents and SRs (Pivot Table)
- **Grouping**: Hierarchical grouping by `Incident` and `Service`.
- **Columns**:
  - `CATEGORY`
  - `TRANS_NON`
  - `TRANS_WORK`
  - `CANCEL`
  - `OPEN`
  - `CLOSE`
  - `Grand Total`
- **Mapping Logic**: Needs to be defined (e.g., Status mapping to pivot columns).

#### 4. Total Incidents and SRs (Pie Chart)
- **Purpose**: Percentage distribution of categories for the selected month.
- **Data**: All ticket categories included.

---

### 📝 Strategic Questions for "The Boss" (Tatsuo-san)

To ensure perfect alignment before we begin coding, I have a few missing pieces to confirm:

1. **Incident vs Service**: 
   - Image 4 shows categories like "Reset password" appearing under both `Incident` and `Service`. 
   - **Question**: Is there a specific column in the Excel file that tells us if a ticket is an "Incident" or a "Service"? Or should the AI determine this based on other keywords?

2. **Column Mapping (TRANS_NON, TRANS_WORK, etc.)**:
   - **Question**: How do we map the Excel "Status" to columns like `TRANS_NON` and `TRANS_WORK`? 
   - Usually, `CLOSE` means the status is 'CLOSE', but the others are less obvious.

3. **Development Hours (TOC Item 2)**:
   - **Question**: You mentioned focus on the first 5 slides, but the TOC lists "Development hours". Is this information currently in the Excel file, or is this a manual entry for later?

4. **Status of Jan 2026**:
   - Since January isn't finished, should we use **December 2025** as our primary test case for the first "Master Monthly Report"?

---

### 🚀 Implementation Strategy

1. **Create `app/services/monthly_report_service.py`**: Handle aggregation (Monthly & Annual).
2. **Create `app/infra/chart_renderer.py`**: Using `matplotlib` to replicate the specific colors and stacked bar/pie chart styles.
3. **Draft `Monthly_Report_Template.pdf`**: Matching the branding and layout seen in the slides.


