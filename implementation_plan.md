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

## 14. Future Expansion: Monthly Report (Planned)

> **Status**: Requirements gathering phase  
> **Current State**: Manually created each month  
> **Goal**: Fully automated generation like weekly report

### Overview

The monthly report provides aggregated analytics and visualizations for management review. Unlike the weekly report (which shows raw ticket data), the monthly report focuses on:
- **Trend analysis** over time
- **Category breakdowns** with percentages
- **Development effort tracking**
- **SLA compliance monitoring**

### Data Sources

#### Primary Source
- Same `NA Daily work.xlsx` file as weekly report
- Aggregated over full calendar month period

#### External Parameters (Manual Input)
The following data comes from outside sources and should be configurable:

1. **Development Efforts** section:
   - `ME Hours` (monthly values)
   - `SOW Planned` (monthly values)
   - `Carry Forward` (monthly values)
   - Development ticket details (Ticket No, Phase, Description, Plan Effort, Plan Close)

2. **Configuration**:
   - Month/Year selection
   - Custom category mappings (if different from weekly)

### Report Sections

#### 1. Total Incidents and SRs (Pie Chart)
**Purpose**: Visual breakdown of ticket categories

**Data**:
- Categories: Deactivate account, Reset password, Transfer to another group, Permissions control, Create account, Cancel, Miscellaneous
- Shows percentage distribution

**Chart Type**: Pie chart (programmatic generation using matplotlib/plotly)

**Business Rules**:
- Count all tickets for the selected month
- Group by ticket Type or Category field
- Calculate percentages

#### 2. Total Incidents and SRs (Pivot Table)
**Purpose**: Detailed status breakdown by category

**Columns**:
- `CATEGORY` (dropdown filter in PDF)
- `TRANS_NON`, `TRANS_WORK`, `CANCEL`, `OPEN`, `CLOSE`
- `Grand Total`

**Rows**:
- Grouped into "Incident" and "Service"
- Individual categories under each group
- Grand Total row

**Business Rules**:
- Aggregate by Category and Status
- Hierarchical grouping (Incident vs Service)

#### 3. Development Efforts
**Purpose**: Track development work hours

**Table 1 - Monthly Hours**:
| Metric | Jan | Feb | Mar | ... | Dec | Total |
|--------|-----|-----|-----|-----|-----|-------|
| ME Hours | (external param) |
| SOW Planned | (external param) |
| Carry Forward | (external param) |

**Table 2 - Ongoing/Recently Closed**:
| Ticket No | Phase | Description | Plan Effort | Plan Close |
|-----------|-------|-------------|-------------|------------|
| ... | ... | ... | ... | ... |

**Data Source**: 
- Monthly metrics: External parameters (manual input)
- Ticket details: External parameters or filtered from Excel

#### 4. SLA Breach
**Purpose**: Highlight tickets that exceeded SLA

**Columns**:
- Ticket No.
- Requested for
- Description
- Time-Arrive
- Business elapsed time
- Remarks

**Business Rules**:
- Show tickets where elapsed time > SLA threshold
- Display "No SLA Breached tickets" if none found
- SLA threshold may be a configurable parameter

#### 5. Annual Summary (Stacked Bar Chart)
**Purpose**: Year-over-year trend visualization

**Chart Type**: Stacked bar chart by month

**Data**:
- X-axis: Months (Jan-Dec)
- Y-axis: Ticket count
- Stacks: Categories (color-coded like pie chart)

**Business Rules**:
- Aggregate tickets by month for current year
- Stack by category
- Include data table below chart

### Technical Implementation Strategy

#### Chart Generation
**Library**: `matplotlib` (recommended for PDF embedding)
- Alternative: `plotly` for web view interactivity

**Approach**:
1. Generate charts as PNG images
2. Embed in PDF using ReportLab's `Image` class
3. For web view: Generate interactive charts with plotly/Chart.js

#### Architecture

```
/app/services/
  ├── weekly_report_parser.py    # Rename current report_parser.py
  ├── monthly_report_parser.py   # New: Monthly aggregation logic
  └── chart_generator.py         # New: Shared chart utilities

/app/api/
  ├── weekly_report.py           # Rename current report.py
  └── monthly_report.py          # New: Monthly endpoint

/app/infra/
  ├── pdf_renderer.py            # Existing (reuse)
  └── chart_renderer.py          # New: matplotlib/plotly wrapper
```

#### API Design

**Endpoint**: `POST /monthly`

**Request**:
```json
{
  "file": "<Excel file upload>",
  "year": 2025,
  "month": 12,
  "development_efforts": {
    "me_hours": [8, 13, 12, ...],
    "sow_planned": [16, 16, 16, ...],
    "carry_forward": [8, 3, 4, ...]
  },
  "dev_tickets": [
    {
      "ticket_no": "SCTASK0929134",
      "phase": "OPEN",
      "description": "...",
      "plan_effort": 4,
      "plan_close": "..."
    }
  ],
  "sla_threshold_hours": 24 // Optional
}
```

**Response**:
- Web: JSON with chart data + aggregated tables
- PDF: Binary PDF file

#### Additional Dependencies

Add to `requirements.txt`:
```
matplotlib>=3.8.0
pandas>=2.1.0  # Already included
numpy>=1.24.0  # For chart calculations
```

### Validation & Testing Strategy

1. **Data Validation**:
   - Ensure month/year are valid
   - Validate development effort arrays (12 months)
   - Check SLA threshold is positive number

2. **Chart Generation**:
   - Unit tests for chart_generator.py
   - Ensure charts render correctly with various data sizes
   - Handle edge cases (no data, single category, etc.)

3. **PDF Layout**:
   - Verify charts fit on landscape pages
   - Test with min/max data scenarios

### Open Questions

- [ ] Should the web view support interactive filtering (e.g., click pie chart segment to filter table)?
- [ ] Do we need Excel export for monthly report?
- [ ] Should historical monthly reports be stored/archived?
- [ ] Any custom branding/styling differences from weekly report?

### Implementation Priority

**Phase 1** (Core functionality):
1. Monthly data aggregation service
2. Pivot table generation
3. Basic chart generation (pie, stacked bar)

**Phase 2** (Enhanced features):
4. Development Efforts section
5. SLA Breach detection
6. Web view with interactive charts

**Phase 3** (Polish):
7. Chart customization (colors, fonts)
8. PDF layout optimization
9. Error handling & validation


