# 📄 Statement of Work (SOW): Marelli IFS Monthly KPI Automation
**Project Version:** 1.0 (MVP)
**Drafted by:** Codex (Commander Antigravity)
**Status:** Proposition for Review

---

## 1. Project Objective
To automate the generation of the "Marelli IFS Monthly KPI" report, transitioning from manual compilation to a high-fidelity, data-driven system. The primary focus is on visual replication of legacy management reports using raw Excel data from `NA Daily work.xlsx`.

## 2. Technical Architecture (3-Layer Mode)
- **Service Layer (`monthly_report_service.py`)**: Responsible for hierarchical data aggregation (Incident vs Service) and monthly/annual pivoting.
- **Infrastructure Layer (`chart_renderer.py`)**: Utilizes `Matplotlib` to generate professional-grade charts.
- **API/UI Layer**: Extension of the current FastAPI dashboard to support Monthly Report selection and PDF rendering.

## 3. Core Requirements & Business Logic

### 3.1 Data Classification (Hierarchical)
The system shall implement a two-tier filtering logic:
1.  **Tier 1 (Type)**: Distinct separation into `Incident` and `Service` categories based on the Excel `Type` column.
2.  **Tier 2 (Category)**: Grouping tickets by fixed dropdown values (e.g., *Reset password, Deactivate account, Create account*).
    - *Note:* The same Category may exist under both Incident and Service; the system must maintain this distinction in the pivot tables.

### 3.2 Status Mapping (Pivot Table)
The Monthly Pivot Table shall map the Excel `Status` column to the following report headers:
- `TRANS_NON`, `TRANS_WORK`, `CANCEL`, `OPEN`, `CLOSE`.
- Mapping is direct based on exact string matches from the source dropdown values.

### 3.3 Visual Components (High-Fidelity)
- **Annual Summary Chart**: 
  - Type: Stacked Bar Chart.
  - Granularity: Monthly (Jan - Dec).
  - Colors: Exact match to management palette (Navy, Orange, Teal, Purple, Green, Red/Brown, Blue).
  - Support: Data table integrated below the X-axis.
- **Monthly Distribution Chart**:
  - Type: Pie Chart.
  - Content: Category percentage distribution for the selected month.

## 4. Implementation Scope

| In Scope | Out of Scope (Phase 1) |
| :--- | :--- |
| Automatic ticket counting | **Development Hours (ME Hours)** |
| Incident/Service Partitioning | **Planned Effort / Carrying Forward** |
| Stacked Bar & Pie Chart Rendering | **Interactive Chart Filtering** |
| PDF Export (Fidelity Focused) | **Automatic SFTP Emailing** |

---

## 5. Success Criteria for MVP
- **Accuracy**: Data counts must match manual Excel pivot calculations for December 2025.
- **Visuals**: Charts must be indistinguishable from the provided slide references in terms of color and layout.
- **Stability**: System handles missing/empty data rows gracefully without crashing.

---
**Commander's Note:** This SOW prioritizes structural integrity and visual professional-grade output. Logic for "Development Hours" is prepared but disabled until source data is available in Excel.
