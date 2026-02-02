# 🏦 Monthly Report Project: Intelligence Package (V1)
**Date:** 2026-01-29
**Commander:** Antigravity
**Directives for:** Codex (Gemini CLI)

## 1. Project Overview
Automate the "Marelli IFS Monthly KPI" report generation from `NA Daily work.xlsx`. The goal is to replicate a professional presentation (as seen in the provided 5 slides) with high-fidelity charts and pivot tables.

## 2. Collected Intelligence (From Tatsuo-san)

### A. Classification Logic (The Type Column)
- **Primary Tier:** Tickets must first be categorized by the `Type` column into **Incident** or **Service**.
- **Secondary Tier:** Under each tier, tickets are split into specific categories (e.g., Reset password, Create account).
- **Rule:** The same category (e.g., "Reset password") can appear in both Incident and Service tiers depending on the `Type` column. Do not merge them.

### B. Status Mapping (Pivot Table Columns)
- **Source:** Excel `Status` column.
- **Mapping:** Directly map status strings to report columns: `TRANS_NON`, `TRANS_WORK`, `CANCEL`, `OPEN`, `CLOSE`.
- **Validation:** Data is entered via dropdowns in Excel, so strings will match the slide labels exactly.

### C. Visual Identity (The Graph)
- **Requirement:** Replicate the "Annual Summary" (Stacked Bar Chart) with near-pixel perfection.
- **Colors:**
  - Miscellaneous: Navy
  - Development: Orange
  - Transfer to another group: Teal
  - Permissions control: Purple
  - Create account: Olive/Green
  - Reset password: Red/Brown
  - Deactivate account: Blue
- **Layout:** Annual trend (Jan-Dec) with a data table below the chart.

### D. Scope Constraints
- **In Scope:** Automated counts, pivot tables, annual summary chart, monthly pie chart.
- **Out of Scope (For now):** Development hours (ME Hours/Planned Effort) - these are NOT in Excel yet.
- **Target Data:** December 2025 as the master test model.

---
**Instruction for Codex:**
Using the intelligence above and the conversation history, draft a professional **Statement of Work (SOW)** or **Implementation Plan**. Focus on architecture, data processing logic, and visual fidelity. Antigravity will judge the quality of your output.
