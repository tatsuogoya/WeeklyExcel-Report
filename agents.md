# Agent Instructions – Alphast Weekly Report (Web + PDF)

These instructions apply to all AI agents (Antigravity, Copilot, etc.)
working on this repository.

## 1. Source of Truth

The **only** authoritative specs are:

- `README.md`
- `implementation_plan.md`

Agents **must not**:

- Change behavior that contradicts these files
- Add new features that are not described there
- Change business rules (filters, sorting, counts) on their own

If something is unclear or inconsistent, **ask for clarification first**  
instead of making assumptions.

---

## 2. Project Architecture

This project follows a 3-layer architecture:

- **API Layer (`/app/api`)**
  - FastAPI endpoints
  - Request validation
  - HTTP error mapping

- **Service Layer (`/app/services`)**
  - Business logic
  - Date filtering
  - Status handling
  - Summary calculation
  - Section separation (Weekly Activity / Open Backlog)

- **Infrastructure Layer (`/app/infra`)**
  - Excel I/O (pandas + openpyxl)
  - PDF rendering (reportlab)
  - File system access (if any)

Agents must respect these boundaries:

- Do **not** put business logic into:
  - API routes
  - PDF rendering code
  - Excel repository code

- Service Layer is the **only** place for:
  - Data selection rules
  - Sorting rules
  - Summary count rules

---

## 3. Responsibilities of Agents

When modifying or generating code, agents should:

1. **Check existing modules first**
   - Reuse and extend existing functions
   - Avoid creating parallel/duplicate logic

2. **Keep logic deterministic**
   - All complex behavior must be implemented as normal Python code
   - No “hidden” behavior in prompts, comments, or tools

3. **Avoid scope creep**
   - Do not:
     - Introduce new frameworks (e.g., React, JS SPAs)
     - Add new output formats (Excel, Google Sheets, Slides)
     - Implement monthly reports (out of scope for MVP)

4. **Preserve contracts**
   - Do not change public function signatures without a clear reason
   - Do not change API response shapes without updating README and discussing

---

## 4. PDF Generation Rules

- Use **reportlab only** for PDF generation.
- `pdf_generator` (or equivalent service) must:
  - Receive a prepared DTO (e.g., `WeeklyReportDTO`)
  - Contain **layout/rendering code only**
  - Never re-implement:
    - Date filtering
    - Status logic
    - Sorting
    - Summary calculation

If business rules appear to be missing, fix them in the **Service Layer**,  
not in the PDF generator.

---

## 5. Error Handling and Tests

Agents should:

- Keep existing exceptions:
  - `InputPayloadError`
  - `ExcelSchemaError`
  - `NoDataWarning`
  - `SystemError`

- When fixing bugs:
  - Add or update **pytest tests** under `/tests`
  - Prefer improving existing tests rather than deleting them

No change is “done” without a corresponding test update.

---

## 6. File / Directory Constraints

Agents must **not**:

- Introduce a `directives/` or `execution/` structure
- Add Google-specific files:
  - `credentials.json`
  - `token.json`
- Add `.tmp/`-based flows unless explicitly requested

Current structure (approximate):

- `/app/api` – FastAPI layer
- `/app/services` – business logic
- `/app/infra` – Excel/PDF I/O
- `/excel` – templates, sample files
- `/tests` – pytest

Stick to this organization unless the human owner requests structural changes.

---

## 7. Change Process

For any non-trivial change, agents should:

1. Describe briefly:
   - What changed
   - Why it changed
   - How it was tested
2. Make changes in small, reviewable increments
3. Prefer refactoring over rewriting large sections

If a design change is required (not a pure bug fix),  
agents must propose it first and wait for human approval.

---

## 8. Summary

- Follow `README.md` and `implementation_plan.md` strictly.
- Respect the 3-layer architecture (API / Service / Infra).
- Keep business logic out of PDF and IO layers.
- Do not expand scope (no new features without approval).
- Every fix should come with tests.

Agents are here to **implement and refine**, not to redefine the project.
