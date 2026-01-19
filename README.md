> ⚠️ This repository is under refactoring.
> The current implementation is transitioning from a CLI-based script (v1)
> to a FastAPI-based web application (v2).
> 
# Excel Weekly Report Generator (MVP)

## Project Overview
This tool automates the extraction of ticket data from "NA Daily work.xlsx" and generates a formatted "Weekly Report" based on a specified date range.

## Technical Stack
- Backend: Python (FastAPI)
- Excel Engine: pandas, openpyxl
- Testing: pytest

## System Architecture
The application follows a 3-layer architecture:
- API Layer: FastAPI endpoints, request validation.
- Service Layer: Data filtering, sorting, and business logic.
- Infrastructure Layer: Excel I/O using openpyxl.

## Data Mapping
| Weekly Report | NA Daily work |
|---|---|
| ServiceNow Ticket # | Ticket No. |
| REQ No. | REQ No. |
| Type | Type |
| Description | Request Detail |
| Requested for | Requested for |
| PIC | Assign To |
| Received | Time - Arrive |
| Resolved | Time - Close |
| Remarks | Remarks |

## Processing Flow
1. Validate request and input file.
2. Read target yearly sheet.
3. Filter by date range.
4. Sort by Received, then Ticket No.
5. Write data into Excel template.

## Exception Design (MVP)
- InputPayloadError (400)
- ExcelSchemaError (400)
- NoDataWarning (400)
- SystemError (500)

## Test Plan
- Normal flow
- Sorting validation
- Missing column error
- Date boundary case
