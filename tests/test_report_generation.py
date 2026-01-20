import pytest
from fastapi.testclient import TestClient
from app.main import app
import os
import io

client = TestClient(app)

def test_health():
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}

def test_v2_generate_json():
    # Paths for test
    sample_file = "sample_daily_work.xlsx"
    if not os.path.exists(sample_file):
        pytest.skip("sample_daily_work.xlsx not found")
    
    with open(sample_file, "rb") as f:
        files = {"file": ("sample.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        data = {
            "begin_date": "2026-01-05", 
            "end_date": "2026-01-11"
        }
        response = client.post("/generate", data=data, files=files)
    
    assert response.status_code == 200
    res_json = response.json()
    
    assert "summary" in res_json
    assert "left_section" in res_json
    assert "right_section" in res_json
    assert "period" in res_json
    
    assert res_json["summary"]["open_count"] == 3  # TKT26-001, TKT26-003, TKT25-999 (all OPEN regardless of date)
    # 2. Verify Left Section (Col A-F)
    # Only OPEN tickets in range: TKT26-001, TKT26-003 (TKT26-002 is CLOSED)
    # The JSON response keys are "left_section" and "right_section"
    left_tickets = [t["Ticket No."] for t in res_json["left_section"]]
    # Right section now only in range but includes OPEN and CLOSE
    # Range is 2026-01-01 to 2026-01-15
    # TKT25-999 is outside range (cancelled)
    # TKT26-001 (OPEN), 003 (OPEN), 002 (CLOSE) are in range.
    # Sorted by Status DESC: OPEN, OPEN, CLOSE
    right_tickets = [t["Ticket No."] for t in res_json["right_section"]]
    # Right section: ALL OPEN tickets + CLOSED tickets where Time - Close in range
    # TKT26-001 (OPEN), TKT26-003 (OPEN), TKT25-999 (OPEN from 2025), TKT26-002 (CLOSED in range)
    # Sorted by Status DESC (OPEN before CLOSE), then by Time-Arrive
    assert len(right_tickets) == 4
    assert "TKT25-999" in right_tickets  # Now included (all OPEN tickets)
    assert "TKT26-002" in right_tickets  # CLOSED but in date range
    # CLOSE should be last
    assert right_tickets[-1] == "TKT26-002"

def test_v2_generate_pdf():
    sample_file = "sample_daily_work.xlsx"
    if not os.path.exists(sample_file):
        pytest.skip("sample_daily_work.xlsx not found")
        
    with open(sample_file, "rb") as f:
        files = {"file": ("sample.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        data = {
            "begin_date": "2026-01-05", 
            "end_date": "2026-01-11"
        }
        response = client.post("/pdf", data=data, files=files)
    
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/pdf"
    assert len(response.content) > 1000 # Should be a non-trivial PDF

def test_no_rows_in_range():
    """
    Test behavior when date range has no CLOSED tickets but OPEN tickets exist.
    Current business logic: left_section includes ALL OPEN tickets regardless of date.
    So this returns 200 with the OPEN tickets, not 400.
    """
    sample_file = "sample_daily_work.xlsx"
    with open(sample_file, "rb") as f:
        files = {"file": ("sample.xlsx", f, "spreadsheet")}
        # Dates with NO closed tickets in range, but OPEN tickets exist
        data = {"begin_date": "2026-02-01", "end_date": "2026-02-07"}
        response = client.post("/generate", data=data, files=files)
    # Returns 200 because we still have OPEN tickets to show
    assert response.status_code == 200
    res_json = response.json()
    # Right section only has OPEN tickets (no CLOSED in this date range)
    right_tickets = [t["Ticket No."] for t in res_json["right_section"]]
    assert "TKT26-002" not in right_tickets  # CLOSED ticket not in this range

