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
    
    assert res_json["summary"]["open_count"] == 2
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
    assert right_tickets[0] in ["TKT26-001", "TKT26-003"]
    assert right_tickets[1] in ["TKT26-001", "TKT26-003"]
    assert right_tickets[2] == "TKT26-002" # CLOSE should be last
    assert "TKT25-999" not in right_tickets # Outside range
    assert len(right_tickets) == 3

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
    sample_file = "sample_daily_work.xlsx"
    with open(sample_file, "rb") as f:
        files = {"file": ("sample.xlsx", f, "spreadsheet")}
        # Dates with NO activity
        data = {"begin_date": "2026-02-01", "end_date": "2026-02-07"}
        response = client.post("/generate", data=data, files=files)
    assert response.status_code == 400
    assert response.json()["detail"]["error_code"] == "NO_ROWS_IN_RANGE"
