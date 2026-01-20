import pandas as pd
from datetime import datetime

def create_sample_data():
    # Data for 2026
    data_2026 = {
        "Date": [datetime(2026, 1, 5), datetime(2026, 1, 6), datetime(2026, 1, 7)],
        "Ticket No.": ["TKT26-001", "TKT26-002", "TKT26-003"],
        "REQ No.": ["REQ26-001", "REQ26-002", "REQ26-003"],
        "Type": ["Incident", "Request", "Incident"],
        "Requested for": ["User A", "User B", "User C"],
        "Assign To": ["Tech 1", "Tech 2", "Tech 1"],
        "Request Detail": ["2026-Issue 1", "2026-Issue 2", "2026-Issue 3"],
        "Time - Arrive": [datetime(2026, 1, 5, 9, 0), datetime(2026, 1, 6, 10, 0), datetime(2026, 1, 7, 11, 0)],
        "Time - Close": [None, datetime(2026, 1, 6, 15, 0), None],
        "Remarks": ["", "", ""],
        "Status": ["OPEN", "CLOSE", "OPEN"]
    }
    
    # Data for 2025 (Backlog)
    data_2025 = {
        "Date": [datetime(2025, 12, 20), datetime(2025, 12, 21)],
        "Ticket No.": ["TKT25-999", "TKT25-888"],
        "REQ No.": ["REQ25-999", "REQ25-888"],
        "Type": ["Incident", "Incident"],
        "Requested for": ["User X", "User Y"],
        "Assign To": ["Tech 1", "Tech 1"],
        "Request Detail": ["2025-Old Issue", "2025-Solved Issue"],
        "Time - Arrive": [datetime(2025, 12, 20, 8, 30), datetime(2025, 12, 21, 9, 0)],
        "Time - Close": [None, datetime(2025, 12, 21, 17, 0)],
        "Remarks": ["Stayed open", "Closed in 2025"],
        "Status": ["OPEN", "CLOSE"]
    }

    with pd.ExcelWriter("sample_daily_work.xlsx") as writer:
        pd.DataFrame(data_2026).to_excel(writer, sheet_name="2026", index=False)
        pd.DataFrame(data_2025).to_excel(writer, sheet_name="2025", index=False)
    
    print("V2 Sample data created at sample_daily_work.xlsx (Sheets: 2026, 2025)")

if __name__ == "__main__":
    create_sample_data()
