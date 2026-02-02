import pandas as pd
from app.services.monthly_report_service import MonthlyReportService

def verify_removal():
    service = MonthlyReportService()
    
    # テストデータ作成
    data = [
        {"Ticket No.": "1", "Date Created": "2025-12-01", "Type": "Incident", "Category": "Deactivate account", "Status": "CLOSE"},
        {"Ticket No.": "2", "Date Created": "2025-12-01", "Type": "Incident", "Category": "Deactivates account", "Status": "CANCEL"},
        {"Ticket No.": "3", "Date Created": "2025-12-02", "Type": "Service", "Category": "Cancel", "Status": "CLOSE"}, # Category is Cancel
        {"Ticket No.": "4", "Date Created": "2025-12-03", "Type": "Service", "Category": "Reset password", "Status": "OPEN"}
    ]
    df = pd.DataFrame(data)
    
    print("--- Starting Verification ---")
    
    # 1. Monthly Data Aggregation
    res = service.aggregate_monthly_data(df, 2025, 12)
    pivot_data = res["pivot_data"]
    
    print("\n[Monthly Pivot Check]")
    has_cancel = False
    for key, vals in pivot_data.items():
        if "CANCEL" in vals:
            has_cancel = True
            print(f"FAILED: Found CANCEL in {key}")
    
    if not has_cancel:
        print("SUCCESS: CANCEL column is NOT in pivot data.")
    
    # Category combination check
    deact_keys = [k for k in pivot_data.keys() if "Deactivate" in k]
    print(f"Deactivate keys found: {deact_keys}")
    if len(deact_keys) == 1:
        print("SUCCESS: Deactivate account combined correctly.")
    else:
        print("FAILED: Multiple deactivate keys found.")

    # 2. Annual Summary Check
    ann = service.get_annual_summary_data(df, 2025)
    print("\n[Annual Summary Check]")
    print(f"Categories: {ann['categories']}")
    # Calculate sum of all counts
    total_count = sum([sum(v) for v in ann['data'].values()])
    print(f"Total count in annual summary: {total_count}")
    
    # There are 4 total, 2 are CANCEL. Result should be 2.
    if total_count == 2:
        print("SUCCESS: CANCEL rows excluded from counts.")
    else:
        print(f"FAILED: Expected count 2, but got {total_count}")
        
    print("\n--- Verification Complete ---")

if __name__ == "__main__":
    verify_removal()
