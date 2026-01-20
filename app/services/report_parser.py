import pandas as pd
from datetime import date
from typing import Dict, Any
from fastapi import HTTPException, status

REQUIRED_COLUMNS = [
    "Date", "Ticket No.", "REQ No.", "Type", "Requested for", 
    "Assign To", "Request Detail", "Time - Arrive", "Time - Close", 
    "Remarks", "Status"
]

def process_report_data(all_sheets: Dict[str, pd.DataFrame], begin_date: date, end_date: date) -> Dict[str, Any]:
    """
    Pure Logic: Processes raw sheet data into partitioned dataframes and summaries.
    No direct file system access.
    """
    def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            for col in missing:
                df[col] = None
        df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
        df["Time - Arrive"] = pd.to_datetime(df["Time - Arrive"], errors='coerce')
        df["Status"] = df["Status"].astype(str).str.strip().str.upper()
        return df

    # 1. Normalize all sheet data
    valid_dfs = {}
    for sheet_name, df in all_sheets.items():
        normalized_sheet_name = str(sheet_name).strip()
        df = normalize_df(df)
        if "Date" in df.columns:
            if normalized_sheet_name in valid_dfs:
                valid_dfs[normalized_sheet_name] = pd.concat([valid_dfs[normalized_sheet_name], df], ignore_index=True)
            else:
                valid_dfs[normalized_sheet_name] = df

    if not valid_dfs:
         raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "REQUIRED_COLUMNS_MISSING", "message": "No valid data sheets found with required columns"}
        )

    # 2. Section Partitioning Rules
    full_df = pd.concat(valid_dfs.values(), ignore_index=True)
    
    # For date range filtering
    mask_range_date = (full_df["Date"].dt.date >= begin_date) & (full_df["Date"].dt.date <= end_date)
    mask_range_close = (full_df["Time - Close"].dt.date >= begin_date) & (full_df["Time - Close"].dt.date <= end_date)
    
    # Summary Counts Rules
    # open_count: All OPEN tickets regardless of date
    # closed_count: CLOSED tickets where "Time - Close" falls within the selected date range
    all_open = full_df[full_df["Status"] == "OPEN"]
    open_count = int(len(all_open))
    
    closed_in_range = full_df[(full_df["Status"] == "CLOSE") & mask_range_close]
    closed_count = int(len(closed_in_range))

    # Left Section Rules: All OPEN tickets regardless of date
    left_df = full_df[full_df["Status"] == "OPEN"].copy()
    if left_df.empty:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "NO_OPEN_TICKETS", "message": "No OPEN tickets found"}
        )
    left_df = left_df.sort_values(by=["Time - Arrive", "Ticket No."], ascending=[True, True])

    # Right Section Rules: All OPEN tickets + CLOSED tickets where "Time - Close" is in date range
    open_all = full_df[full_df["Status"] == "OPEN"].copy()
    closed_range = full_df[(full_df["Status"] == "CLOSE") & mask_range_close].copy()
    right_df = pd.concat([open_all, closed_range], ignore_index=True) if not closed_range.empty else open_all.copy()
    
    if right_df.empty:
         raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "NO_RIGHT_TICKETS", "message": "No OPEN or CLOSED tickets found"}
        )
    right_df = right_df.sort_values(by=["Status", "Time - Arrive", "Ticket No."], ascending=[False, True, True])

    return {
        "left_df": left_df,
        "right_df": right_df,
        "summary": {
            "open_count": open_count,
            "closed_count": closed_count
        }
    }
