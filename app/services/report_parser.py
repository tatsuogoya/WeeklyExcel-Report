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
    mask_range = (full_df["Date"].dt.date >= begin_date) & (full_df["Date"].dt.date <= end_date)
    range_df = full_df[mask_range].copy()
    
    # Summary Counts Rules
    range_status = range_df["Status"].astype(str).str.strip().str.upper()
    open_count = int((range_status == "OPEN").sum())
    closed_count = int((range_status == "CLOSE").sum())

    # Left Section Rules
    left_df = range_df[range_df["Status"] == "OPEN"].copy()
    if left_df.empty:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "NO_ROWS_IN_RANGE", "message": f"No activity found between {begin_date} and {end_date}"}
        )
    left_df = left_df.sort_values(by=["Time - Arrive", "Ticket No."], ascending=[True, True])

    # Right Section Rules
    right_df = range_df[range_df["Status"].isin(["OPEN", "CLOSE"])].copy()
    if right_df.empty:
         raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "NO_RIGHT_TICKETS", "message": f"No OPEN or CLOSE tickets found between {begin_date} and {end_date}"}
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
