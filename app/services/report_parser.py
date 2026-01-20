import pandas as pd
from datetime import date, datetime
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
        df["Time - Close"] = pd.to_datetime(df["Time - Close"], errors='coerce')
        df["Status"] = df["Status"].astype(str).str.strip().str.upper()
        return df

    # 1. Filter sheets: Only process current year and previous year
    current_year = datetime.now().year
    previous_year = current_year - 1
    allowed_years = {str(current_year), str(previous_year)}
    
    filtered_sheets = {}
    for sheet_name, df in all_sheets.items():
        sheet_str = str(sheet_name).strip()
        # Only include current year and previous year sheets
        if sheet_str in allowed_years:
            filtered_sheets[sheet_name] = df
    
    if not filtered_sheets:
         raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "NO_YEAR_SHEETS", "message": f"No {previous_year} or {current_year} sheets found in Excel file"}
        )

    # 2. Normalize sheet data
    valid_dfs = {}
    for sheet_name, df in filtered_sheets.items():
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
            detail={"error_code": "REQUIRED_COLUMNS_MISSING", "message": "No valid data in 2025 or 2026 sheets"}
        )

    # 3. Section Partitioning Rules
    full_df = pd.concat(valid_dfs.values(), ignore_index=True)
    
    # For date range filtering
    mask_range_date = (full_df["Date"].dt.date >= begin_date) & (full_df["Date"].dt.date <= end_date)
    # For Time - Close, need to handle NaN values
    mask_range_close = (full_df["Time - Close"].notna()) & (full_df["Time - Close"].dt.date >= begin_date) & (full_df["Time - Close"].dt.date <= end_date)
    
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

    # 4. New Users Section (no date filtering, show all)
    new_users_df = pd.DataFrame()
    if "New Users" in all_sheets:
        new_users_raw = all_sheets["New Users"]
        
        if len(new_users_raw) > 0:
            # Check if columns are "Unnamed: X" format
            if any("Unnamed" in str(col) for col in new_users_raw.columns):
                # First row contains the actual headers
                # Use it as the new column names
                new_header = new_users_raw.iloc[0].to_dict()
                new_users_raw = new_users_raw.rename(columns=new_header)
                # Skip the first row (which was the header row)
                new_users_df = new_users_raw.iloc[1:].copy()
            else:
                new_users_df = new_users_raw.copy()
            
            # Clean up column names (strip whitespace)
            new_users_df.columns = [str(c).strip() for c in new_users_df.columns]
            
            # Ensure Date Created is datetime
            if "Date Created" in new_users_df.columns:
                new_users_df["Date Created"] = pd.to_datetime(new_users_df["Date Created"], errors='coerce')
            
            # Remove empty rows
            new_users_df = new_users_df.dropna(how='all')
            new_users_df = new_users_df.reset_index(drop=True)
    
    return {
        "left_df": left_df,
        "right_df": right_df,
        "new_users_df": new_users_df,
        "summary": {
            "open_count": open_count,
            "closed_count": closed_count
        }
    }
