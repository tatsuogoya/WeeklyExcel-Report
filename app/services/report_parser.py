import pandas as pd
from datetime import date, datetime
from typing import Dict, Any
from fastapi import HTTPException, status
from app.utils.column_utils import sanitize_column_name

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

    # 4. New Users Section
    new_users_df = pd.DataFrame()
    # Find sheet name case-insensitively
    new_users_sheet_name = next((name for name in all_sheets.keys() if name.lower() == "new users"), None)
    
    if new_users_sheet_name:
        new_users_raw = all_sheets[new_users_sheet_name].copy()
        
        if len(new_users_raw) > 0:
            # Check if columns are problematic (Unnamed, NaN, or nan strings)
            cols_are_bad = any(
                "Unnamed" in str(col) or pd.isna(col) or str(col).lower() == 'nan'
                for col in new_users_raw.columns
            )
            
            # Expected column keywords to look for in header row
            expected_keywords = ['ticket', 'date', 'user', 'name', 'email', 'department', 'function']
            
            if cols_are_bad:
                # Scan first 10 rows to find a row that looks like headers
                header_row_idx = None
                for idx in range(min(10, len(new_users_raw))):
                    row_values = [str(v).lower() for v in new_users_raw.iloc[idx] if pd.notna(v)]
                    matches = sum(1 for kw in expected_keywords if any(kw in val for val in row_values))
                    if matches >= 2:  # At least 2 matching keywords
                        header_row_idx = idx
                        break
                
                if header_row_idx is not None:
                    # Use this row as headers
                    new_header = dict(zip(new_users_raw.columns, new_users_raw.iloc[header_row_idx]))
                    new_users_raw = new_users_raw.rename(columns=new_header)
                    # Skip rows up to and including header row
                    new_users_df = new_users_raw.iloc[header_row_idx + 1:].copy()
                else:
                    new_users_df = new_users_raw.copy()
            else:
                new_users_df = new_users_raw.copy()
            
            # Sanitize column names using shared utility
            new_users_df.columns = [sanitize_column_name(c) for c in new_users_df.columns]
            # Remove empty column names
            new_users_df = new_users_df.loc[:, new_users_df.columns != ""]
            
            # Map common column name variations to expected names
            column_mapping = {
                'Ticket No.': 'Ticket No',
                'Email': 'Email address',
                'E-mail address': 'Email address',
                'EmailAddress': 'Email address',
                'Department': 'Function / Department',
                'Dept': 'Function / Department',
                'Function': 'Function / Department',
                'User': 'User Name',
                'Name': 'User Name',
                'UserName': 'User Name',
                'Created': 'Date Created',
                'Date': 'Date Created',
            }
            new_users_df = new_users_df.rename(columns={k: v for k, v in column_mapping.items() if k in new_users_df.columns})
            
            # Ensure Date Created is datetime
            if "Date Created" in new_users_df.columns:
                new_users_df["Date Created"] = pd.to_datetime(new_users_df["Date Created"], errors='coerce')
                
                # Business Rule: Sort by Date Created descending (newest first)
                new_users_df = new_users_df.sort_values(by="Date Created", ascending=False, na_position='last')
                
                # Business Rule: Add a flag for users created within the requested date range
                def check_new(row):
                    dt = row.get("Date Created")
                    if pd.notna(dt) and hasattr(dt, 'date'):
                        return begin_date <= dt.date() <= end_date
                    return False
                
                new_users_df["is_new_user"] = new_users_df.apply(check_new, axis=1)
            
            # Remove empty rows (where key columns are all NaN)
            key_cols = [c for c in ['Ticket No', 'User Name', 'Email address'] if c in new_users_df.columns]
            if key_cols:
                new_users_df = new_users_df.dropna(subset=key_cols, how='all')
            else:
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
