"""
Weekly Report Generator
=======================
Generates weekly status reports from ServiceNow data.

Main Flow:
1. Get user inputs (date range, source sheet)
2. Load and combine data from Excel
3. Filter tickets (only OPEN status)
4. Map columns from source to template
5. Write data to output Excel file
6. Calculate statistics for dashboard
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import shutil
from datetime import datetime, timedelta
import os
from copy import copy

# =============================================================================
# CONFIGURATION
# =============================================================================

SOURCE_FILE = "NA Daily work.xlsx"
TEMPLATE_FILE = "SNOW_report_Template.xlsx"
SOURCE_SHEET_USERS = "New Users"
TARGET_SHEET_TICKETS = "IFS"
TARGET_SHEET_USERS = "New User"
DATA_START_ROW = 3


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def get_last_week_date_range(reference_date=None):
    """Calculate start and end dates for the previous week (Monday-Sunday)."""
    if reference_date is None:
        reference_date = datetime.today()
    start_of_last_week = reference_date - timedelta(days=reference_date.weekday() + 7)
    start_of_last_week = start_of_last_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_last_week = start_of_last_week + timedelta(days=6, hours=23, minutes=59, seconds=59)
    return start_of_last_week, end_of_last_week


def apply_style(cell, style_dict):
    """Apply formatting styles to an Excel cell."""
    if not style_dict: 
        return
    try:
        cell.font = copy(style_dict['font'])
        cell.border = copy(style_dict['border'])
        cell.fill = copy(style_dict['fill'])
        cell.number_format = copy(style_dict['number_format'])
        cell.alignment = copy(style_dict['alignment'])
    except: 
        pass


# =============================================================================
# STEP 1: USER INPUT
# =============================================================================

def get_user_inputs():
    """Prompt user for report parameters."""
    print("--- Weekly Report Generator ---\n")
    
    # Ask for source sheet
    target_year = input("Enter Source Tab Name (default '2026'): ").strip()
    if not target_year: 
        target_year = "2026"
    
    # Ask for reference date
    date_input = input("Enter Reference Date (YYYY-MM-DD, default Today): ").strip()
    if date_input:
        try: 
            ref_date = datetime.strptime(date_input, "%Y-%m-%d")
        except ValueError: 
            print("Invalid date. Using Today.")
            ref_date = datetime.today()
    else: 
        ref_date = datetime.today()
    
    return target_year, ref_date


# =============================================================================
# STEP 2: LOAD DATA
# =============================================================================

def load_ticket_data(target_year):
    """Load and combine ticket data from multiple sheets."""
    print("\n=== Loading Data ===")
    
    # Load both 2025 and 2026 sheets
    sheets_to_load = list(dict.fromkeys([target_year, "2025", "2026"]))
    df_list = []
    
    for sheet in sheets_to_load:
        try:
            df_temp = pd.read_excel(SOURCE_FILE, sheet_name=sheet)
            df_temp.columns = df_temp.columns.astype(str).str.strip()
            df_list.append(df_temp)
            print(f"  ✓ Loaded sheet: {sheet} ({len(df_temp)} rows)")
        except Exception as e:
            print(f"  ✗ Sheet '{sheet}' not found")
    
    if not df_list:
        raise Exception("No data sheets found")
    
    # Combine and remove duplicates
    df_tickets = pd.concat(df_list, ignore_index=True)
    df_tickets = df_tickets.drop_duplicates()
    print(f"  → Total rows: {len(df_tickets)}")
    
    return df_tickets


def normalize_date_column(df):
    """Find and normalize the date column to 'Date Created'."""
    # Try exact matches first
    for c in df.columns:
        if c.lower() in ['date', 'created', 'opened', 'start date']:
            df.rename(columns={c: 'Date Created'}, inplace=True)
            return df
    
    # Try substring matches
    for c in df.columns:
        if 'date' in c.lower():
            df.rename(columns={c: 'Date Created'}, inplace=True)
            return df
    
    # Fallback to column index 1
    if len(df.columns) > 1:
        print("  Warning: Using column index 1 as date")
        df.rename(columns={df.columns[1]: 'Date Created'}, inplace=True)
    
    if 'Date Created' not in df.columns:
        raise Exception("Could not find date column")
    
    df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
    return df


def find_status_column(df):
    """Find the status column in the dataframe."""
    for c in df.columns:
        if 'status' in c.lower(): 
            return c
    return None


# =============================================================================
# STEP 3: FILTER DATA
# =============================================================================

def filter_open_tickets(df, end_date, status_col):
    """Filter to show only OPEN tickets (exclude future tickets)."""
    print("\n=== Filtering Tickets ===")
    
    # Date check: exclude future tickets
    mask_date_valid = df['Date Created'] <= end_date
    
    # Status check: must be OPEN
    if status_col:
        status_clean = df[status_col].astype(str).str.strip().str.lower()
        mask_open = (status_clean == "open")
    else:
        mask_open = pd.Series([False] * len(df))
    
    # Combine filters
    filtered = df[mask_open & mask_date_valid]
    print(f"  → Found {len(filtered)} open tickets")
    
    return filtered


def calculate_week_statistics(df, start_date, end_date, status_col):
    """Calculate open/closed counts for tickets created in the week."""
    mask_week = (df['Date Created'] >= start_date) & (df['Date Created'] <= end_date)
    df_week = df[mask_week]
    
    if status_col:
        stats_status = df_week[status_col].astype(str).str.strip().str.lower()
        count_closed = stats_status[stats_status.str.contains('close')].count()
        count_open = stats_status[stats_status == 'open'].count()
    else:
        count_closed = 0
        count_open = 0
    
    return count_closed, count_open


# =============================================================================
# STEP 4: COLUMN MAPPING
# =============================================================================

def build_column_mapping(df_tickets, template_headers):
    """Map source columns to template columns."""
    print("\n=== Mapping Columns ===")
    
    def get_col(df, candidates):
        """Find first matching column from candidates list."""
        for c in candidates:
            if c in df.columns: 
                return c
        return None
    
    map_tickets = {}
    
    # Special mappings: Source -> Template
    special_mappings = {
        'ServiceNow Ticket #': get_col(df_tickets, ['Ticket No.', 'Ticket No', 'Ticket']),
        'Description': 'Request Detail' if 'Request Detail' in df_tickets.columns else None,
        'Type': 'Type' if 'Type' in df_tickets.columns else None,
        'PIC': 'Assign To' if 'Assign To' in df_tickets.columns else None,
        'Received': 'Time - Arrive' if 'Time - Arrive' in df_tickets.columns else None,
        'Resolved': 'Time - Close' if 'Time - Close' in df_tickets.columns else None,
    }
    
    # Standard mappings (same names with fallbacks)
    standard_mappings = {
        'REQ No.': get_col(df_tickets, ['REQ No.', 'REQ No', 'Req No.']),
        'Contact': get_col(df_tickets, ['Contact', 'Requester']),
        'Team member': get_col(df_tickets, ['Team member', 'PIC', 'Team Member']),
        'Date Created': 'Date Created',
        'Remarks': get_col(df_tickets, ['Remarks'])
    }
    
    # Apply special and standard mappings
    for template_col_idx, template_header in template_headers.items():
        template_header_clean = template_header.strip()
        
        # Check special mappings first
        if template_header_clean in special_mappings:
            source_col = special_mappings[template_header_clean]
            if source_col:
                map_tickets[template_col_idx] = source_col
                print(f"  {template_header} <- {source_col}")
        
        # Check standard mappings
        elif template_header_clean in standard_mappings:
            source_col = standard_mappings[template_header_clean]
            if source_col:
                map_tickets[template_col_idx] = source_col
                print(f"  {template_header} <- {source_col}")
        
        # Auto-map if exact name match
        elif template_header_clean in df_tickets.columns:
            map_tickets[template_col_idx] = template_header_clean
            print(f"  {template_header} <- {template_header_clean} (auto)")
    
    return map_tickets


# =============================================================================
# STEP 5: EXCEL OPERATIONS
# =============================================================================

def clean_and_get_styles(wb, sheet_name, start_row):
    """Clean old data from sheet and extract cell styles."""
    styles = {}
    if sheet_name not in wb.sheetnames: 
        return styles
    
    ws = wb[sheet_name]
    
    # Extract styles from row 3
    for col in range(1, 27):
        cell = ws.cell(row=start_row, column=col)
        styles[col] = {
            'font': copy(cell.font),
            'border': copy(cell.border),
            'fill': copy(cell.fill),
            'number_format': copy(cell.number_format),
            'alignment': copy(cell.alignment)
        }
    
    # Clean data areas
    max_row = ws.max_row
    if sheet_name == TARGET_SHEET_TICKETS:
        # Clean data table (H-Z) from row 3
        if max_row >= start_row:
            for row in range(start_row, max_row + 1):
                for col in range(8, 27):
                    try: 
                        ws.cell(row=row, column=col).value = None
                    except: 
                        pass
        
        # Clean dashboard list (A-G) from row 9
        for row in range(9, max_row + 1):
            for col in range(1, 8):
                try: 
                    ws.cell(row=row, column=col).value = None
                except: 
                    pass
    else:
        # Clean full sheet for other sheets
        if max_row >= start_row:
            for row in range(start_row, max_row + 1):
                for col in range(1, 27):
                    try: 
                        ws.cell(row=row, column=col).value = None
                    except: 
                        pass
    
    return styles


def write_dashboard_headers(ws, start_date, end_date, count_closed, count_open):
    """Write dashboard headers with statistics."""
    ws['C3'] = "IFS AMS Workload Summary"
    ws['D5'] = "For NAFTA Marelli USA"
    
    # Clear old values
    ws['D7'].value = None
    ws['E7'].value = None
    
    # Write period string with stats
    period_str = f"Period: {start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}   {count_closed} closed, {count_open} open"
    ws['D7'] = period_str
    
    print(f"\n=== Dashboard ===")
    print(f"  {period_str}")


def write_tickets_to_excel(ws, filtered_tickets, column_mapping, styles, start_row):
    """Write filtered ticket data to Excel sheet."""
    print("\n=== Writing Tickets ===")
    
    row_idx = start_row
    for _, row in filtered_tickets.iterrows():
        for col_idx in range(8, 27):  # Columns H-Z
            cell = ws.cell(row=row_idx, column=col_idx)
            apply_style(cell, styles.get(col_idx))
            
            if col_idx in column_mapping and column_mapping[col_idx]:
                val = row.get(column_mapping[col_idx])
                if val is not None: 
                    cell.value = val
        
        row_idx += 1
    
    print(f"  → Wrote {len(filtered_tickets)} tickets")
    return row_idx


def process_new_users(wb, source_file, start_date, end_date, user_styles):
    """Process and write new user data."""
    print("\n=== Processing New Users ===")
    
    df_users = pd.read_excel(source_file, sheet_name=SOURCE_SHEET_USERS)
    df_users.columns = df_users.columns.astype(str).str.strip()
    
    # Find date column
    for c in df_users.columns:
        if 'date' in c.lower():
            df_users.rename(columns={c: 'Date Created'}, inplace=True)
            break
    
    if 'Date Created' not in df_users.columns and len(df_users.columns) > 1:
        df_users.rename(columns={df_users.columns[1]: 'Date Created'}, inplace=True)
    
    df_users['Date Created'] = pd.to_datetime(df_users['Date Created'], errors='coerce')
    
    # Filter by date and category
    mask_date = (df_users['Date Created'] >= start_date) & (df_users['Date Created'] <= end_date)
    
    cat_col = None
    for c in df_users.columns:
        if 'category' in c.lower(): 
            cat_col = c
            break
    
    if cat_col:
        mask_cat = df_users[cat_col].astype(str).str.strip().str.lower() == "create account"
    else:
        mask_cat = pd.Series([True] * len(df_users))
    
    filtered_users = df_users[mask_date & mask_cat]
    print(f"  → Found {len(filtered_users)} users")
    
    # Write to Excel
    ws_users = wb[TARGET_SHEET_USERS]
    row_idx = DATA_START_ROW
    for _, row in filtered_users.iterrows():
        for col_idx in range(1, 20):
            cell = ws_users.cell(row=row_idx, column=col_idx)
            apply_style(cell, user_styles.get(col_idx))
        for i, val in enumerate(row):
            cell = ws_users.cell(row=row_idx, column=i+1)
            cell.value = val
        row_idx += 1
    
    return row_idx


def apply_autofilter(wb, last_row_tickets, last_row_users):
    """Apply AutoFilter to both sheets."""
    print("\n=== Applying Filters ===")
    
    last_row_map = {
        TARGET_SHEET_TICKETS: last_row_tickets - 1,
        TARGET_SHEET_USERS: last_row_users - 1
    }
    
    for ws_name, last_data_row in last_row_map.items():
        if ws_name not in wb.sheetnames: 
            continue
        
        ws = wb[ws_name]
        
        # Remove Excel Table objects
        if hasattr(ws, 'tables') and ws.tables:
            for tbl_name in list(ws.tables.keys()):
                del ws.tables[tbl_name]
        
        # Apply standard AutoFilter
        fill_max_row = max(last_data_row, 2)
        ws.auto_filter.ref = f"A2:P{fill_max_row}"
        print(f"  ✓ {ws_name}: A2:P{fill_max_row}")


# =============================================================================
# MAIN PROGRAM
# =============================================================================

def main():
    """
    Main program flow:
    1. Get user inputs
    2. Load and prepare data
    3. Filter tickets
    4. Create output Excel file
    5. Write data and apply formatting
    """
    
    print("\n" + "="*60)
    print("           WEEKLY REPORT GENERATOR")
    print("="*60)
    
    try:
        # STEP 1: Get user inputs
        target_year, ref_date = get_user_inputs()
        start_date, end_date = get_last_week_date_range(ref_date)
        print(f"\nWeek: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
        
        # Validate files exist
        if not os.path.exists(SOURCE_FILE):
            print(f"ERROR: Source file '{SOURCE_FILE}' not found")
            return
        if not os.path.exists(TEMPLATE_FILE):
            print(f"ERROR: Template file '{TEMPLATE_FILE}' not found")
            return
        
        # STEP 2: Load data
        df_tickets = load_ticket_data(target_year)
        df_tickets = normalize_date_column(df_tickets)
        status_col = find_status_column(df_tickets)
        
        # STEP 3: Filter data
        filtered_tickets = filter_open_tickets(df_tickets, end_date, status_col)
        count_closed, count_open = calculate_week_statistics(df_tickets, start_date, end_date, status_col)
        
        # STEP 4: Create output file
        output_filename = f"Weekly_Report_{end_date.strftime('%Y-%m-%d')}.xlsx"
        print(f"\n=== Creating Output ===")
        print(f"  File: {output_filename}")
        
        # Copy template (handle file being open)
        while True:
            try:
                shutil.copyfile(TEMPLATE_FILE, output_filename)
                break
            except PermissionError:
                print(f"\n  ERROR: '{output_filename}' is open.")
                print("  Please CLOSE the file and press Enter...")
                input()
        
        # STEP 5: Write data to Excel
        wb = openpyxl.load_workbook(output_filename)
        
        # Clean sheets and get styles
        ifs_styles = clean_and_get_styles(wb, TARGET_SHEET_TICKETS, DATA_START_ROW)
        user_styles = clean_and_get_styles(wb, TARGET_SHEET_USERS, DATA_START_ROW)
        
        # Write dashboard
        ws_ifs = wb[TARGET_SHEET_TICKETS]
        write_dashboard_headers(ws_ifs, start_date, end_date, count_closed, count_open)
        
        # Build column mapping
        template_headers = {}
        for col_idx in range(1, 27):
            header_cell = ws_ifs.cell(row=2, column=col_idx)
            if header_cell.value:
                template_headers[col_idx] = str(header_cell.value).strip()
        
        column_mapping = build_column_mapping(df_tickets, template_headers)
        
        # Write tickets
        last_row_tickets = write_tickets_to_excel(ws_ifs, filtered_tickets, column_mapping, ifs_styles, DATA_START_ROW)
        
        # Write users
        last_row_users = process_new_users(wb, SOURCE_FILE, start_date, end_date, user_styles)
        
        # Apply filters
        apply_autofilter(wb, last_row_tickets, last_row_users)
        
        # STEP 6: Save file
        print("\n=== Saving File ===")
        while True:
            try:
                wb.save(output_filename)
                print("  ✓ SUCCESS")
                break
            except PermissionError:
                print("  ERROR: File is open. Close it and press Enter...")
                input()
        
        print("\n" + "="*60)
        print(f"Report generated: {output_filename}")
        print("="*60 + "\n")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

