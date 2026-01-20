import pandas as pd
import sys

file_path = "sample_daily_work.xlsx"
try:
    all_sheets = pd.read_excel(file_path, sheet_name=None)
    print(f"Sheets found: {list(all_sheets.keys())}")
    
    new_users_sheet_name = next((name for name in all_sheets.keys() if name.lower() == "new users"), None)
    if new_users_sheet_name:
        print(f"\n--- Data from '{new_users_sheet_name}' sheet ---")
        df = all_sheets[new_users_sheet_name]
        print("\nColumn Names (raw):")
        print(df.columns.tolist())
        
        print("\nFirst 5 rows (raw):")
        print(df.head(5).to_string())
        
        # Check first row if columns are Unnamed
        if any("Unnamed" in str(col) for col in df.columns):
            print("\nColumns starting with 'Unnamed' detected. First row content:")
            print(df.iloc[0].tolist())
            
            # Simulate the normalization logic in report_parser.py
            new_header = dict(zip(df.columns, df.iloc[0]))
            print("\nNew Header mapping:")
            print(new_header)
            
            df_fixed = df.rename(columns=new_header).iloc[1:].copy()
            df_fixed.columns = [str(c).strip() for c in df_fixed.columns]
            print("\nColumns after normalization:")
            print(df_fixed.columns.tolist())
            print("\nFirst 5 rows after normalization:")
            print(df_fixed.head(5).to_string())
    else:
        print("\n'New Users' sheet not found.")
except Exception as e:
    print(f"Error: {e}")
