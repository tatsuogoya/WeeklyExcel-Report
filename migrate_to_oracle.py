"""
Excel to Oracle Data Migration Tool
===================================
Migrates ticket data from 'NA Daily work.xlsx' to the Oracle 'IFS_TICKETS' table.
"""

import pandas as pd
import oracledb
import os
from datetime import datetime

# --- Database Configuration ---
# You can move these to environment variables or a config file later
DB_USER = "YOUR_USERNAME"
DB_PASS = "YOUR_PASSWORD"
DB_DSN = "YOUR_HOST:PORT/SERVICE_NAME"

SOURCE_FILE = "NA Daily work.xlsx"
SHEETS_TO_IMPORT = ["2025", "2026"]

def get_db_connection():
    """Establish connection to Oracle database."""
    try:
        conn = oracledb.connect(
            user=DB_USER,
            password=DB_PASS,
            dsn=DB_DSN
        )
        return conn
    except Exception as e:
        print(f"❌ Database connection failed: {e}")
        return None

def migrate_data():
    """Main migration logic."""
    print("--- Starting Excel to Oracle Migration ---")
    
    if not os.path.exists(SOURCE_FILE):
        print(f"❌ Source file '{SOURCE_FILE}' not found.")
        return

    conn = get_db_connection()
    if not conn:
        return
    
    cursor = conn.cursor()
    
    total_inserted = 0
    total_updated = 0

    for sheet in SHEETS_TO_IMPORT:
        print(f"\nProcessing sheet: {sheet}...")
        try:
            df = pd.read_excel(SOURCE_FILE, sheet_name=sheet)
            df.columns = df.columns.astype(str).str.strip()
            
            # Basic validation/normalization
            # (Similar to generate_report.py logic)
            # ...
            
            for index, row in df.iterrows():
                ticket_no = str(row.get('Ticket No.', '')).strip()
                if not ticket_no or ticket_no == 'nan':
                    continue
                
                # Check if ticket exists (for Upsert)
                cursor.execute("SELECT COUNT(*) FROM IFS_TICKETS WHERE TICKET_NO = :1", [ticket_no])
                exists = cursor.fetchone()[0] > 0
                
                # Prepare data (mapping)
                # Note: Adjust column names based on your EXACT Excel headers
                data = {
                    "ticket_no": ticket_no,
                    "status": str(row.get('Status', '')),
                    "contact": str(row.get('Contact', '')),
                    "team_member": str(row.get('Team member', '')),
                    # Add more mappings here...
                }

                if exists:
                    # UPDATE logic
                    # sql = "UPDATE IFS_TICKETS SET STATUS = :status ... WHERE TICKET_NO = :ticket_no"
                    # cursor.execute(sql, data)
                    total_updated += 1
                else:
                    # INSERT logic
                    # sql = "INSERT INTO IFS_TICKETS (TICKET_NO, STATUS, ...) VALUES (:ticket_no, :status, ...)"
                    # cursor.execute(sql, data)
                    total_inserted += 1

            conn.commit()
            print(f"  ✓ Sheet {sheet} processed.")

        except Exception as e:
            print(f"  ❌ Error processing sheet {sheet}: {e}")

    cursor.close()
    conn.close()
    
    print("\n" + "="*40)
    print(f"Migration Complete!")
    print(f"Inserted: {total_inserted}")
    print(f"Updated:  {total_updated}")
    print("="*40)

if __name__ == "__main__":
    # Note: You need to install python-oracledb: pip install oracledb
    print("⚠️ This is a template. Please update DB_USER, DB_PASS, and DB_DSN.")
    # migrate_data()
