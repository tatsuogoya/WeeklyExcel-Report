import sys
import os
from datetime import date
import pandas as pd

# Add the project root to sys.path
sys.path.append(os.getcwd())

from app.services.pdf_generator import generate_pdf

# Mock data
data = {
    "summary": {"open_count": 5, "closed_count": 10},
    "left_df": pd.DataFrame({
        "Ticket No.": ["T1", "T2"],
        "Request Detail": ["Detail 1", "Detail 2"],
        "Remarks": ["Remark 1", "Remark 2"],
        "Time - Arrive": [pd.Timestamp("2026-01-01"), pd.Timestamp("2026-01-02")],
        "Status": ["OPEN", "OPEN"]
    }),
    "right_df": pd.DataFrame({
        "Ticket No.": ["T1", "T2", "T3"],
        "Status": ["OPEN", "OPEN", "CLOSE"],
        "REQ No.": ["R1", "R2", "R3"],
        "Type": ["Bug", "Task", "Bug"],
        "Request Detail": ["Detail 1", "Detail 2", "Detail 3"],
        "Requested for": ["User 1", "User 2", "User 3"],
        "Assign To": ["PIC 1", "PIC 2", "PIC 3"],
        "Time - Arrive": [pd.Timestamp("2026-01-01"), pd.Timestamp("2026-01-02"), pd.Timestamp("2026-01-03")],
        "Time - Close": [pd.NaT, pd.NaT, pd.Timestamp("2026-01-04")]
    })
}

begin_dt = date(2026, 1, 1)
end_dt = date(2026, 1, 7)

try:
    print("Generating PDF...")
    pdf_path = generate_pdf(data, begin_dt, end_dt)
    print(f"PDF generated successfully: {pdf_path}")
except Exception as e:
    print(f"Error generating PDF: {str(e)}")
    import traceback
    traceback.print_exc()
