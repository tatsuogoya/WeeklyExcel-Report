from datetime import date
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import os
from app.infra.pdf_renderer import render_workload_report

def generate_pdf_service(data: dict, begin_date: date, end_date: date) -> str:
    """
    Data Preparation Service: Prepares the DTO and calls the infrastructure renderer.
    Contains formatting logic (e.g., bolding rows) but NO selection/sorting logic.
    """
    styles = getSampleStyleSheet()
    
    # Selection of columns for different sections (moved to service)
    LEFT_COLUMNS = [("Ticket #", "Ticket No."), ("Description", "Request Detail"), ("Remarks", "Remarks")]
    RIGHT_COLUMNS = [
        ("Ticket #", "Ticket No."), ("Status", "Status"), ("REQ No.", "REQ No."), 
        ("Type", "Type"), ("Description", "Request Detail"), ("Requested for", "Requested for"), 
        ("PIC", "Assign To"), ("Received", "Time - Arrive"), ("Resolved", "Time - Close")
    ]
    NEW_USERS_COLUMNS = [
        ("Ticket No", "Ticket No"), ("Date Created", "Date Created"), ("User Name", "User Name"), 
        ("Function / Department", "Function / Department"), ("Email address", "Email address")
    ]

    def format_val(val):
        if pd.isna(val) or val is None: return ""
        if isinstance(val, (date, pd.Timestamp)): return val.strftime('%m/%d/%Y')
        return str(val)

    def prepare_table_data(df, mapping, bold_keys=None):
        bold_keys = bold_keys or set()
        table_data = [[h for h, _ in mapping]] # Header row
        for _, row in df.iterrows():
            line = []
            is_bold = row.get("Ticket No.") in bold_keys
            for _, source_col in mapping:
                val = format_val(row.get(source_col))
                para_val = f'<b><font color="dodgerblue">{val}</font></b>' if is_bold else val
                line.append(Paragraph(para_val, styles['Normal']))
            table_data.append(line)
        return table_data

    # Prepare Data Structures
    left_df = data["left_df"]
    right_df = data["right_df"]
    new_users_df = data.get("new_users_df", pd.DataFrame())
    left_ticket_nos = set(left_df["Ticket No."].unique()) if not left_df.empty else set()

    summary_items = [
        f"<b>Period:</b> {begin_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}",
        f"<b>Summary:</b> {data['summary']['closed_count']} closed, {data['summary']['open_count']} open."
    ]

    sections = [
        {
            "title": "Weekly Activity (Date Range Based)",
            "data": prepare_table_data(left_df, LEFT_COLUMNS, bold_keys=left_ticket_nos),
            "widths": [1.5*inch, 6.0*inch, 3.3*inch],
            "empty_msg": "No activity found in this period."
        },
        {
            "title": "Weekly Workload Details",
            "data": prepare_table_data(right_df, RIGHT_COLUMNS, bold_keys=left_ticket_nos),
            "widths": [1.1*inch, 0.7*inch, 1.0*inch, 0.7*inch, 3.25*inch, 1.25*inch, 0.95*inch, 0.9*inch, 0.9*inch],
            "empty_msg": "No workload details found."
        }
    ]

    # Add New Users section if data exists
    if not new_users_df.empty:
        sections.append({
            "title": "New Users",
            "data": prepare_table_data(new_users_df, NEW_USERS_COLUMNS),
            "widths": [1.2*inch, 1.0*inch, 2.0*inch, 1.5*inch, 2.3*inch],
            "empty_msg": "No new users found."
        })

    # Logo Path (Service uses paths, Infra renders)
    logo_path = os.path.join(os.path.dirname(__file__), "../../static/logo.jpg")

    # Call Infrastructure Renderer
    return render_workload_report(
        path=os.path.join(os.path.dirname(__file__), "../../report.pdf"), # Temporary file location strategy
        logo_path=logo_path,
        title="IFS AMS Workload Summary",
        subtitle="For NAFTA Marelli USA",
        summary_items=summary_items,
        sections=sections
    )
