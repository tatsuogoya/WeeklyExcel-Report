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
        # Sort by Date Created descending (newest first)
        if "Date Created" in new_users_df.columns:
            new_users_df = new_users_df.sort_values(by="Date Created", ascending=False, na_position='last')
        
        # Dynamically detect columns from the DataFrame
        all_cols = [c for c in new_users_df.columns if c and str(c).lower() != 'nan' and c != '']
        
        # Create dynamic column mapping: (display_name, source_name)
        dynamic_new_users_cols = [(col, col) for col in all_cols]
        
        # Calculate dynamic widths (distribute evenly across page width ~10.8 inches)
        total_width = 10.8 * inch
        num_cols = len(dynamic_new_users_cols)
        if num_cols > 0:
            col_width = total_width / num_cols
            dynamic_widths = [col_width] * num_cols
        else:
            dynamic_widths = []
        
        # Prepare table data with date range highlighting
        def prepare_new_users_table(df, mapping, begin_dt, end_dt):
            table_data = [[h for h, _ in mapping]]  # Header row
            for _, row in df.iterrows():
                line = []
                # Check if this user was created within the date range
                date_created = row.get("Date Created")
                is_in_range = False
                if pd.notna(date_created):
                    try:
                        if isinstance(date_created, (date, pd.Timestamp)):
                            row_date = date_created.date() if hasattr(date_created, 'date') else date_created
                        else:
                            row_date = pd.to_datetime(date_created).date()
                        is_in_range = begin_dt <= row_date <= end_dt
                    except:
                        pass
                
                for _, source_col in mapping:
                    val = format_val(row.get(source_col))
                    if is_in_range:
                        para_val = f'<b><font color="dodgerblue">{val}</font></b>'
                    else:
                        para_val = val
                    line.append(Paragraph(para_val, styles['Normal']))
                table_data.append(line)
            return table_data
        
        sections.append({
            "title": "New Users",
            "data": prepare_new_users_table(new_users_df, dynamic_new_users_cols, begin_date, end_date),
            "widths": dynamic_widths,
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
