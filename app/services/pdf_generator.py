from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import date
import tempfile
import pandas as pd
import os

def generate_pdf(data: dict, begin_date: date, end_date: date) -> str:
    """
    Generates a PDF report using reportlab.
    Returns the path to the temporary PDF file.
    """
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)

    # Use landscape for more horizontal space
    doc = SimpleDocTemplate(path, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    # 1. Logo and Title
    logo_path = os.path.join(os.path.dirname(__file__), "../../static/logo.jpg")
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=1.5*inch, height=0.5*inch)
        logo.hAlign = 'LEFT'
        story.append(logo)
        story.append(Spacer(1, -0.4 * inch)) # Pull title up near logo level if possible or just space it

    title_style = ParagraphStyle(
        name='Title',
        parent=styles['Heading1'],
        fontSize=20, # Reduced a bit to fit logo alignment
        alignment=2, # Right aligned to contrast with logo
        textColor=colors.dodgerblue,
        spaceAfter=15
    )
    story.append(Paragraph("IFS AMS Workload Summary", title_style))
    subtitle_style = ParagraphStyle(
        name='Subtitle',
        parent=styles['Normal'],
        fontSize=12,
        alignment=2, # Right aligned to match title
        spaceAfter=15
    )
    story.append(Paragraph("For NAFTA Marelli USA", subtitle_style))
    heading2_style = ParagraphStyle(
        name='Heading2',
        parent=styles['Heading2'],
        textColor=colors.dodgerblue
    )
    story.append(HRFlowable(width="100%", thickness=2, color=colors.green, spaceBefore=5, spaceAfter=20))
    story.append(Paragraph("Weekly Status Report", heading2_style))
    story.append(Spacer(1, 0.2 * inch))

    heading3_style = ParagraphStyle(
        name='Heading3',
        parent=styles['Heading3'],
        textColor=colors.dodgerblue,
        spaceBefore=12,
        spaceAfter=6
    )

    # 2. Summary Section
    summary = data["summary"]
    period_str = f"<b>Period:</b> {begin_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}"
    counts_str = f"<b>Summary:</b> {summary['closed_count']} closed, {summary['open_count']} open."
    
    story.append(Paragraph(period_str, styles['Normal']))
    story.append(Paragraph(counts_str, styles['Normal']))
    story.append(Spacer(1, 0.3 * inch))

    # Section-specific column mappings
    LEFT_COLUMNS = [
        ("Ticket #", "Ticket No."),
        ("Description", "Request Detail"),
        ("Remarks", "Remarks")
    ]
    
    RIGHT_COLUMNS = [
        ("Ticket #", "Ticket No."),
        ("Status", "Status"),
        ("REQ No.", "REQ No."),
        ("Type", "Type"),
        ("Description", "Request Detail"),
        ("Requested for", "Requested for"),
        ("PIC", "Assign To"),
        ("Received", "Time - Arrive"),
        ("Resolved", "Time - Close")
    ]

    def format_val(val):
        if pd.isna(val) or val is None:
            return ""
        if isinstance(val, (date, pd.Timestamp)):
            return val.strftime('%m/%d/%Y')
        return str(val)

    def create_table_data(df, mapping, bold_keys=None):
        bold_keys = bold_keys or set()
        headers = [h for h, _ in mapping]
        table_data = [headers]
        
        for _, row in df.iterrows():
            line = []
            ticket_no = row.get("Ticket No.")
            is_bold = ticket_no in bold_keys
            for _, source_col in mapping:
                val = format_val(row.get(source_col))
                # Wrap all values in Paragraph to ensure they wrap
                # Use XML tags for bolding and color if ticket is in both sections
                para_val = f'<b><font color="dodgerblue">{val}</font></b>' if is_bold else val
                line.append(Paragraph(para_val, styles['Normal']))
            table_data.append(line)
        return table_data

    # Table styling
    # Unified Blue Theme
    blue_theme = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.dodgerblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ])

    # 3. Weekly Activity
    story.append(Paragraph("Weekly Activity (Date Range Based)", heading3_style))
    left_df = data["left_df"]
    left_ticket_nos = set(left_df["Ticket No."].unique()) if not left_df.empty else set()

    # Weekly Activity table
    if not left_df.empty:
        # Bold all in left section since they are by definition "active"
        left_table_data = create_table_data(left_df, LEFT_COLUMNS, bold_keys=left_ticket_nos)
        # 10.8 inches total. Ticket # (1.5"), Desc (6.0"), Remarks (3.3")
        left_col_widths = [1.5*inch, 6.0*inch, 3.3*inch]
        t1 = Table(left_table_data, colWidths=left_col_widths, repeatRows=1)
        t1.setStyle(blue_theme)
        story.append(t1)
    else:
        story.append(Paragraph("No activity found in this period.", styles['Normal']))
    
    # 4. Page Break before Backlog
    story.append(PageBreak())

    # 5. Weekly Workload Details
    story.append(Paragraph("Weekly Workload Details", heading3_style))
    right_df = data["right_df"]
    if not right_df.empty:
        right_table_data = create_table_data(right_df, RIGHT_COLUMNS, bold_keys=left_ticket_nos)
        # ServiceNow (1.1), Status (0.7), REQ (1.0), Type (0.7), Desc (2.95), ReqFor (1.25), PIC (0.95), Rec (0.8), Res (0.8)
        right_col_widths = [1.1*inch, 0.7*inch, 1.0*inch, 0.7*inch, 3.25*inch, 1.25*inch, 0.95*inch, 0.9*inch, 0.9*inch]
        t2 = Table(right_table_data, colWidths=right_col_widths, repeatRows=1)
        t2.setStyle(blue_theme)
        story.append(t2)
    else:
        story.append(Paragraph("No open tickets found in the backlog.", styles['Normal']))

    # Build PDF
    doc.build(story)
    return path
