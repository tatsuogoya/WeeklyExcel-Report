from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import os

def render_workload_report(
    path: str,
    logo_path: str,
    title: str,
    subtitle: str,
    summary_items: list,
    sections: list
):
    """
    Pure I/O: Renders a PDF using ReportLab primitives.
    No business logic, sorting, or filtering.
    """
    doc = SimpleDocTemplate(path, pagesize=landscape(A4), 
                             rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    # Define Styles
    blue_header_color = colors.dodgerblue
    
    title_style = ParagraphStyle(
        name='Title', parent=styles['Heading1'], fontSize=20, alignment=2,
        textColor=blue_header_color, spaceAfter=15
    )
    subtitle_style = ParagraphStyle(
        name='Subtitle', parent=styles['Normal'], fontSize=12, alignment=2, spaceAfter=15
    )
    heading2_style = ParagraphStyle(
        name='Heading2', parent=styles['Heading2'], textColor=blue_header_color
    )
    heading3_style = ParagraphStyle(
        name='Heading3', parent=styles['Heading3'], textColor=blue_header_color,
        spaceBefore=12, spaceAfter=6
    )

    blue_table_theme = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), blue_header_color),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.aliceblue),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ])

    # 1. Branding
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=1.5*inch, height=0.5*inch)
        logo.hAlign = 'LEFT'
        story.append(logo)
        story.append(Spacer(1, -0.4 * inch))

    story.append(Paragraph(title, title_style))
    story.append(Paragraph(subtitle, subtitle_style))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.green, spaceBefore=5, spaceAfter=20))
    story.append(Paragraph("Weekly Status Report", heading2_style))
    story.append(Spacer(1, 0.2 * inch))

    # 2. Summary Info
    for item in summary_items:
        story.append(Paragraph(item, styles['Normal']))
    story.append(Spacer(1, 0.3 * inch))

    # 3. Dynamic Sections
    for i, section in enumerate(sections):
        story.append(Paragraph(section['title'], heading3_style))
        if section['data']:
            # data is list of lists of Paragraphs
            t = Table(section['data'], colWidths=section.get('widths'), repeatRows=1)
            t.setStyle(blue_table_theme)
            story.append(t)
        else:
            story.append(Paragraph(section.get('empty_msg', 'No data found.'), styles['Normal']))
        
        if i < len(sections) - 1:
            story.append(PageBreak())

    doc.build(story)
    return path
