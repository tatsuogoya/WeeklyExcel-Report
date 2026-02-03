"""
📊 IFS Weekly Report Generator - Streamlit Web App
==================================================

Generate weekly ServiceNow reports for NAFTA Marelli USA

Features:
- Upload Excel files (NA Daily work.xlsx and Template)
- Select date range
- Generate formatted weekly reports
- Download generated reports
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
from report_generator import WeeklyReportGenerator

# Page configuration
st.set_page_config(
    page_title="IFS Weekly Report Generator",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 IFS Weekly Report Generator")
st.markdown("Generate weekly ServiceNow reports for NAFTA Marelli USA")

# Sidebar for configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # Date selection
    ref_date = st.date_input(
        "Reference Date",
        value=datetime.today(),
        help="Select a date to determine the week range (last week from this date)"
    )
    
    # Source sheet selection
    source_sheet = st.text_input(
        "Source Sheet Name",
        value="2025",
        help="Name of the sheet in the source Excel file (e.g., '2025', '2026')"
    )
    
    st.markdown("---")
    st.markdown("### 📝 Instructions")
    st.markdown("""
    1. Upload your **NA Daily work.xlsx** file
    2. Upload the **SNOW_report_Template.xlsx** file
    3. Select a reference date
    4. Click **Generate Report**
    5. Download the generated report
    """)

# Main content area
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 Upload Source Data")
    uploaded_file = st.file_uploader(
        "NA Daily work.xlsx",
        type=['xlsx'],
        help="Upload the Excel file containing ServiceNow ticket data"
    )
    
    if uploaded_file:
        st.success(f"✅ Uploaded: {uploaded_file.name}")

with col2:
    st.subheader("📋 Upload Template")
    template_file = st.file_uploader(
        "SNOW_report_Template.xlsx",
        type=['xlsx'],
        help="Upload the report template file"
    )
    
    if template_file:
        st.success(f"✅ Uploaded: {template_file.name}")

# Generate button
st.markdown("---")

if st.button("🚀 Generate Report", type="primary", use_container_width=True):
    if not uploaded_file:
        st.error("❌ Please upload the source data file (NA Daily work.xlsx)")
    elif not template_file:
        st.error("❌ Please upload the template file (SNOW_report_Template.xlsx)")
    else:
        try:
            with st.spinner("Generating report... This may take a few moments."):
                # Initialize report generator
                generator = WeeklyReportGenerator()
                
                # Convert date to datetime
                ref_datetime = datetime.combine(ref_date, datetime.min.time())
                
                # Generate report
                output_buffer = generator.generate_report(
                    source_file=uploaded_file,
                    template_file=template_file,
                    reference_date=ref_datetime,
                    source_sheet=source_sheet
                )
                
                # Success message
                st.success("✅ Report generated successfully!")
                
                # Calculate week range for filename
                start_date, end_date = generator.get_last_week_date_range(ref_datetime)
                output_filename = f"Weekly_Report_{end_date.strftime('%Y-%m-%d')}.xlsx"
                
                # Download button
                st.download_button(
                    label="📥 Download Report",
                    data=output_buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Display summary
                st.markdown("---")
                st.markdown("### 📊 Report Summary")
                st.info(f"""
                **Report Period**: {start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}
                
                **Source Sheet**: {source_sheet}
                
                **Generated**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                """)
                
        except Exception as e:
            st.error(f"❌ Error generating report: {str(e)}")
            st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>IFS Weekly Report Generator v1.0</p>
    <p>For NAFTA Marelli USA</p>
</div>
""", unsafe_allow_html=True)
