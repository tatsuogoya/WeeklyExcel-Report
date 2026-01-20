from datetime import date
from app.services.excel_reader import read_and_process_v2
from typing import Dict, Any

def get_weekly_report_data(
    daily_excel_path: str,
    begin_date: date,
    end_date: date,
) -> Dict[str, Any]:
    """
    Get the processed report data as a dictionary.
    Used for both JSON responses and PDF generation.
    """
    data = read_and_process_v2(daily_excel_path, begin_date, end_date)
    
    # We might want to convert dataframes to list of dicts for JSON serialization
    # but the service layer should probably return clean data.
    return data
