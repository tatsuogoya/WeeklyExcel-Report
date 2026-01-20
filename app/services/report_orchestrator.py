from datetime import date
from typing import Dict, Any
from app.infra.excel_repository import load_excel_data
from app.services.report_parser import process_report_data

def get_weekly_report_data(
    daily_excel_path: str,
    begin_date: date,
    end_date: date,
) -> Dict[str, Any]:
    """
    Orchestration Service: Coordinates between Infrastructure and Logic.
    1. Loads raw sheets via Infra.
    2. Processes/Filters/Summarizes via Service Logic.
    """
    # 1. Load raw data from Infra
    all_sheets = load_excel_data(daily_excel_path)
    
    # 2. Process data via pure service logic
    data = process_report_data(all_sheets, begin_date, end_date)
    
    return data
