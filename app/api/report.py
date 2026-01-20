from fastapi import APIRouter, UploadFile, File, Form, HTTPException, status
from fastapi.responses import FileResponse, JSONResponse
from datetime import date
import shutil
import os
from tempfile import NamedTemporaryFile
from app.services.report_orchestrator import get_weekly_report_data

router = APIRouter()

def validate_request(file: UploadFile, begin_date: str, end_date: str):
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "INVALID_FILE_EXTENSION", "message": "Only .xlsx files are allowed"}
        )
    
    try:
        begin_dt = date.fromisoformat(begin_date)
        end_dt = date.fromisoformat(end_date)
    except ValueError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "INVALID_DATE_FORMAT", "message": "Dates must be in YYYY-MM-DD format"}
        )
    
    if begin_dt > end_dt:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"error_code": "INVALID_DATE_RANGE", "message": "begin_date must be before or equal to end_date"}
        )
    return begin_dt, end_dt

from fastapi.encoders import jsonable_encoder

import pandas as pd

@router.post("/generate")
async def generate_json_report(
    begin_date: str = Form(...),
    end_date: str = Form(...),
    file: UploadFile = File(...)
):
    begin_dt, end_dt = validate_request(file, begin_date, end_date)
    
    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        shutil.copyfileobj(file.file, tmp_file)
        tmp_path = tmp_file.name

    try:
        data = get_weekly_report_data(tmp_path, begin_dt, end_dt)
        
        # Clean DataFrames: Replace NaT/NaN with None for JSON serialization
        def clean_df(df):
            return df.astype(object).where(pd.notnull(df), None).to_dict(orient="records")

        # Convert to JSON-serializable format
        result = {
            "summary": data["summary"],
            "left_section": clean_df(data["left_df"]),
            "right_section": clean_df(data["right_df"]),
            "new_users_section": clean_df(data["new_users_df"]),
            "period": {
                "begin": begin_dt.isoformat(),
                "end": end_dt.isoformat()
            }
        }
        return JSONResponse(content=jsonable_encoder(result))
    except Exception as e:
        print(f"Error in generate_json_report: {str(e)}")
        import traceback
        traceback.print_exc()
        if isinstance(e, HTTPException): raise e
        raise HTTPException(status_code=500, detail={"error_code": "INTERNAL_ERROR", "message": str(e)})
    finally:
        if os.path.exists(tmp_path): os.remove(tmp_path)

@router.post("/pdf", response_class=FileResponse)
async def generate_pdf_report(
    begin_date: str = Form(...),
    end_date: str = Form(...),
    file: UploadFile = File(...)
):
    begin_dt, end_dt = validate_request(file, begin_date, end_date)
    
    with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
        shutil.copyfileobj(file.file, tmp_file)
        tmp_path = tmp_file.name

    try:
        data = get_weekly_report_data(tmp_path, begin_dt, end_dt)
        
        # We need a PDF generation service
        from app.services.pdf_service import generate_pdf_service
        output_pdf_path = generate_pdf_service(data, begin_dt, end_dt)

        filename = f"alphast_SNOW_report_{end_date.replace('-', '_')}.pdf"
        return FileResponse(
            path=output_pdf_path,
            filename=filename,
            media_type="application/pdf"
        )
    except Exception as e:
        print(f"Error in generate_pdf_report: {str(e)}")
        import traceback
        traceback.print_exc()
        if isinstance(e, HTTPException): raise e
        raise HTTPException(status_code=500, detail={"error_code": "INTERNAL_ERROR", "message": str(e)})
    finally:
        if os.path.exists(tmp_path): os.remove(tmp_path)
