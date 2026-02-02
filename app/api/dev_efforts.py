"""
Development Efforts Data Management API
Monthly ME Hours tracking with SOW Planned and Carry Forward calculation
"""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional
from datetime import datetime
import json
from pathlib import Path

router = APIRouter(prefix="/api/dev-efforts", tags=["Dev Efforts"])

# Data file path
DATA_DIR = Path(__file__).parent.parent.parent / "data"
DEV_EFFORTS_FILE = DATA_DIR / "dev_efforts_data.json"


class DevEffortsData(BaseModel):
    month: str  # Format: "2025-12"
    me_hours: float
    sow_planned: float = 16.0
    carry_forward: float


def ensure_data_file():
    """Ensure data directory and file exist"""
    DATA_DIR.mkdir(exist_ok=True)
    if not DEV_EFFORTS_FILE.exists():
        with open(DEV_EFFORTS_FILE, 'w', encoding='utf-8') as f:
            json.dump({"monthly_data": {}}, f, ensure_ascii=False, indent=2)


def load_data() -> dict:
    """Load JSON data"""
    ensure_data_file()
    try:
        with open(DEV_EFFORTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {"monthly_data": {}}


def save_data(data: dict):
    """Save JSON data"""
    ensure_data_file()
    with open(DEV_EFFORTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@router.get("/{month}")
async def get_dev_efforts(month: str):
    """Get Development Efforts data for specified month"""
    data = load_data()
    month_data = data["monthly_data"].get(month)
    
    if month_data:
        return {
            "month": month,
            "me_hours": month_data.get("me_hours", 0),
            "sow_planned": month_data.get("sow_planned", 16.0),
            "carry_forward": month_data.get("carry_forward", 0)
        }
    return None


@router.post("")
async def save_dev_efforts(dev_data: DevEffortsData):
    """Save Development Efforts data"""
    data = load_data()
    
    # Calculate carry forward
    carry_forward = dev_data.sow_planned - dev_data.me_hours
    
    month_data = {
        "me_hours": dev_data.me_hours,
        "sow_planned": dev_data.sow_planned,
        "carry_forward": carry_forward,
        "last_updated": datetime.now().isoformat()
    }
    
    data["monthly_data"][dev_data.month] = month_data
    save_data(data)
    
    return {
        "status": "success",
        "month": dev_data.month,
        "me_hours": dev_data.me_hours,
        "sow_planned": dev_data.sow_planned,
        "carry_forward": carry_forward
    }


@router.delete("/{month}")
async def delete_dev_efforts(month: str):
    """Delete Development Efforts data for specified month"""
    data = load_data()
    
    if month in data["monthly_data"]:
        del data["monthly_data"][month]
        save_data(data)
        return {"status": "deleted", "month": month}
    
    raise HTTPException(status_code=404, detail="Data not found")
