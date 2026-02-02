"""
SLA Breach Data Management API
月別SLAデータの保存・取得・削除
"""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import List, Optional
from datetime import datetime
import json
import os
from pathlib import Path

router = APIRouter(prefix="/api/sla", tags=["SLA Data"])

# Data file path
DATA_DIR = Path(__file__).parent.parent.parent / "data"
SLA_DATA_FILE = DATA_DIR / "sla_breach_data.json"


class SLABreach(BaseModel):
    id: Optional[int] = None
    ticket_no: str
    requested_for: str
    description: str
    percentage: str
    elapsed_time: str
    remarks: str
    created_at: Optional[str] = None


class MonthSLAData(BaseModel):
    year: int
    month: int
    breaches: List[SLABreach]


def ensure_data_file():
    """データファイルの存在確認・初期化"""
    DATA_DIR.mkdir(exist_ok=True)
    if not SLA_DATA_FILE.exists():
        with open(SLA_DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump({"monthly_data": {}}, f, ensure_ascii=False, indent=2)


def load_data() -> dict:
    """JSONデータ読み込み"""
    ensure_data_file()
    try:
        with open(SLA_DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {"monthly_data": {}}


def save_data(data: dict):
    """JSONデータ保存"""
    ensure_data_file()
    with open(SLA_DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_month_key(year: int, month: int) -> str:
    """月キー生成"""
    return f"{year}-{str(month).zfill(2)}"


@router.get("/{year}/{month}")
async def get_sla_data(year: int, month: int):
    """指定月のSLAデータ取得"""
    data = load_data()
    month_key = get_month_key(year, month)
    month_data = data["monthly_data"].get(month_key, {"breaches": []})
    return {
        "year": year,
        "month": month,
        "breaches": month_data.get("breaches", [])
    }


@router.post("/{year}/{month}")
async def save_sla_data(year: int, month: int, sla_data: MonthSLAData):
    """指定月のSLAデータ保存"""
    data = load_data()
    month_key = get_month_key(year, month)
    
    # Assign IDs and timestamps
    breaches = []
    for i, breach in enumerate(sla_data.breaches):
        breach_dict = breach.dict()
        breach_dict["id"] = i + 1
        if not breach_dict.get("created_at"):
            breach_dict["created_at"] = datetime.now().isoformat()
        breaches.append(breach_dict)
    
    data["monthly_data"][month_key] = {
        "year": year,
        "month": month,
        "breaches": breaches,
        "updated_at": datetime.now().isoformat()
    }
    
    save_data(data)
    return {"status": "success", "message": f"Saved {len(breaches)} SLA breaches for {month_key}"}


@router.delete("/{year}/{month}/{breach_id}")
async def delete_sla_breach(year: int, month: int, breach_id: int):
    """指定SLA Breach削除"""
    data = load_data()
    month_key = get_month_key(year, month)
    
    if month_key not in data["monthly_data"]:
        raise HTTPException(status_code=404, detail="Month data not found")
    
    month_data = data["monthly_data"][month_key]
    breaches = month_data.get("breaches", [])
    
    # Filter out the breach to delete
    new_breaches = [b for b in breaches if b.get("id") != breach_id]
    
    if len(new_breaches) == len(breaches):
        raise HTTPException(status_code=404, detail="Breach not found")
    
    # Reassign IDs
    for i, breach in enumerate(new_breaches):
        breach["id"] = i + 1
    
    month_data["breaches"] = new_breaches
    month_data["updated_at"] = datetime.now().isoformat()
    
    save_data(data)
    return {"status": "success", "message": f"Deleted breach {breach_id}"}


@router.get("/all")
async def get_all_sla_data():
    """全SLAデータ取得"""
    data = load_data()
    return data["monthly_data"]
