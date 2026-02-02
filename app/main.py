from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
import os
from app.api.report import router as report_router
from app.api.monthly_report import router as monthly_router
from app.api.sla_data import router as sla_router
from app.api.dev_efforts import router as dev_efforts_router

app = FastAPI()

# Mount static files
static_dir = os.path.join(os.path.dirname(__file__), "../static")
if not os.path.exists(static_dir):
    os.makedirs(static_dir)

app.mount("/static", StaticFiles(directory=static_dir), name="static")

@app.get("/")
async def read_index():
    index_path = os.path.join(static_dir, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>Welcome to Weekly Excel-Report Generator</h1><p>Frontend not found.</p>")

@app.get("/health")
def health():
    return {"status": "ok"}


app.include_router(report_router)
app.include_router(monthly_router)
app.include_router(sla_router)
app.include_router(dev_efforts_router)

