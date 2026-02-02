from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
import io
import os
import json
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepTogether, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from app.services.monthly_report_service import MonthlyReportService
from app.infra.chart_renderer import ChartRenderer

router = APIRouter(prefix="/monthly", tags=["monthly"])

report_service = MonthlyReportService()
# static/charts への書き出しを想定
chart_renderer = ChartRenderer(output_dir="static/charts")

@router.post("/process")
async def process_monthly_report(
    file: UploadFile = File(...),
    year: int = Form(...),
    month: int = Form(...)
):
    try:
        # Excelファイルの読み込み
        content = await file.read()
        # すべてのシートを読み込む（2025, 2026, New Usersなどを統合するため）
        xls = pd.ExcelFile(io.BytesIO(content))
        
        all_dfs = []
        for sheet_name in xls.sheet_names:
            # 月報の集計に使用するのはチケットデータが載っている年別のシートのみ
            if sheet_name in ["2024", "2025", "2026"]:
                df_tmp = pd.read_excel(xls, sheet_name=sheet_name)
                df_tmp.columns = df_tmp.columns.astype(str).str.strip()
                all_dfs.append(df_tmp)
        
        if not all_dfs:
            raise HTTPException(status_code=400, detail="Required ticket sheets (2024, 2025, or 2026) not found in Excel.")
            
        df = pd.concat(all_dfs, ignore_index=True).drop_duplicates()
        print(f"DEBUG: Combined ticket data rows: {len(df)}")
        
        # 1. 月間集計 (Pivot Table用)
        monthly_stats = report_service.aggregate_monthly_data(df, year, month)
        if "error" in monthly_stats:
            return JSONResponse(status_code=404, content=monthly_stats)
            
        # 2. 年間サマリー集計 (Stacked Bar Chart用)
        # 選択された月までのデータのみを表示（未来の月は表示しない）
        annual_summary = report_service.get_annual_summary_data(df, year, target_month=month)

        # 3. グラフ描画
        # 集計結果を元に画像ファイルを作成
        pie_path = chart_renderer.render_monthly_pie(monthly_stats["pivot_data"], year, month)
        bar_path = chart_renderer.render_annual_summary(annual_summary, year)

        # 4. SLAデータ読み込み（チェックボックス状態に関わらず自動読み込み）
        sla_data = []
        sla_file = Path(__file__).parent.parent.parent / "data" / "sla_breach_data.json"
        if sla_file.exists():
            with open(sla_file, "r", encoding="utf-8") as f:
                all_sla = json.load(f)
                month_key = f"{year}-{str(month).zfill(2)}"
                if month_key in all_sla.get("monthly_data", {}):
                    sla_data = all_sla["monthly_data"][month_key].get("breaches", [])
        
        print(f"DEBUG: Auto-loaded SLA data for {year}-{month:02d}: {len(sla_data)} breaches")

        # 5. Development Effortsデータ読み込み（年間・チェックボックス状態に関わらず自動読み込み）
        dev_efforts_data = {}
        dev_efforts_file = Path(__file__).parent.parent.parent / "data" / "dev_efforts_data.json"
        if dev_efforts_file.exists():
            with open(dev_efforts_file, "r", encoding="utf-8") as f:
                all_dev = json.load(f)
                # 年間の全月データを取得
                for m in range(1, 13):
                    mk = f"{year}-{str(m).zfill(2)}"
                    if mk in all_dev.get("monthly_data", {}):
                        dev_efforts_data[m] = all_dev["monthly_data"][mk]
        
        print(f"DEBUG: Auto-loaded Dev Efforts data for {year}: {len(dev_efforts_data)} months")

        # レスポンス構築
        return {
            "success": True,
            "year": year,
            "month": month,
            "summary": monthly_stats["summary"],
            "pivot_data": monthly_stats["pivot_data"],
            "annual_summary": annual_summary, # UIでのテーブル表示用
            "charts": {
                "pie": f"/static/charts/{os.path.basename(pie_path)}",
                "bar": f"/static/charts/{os.path.basename(bar_path)}"
            },
            "sla_data": sla_data,
            "dev_efforts_data": dev_efforts_data
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/pdf")
async def generate_monthly_pdf(
    file: UploadFile = File(...),
    year: int = Form(...),
    month: int = Form(...)
):
    """Generate PDF for Monthly Report including SLA Breach data"""
    try:
        # まずJSONデータを処理
        content = await file.read()
        xls = pd.ExcelFile(io.BytesIO(content))

        all_dfs = []
        for sheet_name in xls.sheet_names:
            if sheet_name in ["2024", "2025", "2026"]:
                df_tmp = pd.read_excel(xls, sheet_name=sheet_name)
                df_tmp.columns = df_tmp.columns.astype(str).str.strip()
                all_dfs.append(df_tmp)

        if not all_dfs:
            raise HTTPException(status_code=400, detail="Required sheets not found")

        df = pd.concat(all_dfs, ignore_index=True).drop_duplicates()

        # 月間集計
        monthly_stats = report_service.aggregate_monthly_data(df, year, month)
        if "error" in monthly_stats:
            raise HTTPException(status_code=404, detail=monthly_stats["error"])

        # 年間サマリー集計
        # 選択された月までのデータのみを表示（未来の月は表示しない）
        annual_summary = report_service.get_annual_summary_data(df, year, target_month=month)

        # SLAデータ読み込み（チェックボックス状態に関わらず自動読み込み）
        # 指定年月のデータが存在すれば、常にレポートに含める
        sla_data = []
        sla_file = Path(__file__).parent.parent.parent / "data" / "sla_breach_data.json"
        if sla_file.exists():
            with open(sla_file, "r", encoding="utf-8") as f:
                all_sla = json.load(f)
                month_key = f"{year}-{str(month).zfill(2)}"
                if month_key in all_sla.get("monthly_data", {}):
                    sla_data = all_sla["monthly_data"][month_key].get("breaches", [])
        
        print(f"DEBUG: Auto-loaded SLA data for {year}-{month:02d}: {len(sla_data)} breaches")

        # Development Effortsデータ読み込み（年間・チェックボックス状態に関わらず自動読み込み）
        # 指定年のデータが存在すれば、常にレポートに含める
        dev_efforts_data = {}
        dev_efforts_file = Path(__file__).parent.parent.parent / "data" / "dev_efforts_data.json"
        if dev_efforts_file.exists():
            with open(dev_efforts_file, "r", encoding="utf-8") as f:
                all_dev = json.load(f)
                # 年間の全月データを取得
                for m in range(1, 13):
                    mk = f"{year}-{str(m).zfill(2)}"
                    if mk in all_dev.get("monthly_data", {}):
                        dev_efforts_data[m] = all_dev["monthly_data"][mk]
        
        print(f"DEBUG: Auto-loaded Dev Efforts data for {year}: {len(dev_efforts_data)} months")

        # PDF生成
        output_dir = Path("static")
        output_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = str(output_dir / f"Monthly_Report_{year}_{month:02d}.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=letter, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()

        # タイトル
        elements.append(Paragraph(f"<b>Monthly Report - {year}/{month:02d}</b>", styles["Title"]))
        elements.append(Spacer(1, 8))

        # サマリー
        summary = monthly_stats.get("summary", {})
        elements.append(Paragraph(f"Total Tickets: {summary.get('total_tickets', 0)} | Closed: {summary.get('closed_tickets', 0)}", styles["Normal"]))
        elements.append(Spacer(1, 10))

        # Annual Summary Chart
        bar_path = chart_renderer.render_annual_summary(annual_summary, year)
        elements.append(Image(bar_path, width=6.5 * inch, height=3 * inch))
        elements.append(Spacer(1, 8))

        month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        annual_table_data = [["Category"] + month_labels]
        for category in annual_summary.get("categories", []):
            counts = annual_summary.get("data", {}).get(category, [0] * 12)
            annual_table_data.append([category] + counts)

        annual_table = Table(annual_table_data, colWidths=[1.5 * inch] + [0.45 * inch] * 12)
        annual_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#bfdbfe")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (1, 1), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(annual_table)
        elements.append(PageBreak())

        # Monthly Ticket Distribution + KPI Pivot Table (keep on same page)
        page2_elements = []
        page2_elements.append(Paragraph("<b>Monthly Ticket Distribution</b>", styles["Heading2"]))
        page2_elements.append(Spacer(1, 6))
        pie_path = chart_renderer.render_monthly_pie(monthly_stats["pivot_data"], year, month)
        page2_elements.append(Image(pie_path, width=6.5 * inch, height=4.5 * inch))
        page2_elements.append(Spacer(1, 8))

        # Monthly KPI Pivot Table
        page2_elements.append(Paragraph("<b>Monthly KPI Pivot Table</b>", styles["Heading2"]))
        page2_elements.append(Spacer(1, 6))

        pivot_table_data = [["Type", "Category", "TRANS_NON", "TRANS_WORK", "OPEN", "CLOSE"]]
        for key, values in monthly_stats.get("pivot_data", {}).items():
            parts = [p.strip() for p in key.split("|")]
            type_name = parts[0] if len(parts) > 0 else ""
            category_name = parts[1] if len(parts) > 1 else ""
            pivot_table_data.append([
                type_name,
                category_name,
                values.get("TRANS_NON", 0),
                values.get("TRANS_WORK", 0),
                values.get("OPEN", 0),
                values.get("CLOSE", 0)
            ])

        pivot_table = Table(pivot_table_data, colWidths=[1.2 * inch, 2.0 * inch, 0.9 * inch, 0.9 * inch, 0.8 * inch, 0.8 * inch])
        pivot_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#bfdbfe")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ALIGN", (2, 1), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
        ]))
        page2_elements.append(pivot_table)
        elements.append(KeepTogether(page2_elements))
        elements.append(PageBreak())

        # SLA Breachセクション (keep title and table together)
        sla_elements = []
        sla_elements.append(Paragraph("<b>SLA Breach</b>", styles["Heading2"]))
        sla_elements.append(Spacer(1, 6))

        if sla_data:
            sla_table_data = [["Ticket No.", "Requested for", "Description", "Time-Arrive", "Business elapsed time", "Remarks"]]
            for breach in sla_data:
                sla_table_data.append([
                    breach.get("ticket_no", ""),
                    breach.get("requested_for", ""),
                    breach.get("description", ""),
                    f"{year}/{month}/27",
                    f"{breach.get('elapsed_time', '')}({breach.get('percentage', '')}%)",
                    breach.get("remarks", "")
                ])

            sla_table = Table(sla_table_data, colWidths=[1.0 * inch, 1.0 * inch, 1.0 * inch, 0.7 * inch, 1.8 * inch, 2.8 * inch])
            sla_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#bfdbfe")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("WORDWRAP", (0, 0), (-1, -1), True),
            ]))
            sla_elements.append(sla_table)
        else:
            sla_elements.append(Paragraph("No SLA breach data for this month.", styles["Normal"]))
        
        elements.append(KeepTogether(sla_elements))
        elements.append(Spacer(1, 12))

        # Development Effortsセクション
        elements.append(Paragraph("<b>Development Efforts</b>", styles["Heading2"]))
        elements.append(Spacer(1, 10))

        if dev_efforts_data:
            month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total"]
            
            # ME Hours row
            me_hours_row = [str(year)]
            me_total = 0
            for m in range(1, 13):
                val = dev_efforts_data.get(m, {}).get("me_hours", "")
                me_hours_row.append(str(val) if val != "" else "")
                if val != "":
                    me_total += float(val)
            me_hours_row.append(str(me_total) if me_total > 0 else "")
            
            # SOW Planned row (fixed at 16)
            sow_row = ["SOW Planned"]
            sow_total = 0
            for m in range(1, 13):
                if dev_efforts_data.get(m):
                    sow_row.append("16")
                    sow_total += 16
                else:
                    sow_row.append("16")
                    sow_total += 16
            sow_row.append(str(sow_total))
            
            # Carry Forward row
            cf_row = ["Carry Forward"]
            cf_total = 0
            for m in range(1, 13):
                val = dev_efforts_data.get(m, {}).get("carry_forward", "")
                cf_row.append(str(val) if val != "" else "")
                if val != "":
                    cf_total += float(val)
            cf_row.append(str(cf_total) if cf_total != 0 else "")
            
            # Build table
            dev_table_data = [
                [str(year)] + month_labels,
                ["ME Hours"] + me_hours_row[1:],
                sow_row,
                cf_row
            ]
            
            col_widths = [1.0 * inch] + [0.55 * inch] * 13
            dev_table = Table(dev_table_data, colWidths=col_widths)
            dev_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#bfdbfe")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1e3a5f")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]))
            elements.append(dev_table)
        else:
            elements.append(Paragraph("No Development Efforts data for this year.", styles["Normal"]))

        doc.build(elements)

        return FileResponse(
            pdf_path,
            media_type="application/pdf",
            filename=f"Monthly_Report_{year}_{month:02d}.pdf"
        )

    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
