import pandas as pd
from typing import Dict, Any, List, Optional
from datetime import datetime

class MonthlyReportService:
    """
    月報（Monthly Report）の集計・ビジネスロジックを担当するサービス。
    SOWに基づき、Incident/Serviceの2段階集計とStatusマッピングを行う。
    """

    # マネジメント指定のカラーパレット（SOW Section 3.3）
    COLOR_PALETTE = {
        "Miscellaneous": "#000080",            # Navy
        "Development": "#FF8C00",              # DarkOrange
        "Transfer to another group": "#008080", # Teal
        "Permissions control": "#800080",       # Purple
        "Create account": "#808000",            # Olive
        "Reset password": "#A52A2A",            # Brown
        "Deactivate account": "#0000FF"         # Blue
    }

    # Status列からレポート用列名へのマッピング（SOW Section 3.2）
    STATUS_MAP = {
        "TRANS_NON": "TRANS_NON",
        "TRANS_WORK": "TRANS_WORK",
        "CANCEL": "CANCEL",
        "OPEN": "OPEN",
        "CLOSE": "CLOSE",
        "PENDING": "OPEN" # 念のためのフォールバック（現場判断）
    }

    def __init__(self):
        pass

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        重要カラム名を標準化するヘルパー関数。
        """
        df = df.copy()
        
        # 1. Date Created (日付)
        if 'Date Created' not in df.columns:
            for c in df.columns:
                if c.lower() in ['date', 'created', 'opened', 'start date'] or 'date' in c.lower():
                    df.rename(columns={c: 'Date Created'}, inplace=True)
                    break
        
        # 5. Ticket No. (チケット番号)
        if 'Ticket No.' not in df.columns:
            # First, look for exact or ticket-related
            found = False
            for c in df.columns:
                if 'ticket' in c.lower():
                    df.rename(columns={c: 'Ticket No.'}, inplace=True)
                    found = True
                    break
            if not found:
                for c in df.columns:
                    if 'no' in c.lower().split() or 'no.' in c.lower() or 'id' == c.lower():
                        df.rename(columns={c: 'Ticket No.'}, inplace=True)
                        break

        # 2. Type (種別: Incident/Service)
        if 'Type' not in df.columns:
            for c in df.columns:
                if 'type' in c.lower() or 'kind' in c.lower() or '種別' in c:
                    df.rename(columns={c: 'Type'}, inplace=True)
                    break

        # 3. Category (カテゴリ)
        if 'Category' not in df.columns:
            for c in df.columns:
                if 'category' in c.lower() or 'cat' == c.lower() or 'カテゴリ' in c:
                    df.rename(columns={c: 'Category'}, inplace=True)
                    break

        # 4. Status (状況)
        if 'Status' not in df.columns:
            for c in df.columns:
                if 'status' in c.lower() or 'state' in c.lower() or '状況' in c:
                    df.rename(columns={c: 'Status'}, inplace=True)
                    break

        return df

    def aggregate_monthly_data(self, df: pd.DataFrame, year: int, month: int) -> Dict[str, Any]:
        """
        指定された月のデータを集計し、ピボットテーブル用のデータを生成する。
        """
        df = self._normalize_columns(df)

        if 'Date Created' not in df.columns:
             return {"error": "Date column not found in Excel. Please ensure columns like 'Date Created' exist."}

        # 日付フィルタリング
        df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
        mask = (df['Date Created'].dt.year == year) & (df['Date Created'].dt.month == month)
        target_df = df[mask].copy()

        if target_df.empty:
            return {"error": f"No data found for {year}-{month:02d}"}

        # 1. 'CANCEL' ステータスおよび 'Cancel' カテゴリのチケットを徹底除外
        target_df['norm_status'] = target_df['Status'].astype(str).str.strip().str.upper()
        target_df = target_df[target_df['norm_status'] != 'CANCEL'].copy()
        
        if 'Category' in target_df.columns:
            target_df['norm_cat'] = target_df['Category'].astype(str).str.strip().str.upper()
            target_df = target_df[target_df['norm_cat'] != 'CANCEL'].copy()

        # 2. カテゴリの統合
        if 'Category' in target_df.columns:
            target_df['Category'] = target_df['Category'].astype(str).str.strip()
            mask_deact = target_df['Category'].str.contains('Deactivate account', case=False) | \
                         target_df['Category'].str.contains('Deactivates account', case=False)
            target_df.loc[mask_deact, 'Category'] = 'Deactivate account'

        # 2段階集計 (Tier 1: Type, Tier 2: Category)
        # Statusをマッピング（大文字小文字を無視するように正規化）
        target_df['norm_status'] = target_df['Status'].astype(str).str.strip().str.upper()
        target_df['report_status'] = target_df['norm_status'].map(self.STATUS_MAP).fillna('OTHER')

        # 必要な列が欠落している場合に備えてダミー行を作成（集計には影響させない）
        # これにより、ピボットテーブルが期待通りの列を持つようにする
        
        # ピボットテーブルの生成
        pivot = pd.pivot_table(
            target_df,
            values='Ticket No.',
            index=['Type', 'Category'],
            columns='report_status',
            aggfunc='count',
            fill_value=0
        )

        # 必要な列が欠落している場合に備えて補完 (CANCELは除外対象のため含めない)
        for col in ["TRANS_NON", "TRANS_WORK", "OPEN", "CLOSE"]:
            if col not in pivot.columns:
                pivot[col] = 0

        # Grand Totalの計算 (User Request: 不要なため削除)
        # pivot['Grand Total'] = pivot.sum(axis=1)

         # JSONシリアライズ用にMultiIndexを文字列フラットキーに変換
        # (Incident, Category) -> "Incident | Category"
        pivot_data_raw = pivot.to_dict(orient='index')
        pivot_data_serializable = {}
        for (tipo, cat), vals in pivot_data_raw.items():
             # 型安全に文字列化
             key = f"{str(tipo)} | {str(cat)}"
             pivot_data_serializable[key] = vals

        return {
            "year": year,
            "month": month,
            "pivot_data": pivot_data_serializable,
            "summary": {
                "total_tickets": len(target_df),
                "closed_tickets": int(pivot['CLOSE'].sum()) if 'CLOSE' in pivot.columns else 0
            }
        }

    def get_annual_summary_data(self, df: pd.DataFrame, year: int, target_month: Optional[int] = None) -> Dict[str, Any]:
        """
        年間サマリー（Stacked Bar Chart）用のデータを月別に集計する。
        - ユーザー指示に基づき 'CANCEL' ステータスを除外。
        - 'Deactivate account' の表記揺れを統合。
        - target_monthが指定された場合、その月までのデータのみを表示（未来の月は表示しない）
        """
        df = self._normalize_columns(df)

        if 'Date Created' not in df.columns:
            return {"year": year, "months": [], "categories": [], "data": {}}

        df['Date Created'] = pd.to_datetime(df['Date Created'], errors='coerce')
            
        # 指定された年で、target_month以前のデータのみをフィルタ
        if target_month is not None:
            mask = (df['Date Created'].dt.year == year) & (df['Date Created'].dt.month <= target_month)
        else:
            mask = (df['Date Created'].dt.year == year)
        annual_df = df[mask].copy()

        # 1. 'CANCEL' ステータスおよび 'Cancel' カテゴリのチケットを完全に除外
        if 'Status' in annual_df.columns:
            annual_df = annual_df[annual_df['Status'].astype(str).str.strip().str.upper() != 'CANCEL']
        
        if 'Category' in annual_df.columns:
            # カテゴリ名自体が 'Cancel' のものも除外 (user: "remove it from the list")
            annual_df = annual_df[annual_df['Category'].astype(str).str.strip().str.upper() != 'CANCEL']

        # 2. カテゴリの正規化・統合 (user: "combine deactivate account")
        if 'Category' in annual_df.columns:
            # "Deactivate account", "Deactivate Account", "Deactivates account" などに対応
            annual_df['Category'] = annual_df['Category'].astype(str).str.strip()
            mask_deact = annual_df['Category'].str.contains('Deactivate account', case=False) | \
                         annual_df['Category'].str.contains('Deactivates account', case=False)
            annual_df.loc[mask_deact, 'Category'] = 'Deactivate account'
        
        # 月別・カテゴリ別の集計
        ts_data = annual_df.groupby([annual_df['Date Created'].dt.month, 'Category']).size().unstack(fill_value=0)
        
        # target_monthが指定されている場合は、1月からtarget_monthまでの月を確保
        # 指定されていない場合は、1月から12月までの全月を確保
        if target_month is not None:
            ts_data = ts_data.reindex(range(1, target_month + 1), fill_value=0)
        else:
            ts_data = ts_data.reindex(range(1, 13), fill_value=0)
        
        # マネジメントが期待する表示順にカテゴリを並び替え（SOW準拠）
        defined_cats = list(self.COLOR_PALETTE.keys())
        existing_cats = [c for c in defined_cats if c in ts_data.columns]
        other_cats = [c for c in ts_data.columns if c not in defined_cats]
        final_cats = existing_cats + other_cats
        
        ts_data = ts_data[final_cats]

        return {
            "year": year,
            "months": list(ts_data.index),  # 実際の月リスト（1からtarget_monthまで）
            "categories": final_cats,
            "data": ts_data.to_dict(orient='list')
        }
