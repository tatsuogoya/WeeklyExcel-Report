import matplotlib
matplotlib.use('Agg') # サーバー側での描画用にバックエンドを固定
import matplotlib.pyplot as plt
import pandas as pd
import os
from typing import Dict, List, Any
import numpy as np

class ChartRenderer:
    """
    月報用のグラフ描画を担当するクラス。
    SOWに基づき、特定のカラーパレットを使用して高品質な商用グラフを生成する。
    """
    
    # SOW Section 3.3 に基づくカラー設定
    CATEGORY_COLORS = {
        "Miscellaneous": "#000080",            # Navy
        "Development": "#FF8C00",              # DarkOrange
        "Transfer to another group": "#008080", # Teal
        "Permissions control": "#800080",       # Purple
        "Create account": "#808000",            # Olive
        "Reset password": "#A52A2A",            # Brown
        "Deactivate account": "#0000FF"         # Blue
    }

    def __init__(self, output_dir: str = "static/charts"):
        self.output_dir = output_dir
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def render_annual_summary(self, data: Dict[str, Any], year: int) -> str:
        """
        年間サマリーの積み上げ棒グラフを生成する。
        常に1月〜12月の全ての月を表示し、データがある月のみ棒グラフを描画。
        """
        months = data["months"]  # 実際にデータがある月のリスト（例：[1]なら1月のみ）
        categories = data["categories"]
        stats_data = data["data"]

        # 常に12ヶ月分のスペースを確保
        fig, ax = plt.subplots(figsize=(10, 5))
        
        # 常に12ヶ月分のラベルを表示
        all_month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        # 12ヶ月分のゼロ配列を準備
        bottom = np.zeros(12)

        # カテゴリごとに積み上げ
        for cat in categories:
            # 12ヶ月分の配列を作成（デフォルトは0）
            counts_full = [0] * 12
            
            # データがある月のみ値を設定
            cat_data = stats_data.get(cat, [])
            for i, month in enumerate(months):
                if i < len(cat_data):
                    counts_full[month - 1] = cat_data[i]  # month-1でインデックスに変換
            
            color = self.CATEGORY_COLORS.get(cat, "#808080")
            ax.bar(all_month_labels, counts_full, bottom=bottom, label=cat, color=color)
            bottom += np.array(counts_full)

        # 装飾
        ax.set_title(f"Annual Summary - {year}", fontsize=16, fontweight='bold', pad=20)
        ax.set_ylabel("Ticket Count", fontsize=12)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        
        # 月ラベルのフォントサイズを調整
        ax.tick_params(axis='x', labelsize=10)
        ax.tick_params(axis='y', labelsize=10)

        # x軸の範囲を12ヶ月分に固定
        ax.set_xlim(-0.6, 11.6) 
        plt.subplots_adjust(left=0.12, right=0.98, top=0.88, bottom=0.1)

        # レイアウト確定
        
        output_path = os.path.join(self.output_dir, f"annual_summary_{year}.png")
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        return output_path

    def render_monthly_pie(self, pivot_data: Dict[str, Any], year: int, month: int) -> str:
        """
        月間のカテゴリ配分を示すパイチャートを生成する。
        """
        # カテゴリごとの合計を算出
        cat_totals = {}
        target_cols = ["TRANS_NON", "TRANS_WORK", "OPEN", "CLOSE"]
        
        for key, values in pivot_data.items():
            # key format: "Type | Category"
            parts = key.split('|')
            cat = parts[1].strip() if len(parts) > 1 else parts[0].strip()
            
            # 各ステータスの合計を算出 ('Grand Total'が削除されたため手動)
            row_total = sum(values.get(c, 0) for c in target_cols)
            cat_totals[cat] = cat_totals.get(cat, 0) + row_total

        # 集計データが空、または合計が0の場合のハンドリング
        if not cat_totals or sum(cat_totals.values()) == 0:
            # 描画エラーを避けるため、空の画像を生成するか、例外的に処理
            labels = ["No Data"]
            sizes = [1] # ダミー
            colors = ["#e2e8f0"] # Gray
        else:
            labels = list(cat_totals.keys())
            sizes = list(cat_totals.values())
            colors = [self.CATEGORY_COLORS.get(label, "#808080") for label in labels]

        fig, ax = plt.subplots(figsize=(12, 10)) # サイズを拡大
        
        # パイチャートの描画設定の改善
        wedges, texts, autotexts = ax.pie(
            sizes,
            autopct='%1.1f%%',
            startangle=140, 
            colors=colors,
            pctdistance=0.75,
            textprops={'fontsize': 18, 'fontweight': 'bold'}
        )

        # 内部の数値（パーセント）の色を白に変更し、サイズを大幅にアップ
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(24)

        # 凡例をグラフの下部に移動し、フォントサイズを拡大
        ax.legend(wedges, labels,
                  title="Categories",
                  title_fontsize=16,
                  loc="upper center",
                  bbox_to_anchor=(0.5, -0.05),
                  ncol=min(len(labels), 3),
                  fontsize=14)

        ax.set_title(f"Total incidents and SRs - {year}/{month:02d}", fontsize=28, fontweight='bold', pad=40)
        ax.axis('equal') 

        # レイアウトの微調整（凡例が下で切れないように）
        plt.subplots_adjust(left=0.1, right=0.9, top=0.85, bottom=0.2)

        output_path = os.path.join(self.output_dir, f"monthly_pie_{year}_{month:02d}.png")
        plt.savefig(output_path, dpi=300, bbox_inches='tight') # 余白を自動調整して保存
        plt.close()

        return output_path
