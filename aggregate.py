# -*- coding: utf-8 -*-
"""
売上データ集計・可視化スクリプト
---------------------------------
raw/ にある複数月の売上CSV (sales_YYYY-MM.csv) をまとめて読み込み、
月次・カテゴリ別・店舗別に集計して、Excelレポートとグラフ画像を出力する。

出力:
  output/売上集計.xlsx     … 4シート (月次推移 / カテゴリ別 / 店舗別 / 明細)
  output/月次売上推移.png   … 月次売上の棒グラフ + カテゴリ別構成

使い方:
  pip install pandas openpyxl matplotlib
  python aggregate.py
"""

import sys
from pathlib import Path

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd

# Windowsのコンソールで日本語が文字化けしないようUTF-8で出力する
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# グラフの日本語フォント (Windows標準フォントを優先指定)
matplotlib.rcParams["font.family"] = ["Yu Gothic", "MS Gothic", "Meiryo", "sans-serif"]
matplotlib.rcParams["axes.unicode_minus"] = False

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "raw"
OUT_DIR = BASE / "output"
OUT_XLSX = OUT_DIR / "売上集計.xlsx"
OUT_CHART = OUT_DIR / "月次売上推移.png"


def yen(n) -> str:
    """整数を ¥1,234,567 形式に整形 (コンソール表示用)。"""
    return f"¥{int(n):,}"


def load_all() -> pd.DataFrame:
    """raw/ 内のすべての売上CSVを1つのDataFrameに統合する。"""
    files = sorted(RAW_DIR.glob("sales_*.csv"))
    if not files:
        raise SystemExit(f"売上CSVが見つかりません: {RAW_DIR}")

    frames = []
    for f in files:
        df = pd.read_csv(f)
        frames.append(df)
    print(f"読み込んだファイル: {len(files)} 件 ({', '.join(f.name for f in files)})")

    data = pd.concat(frames, ignore_index=True)
    data["日付"] = pd.to_datetime(data["日付"])
    data["月"] = data["日付"].dt.strftime("%Y-%m")
    data["売上金額"] = data["数量"] * data["単価"]
    return data


def main():
    data = load_all()
    print(f"明細行数: {len(data)} 行 / 期間: {data['月'].min()}〜{data['月'].max()}")

    # --- 集計 ---
    monthly = (
        data.groupby("月")
        .agg(売上金額=("売上金額", "sum"), 数量=("数量", "sum"))
        .reset_index()
    )

    by_category = (
        data.groupby("商品カテゴリ")["売上金額"].sum().sort_values(ascending=False).reset_index()
    )
    total = by_category["売上金額"].sum()
    by_category["構成比%"] = (by_category["売上金額"] / total * 100).round(1)

    by_store = (
        data.groupby("店舗")["売上金額"].sum().sort_values(ascending=False).reset_index()
    )

    # --- コンソール要約 ---
    print("\n[月次売上]")
    for _, r in monthly.iterrows():
        print(f"  {r['月']}: {yen(r['売上金額'])}  (数量 {int(r['数量']):,})")
    print(f"  合計: {yen(total)}")
    print("\n[カテゴリ別]")
    for _, r in by_category.iterrows():
        print(f"  {r['商品カテゴリ']}: {yen(r['売上金額'])}  ({r['構成比%']}%)")
    print("\n[店舗別]")
    for _, r in by_store.iterrows():
        print(f"  {r['店舗']}: {yen(r['売上金額'])}")

    # --- Excel出力 (複数シート) ---
    OUT_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        monthly.to_excel(writer, sheet_name="月次推移", index=False)
        by_category.to_excel(writer, sheet_name="カテゴリ別", index=False)
        by_store.to_excel(writer, sheet_name="店舗別", index=False)
        data.drop(columns="日付").to_excel(writer, sheet_name="明細", index=False)

    # --- グラフ出力 ---
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # 左: 月次売上推移 (棒グラフ + 値ラベル)
    bars = ax1.bar(monthly["月"], monthly["売上金額"], color="#4C72B0")
    ax1.set_title("月次売上推移")
    ax1.set_ylabel("売上金額 (円)")
    ax1.bar_label(bars, labels=[f"¥{v:,.0f}" for v in monthly["売上金額"]], padding=3, fontsize=9)
    ax1.margins(y=0.15)

    # 右: カテゴリ別売上 (横棒グラフ)
    cat = by_category.sort_values("売上金額")
    ax2.barh(cat["商品カテゴリ"], cat["売上金額"], color="#55A868")
    ax2.set_title("カテゴリ別売上")
    ax2.set_xlabel("売上金額 (円)")
    for i, v in enumerate(cat["売上金額"]):
        ax2.text(v, i, f" ¥{v:,.0f}", va="center", fontsize=9)
    ax2.margins(x=0.18)

    fig.tight_layout()
    fig.savefig(OUT_CHART, dpi=120)
    plt.close(fig)

    print(f"\nExcelレポート: {OUT_XLSX}")
    print(f"グラフ画像   : {OUT_CHART}")


if __name__ == "__main__":
    main()
