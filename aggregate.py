"""sales.csv を集計して、グラフ付きのExcelレポートを自動生成する"""
import sys
import pandas as pd
from openpyxl.chart import BarChart, Reference

INPUT = sys.argv[1] if len(sys.argv) > 1 else "sales.csv"
OUTPUT = "monthly_report.xlsx"


def load(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, parse_dates=["日付"])
    # データの簡易チェック(欠損・不正値は除外して件数を表示)
    before = len(df)
    df = df.dropna()
    df = df[(df["数量"] > 0) & (df["単価"] > 0)]
    if before != len(df):
        print(f"不正な行を {before - len(df)} 件除外しました")
    df["売上"] = df["数量"] * df["単価"]
    df["月"] = df["日付"].dt.strftime("%Y-%m")
    return df


def summarize(df: pd.DataFrame) -> dict:
    return {
        "月次サマリー": df.groupby("月").agg(
            売上合計=("売上", "sum"),
            注文数=("注文ID", "nunique"),
            販売数量=("数量", "sum"),
        ),
        "店舗別": df.pivot_table(index="月", columns="店舗", values="売上",
                              aggfunc="sum", fill_value=0),
        "カテゴリ別": df.pivot_table(index="月", columns="カテゴリ", values="売上",
                                aggfunc="sum", fill_value=0),
        "商品別ランキング": df.groupby("商品名").agg(
            売上合計=("売上", "sum"), 販売数量=("数量", "sum")
        ).sort_values("売上合計", ascending=False),
    }


def add_chart(ws, n_rows: int, n_cols: int, title: str, anchor: str):
    chart = BarChart()
    chart.title = title
    chart.y_axis.title = "売上(円)"
    data = Reference(ws, min_col=2, max_col=1 + n_cols, min_row=1, max_row=1 + n_rows)
    cats = Reference(ws, min_col=1, min_row=2, max_row=1 + n_rows)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, anchor)


def main():
    df = load(INPUT)
    sheets = summarize(df)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        for name, table in sheets.items():
            table.to_excel(writer, sheet_name=name)
            ws = writer.sheets[name]
            ws.column_dimensions["A"].width = 22
            for col in "BCDE":
                ws.column_dimensions[col].width = 14
            for row in ws.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    cell.number_format = "#,##0"

        # グラフを追加
        add_chart(writer.sheets["店舗別"], len(sheets["店舗別"]),
                  len(sheets["店舗別"].columns), "店舗別の月次売上", "G2")
        add_chart(writer.sheets["カテゴリ別"], len(sheets["カテゴリ別"]),
                  len(sheets["カテゴリ別"].columns), "カテゴリ別の月次売上", "G2")

    total = int(df["売上"].sum())
    print(f"{OUTPUT} を作成しました(全{len(df)}件 / 売上合計 {total:,}円)")


if __name__ == "__main__":
    main()
