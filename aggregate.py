"""売上CSVを検証・集計し、確認用シート付きのExcelを出力する。"""

import argparse
import csv
import math
import sys
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter


INPUT = "sales.csv"
OUTPUT = "monthly_report.xlsx"
COLUMNS = ["日付", "注文ID", "店舗", "商品名", "カテゴリ", "数量", "単価"]
ISSUE_COLUMNS = ["ファイル", "行番号", "区分", "項目", "入力値", "理由"]


def issue(path, line, kind, field, value, reason):
    return (str(path), line, kind, field, value, reason)


def positive_number(value):
    number = pd.to_numeric(value, errors="coerce")
    if not math.isfinite(number) or number <= 0:
        raise ValueError("有限の正の数値を指定してください")
    return number.item()


def read_sales(path, issues):
    """CSVの各レコードを検証し、元ファイルの開始行を記録する。"""
    with path.open(encoding="utf-8-sig", newline="") as file:
        reader = csv.reader(file, strict=True)
        header = [name.strip() for name in next((record for record in reader if record), [])]
        missing = [name for name in COLUMNS if name not in header]
        repeated = [name for name in COLUMNS if header.count(name) > 1]
        if missing or repeated:
            details = []
            if missing:
                details.append("必須列がありません: " + "、".join(missing))
            if repeated:
                details.append("列名が重複しています: " + "、".join(repeated))
            raise ValueError(" / ".join(details))

        positions = [header.index(name) for name in COLUMNS]
        rows = []
        rejected = 0
        while True:
            line = reader.line_num + 1
            try:
                values = next(reader)
            except StopIteration:
                break
            except csv.Error as error:
                raise ValueError(f"{line}行目のCSV形式が不正です: {error}") from error

            if not values:
                continue
            if len(values) != len(header):
                issues.append(issue(path, line, "除外", "行全体", "",
                                    f"列数が異なります（期待: {len(header)}、実際: {len(values)}）"))
                rejected += 1
                continue

            row = {name: values[pos].strip() for name, pos in zip(COLUMNS, positions)}
            errors = {name: "必須項目が空です" for name, value in row.items() if not value}
            date = pd.NaT
            if "日付" not in errors:
                date = pd.to_datetime(row["日付"], errors="coerce")
                if pd.isna(date):
                    errors["日付"] = "有効な日付を指定してください（例: 2026-06-01）"

            numbers = {}
            for name in ("数量", "単価"):
                if name not in errors:
                    try:
                        numbers[name] = positive_number(row[name])
                    except ValueError as error:
                        errors[name] = str(error)
            if not errors:
                sales = numbers["数量"] * numbers["単価"]
                if not math.isfinite(sales):
                    errors["数量・単価"] = "売上金額が計算可能な範囲を超えています"
            if errors:
                for name, reason in errors.items():
                    issues.append(issue(path, line, "除外", name, row.get(name, ""), reason))
                rejected += 1
                continue

            row.update(numbers)
            row.update(日付=date.date(), 月=date.strftime("%Y-%m"), 売上=sales)
            rows.append((line, row))

    return rows, rejected


def import_sales(paths):
    rows, results, issues = [], [], []
    seen_paths = set()
    first_occurrences = {}
    for path in paths:
        file_issues = []
        rejected = 0
        duplicates = 0
        accepted = []
        try:
            resolved = path.resolve()
            if resolved in seen_paths:
                status = "同一ファイル指定のためスキップ"
                file_issues.append(issue(path, None, "スキップ", "ファイル", "", status))
            else:
                seen_paths.add(resolved)
                accepted, rejected = read_sales(path, file_issues)
                if not accepted and not rejected:
                    file_issues.append(issue(path, None, "確認", "ファイル", "", "明細がありません"))
                status = "要確認" if file_issues else "完了"
        except (OSError, UnicodeError, ValueError, csv.Error) as error:
            status = "取込失敗"
            file_issues.append(issue(path, None, "エラー", "ファイル", "", str(error)))

        for line, row in accepted:
            key = tuple(row[name] for name in COLUMNS)
            if key in first_occurrences:
                first_path, first_line = first_occurrences[key]
                reason = f"{first_path} の {first_line} 行目と同じ内容です。集計に含めています"
                file_issues.append(issue(path, line, "重複候補", "注文ID", row["注文ID"], reason))
                duplicates += 1
                status = "要確認"
            else:
                first_occurrences[key] = (path, line)
            rows.append(row)

        issues.extend(file_issues)
        failed = status == "取込失敗"
        results.append({
            "ファイル": str(path),
            "状態": status,
            "読込明細数": None if failed else len(accepted) + rejected,
            "集計明細数": len(accepted),
            "除外明細数": None if failed else rejected,
            "重複候補数（集計に含む）": duplicates,
        })

    data = pd.DataFrame(rows, columns=COLUMNS + ["月", "売上"])
    return data, pd.DataFrame(results), pd.DataFrame(issues, columns=ISSUE_COLUMNS)


# 売上データを月次・店舗別・カテゴリ別・商品別に集計する
def summarize(df: pd.DataFrame) -> dict:
    orders = df.assign(店舗内注文=list(zip(df["店舗"], df["注文ID"])))
    result = {
        "月次サマリー": orders.groupby("月").agg(
            売上合計=("売上", "sum"),
            注文数=("店舗内注文", "nunique"),
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
    return result


# 指定されたシートの集計データから棒グラフを作成して配置する
def add_chart(ws, n_rows: int, n_cols: int, title: str, anchor: str):
    # 棒グラフを作成し、グラフのタイトルと縦軸のタイトルを設定
    chart = BarChart()
    chart.title = title
    chart.y_axis.title = "売上(円)"
    # 見出しを含むB列以降をグラフのデータ範囲に指定
    data = Reference(ws, min_col=2, max_col=1 + n_cols, min_row=1, max_row=1 + n_rows)
    # A列の2行目以降を横軸の項目名（月）の範囲に指定
    cats = Reference(ws, min_col=1, min_row=2, max_row=1 + n_rows)
    # データの1行目を系列名として使用し、横軸の項目名を設定
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    # グラフの幅と高さを設定し、指定されたセル位置に配置
    chart.width, chart.height = 18, 9
    ws.add_chart(chart, anchor)


def excel_value(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub(lambda match: repr(match[0])[1:-1], value)
    return value


def write_sheet(writer, name, table, *, index=True):
    # CSV由来の文字列をExcelの数式やエラー値として扱わせない。
    clean = table.rename(index=excel_value, columns=excel_value)
    for column in clean.select_dtypes(include=["object", "string"]):
        clean[column] = clean[column].map(excel_value)
    clean.to_excel(writer, sheet_name=name, index=index)
    ws = writer.sheets[name]
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type in {"f", "e"}:
                cell.data_type = "s"
    return ws


def write_report(output, sheets, results, issues):
    status = "要確認（暫定集計）" if not issues.empty else "確認事項なし"
    with NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as file:
        temporary = Path(file.name)
    try:
        with pd.ExcelWriter(temporary, engine="openpyxl") as writer:
            for name, table in sheets.items():
                ws = write_sheet(writer, name, table)
                ws.column_dimensions["A"].width = 22
                for column in range(2, ws.max_column + 1):
                    ws.column_dimensions[get_column_letter(column)].width = 16
                for row in ws.iter_rows(min_row=2, min_col=2):
                    for cell in row:
                        if isinstance(cell.value, float) and not cell.value.is_integer():
                            cell.number_format = "#,##0.########"
                        else:
                            cell.number_format = "#,##0"
                if name in ("店舗別", "カテゴリ別") and not table.empty:
                    anchor = get_column_letter(max(7, ws.max_column + 2)) + "2"
                    add_chart(ws, len(table), len(table.columns), f"{name}の月次売上", anchor)

            write_sheet(writer, "取込結果", results.assign(レポート状態=status), index=False)
            write_sheet(writer, "要確認", issues, index=False)
            for name, widths in {
                "取込結果": [45, 29, 15, 15, 15, 27, 25],
                "要確認": [45, 12, 15, 18, 30, 68],
            }.items():
                ws = writer.sheets[name]
                ws.freeze_panes = "B2"
                ws.auto_filter.ref = ws.dimensions
                for column, width in enumerate(widths, start=1):
                    ws.column_dimensions[get_column_letter(column)].width = width
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("csv", nargs="*", type=Path, help="入力CSV（複数指定可。既定: sales.csv）")
    parser.add_argument("-o", "--output", type=Path, default=Path(OUTPUT), help="出力先のExcelファイル")
    args = parser.parse_args()
    paths = args.csv or [Path(INPUT)]
    if args.output.suffix.lower() != ".xlsx":
        parser.error("出力先の拡張子には .xlsx を指定してください")
    input_paths = {path.resolve() for path in paths}
    if args.output.resolve() in input_paths:
        parser.error("入力と出力に同じファイルは指定できません")

    data, results, issues = import_sales(paths)
    output = args.output
    sheets = {}
    if data.empty:
        output = output.with_name(f"{output.stem}_import_errors.xlsx")
        if output.resolve() in input_paths:
            parser.error("入力と確認用Excelに同じファイルは指定できません")
    else:
        sheets = summarize(data)
    try:
        write_report(output, sheets, results, issues)
    except OSError as error:
        print(f"Excelの保存に失敗しました: {error}", file=sys.stderr)
        return 1

    if data.empty:
        print(f"集計できる明細がありません。{output} の「要確認」を確認してください。",
              file=sys.stderr)
        return 1

    total = data["売上"].sum()
    print(f"{args.output} を作成しました（集計 {len(data)} 明細 / 売上合計 {total:,} 円）")
    if not issues.empty:
        print("暫定集計です。「取込結果」「要確認」シートを確認してください。", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
