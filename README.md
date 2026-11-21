# CSV Sales Report

売上データのCSVを月別・店舗別・カテゴリ別に集計し、グラフ付きのExcelレポートを作成するサンプルプロジェクトです。

## 必要な環境

- Python 3
- `pandas`
- `openpyxl`

## セットアップ

プロジェクトのフォルダで、依存ライブラリをインストールします。

```bash
python3 -m pip install -r requirements.txt
```

## 使い方

サンプル売上データを生成する場合:

```bash
python3 make_dummy_data.py
```

`sales.csv` を読み込み、`monthly_report.xlsx` を作成する場合:

```bash
python3 aggregate.py
```

別のCSVを入力する場合は、ファイルパスを指定できます。

```bash
python3 aggregate.py path/to/sales.csv
```

## 入力データ

`sales.csv` はサンプルデータです。`make_dummy_data.py` を実行すると、同じファイルが再生成されます。

## 出力

`aggregate.py` は `monthly_report.xlsx` を作成します。レポートには月次サマリー、店舗別、カテゴリ別、商品別ランキングが含まれます。
