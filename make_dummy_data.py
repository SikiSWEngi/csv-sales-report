"""ダミーの売上データ(sales.csv)を生成するスクリプト"""
import csv
import random
from datetime import date, timedelta

random.seed(42)  # 毎回同じデータになるように固定

STORES = ["渋谷店", "新宿店", "横浜店"]
PRODUCTS = {  # 商品名: (カテゴリ, 単価)
    "コーヒー豆 200g": ("飲料", 1200),
    "紅茶ティーバッグ": ("飲料", 800),
    "マグカップ": ("雑貨", 1500),
    "タンブラー": ("雑貨", 2800),
    "クッキー詰め合わせ": ("食品", 1800),
    "チョコレート": ("食品", 900),
}

start = date(2026, 6, 1)
days = 150  # 約5か月分

with open("sales.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(["日付", "注文ID", "店舗", "商品名", "カテゴリ", "数量", "単価"])
    order_id = 1
    for d in range(days):
        day = start + timedelta(days=d)
        for _ in range(random.randint(3, 8)):
            store = random.choice(STORES)
            product = random.choice(list(PRODUCTS))
            category, price = PRODUCTS[product]
            qty = random.randint(1, 5)
            writer.writerow([day.isoformat(), f"A{order_id:05d}", store,
                             product, category, qty, price])
            order_id += 1

print("sales.csv を作成しました")
