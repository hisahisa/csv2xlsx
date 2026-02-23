import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# 設定
num_rows = 1000000
chains = {
    "C01": "あおぞらあいうえおあいうえおあいうえおあいうえおあいうえおマーケット",
    "C02": "スマイルかきくけこかきくけこかきくけこかきくけこかきくけこかきくけこかきくけこマート",
    "C03": "キッチン・さしすせそさしすせそさしすせそさしすせそさしすせそ・ハブ",
    "C04": "エブリデイ・たちつてとたちつてとたちつてとたちつてとたちつてとたちつてと・ロープライス",
    "C05": "ナチュラル・はひふへほはひふへほはひふへほはひふへほはひふへほはひふへほ・フーズ"
}

# 商品名リスト（適当な商品名）
item_list = [
    "プレミアム食パンプレミアム食パンプレミアム食パンプレミアム食パンプレミアム食パンプレミアム食パン",
    "有機栽培バナナ有機栽培バナナ有機栽培バナナ有機栽培バナナ有機栽培バナナ有機栽培バナナ",
    "特選和牛カルビ特選和牛カルビ特選和牛カルビ特選和牛カルビ特選和牛カルビ特選和牛カルビ特選和牛カルビ",
    "しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳しぼりたて牛乳",
    "新鮮朝採れレタス新鮮朝採れレタス新鮮朝採れレタス新鮮朝採れレタス新鮮朝採れレタス",
    "十勝産こだわり小豆十勝産こだわり小豆十勝産こだわり小豆十勝産こだわり小豆",
    "深煎りブレンドコーヒー深煎りブレンドコーヒー深煎りブレンドコーヒー深煎りブレンドコーヒー深煎りブレンドコーヒー",
    "北海道バター北海道バター北海道バター北海道バター北海道バター",
    "完熟トマトケチャップ完熟トマトケチャップ完熟トマトケチャップ完熟トマトケチャップ完熟トマトケチャップ",
    "無添加ポテトチップス無添加ポテトチップス無添加ポテトチップス無添加ポテトチップス無添加ポテトチップス"
]

# 店舗マスターの作成
store_master = []
for c_cd, c_name in chains.items():
    for i in range(1, 21):  # 各チェーン20店舗
        s_cd = f"{c_cd}-{i:03d}"
        s_name = f"{c_name} {i}号店智"
        store_master.append([c_cd, c_name, s_cd, s_name])

# ランダムサンプリングでベースを作成
indices = np.random.choice(len(store_master), num_rows)
df = pd.DataFrame([store_master[i] for i in indices],
                  columns=['chain_cd', 'chain_name', 'store_cd', 'store_name'])

# --- item列の追加 ---
df['item'] = np.random.choice(item_list, size=num_rows)

# --- kbn列 (0, 1, 2 / 1, 2, 3 のランダム値) ---
df['kbn1'] = np.random.randint(0, 3, size=num_rows)
df['kbn2'] = np.random.randint(1, 4, size=num_rows)

# 日付・数量・金額の生成
start_date = datetime(2023, 1, 1)
random_days = np.random.randint(0, 700, size=num_rows)
df['sales_date'] = [start_date + timedelta(days=int(d)) for d in random_days]

df['qty'] = np.random.randint(1, 15, size=num_rows)
df['amt'] = df['qty'] * np.random.randint(100, 5000, size=num_rows)

# カラムの並び順を指定
df = df[['chain_cd', 'chain_name', 'store_cd', 'store_name', 'item', 'kbn1', 'kbn2', 'sales_date', 'qty', 'amt']]

# ソート
df = df.sort_values(['sales_date', 'chain_cd', 'store_cd', 'kbn1']).reset_index(drop=True)

# CSV保存
output_file = 'sales_data_1m_with_kbn.csv'
df.to_csv(output_file, index=False, encoding='utf-8-sig')
print(f"完了: {output_file} を作成しました。")
