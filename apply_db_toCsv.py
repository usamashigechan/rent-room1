import sqlite3
import pandas as pd
import os

# ===== DBファイルのパス指定 =====
db_path = r"C:\111accommodationDB\apply.db"

# ===== 保存先フォルダ（同じ場所にCSV出力） =====
csv_folder = os.path.dirname(db_path)

# ===== SQLiteデータベースへ接続 =====
conn = sqlite3.connect(db_path)

# ===== 全テーブル名を取得 =====
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()

# ===== 各テーブルをCSVに保存 =====
for table in tables:
    table_name = table[0]
    print(f"処理中テーブル: {table_name}")

    # テーブルを読み込み
    df = pd.read_sql_query(f"SELECT * FROM {table_name};", conn)

    # CSVファイルの保存
    csv_path = os.path.join(csv_folder, f"{table_name}.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
