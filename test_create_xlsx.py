import csv_converter
import time

csv_file = "sales_data_500k_with_kbn.csv"
excel_file = "test.xlsx"
# ドロップダウンのリストとして使うCSVファイル（同じファイルを使う場合は csv_file を指定）

print(f"Converting {csv_file} to {excel_file} ...")

define_output = '''
str,
str,
str,
str,
kbn_list,
date,
int,
int,
'''

try:
    start_time = time.time()

    # 3つ目の引数を追加
    csv_converter.convert(csv_file, excel_file, define_output)

    end_time = time.time()
    print(f"変換完了: {end_time - start_time:.2f} 秒")

except Exception as e:
    print(f"エラーが発生しました: {e}")
