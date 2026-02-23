import csv_converter
import time
import json

csv_file = "sales_data_1m_with_kbn.csv"
excel_file = "test.xlsx"

print(f"Converting {csv_file} to {excel_file} ...")
define_output_dict = [
    {"width":  8, "col_type": "str"},
    {"width": 50, "col_type": "str"},
    {"width": 10, "col_type": "str"},
    {"width": 70, "col_type": "str"},
    {"width": 80, "col_type": "str"},
    {"width":  5, "col_type": "kbn_list1", "kbn_values": [0, 1, 2]},
    {"width":  5, "col_type": "kbn_list2", "kbn_values": [1, 2, 3]},
    {"width": 12, "col_type": "date"},
    {"width": 10, "col_type": "int"},
    {"width": 10, "col_type": "int"},
]
define_output = json.dumps(define_output_dict)

try:
    start_time = time.time()

    # 3つ目の引数を追加
    csv_converter.convert(csv_file, excel_file, define_output)

    end_time = time.time()
    print(f"変換完了: {end_time - start_time:.2f} 秒")

except Exception as e:
    print(f"エラーが発生しました: {e}")
