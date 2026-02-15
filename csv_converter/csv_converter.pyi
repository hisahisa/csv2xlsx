"""
Rust実装による高速CSV-Excel変換モジュールの型定義
"""


def convert(csv_path: str, excel_path: str, define_output: str) -> None:
    """
    CSVファイルを読み込み、指定された型定義に従ってExcelファイルを作成します。
    
    Args:
        csv_path (str): 入力CSVファイルのパス
        excel_path (str): 出力Excelファイルのパス
        define_output (str): 列ごとの型定義（例: "str,int,date,kbn_list"）

    Raises:
        RuntimeError: CSVの読み込み失敗やExcelの保存失敗時に発生
    """
    ...
