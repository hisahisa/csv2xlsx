mod types; // types.rs を読み込む
mod logic; // logic.rs を読み込む

use types::ColType;
use logic::write_field;
use pyo3::prelude::*;
use rust_xlsxwriter::{Workbook, Format};
use std::collections::{HashMap, HashSet};
use std::error::Error;

const HEADER_ROW: u32 = 1;


// 内部ロジック
fn write_csv_to_excel_inner(
    csv_path: &str,
    excel_path: &str,
    define_output: &str,
) -> Result<(), Box<dyn Error>> {

    // 定義文字列から列の型情報を生成
    let col_types = logic::parse_column_types(define_output);

    // CSVファイルを読み込むためのリーダーを準備
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_path(csv_path)?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 日付フォーマットの事前生成
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // ユニーク値保持用
    let mut column_unique_items: HashMap<u16, HashSet<u8>> = HashMap::new();
    let mut max_row_idx = 0;

    // イテレータを取得
    let mut records = rdr.records().enumerate();

    // ヘッダー処理
    logic::write_header_rows(worksheet, &mut records, HEADER_ROW)?;

    // データ行の処理
    for (row_idx, result) in records {
        let record = result?;
        let current_row = row_idx as u32;
        max_row_idx = current_row;

        for (col_idx, field) in record.iter().enumerate() {
            let c_idx = col_idx as u16;

            // enum ColType取得
            let col_type = col_types.get(col_idx).unwrap_or(&ColType::Str);

            // Excelへの書き出し
            write_field(
                worksheet,
                current_row,
                c_idx,
                field,
                col_type,
                &date_format,
                &mut column_unique_items, // 可変参照で渡す
            )?;
        }
    }

    // 収集したユニーク値からドロップダウンリストを適用
    logic::apply_column_validations(worksheet, column_unique_items, HEADER_ROW, max_row_idx)?;

    workbook.save(excel_path)?;
    Ok(())
}

#[pyfunction]
fn convert(csv_path: &str, excel_path: &str, define_output: &str) -> PyResult<()> {
    write_csv_to_excel_inner(csv_path, excel_path, define_output).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(e.to_string())
    })
}

#[pymodule]
fn csv_converter(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(convert, m)?)?;
    Ok(())
}
