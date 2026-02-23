mod types;
mod logic;

use types::{ColDefinition};
use pyo3::prelude::*;
use rust_xlsxwriter::{Workbook, Format};
use std::error::Error;

const HEADER_ROW: u32 = 1;

fn write_csv_to_excel_inner(
    csv_path: &str,
    excel_path: &str,
    col_defs: Vec<ColDefinition>,
) -> Result<(), Box<dyn Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 列幅の設定（書き込み前に実行）
    logic::apply_column_settings(worksheet, &col_defs)?;

    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_path(csv_path)?;

    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // ループの前に「最後に処理した行」を Option で持つ
    let mut last_row = None;

    let mut records = rdr.records().enumerate();

    // ヘッダー処理
    logic::write_header_rows(worksheet, &mut records, HEADER_ROW)?;

    // データ行の処理
    for (row_idx, result) in records {
        let record = result?;
        let current_row = row_idx as u32;
        last_row = Some(current_row); // 処理した行を記憶

        for (col_idx, field) in record.iter().enumerate() {
            let c_idx = col_idx as u16;
            let def: &ColDefinition = &col_defs[col_idx]; // 構造体をそのまま渡す

            logic::write_field(
                worksheet,
                current_row,
                c_idx,
                field,
                def,
                &date_format
            )?;
        }
    }

    // ドロップダウン適用
    if let Some(max_idx) = last_row {
        logic::apply_column_validations(worksheet, &col_defs, HEADER_ROW, max_idx)?;
    }

    workbook.save(excel_path)?;
    Ok(())
}

#[pyfunction]
fn convert(csv_path: &str, excel_path: &str, define_output: &str) -> PyResult<()> {
    let col_defs: Vec<ColDefinition> = serde_json::from_str(define_output)
        .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(e.to_string()))?;

    // Vec のデバッグ表示は {:?}
    // eprintln!("DEBUG: col_defs = {:?}", col_defs);

    write_csv_to_excel_inner(csv_path, excel_path, col_defs).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(e.to_string())
    })
}

#[pymodule]
fn csv_converter(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(convert, m)?)?;
    Ok(())
}
