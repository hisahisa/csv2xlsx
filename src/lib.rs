use pyo3::prelude::*;
use rust_xlsxwriter::{DataValidation, Workbook, Format, ExcelDateTime};
use std::collections::{HashMap, HashSet};
use std::error::Error;

const HEADER_ROW:u32 = 1;


// 内部ロジック
fn write_csv_to_excel_inner(
    csv_path: &str,
    excel_path: &str,
    define_output: &str,
) -> Result<(), Box<dyn Error>> {
    let types: Vec<&str> = define_output.split(',').map(|s| s.trim()).collect();

    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_path(csv_path)?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // ★ 列番号(u16)をキーにして、それぞれのユニークな値を保持する
    let mut column_unique_items: HashMap<u16, HashSet<String>> = HashMap::new();
    let mut max_row_idx = 0;

    for (row_idx, result) in rdr.records().enumerate() {
        let record = result?;
        let r_idx = row_idx as u32;
        max_row_idx = r_idx;

        for (col_idx, field) in record.iter().enumerate() {
            let c_idx = col_idx as u16;
            let col_type = types.get(col_idx).unwrap_or(&"str");

            // n行目はヘッダーなので、無条件に文字列として書く
            if r_idx < HEADER_ROW {
                worksheet.write_string(r_idx, c_idx, field)?;
                continue;
            }

            match *col_type {
                "date" => {
                    let date_str = field.replace("/", "-");
                    let dt = ExcelDateTime::parse_from_str(&date_str)
                        .unwrap_or_else(|_| ExcelDateTime::parse_from_str("1900-01-01").unwrap());
                    worksheet.write_datetime_with_format(r_idx, c_idx, &dt, &date_format)?;
                },
                "kbn_list" => {
                    // ★ その列専用のHashSetに値を追加
                    if !field.is_empty() {
                        column_unique_items
                            .entry(c_idx)
                            .or_insert_with(HashSet::new)
                            .insert(field.to_string());
                    }
                    worksheet.write_string(r_idx, c_idx, field)?;
                },
                "int" => {
                    if let Ok(num) = field.parse::<f64>() {
                        worksheet.write_number(r_idx, c_idx, num)?;
                    } else {
                        worksheet.write_string(r_idx, c_idx, field)?;
                    }
                },
                _ => {
                    worksheet.write_string(r_idx, c_idx, field)?;
                }
            }
        }
    }

    // ★ ドロップダウンを列ごとに適用
    for (c_idx, items) in column_unique_items {
        let mut items_vec: Vec<String> = items.into_iter().collect();
        items_vec.sort();

        // Excelのドロップダウンには件数制限があるため注意（約8000文字程度まで）
        let validation = DataValidation::new().allow_list_strings(&items_vec)?;
        worksheet.add_data_validation(HEADER_ROW, c_idx, max_row_idx, c_idx, &validation)?;
    }

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
