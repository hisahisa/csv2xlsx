use pyo3::prelude::*;
use rust_xlsxwriter::{DataValidation, Workbook, Format, ExcelDateTime};
use std::collections::{HashMap, HashSet};
use std::error::Error;

const HEADER_ROW: u32 = 1;

// 文字列判定を避けるためのEnum
enum ColType {
    Str,
    Int,
    Date,
    KbnList,
}

// 内部ロジック
fn write_csv_to_excel_inner(
    csv_path: &str,
    excel_path: &str,
    define_output: &str,
) -> Result<(), Box<dyn Error>> {

    // 1. 文字列比較をループの外に出す（プリプロセス）
    let col_types: Vec<ColType> = define_output.split(',')
        .map(|s| match s.trim() {
            "int" => ColType::Int,
            "date" => ColType::Date,
            "kbn_list" => ColType::KbnList,
            _ => ColType::Str,
        })
        .collect();

    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_path(csv_path)?;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // 日付フォーマットの事前生成
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // ユニーク値保持用（キー検索を高速化するためにEntry APIなどを活用もできるが、今回はcontainsチェックで最適化）
    let mut column_unique_items: HashMap<u16, HashSet<String>> = HashMap::new();
    let mut max_row_idx = 0;

    // イテレータを取得
    let mut records = rdr.records().enumerate();

    // --- ヘッダー処理（ループ外に出す） ---
    for i in 0..HEADER_ROW {
        if let Some((_, result)) = records.next() {
            let record = result?;
            for (col_idx, field) in record.iter().enumerate() {
                worksheet.write_string(i as u32, col_idx as u16, field)?;
            }
        }
    }

    // --- データ行の処理 ---
    for (row_idx, result) in records {
        let record = result?;
        let current_row = row_idx as u32;
        max_row_idx = current_row;

        for (col_idx, field) in record.iter().enumerate() {
            let c_idx = col_idx as u16;

            // 安全にアクセス
            let col_type = col_types.get(col_idx).unwrap_or(&ColType::Str);

            match col_type {
                ColType::Int => {
                    // 高速なパース (lexicalなどのクレートを使うともっと速いがstdでも十分)
                    if let Ok(num) = field.parse::<f64>() {
                        worksheet.write_number(current_row, c_idx, num)?;
                    } else {
                        worksheet.write_string(current_row, c_idx, field)?;
                    }
                },
                ColType::Date => {
                    // 軽量な判定：Sring割り当てを避ける
                    let result = ExcelDateTime::parse_from_str(field);
                    match result {
                        Ok(dt) => {
                            worksheet.write_datetime_with_format(current_row, c_idx, &dt, &date_format)?;
                        }
                        Err(_) => {
                            // パース失敗時のみコストを払って置換を試す
                            let replaced = field.replace("/", "-");
                            if let Ok(dt) = ExcelDateTime::parse_from_str(&replaced) {
                                worksheet.write_datetime_with_format(current_row, c_idx, &dt, &date_format)?;
                            } else {
                                // どうしてもダメなら文字列
                                worksheet.write_string(current_row, c_idx, field)?;
                            }
                        }
                    }
                },
                ColType::KbnList => {
                    // ★【重要】Allocationの削減
                    // HashSetにデータを入れる際、まず &str で存在確認をする。
                    // 存在しない場合のみ to_string() してメモリ確保を行う。
                    let set = column_unique_items.entry(c_idx).or_insert_with(HashSet::new);
                    if !field.is_empty() && !set.contains(field) {
                        set.insert(field.to_string());
                    }
                    worksheet.write_string(current_row, c_idx, field)?;
                },
                ColType::Str => {
                    worksheet.write_string(current_row, c_idx, field)?;
                }
            }
        }
    }

    // ドロップダウン適用
    for (c_idx, items) in column_unique_items {
        let mut items_vec: Vec<String> = items.into_iter().collect();
        items_vec.sort();

        // 件数が多いとExcelが壊れるためガード
        if !items_vec.is_empty() && items_vec.iter().map(|s| s.len()).sum::<usize>() < 255 * 255 {
            // データ入力規則の設定（エラーハンドリングは要件に合わせて）
            if let Ok(validation) = DataValidation::new().allow_list_strings(&items_vec) {
                // 行全体に適用（ヘッダー除く）
                worksheet.add_data_validation(HEADER_ROW, c_idx, max_row_idx, c_idx, &validation)?;
            }
        }
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
