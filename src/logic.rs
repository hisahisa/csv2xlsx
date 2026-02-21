use crate::types::ColType;
use rust_xlsxwriter::{DataValidation, Format, Worksheet, ExcelDateTime};
use std::collections::{HashMap, HashSet};
use std::error::Error;


// 定義一覧を Vec<ColType> へ変換
pub fn parse_column_types(define_output: &str) -> Vec<ColType> {
    define_output.split(',')
        .map(|s| match s.trim().to_lowercase().as_str() {
            "int" => ColType::Int,
            "date" => ColType::Date,
            "kbn_list" => ColType::KbnList,
            _ => ColType::Str,
        })
        .collect()
}

// ヘッダー処理
pub fn write_header_rows(
    worksheet: &mut Worksheet,
    records: &mut std::iter::Enumerate<csv::StringRecordsIter<std::fs::File>>,
    header_row_count: u32,
) -> Result<(), Box<dyn std::error::Error>> {
    for i in 0..header_row_count {
        if let Some((_, result)) = records.next() {
            let record = result?;
            for (col_idx, field) in record.iter().enumerate() {
                // ヘッダーは常に文字列として書き込む
                worksheet.write_string(i, col_idx as u16, field)?;
            }
        }
    }
    Ok(())
}

// データ行の処理
pub fn write_field(
    worksheet: &mut Worksheet,
    current_row: u32,
    c_idx: u16,
    field: &str,
    col_type: &ColType, // types.rsのEnumを使う
    date_format: &Format,
    column_unique_items: &mut HashMap<u16, HashSet<u8>>,
) -> Result<(), Box<dyn Error>> {
    match col_type {
        ColType::Int => {
            if let Ok(num) = field.parse::<f64>() {
                worksheet.write_number(current_row, c_idx, num)?;
            } else {
                worksheet.write_string(current_row, c_idx, field)?;
            }
        }
        ColType::Date => {
            let result = ExcelDateTime::parse_from_str(field);
            match result {
                Ok(dt) => {
                    worksheet.write_datetime_with_format(current_row, c_idx, &dt, date_format)?;
                }
                Err(_) => {
                    let replaced = field.replace("/", "-");
                    if let Ok(dt) = ExcelDateTime::parse_from_str(&replaced) {
                        worksheet.write_datetime_with_format(current_row, c_idx, &dt, date_format)?;
                    } else {
                        worksheet.write_string(current_row, c_idx, field)?;
                    }
                }
            }
        }
        ColType::KbnList => {
            let set = column_unique_items.entry(c_idx).or_insert_with(HashSet::new);
            let val_u8 = field.trim().parse::<u8>().unwrap_or(0);
            if !field.is_empty() && !set.contains(&val_u8) {
                set.insert(val_u8);
            }
            worksheet.write_number(current_row, c_idx, val_u8 as f64)?;
        }
        ColType::Str => {
            worksheet.write_string(current_row, c_idx, field)?;
        }
    }
    Ok(())
}

// ドロップダウン適用
pub fn apply_column_validations(
    worksheet: &mut Worksheet,
    column_unique_items: HashMap<u16, HashSet<u8>>,
    header_row: u32,
    max_row_idx: u32,
) -> Result<(), Box<dyn std::error::Error>> {
    for (c_idx, items) in column_unique_items {
        let mut items_vec: Vec<u8> = items.into_iter().collect();
        items_vec.sort();

        let items_str_vec: Vec<String> = items_vec.iter().map(|n| n.to_string()).collect();

        if !items_str_vec.is_empty() {
            // 文字数制限のチェック（255文字）を入れるならここが最適
            if let Ok(validation) = DataValidation::new().allow_list_strings(&items_str_vec) {
                worksheet.add_data_validation(header_row, c_idx, max_row_idx, c_idx, &validation)?;
            }
        }
    }
    Ok(())
}
