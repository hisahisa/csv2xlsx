use crate::types::{ColType, ColDefinition};
use rust_xlsxwriter::{DataValidation, Format, Worksheet, ExcelDateTime};
use std::error::Error;

pub fn apply_column_settings(
    worksheet: &mut Worksheet,
    col_defs: &[ColDefinition])
    -> Result<(), Box<dyn Error>> {
    for (i, def) in col_defs.iter().enumerate() {
        worksheet.set_column_width(i as u16, def.width)?;
    }
    Ok(())
}

pub fn write_header_rows(
    worksheet: &mut Worksheet,
    records: &mut std::iter::Enumerate<csv::StringRecordsIter<std::fs::File>>,
    header_row_count: u32,
) -> Result<(), Box<dyn Error>> {
    for i in 0..header_row_count {
        if let Some((_, result)) = records.next() {
            let record = result?;
            for (col_idx, field) in record.iter().enumerate() {
                worksheet.write_string(i, col_idx as u16, field)?;
            }
        }
    }
    Ok(())
}

pub fn write_field(
    worksheet: &mut Worksheet,
    current_row: u32,
    c_idx: u16,
    field: &str,
    col_type: &ColType,
    date_format: &Format
) -> Result<(), Box<dyn Error>> {
    match col_type {
        ColType::Int => {
            if let Ok(num) = field.trim().parse::<f64>() {
                worksheet.write_number(current_row, c_idx, num)?;
            } else {
                worksheet.write_string(current_row, c_idx, field)?;
            }
        }
        ColType::Date => {
            let trimmed = field.trim();
            if let Ok(dt) = ExcelDateTime::parse_from_str(trimmed) {
                worksheet.write_datetime_with_format(current_row, c_idx, &dt, date_format)?;
            } else {
                let replaced = trimmed.replace("/", "-");
                if let Ok(dt) = ExcelDateTime::parse_from_str(&replaced) {
                    worksheet.write_datetime_with_format(current_row, c_idx, &dt, date_format)?;
                } else {
                    worksheet.write_string(current_row, c_idx, field)?;
                }
            }
        }
        ColType::KbnList => {
            let val_u8 = field.trim().parse::<u8>().unwrap_or(0);
            // kbn_list の場合は自動収集はせず、表示のみ（バリデーションは別途 apply_column_validations で実施）
            worksheet.write_number(current_row, c_idx, val_u8 as f64)?;
        }
        ColType::Str => {
            worksheet.write_string(current_row, c_idx, field)?;
        }
    }
    Ok(())
}

pub fn apply_column_validations(
    worksheet: &mut Worksheet,
    col_defs: &[ColDefinition],
    header_row: u32,
    max_row_idx: u32,
) -> Result<(), Box<dyn Error>> {
    for (c_idx, def) in col_defs.iter().enumerate() {
        if let Some(values) = &def.kbn_values {
            let items: Vec<String> = values.iter().map(|n| n.to_string()).collect();
            if !items.is_empty() {
                let validation = DataValidation::new().allow_list_strings(&items)?;
                worksheet.add_data_validation(header_row,
                                              c_idx as u16,
                                              max_row_idx,
                                              c_idx as u16,
                                              &validation)?;
            }
        }
    }
    Ok(())
}
