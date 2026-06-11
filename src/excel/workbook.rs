use calamine::{open_workbook, Reader, Xlsx};
use std::io::{Read, Seek};

use super::sheet::{build_sheet_id, sheet_to_data};
use super::state::is_excluded_sheet;
use super::types::{log_result, ExcelResult};
use super::value::extract_all_enum_values;

pub fn xls_to_file(
    input_file_name: String,
    dst_path: String,
    format: String,
    multi_sheets: bool,
    export_columns: String,
    pbe_path: String,
) -> usize {
    log_result(run_xls_to_file(
        input_file_name,
        dst_path,
        format,
        multi_sheets,
        export_columns,
        pbe_path,
    ))
}

pub fn build_id(input_file_name: String, multi_sheets: bool, export_columns: String) -> usize {
    log_result(run_build_id(input_file_name, multi_sheets, export_columns))
}

fn run_xls_to_file(
    input_file_name: String,
    dst_path: String,
    format: String,
    multi_sheets: bool,
    export_columns: String,
    pbe_path: String,
) -> ExcelResult<()> {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.as_str())
        .map_err(|err| format!("打开Excel失败 [{}]: {}", input_file_name, err))?;
    let sheets = selected_sheet_names(&mut excel, multi_sheets)?;
    let enum_values = extract_all_enum_values(&pbe_path)?;

    let mut exported_count = 0usize;
    for sheet_name in sheets {
        if is_excluded_sheet(&input_file_name, &sheet_name) {
            continue;
        }
        info!("LOADING [{}] [{}] ...", input_file_name, sheet_name);
        let range = worksheet_range(&mut excel, &input_file_name, &sheet_name)?;
        let data = sheet_to_data(
            input_file_name.to_string(),
            &sheet_name,
            &range,
            export_columns.to_string(),
            &enum_values,
        )?;
        if data.mod_name.is_empty() {
            continue;
        }
        data.export(&format, &dst_path, &enum_values)?;
        exported_count += 1;
    }

    if multi_sheets && exported_count > 1 {
        return Err(format!(
            "一个Excel文件最多只能有一个可导出的数据Sheet File: [{}] ExportedSheets: {}",
            input_file_name, exported_count
        ));
    }
    Ok(())
}

fn run_build_id(
    input_file_name: String,
    multi_sheets: bool,
    export_columns: String,
) -> ExcelResult<()> {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.as_str())
        .map_err(|err| format!("打开Excel失败 [{}]: {}", input_file_name, err))?;
    let sheets = selected_sheet_names(&mut excel, multi_sheets)?;
    let mut data_sheet_count = 0usize;

    for sheet_name in sheets {
        let range = worksheet_range(&mut excel, &input_file_name, &sheet_name)?;
        if build_sheet_id(&input_file_name, &sheet_name, &range, &export_columns)? {
            data_sheet_count += 1;
        }
    }

    if multi_sheets && data_sheet_count > 1 {
        return Err(format!(
            "一个Excel文件最多只能有一个可导出的数据Sheet File: [{}] DataSheets: {}",
            input_file_name, data_sheet_count
        ));
    }
    Ok(())
}

fn selected_sheet_names<RS: Read + Seek>(
    excel: &mut Xlsx<RS>,
    multi_sheets: bool,
) -> ExcelResult<Vec<String>> {
    let sheet_names = excel.sheet_names().to_owned();
    if sheet_names.is_empty() {
        return Err("Excel没有任何Sheet".to_string());
    }
    if multi_sheets {
        Ok(sheet_names)
    } else {
        Ok(vec![sheet_names[0].to_string()])
    }
}

fn worksheet_range(
    excel: &mut Xlsx<impl Read + Seek>,
    input_file_name: &str,
    sheet_name: &str,
) -> ExcelResult<calamine::Range<calamine::DataType>> {
    match excel.worksheet_range(sheet_name) {
        Some(Ok(range)) => Ok(range),
        Some(Err(err)) => Err(format!(
            "读取Sheet失败 File: [{}] Sheet: [{}] Err: {}",
            input_file_name, sheet_name, err
        )),
        None => Err(format!(
            "找不到Sheet File: [{}] Sheet: [{}]",
            input_file_name, sheet_name
        )),
    }
}
