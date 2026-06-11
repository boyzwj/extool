use ahash::AHashMap;
use simple_excel_writer::*;
use std::fs;
use std::path::Path;

use super::state::{
    get_all_grouped_mods, get_ids_set, GLOBAL_LANG, GLOBAL_LANG_SOURCE, GLOBAL_PBD,
};
use super::types::{log_result, ExcelResult};
use super::value::to_hash_id;

pub fn create_pbd_file(out_path: &String) -> usize {
    log_result(run_create_pbd_file(out_path))
}

pub fn create_lang_file(out_path: &String) -> usize {
    log_result(run_create_lang_file(out_path))
}

pub fn create_group_files(out_path: &String) -> usize {
    log_result(run_create_group_files(out_path))
}

fn run_create_pbd_file(out_path: &str) -> ExcelResult<()> {
    let pbds = GLOBAL_PBD.read();
    let mut keys: Vec<String> = pbds.keys().map(|key| key.to_string()).collect();
    keys.sort();
    let mut contents = Vec::new();
    for key in keys {
        if let Some(content) = pbds.get(&key) {
            contents.push(content.to_string());
        }
    }
    let content = format!(
        "syntax = \"proto3\";\n\
        package pbd;\n\
        \n\
        {}",
        contents.join("\n\n")
    );

    fs::create_dir_all(out_path)
        .map_err(|err| format!("创建PBD输出目录失败 [{}]: {}", out_path, err))?;
    let pbd_path = format!("{}/pbd.proto", out_path);
    fs::write(&pbd_path, content)
        .map_err(|err| format!("写入PBD协议失败 [{}]: {}", pbd_path, err))?;

    let tmp_dir = Path::new(out_path).join(".extool_tmp");
    if tmp_dir.exists() {
        fs::remove_dir_all(&tmp_dir)
            .map_err(|err| format!("清理PBD临时目录失败 [{}]: {}", tmp_dir.display(), err))?;
    }
    Ok(())
}

fn run_create_lang_file(out_path: &str) -> ExcelResult<()> {
    let data = GLOBAL_LANG.read();
    let data_source = GLOBAL_LANG_SOURCE.read();
    let mut sorted_keys: Vec<String> = data.keys().map(|key| key.to_string()).collect();
    sorted_keys.sort();

    fs::create_dir_all(out_path)
        .map_err(|err| format!("创建多语言输出目录失败 [{}]: {}", out_path, err))?;

    let workbook_name = "D多语言简体中文表.xlsx";
    let workbook_path = format!("{}/{}", out_path, workbook_name);
    let mut workbook = Workbook::create(workbook_path.as_str());
    let mut sheet = workbook.create_sheet("sheet1");
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });

    let write_result = workbook.write_sheet(&mut sheet, |sheet_writer| {
        let writer = sheet_writer;
        writer.append_row(row!["MOD", "Data.LanguagezhCN", ""])?;
        writer.append_row(row!["BACK_TYPE", "", ""])?;
        writer.append_row(row!["FRONT_TYPE", "uint32", "string", "string", "string"])?;
        writer.append_row(row![
            "DES",
            "文本Key的Hash截取",
            "多语言key",
            "中文",
            "来源"
        ])?;
        writer.append_row(row!["NAMES", "hash", "key", "value", "source"])?;
        writer.append_row(row!["ENUM", blank!(2)])?;
        writer.append_row(row!["REF", blank!(2)])?;
        writer.append_row(row!["FORCE_MOD", "Language", ""])?;
        writer.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            let hash = to_hash_id(key);
            let value = data.get(key).cloned().unwrap_or_default();
            let source = data_source.get(key).cloned().unwrap_or_default();
            let source_filename = Path::new(&source)
                .file_name()
                .and_then(|value| value.to_str())
                .unwrap_or(source.as_str())
                .to_string();
            writer.append_row(row![
                "VALUE",
                hash.to_string(),
                key.to_string(),
                value,
                source_filename
            ])?;
        }
        writer.append_row(row![blank!(3)])
    });
    workbook
        .close()
        .map_err(|err| format!("关闭多语言工作簿失败 [{}]: {}", workbook_name, err))?;
    write_result.map_err(|err| format!("写入多语言工作簿失败 [{}]: {}", workbook_name, err))?;
    Ok(())
}

fn run_create_group_files(out_path: &str) -> ExcelResult<()> {
    let grouped = get_all_grouped_mods();
    for (group_name, mod_names) in grouped {
        create_group_file(&group_name, &get_ids_set(&mod_names), out_path)?;
    }
    Ok(())
}

fn create_group_file(
    group_name: &str,
    ids_set: &AHashMap<String, Vec<String>>,
    out_path: &str,
) -> ExcelResult<()> {
    let mut sorted_keys: Vec<String> = ids_set.keys().map(|key| key.to_string()).collect();
    sorted_keys.sort();
    fs::create_dir_all(out_path)
        .map_err(|err| format!("创建GROUP输出目录失败 [{}]: {}", out_path, err))?;

    let workbook_name = format!("Y映射表-{}.xlsx", group_name);
    let workbook_path = format!("{}/{}", out_path, workbook_name);
    let mut workbook = Workbook::create(workbook_path.as_str());
    let mut sheet = workbook.create_sheet("sheet1");
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });

    let write_result = workbook.write_sheet(&mut sheet, |sheet_writer| {
        let writer: &mut SheetWriter<'_, '_> = sheet_writer;
        let mod_name = format!("Data.Group{}", group_name);
        writer.append_row(row!["MOD", mod_name, ""])?;
        writer.append_row(row!["BACK_TYPE", "uint64", "string"])?;
        writer.append_row(row!["FRONT_TYPE", "uint64", "string"])?;
        writer.append_row(row!["DES", "id", "id来源"])?;
        writer.append_row(row!["NAMES", "id", "mod"])?;
        writer.append_row(row!["ENUM", blank!(2)])?;
        writer.append_row(row!["REF", blank!(2)])?;
        writer.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            if let Some(ids) = ids_set.get(key) {
                for id in ids {
                    writer.append_row(row!["VALUE", id.to_string(), key.to_string()])?;
                }
            }
        }
        writer.append_row(row![blank!(3)])
    });
    workbook
        .close()
        .map_err(|err| format!("关闭GROUP工作簿失败 [{}]: {}", workbook_name, err))?;
    write_result.map_err(|err| format!("写入GROUP工作簿失败 [{}]: {}", workbook_name, err))?;
    Ok(())
}
