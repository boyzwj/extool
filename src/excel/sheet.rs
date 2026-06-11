use ahash::AHashSet;
use calamine::{DataType, Range};

use super::state::{
    exclude_sheet, find_group_mod_name_that_contain_id, GLOBAL_BACK_PRIMARYS,
    GLOBAL_FRONT_PRIMARYS, GLOBAL_GROUP_IDS, GLOBAL_GROUP_NAMES, GLOBAL_IDS, GLOBAL_MODS,
    GLOBAL_MOD_IDS,
};
use super::types::{EnumMap, ExcelResult, SheetData};
use super::value::{get_real_type, is_enum_none, is_list_type, normalize_type, to_enum_index};

pub(crate) fn build_sheet_id(
    input_file_name: &str,
    sheet_name: &str,
    sheet: &Range<DataType>,
    export_columns: &str,
) -> ExcelResult<bool> {
    let mut row_num = 0;
    let mut names = AHashSet::new();
    let mut front_primary = String::new();
    let mut back_primary = String::new();
    let mut mod_name = String::new();
    let mut group_name = String::new();
    let mut ids = Vec::new();
    let mut seen_mod = false;
    let mut seen_names = false;
    let mut exportable_sheet = false;

    for row in sheet.rows() {
        row_num += 1;
        let tag = row_tag(row);
        if tag == "MOD" {
            seen_mod = true;
            mod_name = cell_text(row, 1);
            if mod_name.is_empty() {
                exclude_sheet(input_file_name, sheet_name);
                return Ok(false);
            }
            if GLOBAL_MODS.read().contains(&mod_name) {
                return Err(format!(
                    "配置了重复的MOD File: [{}] Sheet: [{}] Mod_name: [{}] Row: {}",
                    input_file_name, sheet_name, mod_name, row_num
                ));
            }
            GLOBAL_MODS.write().insert(mod_name.to_string());
        } else if tag == "BACK_TYPE" {
            back_primary = cell_text(row, 1).to_uppercase();
        } else if tag == "FRONT_TYPE" {
            front_primary = cell_text(row, 1).to_uppercase();
        } else if tag == "NAMES" {
            if !seen_mod || mod_name.is_empty() {
                continue;
            }
            seen_names = true;
            if front_primary.is_empty() && back_primary.is_empty() {
                info!(
                    "SKIPPING for both `FRONT_TYPE` and `BACK_TYPE` is none File: [{}] Sheet: [{}] export_columns: [{}]",
                    input_file_name, sheet_name, export_columns
                );
                exclude_sheet(input_file_name, sheet_name);
                return Ok(false);
            }
            exportable_sheet = true;

            if !back_primary.is_empty() {
                GLOBAL_BACK_PRIMARYS
                    .write()
                    .insert(mod_name.to_string(), normalize_type(&back_primary));
            }
            if !front_primary.is_empty() {
                GLOBAL_FRONT_PRIMARYS
                    .write()
                    .insert(mod_name.to_string(), normalize_type(&front_primary));
            }

            let primary_key = cell_text(row, 1).to_uppercase();
            if primary_key.is_empty() {
                return Err(format!(
                    "NAMES的第二列固定为主键列，不能为空 File: [{}] Sheet: [{}] Mod_name: [{}] Row: {} Column: 1",
                    input_file_name, sheet_name, mod_name, row_num
                ));
            }

            for index in 1..row.len() {
                let field_name = cell_text(row, index).to_uppercase();
                if field_name.is_empty() {
                    continue;
                }
                if field_name
                    .chars()
                    .next()
                    .map(|ch| ch.is_ascii_digit())
                    .unwrap_or(false)
                {
                    return Err(format!(
                        "NAMES 不能存在以数字开头的字段 [{}] File: [{}] Sheet: [{}] Mod_name: [{}] Row: {} Column: {}",
                        field_name, input_file_name, sheet_name, mod_name, row_num, index
                    ));
                }
                if names.contains(&field_name) {
                    return Err(format!(
                        "NAMES 配置了重复的字段 [{}] File: [{}] Sheet: [{}] Mod_name: [{}] Row: {} Column: {}",
                        field_name, input_file_name, sheet_name, mod_name, row_num, index
                    ));
                }
                names.insert(field_name);
            }
        } else if tag == "GROUP" {
            group_name = cell_text(row, 1);
        } else if tag == "VALUE" {
            if mod_name.is_empty() {
                continue;
            }
            let record_id = cell_text(row, 1);
            if record_id.is_empty() {
                return Err(format!(
                    "主键不能为空 File: [{}] Sheet: [{}] Mod_name: [{}] Row: {}",
                    input_file_name, sheet_name, mod_name, row_num
                ));
            }
            let key = format!("{}:{}", mod_name, record_id);
            if GLOBAL_IDS.read().contains(&key) {
                return Err(format!(
                    "配置了重复的键值 File: [{}] Sheet: [{}] Mod_name: [{}] Row: {} Key: {}",
                    input_file_name, sheet_name, mod_name, row_num, key
                ));
            }
            GLOBAL_IDS.write().insert(key);
            ids.push(record_id.to_string());

            if !group_name.is_empty() {
                let group_id_key = format!("{}:{}", group_name, record_id);
                if GLOBAL_GROUP_IDS.read().contains(&group_id_key) {
                    let conflict_mod_names =
                        find_group_mod_name_that_contain_id(&group_name, &record_id);
                    return Err(format!(
                        "配置了重复的GROUP键值 Mod_name: [{}] ConflictModName: [{}] Row: {} group_id_key: {}",
                        mod_name,
                        conflict_mod_names.join(","),
                        row_num,
                        group_id_key
                    ));
                }
                GLOBAL_GROUP_IDS.write().insert(group_id_key);
            }
        }
    }

    if !seen_mod {
        exclude_sheet(input_file_name, sheet_name);
        return Ok(false);
    }
    if !seen_names && !mod_name.is_empty() {
        return Err(format!(
            "数据Sheet缺少NAMES行 File: [{}] Sheet: [{}] Mod_name: [{}]",
            input_file_name, sheet_name, mod_name
        ));
    }

    if !group_name.is_empty() {
        GLOBAL_GROUP_NAMES
            .write()
            .insert(mod_name.to_string(), group_name.to_string());
    }
    if !mod_name.is_empty() {
        GLOBAL_MOD_IDS.write().insert(mod_name, ids);
    }

    Ok(exportable_sheet)
}

pub(crate) fn sheet_to_data(
    input_file_name: String,
    sheet_name: &str,
    sheet: &Range<DataType>,
    export_columns: String,
    enum_values: &EnumMap,
) -> ExcelResult<SheetData> {
    let mut output_file_name = String::new();
    let mut mod_name = String::new();
    let mut values: Vec<Vec<String>> = Vec::new();
    let mut names: Vec<String> = Vec::new();
    let mut front_types: Vec<String> = Vec::new();
    let mut back_types: Vec<String> = Vec::new();
    let mut refs: Vec<String> = Vec::new();
    let mut describes: Vec<String> = Vec::new();
    let mut enum_names: Vec<String> = Vec::new();
    let mut force_mods: Vec<String> = Vec::new();
    let mut row_num = 0;

    for row in sheet.rows() {
        row_num += 1;
        let tag = row_tag(row);
        if tag == "MOD" {
            mod_name = cell_text(row, 1);
            output_file_name = mod_name.replace('.', "_").to_lowercase();
        } else if tag == "NAMES" {
            names = collect_row(row, row.len(), false);
        } else if tag == "FRONT_TYPE" {
            front_types = collect_row(row, row.len(), true);
        } else if tag == "BACK_TYPE" {
            back_types = collect_row(row, row.len(), true);
        } else if tag == "REF" {
            refs = collect_row(row, row.len(), false);
        } else if tag == "DES" {
            describes = collect_row(row, row.len(), false)
                .into_iter()
                .map(|value| value.replace("\r\n", " ").replace('\n', " "))
                .collect();
        } else if tag == "ENUM" {
            enum_names = collect_row(row, row.len(), false);
        } else if tag == "FORCE_MOD" {
            force_mods = collect_row(row, row.len(), false);
        } else if tag == "VALUE" {
            validate_value_row(
                &input_file_name,
                sheet_name,
                row,
                row_num,
                &front_types,
                &back_types,
                &refs,
                &enum_names,
                enum_values,
            )?;
            values.push(collect_row(
                row,
                max_known_len(
                    row.len(),
                    &[
                        &names,
                        &front_types,
                        &back_types,
                        &refs,
                        &enum_names,
                        &describes,
                    ],
                ),
                false,
            ));
        }
    }

    if mod_name.is_empty() {
        return Ok(SheetData {
            input_file_name,
            output_file_name,
            sheet_name: sheet_name.to_string(),
            mod_name,
            names,
            refs,
            types: Vec::new(),
            describes,
            enum_names,
            values,
            force_mods,
            export_columns,
            valid_columns: Vec::new(),
            valid_front_types: Vec::new(),
        });
    }

    if names.len() <= 1 {
        return Err(format!(
            "数据Sheet缺少有效NAMES行 File: [{}] Sheet: [{}] Mod_name: [{}]",
            input_file_name, sheet_name, mod_name
        ));
    }

    let mut types = build_export_types(&front_types, &back_types, &export_columns);
    let column_len = max_vec_len(&[&names, &types, &refs, &describes, &enum_names]);
    ensure_len(&mut names, column_len);
    ensure_len(&mut types, column_len);
    ensure_len(&mut refs, column_len);
    ensure_len(&mut describes, column_len);
    ensure_len(&mut enum_names, column_len);
    ensure_len(&mut force_mods, column_len);
    for row in values.iter_mut() {
        ensure_len(row, column_len);
    }

    let mut valid_columns = Vec::new();
    let mut valid_front_types = Vec::new();
    for index in 1..types.len() {
        let origin_type = types[index].trim();
        let field_name = names[index].trim();
        if origin_type.is_empty() || field_name.is_empty() {
            continue;
        }
        let is_enum = origin_type != "BOOL" && !enum_names[index].is_empty();
        let real_type = get_real_type(
            export_columns.as_str(),
            refs[index].as_str(),
            origin_type,
            is_enum,
        );
        valid_front_types.push(real_type);
        valid_columns.push(index);
    }

    Ok(SheetData {
        input_file_name,
        output_file_name,
        sheet_name: sheet_name.to_string(),
        mod_name,
        names,
        refs,
        types,
        describes,
        enum_names,
        values,
        force_mods,
        export_columns,
        valid_columns,
        valid_front_types,
    })
}

fn validate_value_row(
    input_file_name: &str,
    sheet_name: &str,
    row: &[DataType],
    row_num: usize,
    front_types: &[String],
    back_types: &[String],
    refs: &[String],
    enum_names: &[String],
    enum_values: &EnumMap,
) -> ExcelResult<()> {
    let max_len = max_known_len(row.len(), &[front_types, back_types, refs, enum_names]);
    for index in 1..max_len {
        let value = cell_text(row, index);
        let ref_name = refs.get(index).map(String::as_str).unwrap_or("");
        if !ref_name.is_empty() && !value.is_empty() {
            let is_list_ref = front_types
                .get(index)
                .map(|value| is_list_type(value))
                .unwrap_or(false)
                || back_types
                    .get(index)
                    .map(|value| is_list_type(value))
                    .unwrap_or(false);
            if is_list_ref {
                for item in value
                    .split(',')
                    .map(|item| item.trim())
                    .filter(|item| !item.is_empty())
                {
                    validate_ref(input_file_name, sheet_name, row_num, ref_name, item)?;
                }
            } else {
                validate_ref(input_file_name, sheet_name, row_num, ref_name, &value)?;
            }
        }

        let enum_name = enum_names.get(index).map(String::as_str).unwrap_or("");
        if !enum_name.is_empty() && !is_enum_none(&value) {
            to_enum_index(enum_values, enum_name, &value).map_err(|_| {
                format!(
                    "{} 不在枚举集合 [{}] 中 File: [{}] Sheet: [{}] Row: {} Column: {}",
                    value, enum_name, input_file_name, sheet_name, row_num, index
                )
            })?;
        }
    }
    Ok(())
}

fn validate_ref(
    input_file_name: &str,
    sheet_name: &str,
    row_num: usize,
    ref_name: &str,
    value: &str,
) -> ExcelResult<()> {
    let key = format!("{}:{}", ref_name, value);
    if GLOBAL_IDS.read().contains(&key) {
        Ok(())
    } else {
        Err(format!(
            "没找到引用的键值 File: [{}] Sheet: [{}] Row: {} Key: {}",
            input_file_name, sheet_name, row_num, key
        ))
    }
}

fn build_export_types(
    front_types: &[String],
    back_types: &[String],
    export_columns: &str,
) -> Vec<String> {
    if export_columns == "BOTH" {
        let len = std::cmp::max(front_types.len(), back_types.len());
        let mut result = vec!["BOTH".to_string()];
        for index in 1..len {
            let front_type = front_types
                .get(index)
                .map(String::as_str)
                .unwrap_or("")
                .trim();
            let back_type = back_types
                .get(index)
                .map(String::as_str)
                .unwrap_or("")
                .trim();
            if front_type.is_empty() {
                result.push(normalize_type(back_type));
            } else {
                result.push(normalize_type(front_type));
            }
        }
        result
    } else if export_columns == "BACK" {
        back_types
            .iter()
            .map(|value| normalize_type(value))
            .collect()
    } else {
        front_types
            .iter()
            .map(|value| normalize_type(value))
            .collect()
    }
}

fn row_tag(row: &[DataType]) -> String {
    cell_text(row, 0).to_uppercase()
}

fn cell_text(row: &[DataType], index: usize) -> String {
    row.get(index)
        .map(|value| value.to_string().trim().to_string())
        .unwrap_or_default()
}

fn collect_row(row: &[DataType], len: usize, uppercase: bool) -> Vec<String> {
    let mut result = Vec::new();
    for index in 0..len {
        let mut value = cell_text(row, index);
        if uppercase {
            value.make_ascii_uppercase();
        }
        result.push(value);
    }
    result
}

fn ensure_len(values: &mut Vec<String>, len: usize) {
    while values.len() < len {
        values.push(String::new());
    }
}

fn max_vec_len(groups: &[&Vec<String>]) -> usize {
    groups.iter().map(|group| group.len()).max().unwrap_or(0)
}

fn max_known_len<T>(row_len: usize, groups: &[&[T]]) -> usize {
    groups
        .iter()
        .map(|group| group.len())
        .fold(row_len, std::cmp::max)
}
