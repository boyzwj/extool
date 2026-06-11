use prost_reflect::Value as PValue;
use serde_json::value::Value;
use std::fs;
use std::path::Path;

use super::state::{GLOBAL_BACK_PRIMARYS, GLOBAL_FRONT_PRIMARYS};
use super::types::{EnumMap, EnumValue, ExcelResult};

pub(crate) fn cell_to_json(
    cell: &str,
    row_type: &str,
    filename: &str,
    key: &str,
    sheetname: &str,
) -> ExcelResult<Value> {
    let s = cell.trim().to_string();
    if row_type.starts_with("INT") || row_type.starts_with("UINT") {
        if s.is_empty() {
            return Ok(json!(0));
        }
        parse_i64(&s, filename, key, sheetname).map(|v| json!(v))
    } else if row_type.starts_with("FLOAT") {
        if s.is_empty() {
            return Ok(json!(0));
        }
        parse_f64(&s, filename, key, sheetname).map(|v| json!(v))
    } else if row_type.starts_with("LIST_UINT") || row_type.starts_with("LIST_INT") {
        let values = parse_list_i64(&s, filename, key, sheetname)?;
        Ok(json!(values))
    } else if row_type.starts_with("LIST_FLOAT") {
        let values = parse_list_f64(&s, filename, key, sheetname)?;
        Ok(json!(values))
    } else if row_type.starts_with("LIST_STRING") {
        Ok(json!(split_list(&s)))
    } else if row_type == "STRING_LOC" {
        Ok(json!(to_hash_id(&s)))
    } else if row_type == "BOOL" {
        Ok(serde_json::Value::Bool(parse_bool(&s)))
    } else {
        Ok(json!(s))
    }
}

pub(crate) fn cell_to_string(
    cell: &str,
    row_type: &str,
    filename: &str,
    key: &str,
) -> ExcelResult<String> {
    let s = cell.trim().to_string();
    if row_type.starts_with("INT") || row_type.starts_with("UINT") {
        if s.is_empty() {
            return Ok("0".to_string());
        }
        parse_i64(&s, filename, key, "").map(|_| s)
    } else if row_type.starts_with("FLOAT") {
        if s.is_empty() {
            return Ok("0".to_string());
        }
        parse_f64(&s, filename, key, "").map(|_| s)
    } else if row_type.starts_with("LIST_UINT") || row_type.starts_with("LIST_INT") {
        let values = parse_list_i64(&s, filename, key, "")?;
        Ok(format!("[{}]", join_values(values)))
    } else if row_type.starts_with("LIST_FLOAT") {
        let values = parse_list_f64(&s, filename, key, "")?;
        Ok(format!("[{}]", join_values(values)))
    } else if row_type.starts_with("LIST_STRING") {
        let values = split_list(&s);
        serde_json::to_string(&values).map_err(|err| {
            format!(
                "字符串列表序列化失败 File: [{}] Key: [{}] Err: {}",
                filename, key, err
            )
        })
    } else if row_type == "STRING" {
        serde_json::to_string(&s).map_err(|err| {
            format!(
                "字符串序列化失败 File: [{}] Key: [{}] Err: {}",
                filename, key, err
            )
        })
    } else if row_type == "STRING_LOC" {
        Ok(to_hash_id(&s).to_string())
    } else if row_type == "BOOL" {
        Ok(if parse_bool(&s) { "true" } else { "false" }.to_string())
    } else {
        Ok(s)
    }
}

pub(crate) fn cell_to_lua_string(
    cell: &str,
    row_type: &str,
    filename: &str,
    key: &str,
) -> ExcelResult<String> {
    let s = cell.trim().to_string();
    if row_type.starts_with("LIST_UINT") || row_type.starts_with("LIST_INT") {
        let values = parse_list_i64(&s, filename, key, "")?;
        Ok(format!("{{{}}}", join_values(values)))
    } else if row_type.starts_with("LIST_FLOAT") {
        let values = parse_list_f64(&s, filename, key, "")?;
        Ok(format!("{{{}}}", join_values(values)))
    } else if row_type.starts_with("LIST_STRING") {
        let values = split_list(&s);
        let mut quoted = Vec::new();
        for value in values {
            quoted.push(serde_json::to_string(&value).map_err(|err| {
                format!(
                    "Lua字符串列表序列化失败 File: [{}] Key: [{}] Err: {}",
                    filename, key, err
                )
            })?);
        }
        Ok(format!("{{{}}}", quoted.join(",")))
    } else {
        cell_to_string(cell, row_type, filename, key)
    }
}

pub(crate) fn cell_to_pvalue(
    cell: &str,
    row_type: &str,
    filename: &str,
    sheetname: &str,
    key: &str,
) -> ExcelResult<PValue> {
    let mut s = cell.trim().to_string();
    if row_type.contains("INT") && s.is_empty() {
        s = "0".to_string();
    }
    if row_type.contains("FLOAT") && s.is_empty() {
        s = "0.0".to_string();
    }

    if row_type == "UINT32" {
        parse_u32(&s, filename, key, sheetname).map(PValue::U32)
    } else if row_type == "UINT64" {
        parse_u64(&s, filename, key, sheetname).map(PValue::U64)
    } else if row_type == "INT32" {
        parse_i32(&s, filename, key, sheetname).map(PValue::I32)
    } else if row_type == "INT64" {
        parse_i64(&s, filename, key, sheetname).map(PValue::I64)
    } else if row_type == "FLOAT" {
        parse_f32(&s, filename, key, sheetname).map(PValue::F32)
    } else if row_type == "LIST_UINT32" {
        list_to_pvalues(&s, |item| {
            parse_u32(item, filename, key, sheetname).map(PValue::U32)
        })
    } else if row_type == "LIST_UINT64" {
        list_to_pvalues(&s, |item| {
            parse_u64(item, filename, key, sheetname).map(PValue::U64)
        })
    } else if row_type == "LIST_INT32" {
        list_to_pvalues(&s, |item| {
            parse_i32(item, filename, key, sheetname).map(PValue::I32)
        })
    } else if row_type == "LIST_INT64" {
        list_to_pvalues(&s, |item| {
            parse_i64(item, filename, key, sheetname).map(PValue::I64)
        })
    } else if row_type == "LIST_FLOAT" {
        list_to_pvalues(&s, |item| {
            parse_f32(item, filename, key, sheetname).map(PValue::F32)
        })
    } else if row_type == "LIST_STRING" {
        Ok(PValue::List(
            split_list(&s).into_iter().map(PValue::String).collect(),
        ))
    } else if row_type == "STRING_LOC" {
        Ok(PValue::U32(to_hash_id(&s)))
    } else if row_type == "STRING" {
        Ok(PValue::String(s))
    } else if row_type == "BOOL" {
        Ok(PValue::Bool(parse_bool(&s)))
    } else {
        Err(format!(
            "不支持的PBD字段类型 [{}] File: [{}] Sheet: [{}] Key: [{}] Val: [{}]",
            row_type, filename, sheetname, key, s
        ))
    }
}

pub(crate) fn get_real_type(
    export_columns: &str,
    ref_name: &str,
    origin_type: &str,
    is_enum: bool,
) -> String {
    if is_enum {
        return "ENUM".to_string();
    }
    if !ref_name.trim().is_empty() {
        return get_ref_primary_type(export_columns, ref_name, origin_type);
    }
    normalize_type(origin_type)
}

pub(crate) fn get_ref_primary_type(
    export_columns: &str,
    ref_name: &str,
    origin_type: &str,
) -> String {
    let primary_type = match export_columns {
        "BACK" => GLOBAL_BACK_PRIMARYS.read().get(ref_name).cloned(),
        "FRONT" => GLOBAL_FRONT_PRIMARYS.read().get(ref_name).cloned(),
        _ => None,
    };

    match primary_type {
        Some(value) if is_list_type(origin_type) => {
            format!("LIST_{}", normalize_scalar_type(&value))
        }
        Some(value) => normalize_scalar_type(&value),
        None => normalize_type(origin_type),
    }
}

pub(crate) fn normalize_type(origin_type: &str) -> String {
    let upper = origin_type.trim().to_uppercase();
    if upper == "LIST_INT" || upper == "LIST" {
        "LIST_UINT32".to_string()
    } else if upper == "INT" {
        "UINT32".to_string()
    } else {
        upper
    }
}

pub(crate) fn normalize_scalar_type(origin_type: &str) -> String {
    let upper = origin_type.trim().to_uppercase();
    if upper == "INT" {
        "UINT32".to_string()
    } else {
        upper
    }
}

pub(crate) fn is_list_type(row_type: &str) -> bool {
    row_type.trim().to_uppercase().contains("LIST")
}

pub(crate) fn to_hash_id(key: &str) -> u32 {
    if key.is_empty() {
        return 0;
    }
    let digest = md5::compute(key);
    u32::from_be_bytes([digest.0[12], digest.0[13], digest.0[14], digest.0[15]])
}

pub(crate) fn is_enum_none(value: &str) -> bool {
    value.trim().is_empty() || value.to_uppercase() == "NONE"
}

pub(crate) fn extract_all_enum_values(file_path: &str) -> ExcelResult<EnumMap> {
    if !Path::new(file_path).exists() {
        info!("枚举文件不存在，跳过枚举预加载: {}", file_path);
        return Ok(EnumMap::new());
    }

    let file_contents = fs::read_to_string(file_path)
        .map_err(|err| format!("读取枚举文件失败 [{}]: {}", file_path, err))?;
    let mut enum_values = EnumMap::new();

    let enum_regex = regex::Regex::new(r#"enum\s+(\w+)\s*\{([^}]+)\}"#)
        .map_err(|err| format!("枚举正则初始化失败: {}", err))?;
    let comment_regex =
        regex::Regex::new(r#"//(.*)$"#).map_err(|err| format!("注释正则初始化失败: {}", err))?;
    let value_regex = regex::Regex::new(r#"(\w+)\s*=\s*(\d+);"#)
        .map_err(|err| format!("枚举值正则初始化失败: {}", err))?;

    for enum_match in enum_regex.captures_iter(&file_contents) {
        let enum_name = enum_match[1].to_owned();
        let enum_body = enum_match[2].to_owned();
        let mut values = Vec::new();

        for line in enum_body.lines() {
            if let Some(comment_match) = comment_regex.captures(line) {
                if let Some(value_match) = value_regex.captures(line) {
                    let index = value_match[2].parse::<i32>().map_err(|err| {
                        format!(
                            "枚举值解析失败 Enum: [{}] Line: [{}] Err: {}",
                            enum_name, line, err
                        )
                    })?;
                    let comment = comment_match[1].trim().to_owned();
                    values.push(EnumValue { index, comment });
                }
            }
        }

        if !values.is_empty() {
            enum_values.insert(enum_name, values);
        }
    }

    Ok(enum_values)
}

pub(crate) fn to_enum_index(
    enum_values: &EnumMap,
    enum_name: &str,
    comment: &str,
) -> ExcelResult<i32> {
    match enum_values.get(enum_name) {
        Some(values) => {
            for info in values {
                if info.comment == comment {
                    return Ok(info.index);
                }
            }
            Err(format!(
                "EnumName: [{}], Comment: [{}] not found",
                enum_name, comment
            ))
        }
        None => Err(format!("EnumName: [{}] not found", enum_name)),
    }
}

fn split_list(s: &str) -> Vec<String> {
    if s.trim().is_empty() {
        return Vec::new();
    }
    s.split(',').map(|item| item.trim().to_string()).collect()
}

fn join_values<T: ToString>(values: Vec<T>) -> String {
    values
        .into_iter()
        .map(|value| value.to_string())
        .collect::<Vec<String>>()
        .join(",")
}

fn list_to_pvalues<F>(s: &str, parser: F) -> ExcelResult<PValue>
where
    F: Fn(&str) -> ExcelResult<PValue>,
{
    if s.trim().is_empty() {
        return Ok(PValue::List(Vec::new()));
    }
    let mut result = Vec::new();
    for item in s.split(',') {
        result.push(parser(item.trim())?);
    }
    Ok(PValue::List(result))
}

fn parse_list_i64(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<Vec<i64>> {
    if s.trim().is_empty() {
        return Ok(Vec::new());
    }
    let mut result = Vec::new();
    for item in s.split(',') {
        result.push(parse_i64(item.trim(), filename, key, sheetname)?);
    }
    Ok(result)
}

fn parse_list_f64(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<Vec<f64>> {
    if s.trim().is_empty() {
        return Ok(Vec::new());
    }
    let mut result = Vec::new();
    for item in s.split(',') {
        result.push(parse_f64(item.trim(), filename, key, sheetname)?);
    }
    Ok(result)
}

fn parse_bool(s: &str) -> bool {
    s == "是" || s.to_uppercase() == "TRUE"
}

fn parse_u32(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<u32> {
    s.parse::<u32>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_u64(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<u64> {
    s.parse::<u64>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_i32(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<i32> {
    s.parse::<i32>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_i64(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<i64> {
    s.parse::<i64>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_f32(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<f32> {
    s.parse::<f32>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_f64(s: &str, filename: &str, key: &str, sheetname: &str) -> ExcelResult<f64> {
    s.parse::<f64>()
        .map_err(|_| parse_err(filename, key, s, sheetname))
}

fn parse_err(filename: &str, key: &str, s: &str, sheetname: &str) -> String {
    format!(
        "数据类型转换出错 File: [{}] Sheet: [{}] Key: [{}] Content: [{}]",
        filename, sheetname, key, s
    )
}
