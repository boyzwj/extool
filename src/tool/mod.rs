use ahash::{AHashMap, AHashSet};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use inflector::Inflector;
use serde_json;
use serde_json::value::Value;
use serde_json::Map;

//use std::collections::HashMap;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::string::String;

use static_init::dynamic;
#[dynamic]
static mut GLOBAL_IDS: AHashSet<String> = AHashSet::new();

// static mut
pub struct SheetData<'a> {
    input_file_name: String,
    output_file_name: String,
    sheet_name: &'a String,
    mod_name: String,
    names: Vec<String>,
    refs: Vec<String>,
    front_types: Vec<String>,
    back_types: Vec<String>,
    enums: Vec<AHashMap<String, usize>>,
    values: Vec<Vec<&'a DataType>>,
}

impl<'a> SheetData<'_> {
    pub fn export(&self, format: &String, dst_path: &String) {
        if format == "JSON" {
            self.data_to_json(&dst_path);
        } else if format == "LUA" {
            self.data_to_lua(&dst_path);
        } else if format == "EX" {
            self.data_to_ex(&dst_path);
        }
    }

    pub fn write_file(&self, path_str: &str, content: &String) {
        match File::create(path_str) {
            Err(e) => {
                error!("创建导出文件失败 : {:?}", e);
            }
            Ok(f) => {
                let mut bw = BufWriter::new(f);
                match bw.write_all(content.as_bytes()) {
                    Err(e) => {
                        error!("写入导出内容失败 : {:?}", e);
                    }
                    Ok(_) => {
                        info!(
                            "成功 [{}] [{}] 导出 [{}] 条记录",
                            self.input_file_name,
                            self.sheet_name,
                            self.values.len()
                        );
                    }
                }
            }
        }
    }

    pub fn data_to_json(&self, out_path: &String) {
        let mut res: Map<String, Value> = Map::new();
        for rv in &self.values {
            let mut map: Map<String, Value> = Map::new();
            for i in 1..self.front_types.len() {
                if self.front_types[i] != "" {
                    let column_name = &self.names[i];
                    let row_type = &self.front_types[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let value =
                            cell_to_json(&rv[i], &row_type, &self.output_file_name, &column_name);
                        map.insert(column_name.to_string(), value);
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if let Some(x) = &self.enums[i].get(value) {
                            map.insert(column_name.to_string(), json!(x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                                column_name,
                                &rv[1],
                                value,
                                &self.enums[i].keys()
                            )
                        }
                    }
                }
            }

            if map.len() > 0 {
                let key_str = &rv[1];
                res.insert(key_str.to_string(), json!(map));
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return;
        }
        let obj = json!(res);
        let json = serde_json::to_string_pretty(&obj).unwrap();
        let path_str = format!("{}/{}.json", out_path, self.output_file_name);
        self.write_file(&path_str, &json);
    }

    pub fn data_to_lua(&self, out_path: &String) {
        let mut res: Vec<String> = vec![];
        let mut kns: Vec<String> = vec![];
        let mut j = 1;
        for i in 1..self.front_types.len() {
            if self.front_types[i] != "" {
                let column_name = &self.names[i];
                kns.push(format!("{} = {}", column_name, j.to_string()));
                j = j + 1;
            }
        }
        for rv in &self.values {
            let mut columns: Vec<String> = vec![];
            for i in 1..self.front_types.len() {
                if self.front_types[i] != "" {
                    let column_name = &self.names[i];
                    let row_type = &self.front_types[i];
                    let dic = &self.enums[i];
                    if dic.len() == 0 {
                        let value =
                            cell_to_string(&rv[i], row_type, &self.output_file_name, &column_name);
                        columns.push(format!("{}", value.replace("[", "{").replace("]", "}")));
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if let Some(x) = dic.get(value) {
                            columns.push(format!("{}", x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                                column_name,
                                &rv[1],
                                value,
                                dic.keys()
                            );
                            panic!("abort")
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let keyvalue = cell_to_string(
                    &rv[1],
                    &self.front_types[1],
                    &self.output_file_name,
                    &self.names[1],
                );
                res.push(format!("\t[{}] = {{ {} }}", keyvalue, columns.join(",\n")));
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return;
        }
        let out = format!(
            "local KT = {{ {} }}\n\
    local data = {{ \n {}\n}}\n\
    do\n\
    \tlocal base = {{\n\
    \t\t__index = function(table,key)\n\
    \t\t\tlocal ki = KT[key]\n\
    \t\t\tif not ki then\n\
    \t\t\t\treturn nil\n\
    \t\t\tend\n\
    \t\t\treturn table[ki]
    \t\tend,\n\
    \t\t__newindex = function()\n\
    \t\t\terror([[Attempt to modify read-only table]])\n\
    \t\tend\n\
    \t}}\n\
    \tfor k, v in pairs(data) do\n\
    \t\tsetmetatable(v, base)\n\
    \tend\n\
    \tbase.__metatable = false\n\
    end\n\
    return data",
            kns.join(","),
            res.join(",\n"),
        );
        let path_str = format!("{}/{}.lua", out_path, self.output_file_name);
        self.write_file(&path_str, &out);
    }

    pub fn data_to_ex(&self, dst_path: &String) {
        let mut res: Vec<String> = vec![];
        let mut ids: Vec<String> = vec![];
        for rv in &self.values {
            let mut columns: Vec<String> = vec![];
            for i in 1..self.back_types.len() {
                if self.back_types[i] != "" {
                    let column_name = &self.names[i];
                    let row_type = &self.back_types[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let value =
                            cell_to_string(&rv[i], row_type, &self.output_file_name, &column_name);
                        columns.push(format!("\t\t\t{}: {}", column_name, value));
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if let Some(x) = &self.enums[i].get(value) {
                            columns.push(format!("\t\t\t{}: {}", column_name, x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                                column_name,
                                &rv[1],
                                value,
                                &self.enums[i].keys()
                            );
                            panic!("abort")
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let keyvalue = cell_to_string(
                    &rv[1],
                    &self.back_types[1],
                    &self.output_file_name,
                    &self.names[1],
                );
                res.push(format!(
                    "\tdef get({}) do\n\t\t%{{\n{}\n\t}}",
                    keyvalue,
                    columns.join(",\n")
                ));
                ids.push(keyvalue.to_string());
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return;
        }
        let module_name = get_module_name(self.output_file_name.clone());
        let out = format!(
            "defmodule {} do\n\
             \t## SOURCE:\"{}\" SHEET:\"{}\"\n\n\
             \tdef ids() do\n\
             \t\t[{}]\n\
             \tend\n\n\
             \tdef all(), do: for id <- ids(), do: get(id)\n\n\
             \tdef query(q), do: for data <- all(), q.(data), do: data\n\
             \n\
             {}\n\
             \tdef get(_), do: nil\n\
             end",
            module_name,
            self.input_file_name,
            self.sheet_name,
            ids.join(", "),
            res.join("\n")
        );
        let path_str = format!("{}/{}.ex", dst_path, self.output_file_name);
        self.write_file(&path_str, &out);
    }
}

pub fn xls_to_file(input_file_name: String, dst_path: String, format: String) {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.clone()).unwrap();
    let sheets = excel.sheet_names().to_owned();
    for sheet in sheets {
        info!("LOADING [{}] [{}] ...", input_file_name, sheet);
        if let Some(Ok(r)) = excel.worksheet_range(&sheet) {
            let data = sheet_to_data(input_file_name.clone(), &sheet, &r);
            data.export(&format, &dst_path);
        }
    }
}

pub fn build_id(input_file_name: String) {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.clone()).unwrap();
    for sheet in excel.sheet_names().to_owned() {
        if let Some(Ok(r)) = excel.worksheet_range(&sheet) {
            let mut mod_name: String = String::new();
            let mut row_num = 0;
            for row in r.rows() {
                row_num = row_num + 1;
                let mut st = row[0].to_string().trim().to_string();
                st.make_ascii_uppercase();
                if st == "MODE_NAME" {
                    mod_name = row[1].to_string().trim().to_string();
                } else if st == "VALUE" {
                    let key = format!("{}:{}", mod_name, row[1].to_string().trim().to_string());
                    if GLOBAL_IDS.read().contains(&key) {
                        error!(
                            "配置了重复的键值!! File: [{}] Sheet: [{}], Row: {} Key: {} \n",
                            &input_file_name, &sheet, row_num, &key
                        );
                        panic!("abort");
                    } else {
                        GLOBAL_IDS.write().insert(key);
                    }
                }
            }
        }
    }
}

pub fn sheet_to_data<'a>(
    input_file_name: String,
    sheet_name: &'a String,
    sheet: &'a Range<DataType>,
) -> SheetData<'a> {
    let mut output_file_name: String = String::new();
    let mut mod_name: String = String::new();
    let mut values: Vec<Vec<&DataType>> = vec![];
    let mut names: Vec<String> = vec![];
    let mut front_types: Vec<String> = vec![];
    let mut back_types: Vec<String> = vec![];
    let mut refs: Vec<String> = vec![];
    let mut enums: Vec<AHashMap<String, usize>> = vec![];
    let mut row_num: usize = 0;
    for row in sheet.rows() {
        row_num = row_num + 1;
        let mut st = row[0].to_string().trim().to_string();
        st.make_ascii_uppercase();
        if st == "MODE_NAME" {
            mod_name = row[1].to_string().trim().to_string();
            output_file_name = row[1].to_string().trim().replace(".", "_").to_lowercase();
        } else if st == "NAMES" {
            for v in row {
                names.push(v.to_string().trim().to_string());
            }
        } else if st == "FRONT_TYPE" {
            for v in row {
                let mut upper_type = v.to_string().trim().to_string();
                upper_type.make_ascii_uppercase();
                front_types.push(upper_type);
            }
        } else if st == "BACK_TYPE" {
            for v in row {
                let mut upper_type = v.to_string().trim().to_string();
                upper_type.make_ascii_uppercase();
                back_types.push(upper_type);
            }
        } else if st == "REF" {
            for v in row {
                let mod_name = v.to_string().trim().to_string();
                refs.push(mod_name);
            }
        } else if st == "ENUM" {
            for v in row {
                let enum_words = v.to_string().trim().to_string();
                let mut dic: AHashMap<String, usize> = AHashMap::new();
                if enum_words == "" {
                    enums.push(dic);
                } else {
                    let words: Vec<&str> = enum_words.split('|').collect();
                    for i in 0..words.len() {
                        dic.insert(words[i].to_string(), i);
                    }
                    enums.push(dic);
                }
            }
        } else if st == "VALUE" {
            let mut row_value: Vec<&DataType> = vec![];
            for i in 0..row.len() {
                let value = row[i].to_string().trim().to_string();
                if refs.len() > 0 && refs[i] != "" && i > 0 {
                    let key = format!("{}:{}", refs[i], &value);
                    //检查引用
                    if !GLOBAL_IDS.read().contains(&key) {
                        error!(
                            "没找到引用的键值!File: {},Sheet: {},Row: {}, Key: {}",
                            input_file_name, sheet_name, row_num, key
                        );
                        panic!("abort")
                    }
                }
                if enums.len() > 0 && enums[i].len() > 0 && i > 0 {
                    if !enums[i].contains_key(&value) {
                        error!(
                            "{} 不在枚举类型 {:?} 中! File: [{}],Sheet: [{}], Row: [{}]",
                            value, &enums[i], input_file_name, sheet_name, row_num
                        );
                        panic!("abort")
                    }
                }
                row_value.push(&row[i]);
            }
            values.push(row_value);
        }
    }

    let info: SheetData = SheetData {
        input_file_name: input_file_name,
        output_file_name: output_file_name,
        sheet_name: sheet_name,
        mod_name: mod_name,
        names: names,
        front_types: front_types,
        back_types: back_types,
        values: values,
        refs: refs,
        enums: enums,
    };
    return info;
}

fn cell_to_json(cell: &DataType, row_type: &String, filename: &String, key: &String) -> Value {
    match cell {
        &DataType::Float(f) if row_type == "INT" => json!(f as i64),

        &DataType::String(ref s) if row_type == "INT" => match s.parse::<i64>() {
            Ok(x) => json!(x),
            Err(e) => {
                error!("ParseError:{:?}, {}, {:?}", e, filename, cell);
                Value::Null
            }
        },
        &DataType::String(ref s) if row_type == "FLOAT" => json!(s
            .parse::<f64>()
            .ok()
            .expect(parse_err(filename, key, s).as_str())),
        &DataType::String(ref s) if row_type == "INT" => json!(s
            .parse::<i64>()
            .ok()
            .expect(parse_err(filename, key, s).as_str())),
        &DataType::String(ref s)
            if row_type == "STRING" || row_type == "CODE" || row_type == "JSON" =>
        {
            json!(s)
        }
        &DataType::Empty if row_type == "LIST" => json!([]),
        &DataType::String(ref s) if row_type == "JSON" => json!(s),
        &DataType::String(ref s) if row_type == "LIST" => {
            let final_str = format!("[{}]", s);
            let data: Value = serde_json::from_str(final_str.as_str())
                .ok()
                .expect(parse_err(filename, key, s).as_str());
            json!(data)
        }

        &DataType::Float(f) if row_type == "LIST" => {
            let final_str = format!("[{}]", f);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data)
        }

        &DataType::Int(f) if row_type == "LIST" => {
            let final_str = format!("[{}]", f);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data)
        }
        &DataType::DateTime(x) => json!(x),
        &DataType::Empty if row_type == "FLOAT" || row_type == "INT" => json!(0),
        &DataType::Empty if row_type == "STRING" => json!(""),
        &DataType::String(ref s) => json!(s),
        &DataType::Bool(b) => json!(b),
        &DataType::Float(f) => json!(f),
        &DataType::Int(i) => json!(i),
        &DataType::Empty => Value::Null,
        &DataType::Error(_) => Value::Null,
    }
}

fn cell_to_string(cell: &DataType, row_type: &String, filename: &String, key: &String) -> String {
    match cell {
        &DataType::Float(f) if row_type == "INT" => json!(f as i64).to_string(),

        &DataType::String(ref s) if row_type == "INT" => match s.parse::<i64>() {
            Ok(x) => json!(x).to_string(),
            Err(_) => {
                error!("{},{:?}", filename, cell);
                Value::Null.to_string()
            }
        },
        &DataType::String(ref s) if row_type == "FLOAT" => json!(s
            .parse::<f64>()
            .ok()
            .expect(parse_err(filename, key, s).as_str()))
        .to_string(),
        &DataType::String(ref s) if row_type == "INT" => json!(s
            .parse::<i64>()
            .ok()
            .expect(parse_err(filename, key, s).as_str()))
        .to_string(),
        &DataType::String(ref s) if row_type == "ATOM" => format!(":{}", s).to_string(),

        &DataType::String(ref s)
            if row_type == "STRING" || row_type == "CODE" || row_type == "JSON" =>
        {
            json!(s).to_string()
        }
        &DataType::Empty if row_type == "LIST" => json!([]).to_string(),
        &DataType::String(ref s) if row_type == "JSON" => json!(s).to_string(),
        &DataType::String(ref s) if row_type == "LIST" => {
            let final_str = format!("[{}]", s);
            let data: Value = serde_json::from_str(final_str.as_str())
                .ok()
                .expect(parse_err(filename, key, s).as_str());
            json!(data).to_string()
        }

        &DataType::Float(f) if row_type == "LIST" => {
            let final_str = format!("[{}]", f);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data).to_string()
        }

        &DataType::Int(f) if row_type == "LIST" => {
            let final_str = format!("[{}]", f);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data).to_string()
        }
        &DataType::DateTime(x) => json!(x).to_string(),
        &DataType::Empty if row_type == "FLOAT" || row_type == "INT" => json!(0).to_string(),
        &DataType::Empty if row_type == "STRING" => json!("").to_string(),
        &DataType::String(ref s) => json!(s).to_string(),
        &DataType::Bool(b) => json!(b).to_string(),
        &DataType::Float(f) => json!(f).to_string(),
        &DataType::Int(i) => json!(i).to_string(),
        &DataType::Empty => "nil".to_string(),
        &DataType::Error(_) => "nil".to_string(),
    }
}

fn parse_err(filename: &String, key: &String, s: &String) -> String {
    let error = format!(
        "[ {} ] [ {} ]\nContent: {} ",
        filename.as_str(),
        key.as_str(),
        s.as_str()
    );
    return error;
}

fn get_module_name(fname: String) -> String {
    let a = fname
        .replace(".ex", "")
        .as_str()
        .to_train_case()
        .replace("-", ".");
    return a;
}
