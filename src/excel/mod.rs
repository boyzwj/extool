use ahash::{AHashMap, AHashSet};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use inflector::Inflector;
use serde_json;
use serde_json::value::Value;
use serde_json::Map;
//use std::collections::HashMap;
use prost::Message;
use prost_reflect::{DescriptorPool, DynamicMessage};
use prost_reflect::{MapKey, Value as PValue};
use regex::Regex;
use static_init::dynamic;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::process::Command;
use std::string::String;
use std::{env, fs};
#[dynamic]
static mut GLOBAL_IDS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_FRONT_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_BACK_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_PBD: Vec<String> = Vec::new();

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
    describes: Vec<String>,
    enums: Vec<AHashMap<String, usize>>,
    values: Vec<Vec<&'a DataType>>,
}

impl<'a> SheetData<'_> {
    pub fn export(&self, format: &String, dst_path: &String) {
        let temp = format.to_uppercase();
        if temp == "JSON" {
            self.data_to_json(&dst_path);
        } else if temp == "LUA" {
            self.data_to_lua(&dst_path);
        } else if temp == "EX" {
            self.data_to_ex(&dst_path);
        } else if temp == "PBD" {
            self.data_to_pbd(&dst_path);
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
                        let real_type = &get_real_front_type(&self.refs[i], row_type);
                        let value =
                            cell_to_json(&rv[i], &real_type, &self.output_file_name, &column_name);
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
                    // let dic = &self.enums[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let real_type = &get_real_front_type(&self.refs[i], row_type);
                        let value =
                            cell_to_string(&rv[i], real_type, &self.output_file_name, &column_name);
                        columns.push(format!("{}", value.replace("[", "{").replace("]", "}")));
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if let Some(x) = &self.enums[i].get(value) {
                            columns.push(format!("{}", x));
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
                    &self.front_types[1],
                    &self.output_file_name,
                    &self.names[1],
                );
                res.push(format!("\t[{}] = {{{}}}", keyvalue, columns.join(",")));
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return;
        }
        let out = format!(
            "local KT = {{{}}}\n\
    local data = {{ \n {}\n}}\n\
    do\n\
    \tlocal base = {{\n\
    \t\t__index = function(table,key)\n\
    \t\t\tlocal ki = KT[key]\n\
    \t\t\tif not ki then\n\
    \t\t\t\treturn nil\n\
    \t\t\tend\n\
    \t\t\treturn table[ki]
    \tend,\n\
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
                        let real_type = &get_real_back_type(&self.refs[i], row_type);

                        let value =
                            cell_to_string(&rv[i], real_type, &self.output_file_name, &column_name);
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

    pub fn data_to_pbd(&self, out_path: &String) {
        if self.mod_name == "" {
            return;
        }
        let mut field_schemas: Vec<String> = vec![];
        let msg_name = self.mod_name.to_string().replace("Data.", "");
        let mut n = 1;
        let float_reg = Regex::new(r"^-?([1-9]\d*\.\d*|0\.\d*[1-9]\d*|0?\.0+|0)$").unwrap();
        let integer_reg = Regex::new(r"^-?[0-9]*$").unwrap();
        let mut valid_columns: Vec<usize> = vec![];
        let mut valid_front_types: Vec<String> = vec![];
        for i in 1..self.front_types.len() {
            let ft = &self.front_types[i];
            let fk = &self.names[i];
            let des = &self.describes[i];
            let enu = &self.enums[i];
            if ft != "" && fk != "" {
                let (mut field_schema, mut front_type) = match ft.as_str() {
                    "INT" => ("int64".to_string(), "INT64".to_string()),
                    "FLOAT" => ("double".to_string(), "FLOAT".to_string()),
                    "LIST" => {
                        let mut dt = "LIST,STRING".to_string();
                        let mut fs = format!("repeated {}", "string");
                        for j in 0..self.values.len() {
                            let temp = &self.values[j][i].to_string();
                            if temp == "" {
                                continue;
                            }
                            let list: Vec<&str> = temp.split(',').collect();
                            let v1 = list[0];
                            if float_reg.is_match(v1) {
                                dt = "LIST,FLOAT".to_string();
                                fs = format!("repeated {}", "double");
                                break;
                            }

                            if integer_reg.is_match(v1) {
                                dt = "LIST,INT64".to_string();
                                fs = format!("repeated {}", "int64");
                                break;
                            }
                            break;
                        }
                        (fs, dt)
                    }
                    "LOC_STRING" => ("string".to_string(), "STRING".to_string()),
                    column_type => (column_type.to_lowercase(), column_type.to_string()),
                };
                if !front_type.contains("LIST") && !enu.is_empty() {
                    front_type = "ENUM".to_string();
                    field_schema = "uint32".to_string();
                }
                let fr = &self.refs[i];
                if enu.is_empty() && !fr.is_empty() {
                    front_type = get_real_front_type(fr, &front_type);
                    field_schema = front_type.to_lowercase();
                }
                field_schemas.push(format!("\t{} {} = {}; //{}", &field_schema, fk, n, des));
                valid_columns.push(i);
                valid_front_types.push(front_type);
                n = n + 1;
            }
        }
        if field_schemas.len() == 0 {
            return;
        }
        let msg_schema = format!(
            "message {}{{\n\
            {}\n\
            }}",
            msg_name,
            field_schemas.join("\n")
        );
        let key_name = &self.names[valid_columns[0]];
        if !valid_front_types[0].contains("INT") {
            error!(
                "主键 [{}] 仅支持整型类型! File: [{}] Sheet: [{}],Mod_name: [{}], Key: {}, Type: {}\n",
                key_name, &self.input_file_name, &self.sheet_name, &self.mod_name, key_name, valid_front_types[0]
            );
            panic!("abort");
        }
        let out = format!(
            "message {}Info{{\n\
             \tmap<{},{}> data = 1;\n\
            }}\n\
            {}",
            msg_name,
            &valid_front_types[0].to_lowercase(),
            msg_name,
            msg_schema
        );

        let content = format!(
            "syntax = \"proto3\";\n\
            package pbd;\n\
            \n\
            {}",
            out
        );
        let path_str = format!("{}/{}.proto", out_path, msg_name);
        self.write_file(&path_str, &content);
        GLOBAL_PBD.write().push(out);

        // builder bin data
        let mut builder = prost_reflect_build::Builder::new();
        let bin_path = format!("{}/{}.bin", out_path, msg_name);
        builder.file_descriptor_set_path(&bin_path);
        env::set_var("OUT_DIR", out_path);
        builder.compile_protos(&[&path_str], &["."]).unwrap();
        // new  messagedescriptor
        let bytes = fs::read(&bin_path).unwrap();
        let pool = DescriptorPool::decode(bytes.as_ref()).unwrap();
        let n1 = format!("pbd.{}Info", msg_name);
        let n2 = format!("pbd.{}", msg_name);
        let info_des = pool.get_message_by_name(&n1).unwrap();
        let mut info_dm = DynamicMessage::new(info_des);
        let msg_des = pool.get_message_by_name(&n2).unwrap();
        let mut data: HashMap<MapKey, PValue> = HashMap::new();
        let mut row = 1;
        for x in 0..self.values.len() {
            let mut dm = DynamicMessage::new(msg_des.clone());
            for y in 0..valid_columns.len() {
                let i = valid_columns[y];
                let ft = &valid_front_types[y];
                let fk = &self.names[i];
                let fv = self.values[x][i];
                let enu_val = &fv.to_string().trim().to_string();
                if ft == "ENUM" {
                    if let Some(x) = self.enums[i].get(enu_val) {
                        dm.set_field_by_name(fk, PValue::U32(*x as u32));
                    } else {
                        error!(
                            "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                            fk,
                            &self.values[x][1],
                            enu_val,
                            &self.enums[i].keys()
                        );
                    }
                } else {
                    let fr = &self.refs[i];
                    if !fr.is_empty() {
                        let new_ft = GLOBAL_FRONT_PRIMARYS.read().get(fr).unwrap().to_string();
                        let p_val = cell_to_pvalue(fv, &new_ft, &self.mod_name, fk);
                        dm.set_field_by_name(fk, p_val);
                    } else {
                        let p_val = cell_to_pvalue(fv, &ft.to_string(), &self.mod_name, fk);
                        dm.set_field_by_name(fk, p_val);
                    };
                    // println!("ft: {},fk: {},fv: {:?},p_val: {:?}", ft, fk, fv, p_val);
                }
            }
            let key_val = dm.get_field_by_name_mut(key_name).unwrap().clone();
            let dy_msg = PValue::Message(dm);
            match key_val {
                PValue::I32(ref s) => data.insert(MapKey::I32(*s), dy_msg),
                PValue::I64(ref s) => data.insert(MapKey::I64(*s), dy_msg),
                PValue::U32(ref s) => data.insert(MapKey::U32(*s), dy_msg),
                PValue::U64(ref s) => data.insert(MapKey::U64(*s), dy_msg),
                PValue::String(ref s) => data.insert(MapKey::String(s.to_string()), dy_msg),
                _ => {
                    error!("键值的数据类型不对! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Key: {}\n", &self.input_file_name,&self.sheet_name,&self.mod_name,row,key_name);
                    panic!("abort");
                }
            };
            row = row + 1;
        }
        info_dm.set_field_by_name("data", PValue::Map(data));
        let mut buf = vec![];
        info_dm.encode(&mut buf).unwrap();
        let out_pbd_path = format!("{}/{}.pbd", out_path, msg_name);
        fs::write(out_pbd_path, buf).unwrap();
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
                if st == "MOD" {
                    mod_name = row[1].to_string().trim().to_string();
                } else if st == "VALUE" {
                    let key = format!("{}:{}", mod_name, row[1].to_string().trim().to_string());
                    if GLOBAL_IDS.read().contains(&key) {
                        error!(
                            "配置了重复的键值!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Key: {} \n",
                            &input_file_name, &sheet,&mod_name, row_num, &key
                        );
                        panic!("abort");
                    } else {
                        GLOBAL_IDS.write().insert(key);
                    }
                } else if st == "BACK_TYPE" {
                    for i in 1..row.len() {
                        let rv = row[i].clone().to_string().trim().to_uppercase();
                        if !rv.is_empty() {
                            GLOBAL_BACK_PRIMARYS
                                .write()
                                .insert(mod_name.to_string(), rv);
                            break;
                        }
                    }
                } else if st == "FRONT_TYPE" {
                    for i in 1..row.len() {
                        let rv = row[i].clone().to_string().trim().to_uppercase();
                        if !rv.is_empty() {
                            GLOBAL_FRONT_PRIMARYS
                                .write()
                                .insert(mod_name.to_string(), rv);
                            break;
                        }
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
    let mut describes: Vec<String> = vec![];
    let mut enums: Vec<AHashMap<String, usize>> = vec![];
    let mut row_num: usize = 0;
    for row in sheet.rows() {
        row_num = row_num + 1;
        let mut st = row[0].to_string().trim().to_string();
        st.make_ascii_uppercase();
        if st == "MOD" {
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
        } else if st == "DES" {
            for v in row {
                describes.push(v.to_string().trim().to_string());
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
                if refs.len() > 0 && refs[i] != "" && i > 0 && value != "" {
                    //检查引用
                    if front_types[i] == "LIST" || back_types[i] == "LIST" {
                        let vals: Vec<&str> = value.split(",").collect();
                        for j in 0..vals.len() {
                            let v = vals[j];
                            let key = format!("{}:{}", refs[i], &v);
                            if !GLOBAL_IDS.read().contains(&key) {
                                error!(
                                    "没找到引用的键值!File: {},Sheet: {},Row: {}, Key: {}",
                                    input_file_name, sheet_name, row_num, key
                                );
                                panic!("abort")
                            }
                        }
                    } else {
                        let key = format!("{}:{}", refs[i], &value);
                        if !GLOBAL_IDS.read().contains(&key) {
                            error!(
                                "没找到引用的键值!File: {},Sheet: {},Row: {}, Key: {}",
                                input_file_name, sheet_name, row_num, key
                            );
                            panic!("abort")
                        }
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
        describes: describes,
        enums: enums,
    };
    return info;
}

pub fn create_pbd_file(out_path: &String) {
    let content = format!(
        "syntax = \"proto3\";\n\
        package pbd;\n\
        \n\
        {}",
        GLOBAL_PBD.read().join("\n\n")
    );

    let objects = fs::read_dir(out_path).unwrap();

    for obj in objects {
        let path = obj.unwrap().path();
        let temp_path = path.display().to_string();
        if temp_path.ends_with(".bin")
            || temp_path.ends_with(".rs")
            || temp_path.ends_with(".proto")
        {
            fs::remove_file(path).ok();
        }
    }
    let pbd_path = format!("{}/pbd.proto", out_path);
    fs::write(&pbd_path, content).unwrap();

    Command::new("protoc")
        .arg(&pbd_path)
        .arg(format!("--csharp_out={}", out_path))
        .output()
        .expect("failed to execute process");
}

fn is_number_str(value: &str) -> bool {
    match value.parse::<i64>() {
        Ok(_) => true,
        Err(_) => match value.parse::<f64>() {
            Ok(_) => true,
            Err(_) => false,
        },
    }
}

fn cell_to_json(cell: &DataType, row_type: &String, filename: &String, key: &String) -> Value {
    let s = cell.to_string().trim().to_string();
    if row_type.starts_with("INT") || row_type.starts_with("UINT") {
        if s == "" {
            return json!(0);
        }
        json!(s
            .parse::<i64>()
            .ok()
            .expect(parse_err(filename, row_type, &s).as_str()))
    } else if row_type.starts_with("FLOAT") {
        if s == "" {
            return json!(0);
        }
        json!(s
            .parse::<f64>()
            .ok()
            .expect(parse_err(filename, key, &s).as_str()))
    } else if row_type.starts_with("LIST_UINT")
        || row_type.starts_with("LIST_INT")
        || row_type.starts_with("LIST_FLOAT")
    {
        if s == "" {
            return json!([]);
        }

        let final_str = format!("[{}]", s);
        let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
        json!(data)
    } else if row_type.starts_with("LIST_STRING") {
        if s == "" {
            return json!([]);
        }
        let mut data: Vec<Value> = vec![];
        let list: Vec<&str> = s.split(',').collect();
        for val in list {
            data.push(json!(val))
        }
        json!(data)
    } else if row_type == "LIST" {
        if s == "" {
            return json!([]);
        }
        let list: Vec<&str> = s.split(',').collect();
        if list.clone().into_iter().any(|x| is_number_str(x) == false) {
            let mut data: Vec<Value> = vec![];
            for val in list {
                data.push(json!(val))
            }
            json!(data)
        } else {
            let final_str = format!("[{}]", s);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data)
        }
    } else {
        json!(s)
    }
}

fn cell_to_string(cell: &DataType, row_type: &String, filename: &String, key: &String) -> String {
    let s = cell.to_string().trim().to_string();
    let s = cell.to_string().trim().to_string();
    if row_type.starts_with("INT") || row_type.starts_with("UINT") || row_type.starts_with("FLOAT")
    {
        if s == "" {
            return "0".to_string();
        }
        return s;
    } else if row_type.starts_with("LIST_UINT")
        || row_type.starts_with("LIST_INT")
        || row_type.starts_with("LIST_FLOAT")
    {
        if s == "" {
            return "[]".to_string();
        }

        return format!("[{}]", s);
    } else if row_type.starts_with("LIST_STRING") {
        if s == "" {
            return "[]".to_string();
        }
        let mut data: Vec<Value> = vec![];
        let list: Vec<&str> = s.split(',').collect();
        for val in list {
            data.push(json!(val))
        }
        json!(data).to_string()
    } else if row_type == "LIST" {
        if s == "" {
            return "[]".to_string();
        }
        let list: Vec<&str> = s.split(',').collect();
        if list.clone().into_iter().any(|x| is_number_str(x) == false) {
            let mut data: Vec<Value> = vec![];
            for val in list {
                data.push(json!(val))
            }
            json!(data).to_string()
        } else {
            let final_str = format!("[{}]", s);
            let data: Value = serde_json::from_str(final_str.as_str()).unwrap();
            json!(data).to_string()
        }
    } else if row_type == "STRING" {
        if s == "" {
            return "\"\"".to_string();
        }
        return format!("\"{}\"", s);
    } else {
        s
    }
}

fn cell_to_pvalue(cell: &DataType, row_type: &String, filename: &String, key: &String) -> PValue {
    match cell {
        // INT32
        &DataType::Int(ref s) if row_type == "INT32" => PValue::I32(*s as i32),
        &DataType::Float(f) if row_type == "INT32" => PValue::I32(f as i32),
        &DataType::String(ref s) if row_type == "INT32" => PValue::I32(
            s.parse::<i32>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // INT64
        &DataType::Int(ref s) if row_type == "INT64" || row_type == "INT" => PValue::I64(*s),
        &DataType::Float(f) if row_type == "INT64" || row_type == "INT" => PValue::I64(f as i64),
        &DataType::String(ref s) if row_type == "INT64" || row_type == "INT" => PValue::I64(
            s.parse::<i64>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // UINT32
        &DataType::Int(ref s) if row_type == "UINT32" => PValue::U32(*s as u32),
        &DataType::Float(f) if row_type == "UINT32" => PValue::U32(f as u32),
        &DataType::String(ref s) if row_type == "UINT32" => PValue::U32(
            s.parse::<u32>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // UINT64
        &DataType::Int(ref s) if row_type == "UINT64" => PValue::U64(*s as u64),
        &DataType::Float(f) if row_type == "UINT64" => PValue::U64(f as u64),
        &DataType::String(ref s) if row_type == "UINT64" => PValue::U64(
            s.parse::<u64>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // FLOAT
        &DataType::Int(ref s) if row_type == "FLOAT" => PValue::F64(*s as f64),
        &DataType::Float(ref s) if row_type == "FLOAT" => PValue::F64(*s),
        &DataType::String(ref s) if row_type == "FLOAT" => PValue::F64(
            s.parse::<f64>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // STRING
        &DataType::Int(ref s) if row_type == "STRING" => PValue::String(s.to_string()),
        &DataType::Float(ref s) if row_type == "STRING" => PValue::String(s.to_string()),
        &DataType::String(ref s) if row_type == "STRING" => PValue::String(s.to_string()),

        // LIST
        &DataType::Empty if row_type.contains("LIST") => PValue::List([].to_vec()),
        &DataType::Int(f) if row_type == "LIST,INT64" => PValue::List(vec![PValue::I64(f)]),
        &DataType::Float(f) if row_type == "LIST,FLOAT" => PValue::List(vec![PValue::F64(f)]),
        &DataType::String(ref s) if row_type == "LIST,STRING" => {
            let list: Vec<&str> = s.split(',').collect();
            let mut result: Vec<PValue> = vec![];
            for i in 0..list.len() {
                result.push(PValue::String(list[i].to_string()))
            }
            PValue::List(result)
        }
        &DataType::String(ref s) if row_type == "LIST,FLOAT" => {
            let list: Vec<&str> = s.split(',').collect();
            let mut result: Vec<PValue> = vec![];
            for i in 0..list.len() {
                result.push(PValue::F64(
                    list[i]
                        .parse::<f64>()
                        .ok()
                        .expect(parse_err(filename, key, s).as_str()),
                ))
            }
            PValue::List(result)
        }
        &DataType::String(ref s) if row_type == "LIST,INT64" => {
            let list: Vec<&str> = s.split(',').collect();
            let mut result: Vec<PValue> = vec![];
            for i in 0..list.len() {
                result.push(PValue::I64(
                    list[i]
                        .parse::<i64>()
                        .ok()
                        .expect(parse_err(filename, key, s).as_str()),
                ))
            }
            PValue::List(result)
        }

        &DataType::DateTime(x) => PValue::String(x.to_string()),
        &DataType::Empty if row_type == "INT32" => PValue::I32(0),
        &DataType::Empty if row_type == "INT64" || row_type == "INT" => PValue::I64(0),
        &DataType::Empty if row_type == "UINT32" => PValue::U32(0),
        &DataType::Empty if row_type == "UINT64" => PValue::U64(0),
        &DataType::Empty if row_type == "FLOAT" => PValue::F64(0.0),
        &DataType::Empty if row_type == "STRING" => PValue::String("".to_string()),
        &DataType::String(ref s) => PValue::String(s.to_string()),
        &DataType::Bool(b) => PValue::Bool(b),
        &DataType::Float(f) => PValue::F64(f),
        &DataType::Int(i) => PValue::I64(i),
        &DataType::Empty => PValue::String("".to_string()),
        &DataType::Error(_) => PValue::String("".to_string()),
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

fn get_real_front_type(mod_name: &String, old_type: &String) -> String {
    if mod_name.trim().is_empty() {
        return old_type.to_string();
    }
    if let Some(x) = GLOBAL_FRONT_PRIMARYS.read().get(mod_name) {
        if old_type == "LIST" {
            return format!("LIST_{}", x);
        } else {
            return x.to_string();
        }
    } else {
        return old_type.to_string();
    }
}

fn get_real_back_type(mod_name: &String, old_type: &String) -> String {
    if mod_name.trim().is_empty() {
        return old_type.to_string();
    }
    if let Some(x) = GLOBAL_BACK_PRIMARYS.read().get(mod_name) {
        if old_type == "LIST" {
            return format!("LIST_{}", x);
        } else {
            return x.to_string();
        }
    } else {
        return old_type.to_string();
    }
}
