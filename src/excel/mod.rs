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
use std::env;
use std::fs;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::process::Command;
use std::string::String;
#[dynamic]
static mut GLOBAL_IDS: AHashSet<String> = AHashSet::new();

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
        if format == "JSON" {
            self.data_to_json(&dst_path);
        } else if format == "LUA" {
            self.data_to_lua(&dst_path);
        } else if format == "EX" {
            self.data_to_ex(&dst_path);
        } else if format == "PBD" {
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
                    // let dic = &self.enums[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let value =
                            cell_to_string(&rv[i], row_type, &self.output_file_name, &column_name);
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
        let mut valid_front_types: Vec<&str> = vec![];
        for i in 1..self.front_types.len() {
            let ft = &self.front_types[i];
            let fk = &self.names[i];
            let des = &self.describes[i];

            if ft != "" && fk != "" {
                let (field_schema, front_type) = match ft.as_str() {
                    "STRING" => ("string".to_string(), "string"),
                    "INT" => ("int64".to_string(), "int64"),
                    "INT32" => ("int64".to_string(), "int64"),
                    "INT64" => ("int64".to_string(), "int64"),
                    "UINT32" => ("int64".to_string(), "int64"),
                    "UINT64" => ("int64".to_string(), "int64"),
                    "FLOAT" => ("double".to_string(), "float"),
                    "LIST" => {
                        let mut dt = "list,string";
                        let mut fs = format!("repeated {}", "string");
                        for j in 0..self.values.len() {
                            let temp = &self.values[j][i].to_string();
                            if temp == "" {
                                continue;
                            }
                            let list: Vec<&str> = temp.split(',').collect();
                            let v1 = list[0];
                            if float_reg.is_match(v1) {
                                dt = "list,float";
                                fs = format!("repeated {}", "double");
                                break;
                            }

                            if integer_reg.is_match(v1) {
                                dt = "list,int64";
                                fs = format!("repeated {}", "int64");
                                break;
                            }
                            break;
                        }
                        (fs, dt)
                    }
                    _ => ("string".to_string(), "string"),
                };
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
        if valid_front_types[0] == "float" {
            error!(
                "主键 [{}] 不支持double类型! File: [{}] Sheet: [{}],Mod_name: [{}] Key: {}\n",
                key_name, &self.input_file_name, &self.sheet_name, &self.mod_name, key_name
            );
            panic!("abort");
        }
        let out = format!(
            "message {}Info{{\n\
             \tmap<{},{}> data = 1;\n\
            }}\n\
            {}",
            msg_name, &valid_front_types[0], msg_name, msg_schema
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
        let out_dir = env::var_os("OUT_DIR")
            .unwrap()
            .to_str()
            .unwrap()
            .to_string();
        let bin_path = format!("{}/{}.bin", out_dir, msg_name);
        builder.file_descriptor_set_path(&bin_path);
        builder.compile_protos(&[&path_str], &[out_path]).unwrap();

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
                let ft = valid_front_types[y];
                let fk = &self.names[i];
                let fv = self.values[x][i];
                let p_val = cell_to_pvalue(fv, ft, &self.mod_name, fk);
                // println!("ft: {},fk: {},fv: {:?},p_val: {:?}", ft, fk, fv, p_val);
                dm.set_field_by_name(fk, p_val)
            }
            let key_val = dm.get_field_by_name_mut(key_name).unwrap().clone();
            let dy_msg = PValue::Message(dm);
            match key_val {
                PValue::I64(ref s) => data.insert(MapKey::I64(*s), dy_msg),
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
        if path.display().to_string().ends_with(".proto") {
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

fn cell_to_json(cell: &DataType, row_type: &String, filename: &String, key: &String) -> Value {
    match cell {
        &DataType::Float(f) if row_type.contains("INT") => json!(f as i64),

        &DataType::String(ref s) if row_type.contains("INT") => match s.parse::<i64>() {
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
        &DataType::String(ref s) if row_type.contains("INT") => json!(s
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
        &DataType::Empty if row_type == "FLOAT" || row_type.contains("INT") => json!(0),
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
        &DataType::Float(f) if row_type.contains("INT") => json!(f as i64).to_string(),

        &DataType::String(ref s) if row_type.contains("INT") => match s.parse::<i64>() {
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
        &DataType::String(ref s) if row_type.contains("INT") => json!(s
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
        &DataType::Empty if row_type == "FLOAT" || row_type.contains("INT") => json!(0).to_string(),
        &DataType::Empty if row_type == "STRING" => json!("").to_string(),
        &DataType::String(ref s) => json!(s).to_string(),
        &DataType::Bool(b) => json!(b).to_string(),
        &DataType::Float(f) => json!(f).to_string(),
        &DataType::Int(i) => json!(i).to_string(),
        &DataType::Empty => "nil".to_string(),
        &DataType::Error(_) => "nil".to_string(),
    }
}

fn cell_to_pvalue(cell: &DataType, row_type: &str, filename: &String, key: &String) -> PValue {
    match cell {
        // int
        &DataType::Int(ref s) if row_type == "int64" => PValue::I64(*s),
        &DataType::Float(f) if row_type == "int64" => PValue::I64(f as i64),
        &DataType::String(ref s) if row_type == "int64" => PValue::I64(
            s.parse::<i64>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // float
        &DataType::Int(ref s) if row_type == "float" => PValue::F64(*s as f64),
        &DataType::Float(ref s) if row_type == "float" => PValue::F64(*s),
        &DataType::String(ref s) if row_type == "float" => PValue::F64(
            s.parse::<f64>()
                .ok()
                .expect(parse_err(filename, key, s).as_str()),
        ),

        // string
        &DataType::Int(ref s) if row_type == "string" => PValue::String(s.to_string()),
        &DataType::Float(ref s) if row_type == "string" => PValue::String(s.to_string()),
        &DataType::String(ref s) if row_type == "string" => PValue::String(s.to_string()),

        // list
        &DataType::Empty if row_type.contains("list") => PValue::List([].to_vec()),
        &DataType::Int(f) if row_type == "list,int64" => PValue::List(vec![PValue::I64(f)]),
        &DataType::Float(f) if row_type == "list,float" => PValue::List(vec![PValue::F64(f)]),
        &DataType::String(ref s) if row_type == "list,string" => {
            let list: Vec<&str> = s.split(',').collect();
            let mut result: Vec<PValue> = vec![];
            for i in 0..list.len() {
                result.push(PValue::String(list[i].to_string()))
            }
            PValue::List(result)
        }
        &DataType::String(ref s) if row_type == "list,float" => {
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
        &DataType::String(ref s) if row_type == "list,int64" => {
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
        &DataType::Empty if row_type == "int64" => PValue::I64(0),
        &DataType::Empty if row_type == "float" => PValue::F64(0.0),
        &DataType::Empty if row_type == "string" => PValue::String("".to_string()),
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
