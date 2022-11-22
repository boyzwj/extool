use ahash::{AHashMap, AHashSet};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use indexmap::IndexMap;
use inflector::Inflector;
use pinyin::ToPinyin;
use prost::Message;
use prost_reflect::{DescriptorPool, DynamicMessage};
use prost_reflect::{MapKey, Value as PValue};
use serde_json;
use serde_json::value::Value;
use serde_json::Map;
use simple_excel_writer::*;
use static_init::dynamic;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::io::{Error, ErrorKind};
use std::string::String;
use std::{env, fs};

#[dynamic]
static mut GLOBAL_IDS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_FRONT_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_BACK_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_PBD: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_LANG: AHashMap<String, String> = AHashMap::new();
#[dynamic]
static mut GLOBAL_EXCLUDE_SHEETS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_MODS: AHashSet<String> = AHashSet::new();

// static mut
pub struct SheetData<'a> {
    input_file_name: String,
    output_file_name: String,
    sheet_name: &'a String,
    mod_name: String,
    names: Vec<String>,
    refs: Vec<String>,
    types: Vec<String>,
    describes: Vec<String>,
    enums: Vec<AHashMap<String, usize>>,
    values: Vec<Vec<&'a DataType>>,
    force_mods: Vec<String>,
}

impl<'a> SheetData<'_> {
    pub fn export(&self, format: &String, dst_path: &String) -> Result<usize, usize> {
        let temp = format.to_uppercase();
        if temp == "JSON" {
            return self.data_to_json(&dst_path);
        } else if temp == "LUA" {
            return self.data_to_lua(&dst_path);
        } else if temp == "EX" {
            return self.data_to_ex(&dst_path);
        } else if temp == "PBD" {
            return self.data_to_pbd(&dst_path);
        } else if temp == "LANG" {
            return self.data_to_lang(&dst_path);
        }
        return Ok(0);
    }

    pub fn write_file(&self, path_str: &str, content: &String) -> usize {
        match File::create(path_str) {
            Err(e) => {
                error!("创建导出文件失败 : {:?}", e);
                return 1;
            }
            Ok(f) => {
                let mut bw = BufWriter::new(f);
                match bw.write_all(content.as_bytes()) {
                    Err(e) => {
                        error!("写入导出内容失败 : {:?}", e);
                        return 1;
                    }
                    Ok(_) => {
                        info!(
                            "成功 [{}] [{}] 导出 [{}] 条记录",
                            self.input_file_name,
                            self.sheet_name,
                            self.values.len()
                        );
                        return 0;
                    }
                }
            }
        }
    }

    pub fn data_to_json(&self, out_path: &String) -> Result<usize, usize> {
        let mut res: Map<String, Value> = Map::new();
        for rv in &self.values {
            let mut map: Map<String, Value> = Map::new();
            for i in 1..self.types.len() {
                if self.types[i] != "" {
                    let column_name = &self.names[i];
                    let origin_type = &self.types[i];
                    let ref_name = &self.refs[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let real_type = &get_real_front_type(ref_name, origin_type, false);
                        let value = cell_to_json(
                            &rv[i],
                            &real_type,
                            &self.output_file_name,
                            &column_name,
                            &self.sheet_name,
                        );
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
                            );
                            return Err(1);
                        }
                    }
                }
            }

            if map.len() > 0 {
                let key_is_enum = self.enums[1].is_empty() == false;
                let enum_value = rv[1].to_string().trim().to_string();
                let mut keyvalue = enum_value.to_string();
                if key_is_enum {
                    keyvalue = self.enums[1].get(&enum_value).unwrap().to_string();
                }

                res.insert(keyvalue, json!(map));
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return Ok(0);
        }
        let obj = json!(res);
        let json = serde_json::to_string_pretty(&obj).unwrap();
        let path_str = format!("{}/{}.json", out_path, self.output_file_name);
        match self.write_file(&path_str, &json) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_lua(&self, out_path: &String) -> Result<usize, usize> {
        let mut res: Vec<String> = vec![];
        let mut kns: Vec<String> = vec![];
        let mut j = 1;
        for i in 1..self.types.len() {
            if self.types[i] != "" {
                let column_name = &self.names[i];
                kns.push(format!("{} = {}", column_name, j.to_string()));
                j = j + 1;
            }
        }
        for rv in &self.values {
            let mut columns: Vec<String> = vec![];
            for i in 1..self.types.len() {
                if self.types[i] != "" {
                    let column_name = &self.names[i];
                    let origin_type = &self.types[i];
                    // let dic = &self.enums[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let real_type = &get_real_front_type(&self.refs[i], origin_type, false);
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
                            return Err(1);
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let key_is_enum = self.enums[1].is_empty() == false;
                let mut keyvalue = cell_to_string(
                    &rv[1],
                    &self.types[1],
                    &self.output_file_name,
                    &self.names[1],
                );
                if key_is_enum {
                    let enum_value = &rv[1].to_string().trim().to_string();
                    keyvalue = self.enums[1].get(enum_value).unwrap().to_string();
                }

                res.push(format!("\t[{}] = {{{}}}", keyvalue, columns.join(",")));
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return Ok(0);
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
        match self.write_file(&path_str, &out) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_ex(&self, dst_path: &String) -> Result<usize, usize> {
        let mut res: Vec<String> = vec![];
        let mut ids: Vec<String> = vec![];
        for rv in &self.values {
            let mut columns: Vec<String> = vec![];
            for i in 1..self.types.len() {
                if self.types[i] != "" {
                    let column_name = &self.names[i];
                    let origin_type = &self.types[i];
                    if self.enums.len() == 0 || self.enums[i].len() == 0 {
                        let real_type = &get_real_back_type(&self.refs[i], origin_type, false);

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
                            return Err(1);
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let key_is_enum = self.enums[1].is_empty() == false;
                let mut keyvalue = cell_to_string(
                    &rv[1],
                    &self.types[1],
                    &self.output_file_name,
                    &self.names[1],
                );
                if key_is_enum {
                    let enum_value = &rv[1].to_string().trim().to_string();
                    keyvalue = self.enums[1].get(enum_value).unwrap().to_string();
                }

                res.push(format!(
                    "\tdef get({}) do\n\t\t%{{\n{}\n\t\t}}\n\tend\n",
                    keyvalue,
                    columns.join(",\n")
                ));
                ids.push(keyvalue.to_string());
            }
        }

        if self.values.len() == 0 || self.output_file_name.len() <= 0 {
            return Ok(0);
        }
        let module_name = self.mod_name.clone();
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
        match self.write_file(&path_str, &out) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_pbd(&self, out_path: &String) -> Result<usize, usize> {
        if self.mod_name == "" {
            return Ok(0);
        }
        let mut field_schemas: Vec<String> = vec![];
        let msg_name = self.mod_name.to_string().replace("Data.", "");
        let mut n = 1;
        let mut valid_columns: Vec<usize> = vec![];
        let mut valid_front_types: Vec<String> = vec![];
        for i in 1..self.types.len() {
            let origin_type = &self.types[i];
            let fk = &self.names[i];
            let des = &self.describes[i];
            let is_enum = self.enums[i].is_empty() == false;
            let ref_name = &self.refs[i];
            if origin_type != "" && fk != "" {
                let ft = get_real_front_type(ref_name, origin_type, is_enum);
                let field_schema = match ft.as_str() {
                    "LIST_UINT32" => "repeated uint32".to_string(),
                    "LIST_UINT64" => "repeated uint64".to_string(),
                    "LIST_INT32" => "repeated int32".to_string(),
                    "LIST_INT64" => "repeated int64".to_string(),
                    "LIST_FLOAT" => "repeated float".to_string(),
                    "LIST_STRING" => "repeated string".to_string(),
                    "ENUM" => "uint32".to_string(),
                    // "ENUM" => convert_enum_type(fk),
                    "STRING_LOC" => "uint32".to_string(),
                    _ => ft.to_lowercase(),
                };
                field_schemas.push(format!("\t{} {} = {}; //{}", &field_schema, fk, n, des));
                valid_columns.push(i);
                valid_front_types.push(ft);
                n = n + 1;
            }
        }
        if field_schemas.len() == 0 {
            return Ok(0);
        }

        let key_name = &self.names[valid_columns[0]];
        if valid_front_types[0].contains("LIST") || valid_front_types[0] == "FLOAT" {
            error!(
                "主键仅支持整型类型和字符串类型!请确定字段[{}]的类型  File: [{}] Sheet: [{}],Mod_name: [{}], Key: {}, Type: {}\n",
                key_name, &self.input_file_name, &self.sheet_name, &self.mod_name, key_name, valid_front_types[0]
            );
            return Err(1);
        }
        let mut map_key_type = valid_front_types[0].to_string();
        if map_key_type == "ENUM" {
            map_key_type = "UINT32".to_string();
        }

        if map_key_type == "STRING_LOC" {
            map_key_type = "UINT32".to_string();
        }

        let mut class_name = msg_name.to_string();
        if self.force_mods.len() > 0 && !self.force_mods[1].is_empty() {
            class_name = self.force_mods[1].to_string();
        }

        let mut enum_msgs = vec![];
        for y in 0..valid_columns.len() {
            let i = valid_columns[y];
            let ft = &valid_front_types[y];
            if ft == "ENUM" {
                match convert_to_pinyin(&self.enums[i]) {
                    Ok(arr) => {
                        let mut enum_schemas = vec![];
                        let fk = self.names[i].to_lowercase();
                        for (k, v, comment) in arr {
                            enum_schemas.push(format!("\t\t{}_{} = {};//{}", fk, v, k, comment))
                        }
                        let msg_enum = format!(
                            "\tenum {}
    {{\n\
                            {}\n\
                            \t}}",
                            convert_enum_type(&self.names[i]),
                            enum_schemas.join("\n")
                        );
                        enum_msgs.push(msg_enum)
                    }
                    Err(txt) => {
                        error!(
                            "Enum的值[{}]只能纯汉字或者纯英文  File: [{}] Sheet: [{}],Mod_name: [{}], Key: {}, Type: {}\n",
                            txt, &self.input_file_name, &self.sheet_name, &self.mod_name,  &self.names[i], ft
                        );
                        return Err(1);
                    }
                }
            }
        }

        // let msg_enum = format!("{}", "");

        let msg_schema = format!(
            "message {}{{\n\
            {}\n\
            {}\n\
            }}",
            class_name,
            field_schemas.join("\n"),
            enum_msgs.join("\n")
        );

        let out = format!(
            "message Data{}{{\n\
             \tmap<{},{}> data = 1;\n\
            }}\n\
            {}",
            class_name,
            &map_key_type.to_lowercase(),
            class_name,
            msg_schema
        );

        let content = format!(
            "syntax = \"proto3\";\n\
            package pbd;\n\
            \n\
            {}",
            out
        );
        let path_str = format!("{}/tmp_{}.proto", out_path, msg_name.to_lowercase());
        match self.write_file(&path_str, &content) {
            0 => (),
            _ => return Err(1),
        }
        GLOBAL_PBD.write().insert(class_name.to_string(), out);

        // builder bin data
        let mut builder = prost_reflect_build::Builder::new();
        let bin_path = format!("{}/tmp_{}.bin", out_path, msg_name.to_lowercase());
        builder.file_descriptor_set_path(&bin_path);
        env::set_var("OUT_DIR", out_path);
        match builder.compile_protos(&[&path_str], &["."]) {
            Err(e) => {
                error!(
                    "协议文件编译错误 ,File: [{}],Sheet: [{}] ERR: {}",
                    &self.input_file_name, &self.sheet_name, e
                );
                return Err(1);
            }

            _ => (),
        };
        // new  messagedescriptor
        let bytes = fs::read(&bin_path).unwrap();
        let pool = DescriptorPool::decode(bytes.as_ref()).unwrap();
        let n1 = format!("pbd.Data{}", &class_name);
        let n2 = format!("pbd.{}", &class_name);
        let info_des = pool.get_message_by_name(&n1).unwrap();
        let mut info_dm = DynamicMessage::new(info_des);
        let msg_des = pool.get_message_by_name(&n2).unwrap();
        let mut data: IndexMap<MapKey, PValue> = IndexMap::new();
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
                        return Err(1);
                    }
                } else {
                    let p_val =
                        cell_to_pvalue(fv, &ft, &self.input_file_name, &self.sheet_name, fk);
                    // println!("ft: {},fk: {},fv: {:?},p_val: {:?}", ft, fk, fv, p_val);
                    dm.set_field_by_name(fk, p_val);
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
                    return Err(1);
                }
            };
            row = row + 1;
        }
        info_dm.set_field_by_name("data", PValue::Map(data));
        let mut buf = vec![];
        info_dm.encode(&mut buf).unwrap();
        let out_pbd_path = format!("{}/data_{}.bytes", out_path, msg_name.to_lowercase());
        fs::write(out_pbd_path, buf).unwrap();
        return Ok(0);
    }

    pub fn data_to_lang(&self, _out_path: &String) -> Result<usize, usize> {
        for i in 1..self.types.len() {
            if &self.types[i] == "STRING_LOC" {
                let lang_field_name = format!("{}_local", self.names[i].trim());
                let lang_field_index = self
                    .names
                    .iter()
                    .position(|r| r == &lang_field_name)
                    .expect(
                        format!(
                            "没有找到字段名: `{lang_field_name}`, 需要设置！File: [{}] Sheet: [{}]",
                            self.input_file_name, self.sheet_name
                        )
                        .as_str(),
                    );
                for rows in &self.values {
                    let lang_key = rows[i].to_string();
                    let lang_val = rows[lang_field_index].to_string();

                    GLOBAL_LANG.write().insert(lang_key, lang_val);
                }
            }
        }
        return Ok(0);
    }
}

pub fn xls_to_file(
    input_file_name: String,
    dst_path: String,
    format: String,
    multi_sheets: bool,
    export_columns: String,
) -> usize {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.clone()).unwrap();
    let mut sheets = excel.sheet_names().to_owned();

    if !multi_sheets {
        sheets = vec![excel.sheet_names()[0].to_string()];
    }
    for sheet in sheets {
        let sheet_path = format!("{}.{}", &input_file_name, &sheet);
        if GLOBAL_EXCLUDE_SHEETS.read().contains(&sheet_path) {
            continue;
        }
        info!("LOADING [{}] [{}] ...", input_file_name, sheet);
        if let Some(Ok(r)) = excel.worksheet_range(&sheet) {
            let result = sheet_to_data(
                input_file_name.clone(),
                &sheet,
                &r,
                export_columns.to_string(),
            );

            match result {
                Ok(data) => match data.export(&format, &dst_path) {
                    Ok(code) => return code,
                    Err(err_code) => return err_code,
                },
                Err(err_code) => return err_code,
            }
        }
    }
    return 0;
}

pub fn build_id(input_file_name: String, multi_sheets: bool, export_columns: String) -> usize {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.clone()).unwrap();
    let mut sheets = excel.sheet_names().to_owned();

    if !multi_sheets {
        sheets = vec![excel.sheet_names()[0].to_string()];
    }
    for sheet in sheets {
        let sheet_path = format!("{}.{}", &input_file_name, &sheet);

        if let Some(Ok(r)) = excel.worksheet_range(&sheet) {
            let mut row_num = 0;
            let mut names: AHashSet<String> = AHashSet::new();
            let mut front_primary = String::new();
            let mut back_primary = String::new();
            let mut mod_name: String = String::new();
            for row in r.rows() {
                row_num = row_num + 1;
                let mut st = row[0].to_string().trim().to_string();
                st.make_ascii_uppercase();
                if st == "MOD" {
                    mod_name = row[1].to_string().trim().to_string();
                    if mod_name.is_empty() {
                        GLOBAL_EXCLUDE_SHEETS.write().insert(sheet_path);
                        break;
                    }
                    if GLOBAL_MODS.read().contains(&mod_name) {
                        error!(
                            "配置了重复的MOD!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} \n",
                            &input_file_name, &sheet, &mod_name, row_num
                        );
                        return 1;
                    } else {
                        GLOBAL_MODS.write().insert(mod_name.to_string());
                    }
                } else if st == "VALUE" {
                    let key = format!("{}:{}", mod_name, row[1].to_string().trim().to_string());
                    if GLOBAL_IDS.read().contains(&key) {
                        error!(
                            "配置了重复的键值!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Key: {} \n",
                            &input_file_name, &sheet,&mod_name, row_num, &key
                        );
                        return 1;
                    } else {
                        GLOBAL_IDS.write().insert(key);
                    }
                } else if st == "BACK_TYPE" {
                    back_primary = row[1].clone().to_string().trim().to_uppercase();
                } else if st == "FRONT_TYPE" {
                    front_primary = row[1].clone().to_string().trim().to_uppercase();
                } else if st == "NAMES" {
                    // if export_columns == "BACK" && back_primary.is_empty() {
                    //     info!(
                    //         "SKIPPING for {}_TYPE is none！File: [{}] Sheet: [{}],export_columns: [{}]\n",
                    //         &export_columns, &input_file_name, &sheet, &export_columns
                    //     );
                    //     GLOBAL_EXCLUDE_SHEETS.write().insert(sheet_path.to_string());
                    // } else if export_columns == "FRONT" && front_primary.is_empty() {
                    //     info!(
                    //         "SKIPPING for {}_TYPE is none！File: [{}] Sheet: [{}],export_columns: [{}]\n",
                    //         &export_columns, &input_file_name, &sheet, &export_columns
                    //     );
                    //     GLOBAL_EXCLUDE_SHEETS.write().insert(sheet_path.to_string());
                    // }

                    if front_primary.is_empty() && back_primary.is_empty() {
                        info!(
                            "SKIPPING for both `FRONT_TYPE` and `BACK_TYPE`  is none！ File: [{}] Sheet: [{}],export_columns: [{}]\n",
                            &&input_file_name, &sheet,&export_columns
                        );
                        GLOBAL_EXCLUDE_SHEETS.write().insert(sheet_path.to_string());
                    }
                    if !back_primary.is_empty() {
                        GLOBAL_BACK_PRIMARYS
                            .write()
                            .insert(mod_name.to_string(), back_primary.to_string());
                    }
                    if !front_primary.is_empty() {
                        GLOBAL_FRONT_PRIMARYS
                            .write()
                            .insert(mod_name.to_string(), front_primary.to_string());
                    }

                    let primary_key = row[1].clone().to_string().trim().to_uppercase();
                    if primary_key.is_empty() {
                        error!(
                        "NAMES的第二列固定为主键列，不能为空!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Column: {}\n",
                        &input_file_name, &sheet,&mod_name, row_num,1);
                        return 1;
                    }

                    for i in 1..row.len() {
                        let rv = row[i].clone().to_string().trim().to_uppercase();
                        if !rv.is_empty() {
                            let first_char = &rv.chars().next().unwrap();
                            if first_char.is_ascii_digit() {
                                error!(
                                    "NAMES 不能存在以数字开头的字段【{}】!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Column: {}\n",&rv,
                                    &input_file_name, &sheet,&mod_name, row_num, i
                                );
                                return 1;
                            }
                            if names.contains(&rv) {
                                error!(
                                    "NAMES 配置了重复的字段【{}】!! File: [{}] Sheet: [{}],Mod_name: [{}] Row: {} Column: {}\n",&rv,
                                    &input_file_name, &sheet,&mod_name, row_num, i
                                );
                                return 1;
                            }
                            names.insert(rv);
                        }
                    }
                }
            }
        }
    }
    return 0;
}

pub fn sheet_to_data<'a>(
    input_file_name: String,
    sheet_name: &'a String,
    sheet: &'a Range<DataType>,
    export_columns: String,
) -> Result<SheetData<'a>, usize> {
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
    let mut force_mods: Vec<String> = vec![];
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
                describes.push(v.to_string().trim().replace('\n', ""));
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
        } else if st == "FORCE_MOD" {
            for v in row {
                let mod_name = v.to_string().trim().to_string();
                force_mods.push(mod_name);
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
                                return Err(1);
                            }
                        }
                    } else {
                        let key = format!("{}:{}", refs[i], &value);
                        if !GLOBAL_IDS.read().contains(&key) {
                            error!(
                                "没找到引用的键值!File: {},Sheet: {},Row: {}, Key: {}",
                                input_file_name, sheet_name, row_num, key
                            );
                            return Err(1);
                        }
                    }
                }
                if enums.len() > 0 && enums[i].len() > 0 && i > 0 {
                    if !enums[i].contains_key(&value) {
                        error!(
                            "{} 不在枚举类型 {:?} 中! File: [{}],Sheet: [{}], Row: [{}]",
                            value, &enums[i], input_file_name, sheet_name, row_num
                        );
                        return Err(1);
                    }
                }
                row_value.push(&row[i]);
            }
            values.push(row_value);
        }
    }

    let mut types: Vec<String> = front_types;

    if export_columns == "BOTH" {
        let mut temp_types: Vec<String> = vec!["BOTH".to_string()];
        for n in 1..types.len() {
            let front_type = types[n].trim();
            if front_type.is_empty() {
                temp_types.push(back_types[n].trim().to_string());
            } else {
                temp_types.push(front_type.to_string());
            }
        }
        types = temp_types;
    } else if export_columns == "BACK" {
        types = back_types;
    }

    if enums.is_empty() {
        enums = vec![AHashMap::new(); types.len()];
    }

    if refs.is_empty() {
        refs = vec![String::new(); types.len()];
    }

    let info: SheetData = SheetData {
        input_file_name: input_file_name,
        output_file_name: output_file_name,
        sheet_name: sheet_name,
        mod_name: mod_name,
        names: names,
        types: types,
        values: values,
        refs: refs,
        describes: describes,
        enums: enums,
        force_mods: force_mods,
    };
    return Ok(info);
}

pub fn create_pbd_file(out_path: &String) -> usize {
    let pbds = GLOBAL_PBD.read();
    let keys = pbds.keys();
    let mut sorted_keys: Vec<String> = vec![];
    let mut contents: Vec<String> = vec![];
    for key in keys {
        sorted_keys.push(key.to_string());
    }
    sorted_keys.sort();
    for k in sorted_keys {
        contents.push(pbds.get(&k).unwrap().to_string());
    }
    let content = format!(
        "syntax = \"proto3\";\n\
        package pbd;\n\
        \n\
        {}",
        contents.join("\n\n")
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
    return 0;
}

pub fn create_lang_file(out_path: &String) -> usize {
    let data = GLOBAL_LANG.read();
    let mut sorted_keys: Vec<String> = vec![];
    let mut hash_ids = vec![];
    for key in data.keys() {
        sorted_keys.push(key.to_string());
    }
    sorted_keys.sort();

    let wbk_name = "D多语言Key表.xlsx";
    let mut wbk = Workbook::create(format!("{out_path}/{wbk_name}").as_str());
    let mut sheetk = wbk.create_sheet("sheet1");
    sheetk.add_column(Column { width: 30.0 });
    sheetk.add_column(Column { width: 30.0 });
    sheetk.add_column(Column { width: 80.0 });

    let write_result = wbk.write_sheet(&mut sheetk, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["MOD", "Data.LanguageKey", ""])?;
        sw.append_row(row!["BACK_TYPE", "", ""])?;
        sw.append_row(row!["FRONT_TYPE", "uint32", "string"])?;
        sw.append_row(row!["DES", "文本Key的Hash截取", "文本Key D_XXX"])?;
        sw.append_row(row!["NAMES", "hash", "value"])?;
        sw.append_row(row!["ENUM", blank!(2)])?;
        sw.append_row(row!["REF", blank!(2)])?;
        sw.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            let hash = to_hash_id(key);
            if hash_ids.contains(&hash) {
                error!("key: {key} 的hash值重复 hash: {hash}");
                return Err(Error::new(ErrorKind::Other, "key is repeat"));
            } else {
                hash_ids.push(hash);
            }
            sw.append_row(row!["VALUE", hash.to_string(), key.to_string()])?;
        }
        sw.append_row(row![blank!(3)])
    });
    wbk.close()
        .expect(format!("close {wbk_name} error!").as_str());

    match write_result {
        Ok(_) => (),
        Err(_) => return 1,
    };
    let wbl_name = "D多语言简体中文表.xlsx";
    let mut wbl = Workbook::create(format!("{out_path}/{wbl_name}").as_str());
    let mut sheetl = wbl.create_sheet("sheet1");
    sheetl.add_column(Column { width: 30.0 });
    sheetl.add_column(Column { width: 30.0 });
    sheetl.add_column(Column { width: 80.0 });

    let write_result = wbl.write_sheet(&mut sheetl, |sheet_writer| {
        let sw = sheet_writer;
        sw.append_row(row!["MOD", "Data.LanguagezhCN", ""])?;
        sw.append_row(row!["BACK_TYPE", "", ""])?;
        sw.append_row(row!["FRONT_TYPE", "uint32", "string"])?;
        sw.append_row(row!["DES", "文本Key的Hash截取", "中文"])?;
        sw.append_row(row!["NAMES", "hash", "value"])?;
        sw.append_row(row!["ENUM", blank!(2)])?;
        sw.append_row(row!["REF", blank!(2)])?;
        sw.append_row(row!["FORCE_MOD", "Language", ""])?;
        sw.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            let hash = to_hash_id(key);
            let val = data.get(key).unwrap();
            sw.append_row(row!["VALUE", hash.to_string(), val.to_string()])?;
        }
        sw.append_row(row![blank!(3)])
    });
    wbl.close()
        .expect(format!("close {wbl_name} error!").as_str());
    match write_result {
        Ok(_) => (),
        Err(_) => return 1,
    };
    return 0;
}

fn cell_to_json(
    cell: &DataType,
    row_type: &String,
    filename: &String,
    key: &String,
    sheetname: &String,
) -> Value {
    let s = cell.to_string().trim().to_string();
    if row_type.starts_with("INT") || row_type.starts_with("UINT") {
        if s == "" {
            return json!(0);
        }
        json!(s
            .parse::<i64>()
            .ok()
            .expect(parse_err(filename, row_type, &s, sheetname).as_str()))
    } else if row_type.starts_with("FLOAT") {
        if s == "" {
            return json!(0);
        }
        json!(s
            .parse::<f64>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str()))
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
    } else if row_type == "STRING_LOC" {
        json!(to_hash_id(&s))
    } else if row_type == "BOOL" {
        if s == "是" || s.to_uppercase() == "TRUE" {
            serde_json::Value::Bool(true)
        } else {
            serde_json::Value::Bool(false)
        }
    } else {
        json!(s)
    }
}

fn cell_to_string(cell: &DataType, row_type: &String, _filename: &String, _key: &String) -> String {
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
    } else if row_type == "STRING" {
        if s == "" {
            return "\"\"".to_string();
        }
        return format!("\"{}\"", s);
    } else if row_type == "STRING_LOC" {
        if s == "" {
            return "".to_string();
        }
        return format!("{}", to_hash_id(&s));
    } else if row_type == "BOOL" {
        if s == "是" || s.to_uppercase() == "TRUE" {
            "true".to_string()
        } else {
            "false".to_string()
        }
    } else {
        s
    }
}

fn cell_to_pvalue(
    cell: &DataType,
    row_type: &String,
    filename: &String,
    sheetname: &String,
    key: &String,
) -> PValue {
    let mut s = cell.to_string().trim().to_string();
    if row_type.contains("INT") && s.is_empty() {
        s = "0".to_string();
    }
    if row_type.contains("FLOAT") && s.is_empty() {
        s = "0.0".to_string();
    }

    if row_type == "UINT32" {
        let val = s
            .parse::<u32>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str());
        return PValue::U32(val);
    } else if row_type == "UINT64" {
        let val = s
            .parse::<u64>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str());
        return PValue::U64(val);
    } else if row_type == "INT32" {
        let val = s
            .parse::<i32>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str());
        return PValue::I32(val);
    } else if row_type == "INT64" {
        let val = s
            .parse::<i64>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str());
        return PValue::I64(val);
    } else if row_type == "FLOAT" {
        let val = s
            .parse::<f32>()
            .ok()
            .expect(parse_err(filename, key, &s, sheetname).as_str());
        return PValue::F32(val);
    } else if row_type == "LIST_UINT32" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            let val = list[i]
                .parse::<u32>()
                .ok()
                .expect(parse_err(filename, key, &s, sheetname).as_str());
            result.push(PValue::U32(val));
        }
        return PValue::List(result);
    } else if row_type == "LIST_UINT64" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            let val = list[i]
                .parse::<u64>()
                .ok()
                .expect(parse_err(filename, key, &s, sheetname).as_str());
            result.push(PValue::U64(val));
        }
        return PValue::List(result);
    } else if row_type == "LIST_INT32" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            let val = list[i]
                .parse::<i32>()
                .ok()
                .expect(parse_err(filename, key, &s, sheetname).as_str());
            result.push(PValue::I32(val));
        }
        return PValue::List(result);
    } else if row_type == "LIST_INT64" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            let val = list[i]
                .parse::<i64>()
                .ok()
                .expect(parse_err(filename, key, &s, sheetname).as_str());
            result.push(PValue::I64(val));
        }
        return PValue::List(result);
    } else if row_type == "LIST_FLOAT" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            let val = list[i]
                .parse::<f32>()
                .ok()
                .expect(parse_err(filename, key, &s, sheetname).as_str());
            result.push(PValue::F32(val));
        }
        return PValue::List(result);
    } else if row_type == "LIST_STRING" {
        let list: Vec<&str> = s.split(',').collect();
        let mut result: Vec<PValue> = vec![];
        for i in 0..list.len() {
            result.push(PValue::String(list[i].to_string()))
        }
        return PValue::List(result);
    } else if row_type == "STRING_LOC" {
        return PValue::U32(to_hash_id(&s));
    } else if row_type == "STRING" {
        return PValue::String(s);
    } else if row_type == "BOOL" {
        if s == "是" || s.to_uppercase() == "TRUE" {
            return PValue::Bool(true);
        } else {
            return PValue::Bool(false);
        }
    } else {
        error!(
            "cell_to_pvalue failed,unsupport front_type [{}]! File: [{}] Key: [{}] Val: [{}]",
            row_type, &filename, &key, &s
        );
        return PValue::String(s);
    }
}

fn parse_err(filename: &String, key: &String, s: &String, sheetname: &String) -> String {
    let error = format!(
        "数据类型转换出错[ {} ] sheetname: [ {} ] key: [ {} ] Content: {} ",
        filename, sheetname, key, s
    );
    return error;
}

// fn get_module_name(fname: String) -> String {
//     let a = fname
//         .replace(".ex", "")
//         .as_str()
//         .to_train_case()
//         .replace("-", ".");
//     return a;
// }
fn get_real_front_type(ref_name: &String, origin_type: &String, is_enum: bool) -> String {
    if is_enum {
        return "ENUM".to_string();
    }
    if !ref_name.trim().is_empty() {
        if let Some(ref_primary_type) = GLOBAL_FRONT_PRIMARYS.read().get(ref_name) {
            if origin_type.contains("LIST") {
                if origin_type == "INT" {
                    "LIST_UINT32".to_string();
                } else {
                    return format!("LIST_{}", ref_primary_type);
                }
            } else {
                return ref_primary_type.to_string();
            }
        }
    }
    if origin_type == "LIST_INT" || origin_type == "LIST" {
        return "LIST_UINT32".to_string();
    } else if origin_type == "INT" {
        return "UINT32".to_string();
    } else {
        return origin_type.to_string();
    }
}

fn get_real_back_type(ref_name: &String, origin_type: &String, is_enum: bool) -> String {
    if is_enum {
        return "ENUM".to_string();
    }
    if !ref_name.trim().is_empty() {
        if let Some(ref_primary_type) = GLOBAL_BACK_PRIMARYS.read().get(ref_name) {
            if origin_type.contains("LIST") {
                if origin_type == "INT" {
                    "LIST_UINT32".to_string();
                } else {
                    return format!("LIST_{}", ref_primary_type);
                }
            } else {
                return ref_primary_type.to_string();
            }
        }
    }
    if origin_type == "LIST_INT" || origin_type == "LIST" {
        return "LIST_UINT32".to_string();
    } else if origin_type == "INT" {
        return "UINT32".to_string();
    } else {
        return origin_type.to_string();
    }
}

fn to_hash_id(key: &String) -> u32 {
    let digest = md5::compute(key);
    return (u128::from_str_radix(&format!("{:x}", digest), 16).unwrap() % 4294967296) as u32;
}

fn convert_enum_type(field_name: &String) -> String {
    return field_name.to_class_case();
}

// input: {"是": 1, "否": 0}
// output: {0: "shi", 1: "fou"}
fn convert_to_pinyin(
    enum_keys: &AHashMap<String, usize>,
) -> Result<Vec<(usize, String, String)>, String> {
    let mut arr: Vec<(usize, String, String)> = vec![];
    for (k, i) in enum_keys {
        let text = k.as_str();
        let mut enum_elems: Vec<String> = vec![];
        let mut is_pure_hanzi = true;
        for option_pinyin in text.to_pinyin() {
            match option_pinyin {
                Some(pinyin) => {
                    // let py = pinyin.plain().to_string();
                    let py = pinyin.plain().replace('ü', "v");
                    enum_elems.push(py);
                }
                None => {
                    is_pure_hanzi = false;
                }
            }
        }
        if is_pure_hanzi == false {
            if !enum_elems.is_empty() {
                return Err(text.to_string());
            }
            enum_elems.push(k.to_string());
        }
        let temp = enum_elems.join("_");
        arr.push((*i, temp, k.to_string()));
    }
    arr.sort();
    return Ok(arr);
}
