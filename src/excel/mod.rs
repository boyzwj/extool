use ahash::{AHashMap, AHashSet};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use indexmap::IndexMap;
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
use std::path::Path;
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
static mut GLOBAL_LANG_SOURCE: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_EXCLUDE_SHEETS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_MODS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_GROUP_NAMES: AHashMap<String, String> = AHashMap::new();

#[dynamic]
static mut GLOBAL_GROUP_IDS: AHashSet<String> = AHashSet::new();

#[dynamic]
static mut GLOBAL_MOD_IDS: AHashMap<String, Vec<String>> = AHashMap::new();

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
    enum_names: Vec<String>,
    values: Vec<Vec<&'a DataType>>,
    force_mods: Vec<String>,
    export_columns: String,
    valid_columns: Vec<usize>,
    valid_front_types: Vec<String>,
}

impl<'a> SheetData<'_> {
    pub fn export(
        &self,
        format: &String,
        dst_path: &String,
        enum_values: &HashMap<String, Vec<EnumValue>>,
    ) -> Result<usize, usize> {
        let temp = format.to_uppercase();
        if temp == "JSON" {
            return self.data_to_json(&dst_path, enum_values);
        } else if temp == "LUA" {
            return self.data_to_lua(&dst_path, enum_values);
        } else if temp == "EX" {
            return self.data_to_ex(&dst_path, enum_values);
        } else if temp == "PBD" {
            return self.data_to_pbd(&dst_path, enum_values);
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

    pub fn data_to_json(
        &self,
        out_path: &String,
        enum_values: &HashMap<String, Vec<EnumValue>>,
    ) -> Result<usize, usize> {
        let file_name = &self.output_file_name;
        let export_columns = &self.export_columns;
        let mut res: Map<String, Value> = Map::new();
        for rv in &self.values {
            let mut map: Map<String, Value> = Map::new();
            for i in 1..self.types.len() {
                if self.types[i] != "" {
                    let column_name = &self.names[i];
                    let origin_type = &self.types[i];
                    let ref_name = &self.refs[i];
                    let enum_name = &self.enum_names[i];
                    if self.enum_names.len() == 0 || self.enum_names[i].len() == 0 {
                        let real_type = get_real_type(export_columns, ref_name, origin_type, false);
                        let value = cell_to_json(
                            &rv[i],
                            &real_type,
                            file_name,
                            &column_name,
                            &self.sheet_name,
                        );
                        map.insert(column_name.to_string(), value);
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if is_enum_none(value) {
                            map.insert(column_name.to_string(), json!(0));
                        } else if let Ok(x) = to_enum_index(enum_values, enum_name, value) {
                            map.insert(column_name.to_string(), json!(x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {} 中",
                                column_name, &rv[1], value, &self.enum_names[i]
                            );
                            return Err(1);
                        }
                    }
                }
            }

            if map.len() > 0 {
                let key_is_enum = self.enum_names[1].is_empty() == false;
                let enum_value = rv[1].to_string().trim().to_string();
                let mut keyvalue = enum_value.to_string();
                if key_is_enum {
                    keyvalue = to_enum_index(enum_values, &self.enum_names[1], &keyvalue)
                        .unwrap()
                        .to_string();
                }

                res.insert(keyvalue, json!(map));
            }
        }

        if self.values.len() == 0 || file_name.len() <= 0 {
            return Ok(0);
        }
        let obj = json!(res);
        let json = serde_json::to_string_pretty(&obj).unwrap();
        let path_str = format!("{}/{}.json", out_path, file_name);
        match self.write_file(&path_str, &json) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_lua(
        &self,
        out_path: &String,
        enum_values: &HashMap<String, Vec<EnumValue>>,
    ) -> Result<usize, usize> {
        let file_name = &self.output_file_name;
        let export_columns = &self.export_columns;
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
                    let enum_name = &self.enum_names[i];
                    if self.enum_names.len() == 0 || self.enum_names[i].len() == 0 {
                        let real_type =
                            get_real_type(export_columns, &self.refs[i], origin_type, false);
                        let value = cell_to_string(&rv[i], &real_type, file_name, &column_name);
                        columns.push(format!("{}", value.replace("[", "{").replace("]", "}")));
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if is_enum_none(value) {
                            columns.push(format!("{}", 0));
                        } else if let Ok(x) = to_enum_index(enum_values, enum_name, value) {
                            columns.push(format!("{}", x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                                column_name, &rv[1], value, &self.enum_names[i]
                            );
                            return Err(1);
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let enum_name = &self.enum_names[1];
                let key_is_enum = enum_name.is_empty() == false;
                let mut keyvalue =
                    cell_to_string(&rv[1], &self.types[1], file_name, &self.names[1]);
                if key_is_enum {
                    let enum_value = &rv[1].to_string().trim().to_string();
                    keyvalue = to_enum_index(enum_values, enum_name, enum_value)
                        .unwrap()
                        .to_string();
                }

                res.push(format!("\t[{}] = {{{}}}", keyvalue, columns.join(",")));
            }
        }

        if self.values.len() == 0 || file_name.len() <= 0 {
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
        let path_str = format!("{}/{}.lua", out_path, file_name);
        match self.write_file(&path_str, &out) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_ex(
        &self,
        dst_path: &String,
        enum_values: &HashMap<String, Vec<EnumValue>>,
    ) -> Result<usize, usize> {
        let file_name = &self.output_file_name;
        let export_columns = &self.export_columns;
        let mut res: Vec<String> = vec![];
        let mut ids: Vec<String> = vec![];
        for rv in &self.values {
            let mut columns: Vec<String> = vec![];
            for i in 1..self.types.len() {
                if self.types[i] != "" {
                    let column_name = &self.names[i];
                    let origin_type = &self.types[i];
                    let enum_name = &self.enum_names[i];
                    if self.enum_names.len() == 0 || self.enum_names[i].len() == 0 {
                        let real_type =
                            get_real_type(export_columns, &self.refs[i], origin_type, false);
                        let value = cell_to_string(&rv[i], &real_type, file_name, &column_name);
                        columns.push(format!("\t\t\t{}: {}", column_name.to_lowercase(), value));
                    } else {
                        let value = &rv[i].to_string().trim().to_string();
                        if is_enum_none(value) {
                            columns.push(format!("\t\t\t{}: {}", column_name.to_lowercase(), 0));
                        } else if let Ok(x) = to_enum_index(enum_values, enum_name, value) {
                            columns.push(format!("\t\t\t{}: {}", column_name.to_lowercase(), x));
                        } else {
                            error!(
                                "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {} 中",
                                column_name, &rv[1], value, &self.enum_names[i]
                            );
                            return Err(1);
                        }
                    }
                }
            }
            if columns.len() > 0 {
                let enum_name = &self.enum_names[1];
                let key_is_enum = enum_name.is_empty() == false;
                let mut keyvalue =
                    cell_to_string(&rv[1], &self.types[1], file_name, &self.names[1]);
                if key_is_enum {
                    let enum_value = &rv[1].to_string().trim().to_string();
                    keyvalue = to_enum_index(enum_values, enum_name, enum_value)
                        .unwrap()
                        .to_string();
                }

                res.push(format!(
                    "  def get({}) do
    %{{
{}
    }}
  end",
                    keyvalue,
                    columns.join(",\n")
                ));
                ids.push(keyvalue.to_string());
            }
        }

        if self.values.len() == 0 || file_name.len() <= 0 {
            return Ok(0);
        }
        let module_name = self.mod_name.clone();
        let out = format!(
            "defmodule {} do
  ## SOURCE:\"{}\" SHEET:\"{}\"

  def ids() do
    [{}]
  end

  def all() do
    ids() |> Enum.map(&get/1)
  end

  def query(q) do
    all() |> Enum.filter(q)
  end

{}
  
  def get(_), do: nil\n\
end",
            module_name,
            self.input_file_name,
            self.sheet_name,
            ids.join(", "),
            res.join("\n\n")
        );
        let path_str = format!("{}/{}.ex", dst_path, file_name);
        match self.write_file(&path_str, &out) {
            0 => (),
            _ => return Err(1),
        }
        return Ok(0);
    }

    pub fn data_to_pbd(
        &self,
        out_path: &String,
        enum_values: &HashMap<String, Vec<EnumValue>>,
    ) -> Result<usize, usize> {
        if self.mod_name == "" {
            return Ok(0);
        }
        let valid_columns = &self.valid_columns;
        if valid_columns.len() == 0 {
            return Ok(0);
        }
        let valid_front_types = &self.valid_front_types;

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

        let msg_name = self.mod_name.to_string().replace("Data.", "");
        let mut class_name = msg_name.to_string();
        if self.force_mods.len() > 0 && !self.force_mods[1].is_empty() {
            class_name = self.force_mods[1].to_string();
        }

        // let msg_enum = format!("{}", "");
        let msg_schema = self.pbd_msg_schema(&class_name);

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
                let enum_name = &self.enum_names[i];
                if ft == "ENUM" {
                    if enu_val == "0" || enu_val.trim().is_empty() {
                        dm.set_field_by_name(fk, PValue::U32(0 as u32));
                    } else if let Ok(x) = to_enum_index(enum_values, enum_name, enu_val) {
                        dm.set_field_by_name(fk, PValue::U32(x as u32));
                    } else {
                        error!(
                            "列 {},ID: {} 存在非法枚举值, \"{}\" 不在 {:?} 中",
                            fk, &self.values[x][1], enu_val, &self.enum_names[i]
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

                    if !lang_key.is_empty() {
                        GLOBAL_LANG.write().insert(lang_key.to_string(), lang_val);
                        GLOBAL_LANG_SOURCE
                            .write()
                            .insert(lang_key.to_string(), self.input_file_name.to_string());
                    }
                }
            }
        }
        return Ok(0);
    }

    fn pbd_msg_schema(&self, class_name: &String) -> String {
        let valid_columns = &self.valid_columns;
        let valid_front_types = &self.valid_front_types;
        let names = &self.names;
        let describes = &self.describes;
        let mut field_schemas: Vec<String> = vec![];
        for y in 0..valid_columns.len() {
            let i = valid_columns[y];
            let ft = valid_front_types[y].to_string();
            let fk = names[i].to_string();
            let des = describes[i].to_string();
            let field_schema = match ft.as_str() {
                "LIST_UINT32" => "repeated uint32".to_string(),
                "LIST_UINT64" => "repeated uint64".to_string(),
                "LIST_INT32" => "repeated int32".to_string(),
                "LIST_INT64" => "repeated int64".to_string(),
                "LIST_FLOAT" => "repeated float".to_string(),
                "LIST_STRING" => "repeated string".to_string(),
                "ENUM" => "uint32".to_string(),
                "STRING_LOC" => "uint32".to_string(),
                _ => ft.to_lowercase(),
            };
            field_schemas.push(format!("\t{} {} = {}; //{}", &field_schema, fk, y + 1, des));
        }

        return format!(
            "message {}{{\n\
            {}\n\
            }}",
            class_name,
            field_schemas.join("\n"),
        );
    }
}

pub fn xls_to_file(
    input_file_name: String,
    dst_path: String,
    format: String,
    multi_sheets: bool,
    export_columns: String,
    pbe_path: String,
) -> usize {
    let mut excel: Xlsx<_> = open_workbook(input_file_name.clone()).unwrap();
    let mut sheets = excel.sheet_names().to_owned();

    let enum_values = extract_all_enum_values(&pbe_path);

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
                &enum_values,
            );

            match result {
                Ok(data) => match data.export(&format, &dst_path, &enum_values) {
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
            let mut group_name: String = String::new();
            let mut ids: Vec<String> = vec![];

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
                } else if st == "BACK_TYPE" {
                    back_primary = row[1].clone().to_string().trim().to_uppercase();
                } else if st == "FRONT_TYPE" {
                    front_primary = row[1].clone().to_string().trim().to_uppercase();
                } else if st == "NAMES" {
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
                } else if st == "GROUP" {
                    group_name = row[1].to_string().trim().to_string();
                } else if st == "VALUE" {
                    let record_id = row[1].to_string().trim().to_string();
                    let key = format!("{}:{}", mod_name, &record_id);
                    if GLOBAL_IDS.read().contains(&key) {
                        error!(
                            "配置了重复的键值!! File: [{}] Sheet: [{}],Mod_Name: [{}] Row: {} Key: {} \n",
                            &input_file_name, &sheet,&mod_name, row_num, &key
                        );
                        return 1;
                    } else {
                        GLOBAL_IDS.write().insert(key);
                        ids.push(record_id.to_string());
                    }

                    if !group_name.is_empty() {
                        let group_id_key = format!("{}:{}", &group_name, &record_id);
                        if GLOBAL_GROUP_IDS.read().contains(&group_id_key) {
                            let conflict_mod_names =
                                find_group_mod_name_that_contain_id(&group_name, &record_id);
                            error!(
                                "配置了重复的键值!! Mod_Name: [{}],ConfictModName: [{}] Row: {} group_id_key: {} \n",
                                &mod_name,
                                conflict_mod_names.join(","),
                                row_num,
                                &group_id_key
                            );
                            return 1;
                        } else {
                            GLOBAL_GROUP_IDS.write().insert(group_id_key);
                        }
                    }
                }
            }

            if !group_name.is_empty() {
                GLOBAL_GROUP_NAMES
                    .write()
                    .insert(mod_name.to_string(), group_name.to_string());
            }
            GLOBAL_MOD_IDS.write().insert(mod_name.to_string(), ids);
        }
    }
    return 0;
}

fn find_group_mod_names(group_name: &String) -> Vec<String> {
    let group_names = GLOBAL_GROUP_NAMES.read();
    let mut contents = vec![];
    for (mod_name, v) in group_names.iter() {
        if v == group_name {
            contents.push(mod_name.to_string());
        }
    }
    return contents;
}

fn find_group_mod_name_that_contain_id(group_name: &String, id: &String) -> Vec<String> {
    let mod_names = find_group_mod_names(group_name);
    let mut contents = vec![];
    for mod_name in mod_names {
        let ids = GLOBAL_MOD_IDS.read().get(&mod_name).unwrap().to_vec();
        if ids.contains(id) {
            contents.push(mod_name.to_string());
        }
    }
    return contents;
}

pub fn sheet_to_data<'a>(
    input_file_name: String,
    sheet_name: &'a String,
    sheet: &'a Range<DataType>,
    export_columns: String,
    enum_values: &HashMap<String, Vec<EnumValue>>,
) -> Result<SheetData<'a>, usize> {
    let mut output_file_name: String = String::new();
    let mut mod_name: String = String::new();
    let mut values: Vec<Vec<&DataType>> = vec![];
    let mut names: Vec<String> = vec![];
    let mut front_types: Vec<String> = vec![];
    let mut back_types: Vec<String> = vec![];
    let mut refs: Vec<String> = vec![];
    let mut describes: Vec<String> = vec![];
    let mut enum_names: Vec<String> = vec![];
    let mut row_num: usize = 0;
    let mut force_mod: String = String::new();
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
                describes.push(v.to_string().trim().replace("\r\n", " ").replace('\n', " "));
            }
        } else if st == "ENUM" {
            for v in row {
                let enum_name = v.to_string().trim().to_string();
                enum_names.push(enum_name);
            }
        } else if st == "FORCE_MOD" {
            for v in row {
                force_mod = v.to_string().trim().to_string();
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
                if enum_names.len() > 0 && !enum_names[i].is_empty() && i > 0 {
                    if is_enum_none(&value) {
                    } else {
                        match to_enum_index(enum_values, &enum_names[i], &value) {
                            Ok(_index) => (),
                            Err(_) => {
                                error!(
                                    "{} 不在枚举集合 {:?} 中! File: [{}],Sheet: [{}], Row: [{}]",
                                    value, &enum_names[i], input_file_name, sheet_name, row_num
                                );
                                return Err(1);
                            }
                        }
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

    if enum_names.is_empty() {
        enum_names = vec![String::new(); types.len()];
    }

    if refs.is_empty() {
        refs = vec![String::new(); types.len()];
    }

    let mut valid_columns: Vec<usize> = vec![];
    let mut valid_front_types: Vec<String> = vec![];
    for i in 1..types.len() {
        let origin_type = &types[i];
        let fk = &names[i];
        if origin_type == "" || fk == "" {
            continue;
        }
        let is_enum = origin_type != "BOOL" && !enum_names[i].is_empty();
        let ref_name = &refs[i];
        let ft = get_real_type(&export_columns, ref_name, &origin_type, is_enum);
        valid_front_types.push(ft);
        valid_columns.push(i);
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
        enum_names: enum_names,
        force_mod: force_mod,
        export_columns: export_columns,
        valid_columns,
        valid_front_types,
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
    let data_source = GLOBAL_LANG_SOURCE.read();
    let mut sorted_keys: Vec<String> = vec![];
    for key in data.keys() {
        sorted_keys.push(key.to_string());
    }
    sorted_keys.sort();

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
        sw.append_row(row!["FRONT_TYPE", "uint32", "string", "string", "string"])?;
        sw.append_row(row![
            "DES",
            "文本Key的Hash截取",
            "多语言key",
            "中文",
            "来源"
        ])?;
        sw.append_row(row!["NAMES", "hash", "key", "value", "source"])?;
        sw.append_row(row!["ENUM", blank!(2)])?;
        sw.append_row(row!["REF", blank!(2)])?;
        sw.append_row(row!["FORCE_MOD", "Language", ""])?;
        sw.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            let hash = to_hash_id(key);
            let val = data.get(key).unwrap();
            let source = data_source.get(key).unwrap();
            let p1 = Path::new(&source);
            let source_filename = p1.file_name().unwrap().to_str().unwrap();
            sw.append_row(row![
                "VALUE",
                hash.to_string(),
                key.to_string(),
                val.to_string(),
                source_filename.to_string()
            ])?;
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

pub fn create_group_files(out_path: &String) -> usize {
    let kv: HashMap<String, Vec<String>> = get_all_grouped_mods();
    // GLOBAL_GROUP_NAMES
    for (group_name, mod_names) in kv {
        let result = create_group_file(&group_name, &get_ids_set(&mod_names), out_path);
        if result > 0 {
            return result;
        }
    }

    return 0;
}

fn get_all_grouped_mods() -> HashMap<String, Vec<String>> {
    let group_names = GLOBAL_GROUP_NAMES.read();
    let mut uniq_groups: Vec<String> = vec![];
    for (_mod_name, v) in group_names.iter() {
        if !uniq_groups.contains(v) {
            uniq_groups.push(v.to_string())
        }
    }

    let mut contents: HashMap<String, Vec<String>> = HashMap::new();
    for group_name in uniq_groups {
        contents.insert(group_name.to_string(), find_group_mod_names(&group_name));
    }
    return contents;
}

fn get_ids_set(mod_names: &Vec<String>) -> AHashMap<String, Vec<String>> {
    let mut ret: AHashMap<String, Vec<String>> = AHashMap::new();
    for mod_name in mod_names {
        let ids = GLOBAL_MOD_IDS.read().get(mod_name).unwrap().to_vec();
        ret.insert(mod_name.to_string(), ids);
    }
    return ret;
}

fn create_group_file(
    group_name: &String,
    ids_set: &AHashMap<String, Vec<String>>,
    out_path: &String,
) -> usize {
    let mut sorted_keys: Vec<String> = vec![];
    for mod_name in ids_set.keys() {
        sorted_keys.push(mod_name.to_string());
    }
    sorted_keys.sort();
    let wbl_name = format!("Y映射表-{}.xlsx", group_name);
    let mut wbl = Workbook::create(format!("{out_path}/{wbl_name}").as_str());
    let mut sheetl = wbl.create_sheet("sheet1");
    sheetl.add_column(Column { width: 30.0 });
    sheetl.add_column(Column { width: 30.0 });
    sheetl.add_column(Column { width: 80.0 });

    let write_result = wbl.write_sheet(&mut sheetl, |sheet_writer| {
        let sw: &mut SheetWriter<'_, '_> = sheet_writer;
        let mod_name = format!("Data.Gruop{}", group_name);
        sw.append_row(row!["MOD", mod_name.to_string(), ""])?;
        sw.append_row(row!["BACK_TYPE", "uint64", "string"])?;
        sw.append_row(row!["FRONT_TYPE", "uint64", "string"])?;
        sw.append_row(row!["DES", "id", "id来源"])?;
        sw.append_row(row!["NAMES", "id", "mod"])?;
        sw.append_row(row!["ENUM", blank!(2)])?;
        sw.append_row(row!["REF", blank!(2)])?;
        sw.append_row(row![blank!(3)])?;
        for key in &sorted_keys {
            let ids = ids_set.get(key).unwrap();
            for id in ids {
                sw.append_row(row!["VALUE", id.to_string(), key.to_string()])?;
            }
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
fn get_real_type(
    export_columns: &String,
    ref_name: &String,
    origin_type: &String,
    is_enum: bool,
) -> String {
    if is_enum {
        return "ENUM".to_string();
    }
    if !ref_name.trim().is_empty() {
        return get_ref_primary_type(export_columns, ref_name, origin_type);
    }
    if origin_type == "LIST_INT" || origin_type == "LIST" {
        return "LIST_UINT32".to_string();
    } else if origin_type == "INT" {
        return "UINT32".to_string();
    } else {
        return origin_type.to_string();
    }
}

fn get_ref_primary_type(
    export_columns: &String,
    ref_name: &String,
    origin_type: &String,
) -> String {
    match export_columns.as_str() {
        "BACK" => {
            if let Some(ref_primary_type) = GLOBAL_BACK_PRIMARYS.read().get(ref_name) {
                if origin_type.contains("LIST") {
                    if origin_type == "INT" {
                        return "LIST_UINT32".to_string();
                    } else {
                        return format!("LIST_{}", ref_primary_type);
                    }
                } else {
                    return ref_primary_type.to_string();
                }
            }
            return origin_type.to_string();
        }
        "FRONT" => {
            if let Some(ref_primary_type) = GLOBAL_FRONT_PRIMARYS.read().get(ref_name) {
                if origin_type.contains("LIST") {
                    if origin_type == "INT" {
                        return "LIST_UINT32".to_string();
                    } else {
                        return format!("LIST_{}", ref_primary_type);
                    }
                } else {
                    return ref_primary_type.to_string();
                }
            }
            return origin_type.to_string();
        }
        _ => return origin_type.to_string(),
    };
}

fn to_hash_id(key: &String) -> u32 {
    if !key.is_empty() {
        let digest = md5::compute(key);
        return (u128::from_str_radix(&format!("{:x}", digest), 16).unwrap() % 4294967296) as u32;
    } else {
        return 0;
    }
}

fn is_enum_none(value: &String) -> bool {
    value.trim().is_empty() || value.to_uppercase() == "NONE"
}

pub struct EnumValue {
    index: i32,
    comment: String,
}

use std::collections::HashMap;
pub fn extract_all_enum_values(file_path: &str) -> HashMap<String, Vec<EnumValue>> {
    let file_contents = fs::read_to_string(file_path).unwrap();
    let mut enum_values = HashMap::new();

    let enum_regex = regex::Regex::new(r#"enum\s+(\w+)\s*\{([^}]+)\}"#).unwrap();
    let comment_regex = regex::Regex::new(r#"//(.*)$"#).unwrap();
    let value_regex = regex::Regex::new(r#"(\w+)\s*=\s*(\d+);"#).unwrap();

    for enum_match in enum_regex.captures_iter(&file_contents) {
        let enum_name = enum_match[1].to_owned();
        let enum_body = enum_match[2].to_owned();
        let mut values = Vec::new();

        for line in enum_body.lines() {
            if let Some(comment_match) = comment_regex.captures(line) {
                if let Some(value_match) = value_regex.captures(line) {
                    // let key = value_match[1].to_owned();
                    let index = value_match[2].parse().unwrap();
                    let comment = comment_match[1].trim().to_owned();
                    values.push(EnumValue { index, comment });
                }
            }
        }

        if !values.is_empty() {
            enum_values.insert(enum_name, values);
        }
    }

    enum_values
}

fn to_enum_index(
    enum_values: &HashMap<String, Vec<EnumValue>>,
    enum_name: &String,
    comment: &String,
) -> Result<i32, String> {
    match enum_values.get(enum_name) {
        Some(arr) => {
            for info in arr {
                if &info.comment == comment {
                    return Ok(info.index);
                }
            }
            return Err(format!(
                "EnumName: {},Comment: {} not found",
                enum_name, comment
            ));
        }
        None => return Err(format!("EnumName: {} not found", enum_name)),
    }
}
