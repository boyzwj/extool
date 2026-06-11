use indexmap::IndexMap;
use prost::Message;
use prost_reflect::{DescriptorPool, DynamicMessage, MapKey, Value as PValue};
use serde_json::value::Value;
use serde_json::Map;
use std::fs;
use std::io::{BufWriter, Write};
use std::path::Path;

use super::state::{GLOBAL_LANG, GLOBAL_LANG_HASHES, GLOBAL_LANG_SOURCE, GLOBAL_PBD};
use super::types::{EnumMap, ExcelResult, SheetData};
use super::value::{
    cell_to_json, cell_to_lua_string, cell_to_pvalue, cell_to_string, get_real_type, is_enum_none,
    to_enum_index, to_hash_id,
};

impl SheetData {
    pub(crate) fn export(
        &self,
        format: &str,
        dst_path: &str,
        enum_values: &EnumMap,
    ) -> ExcelResult<()> {
        match format.to_uppercase().as_str() {
            "JSON" => self.data_to_json(dst_path, enum_values),
            "LUA" => self.data_to_lua(dst_path, enum_values),
            "EX" => self.data_to_ex(dst_path, enum_values),
            "PBD" => self.data_to_pbd(dst_path, enum_values),
            "LANG" => self.data_to_lang(),
            _ => Ok(()),
        }
    }

    fn write_file(&self, path_str: &str, content: &str) -> ExcelResult<()> {
        if let Some(parent) = Path::new(path_str).parent() {
            fs::create_dir_all(parent)
                .map_err(|err| format!("创建导出目录失败 [{}]: {}", parent.display(), err))?;
        }
        let file = fs::File::create(path_str)
            .map_err(|err| format!("创建导出文件失败 [{}]: {}", path_str, err))?;
        let mut writer = BufWriter::new(file);
        writer
            .write_all(content.as_bytes())
            .map_err(|err| format!("写入导出内容失败 [{}]: {}", path_str, err))?;
        info!(
            "成功 [{}] [{}] 导出 [{}] 条记录",
            self.input_file_name,
            self.sheet_name,
            self.values.len()
        );
        Ok(())
    }

    fn data_to_json(&self, out_path: &str, enum_values: &EnumMap) -> ExcelResult<()> {
        let file_name = &self.output_file_name;
        let mut result: Map<String, Value> = Map::new();
        for row in &self.values {
            let mut map: Map<String, Value> = Map::new();
            for index in 1..self.types.len() {
                if self.types[index].is_empty() {
                    continue;
                }
                let column_name = &self.names[index];
                let origin_type = &self.types[index];
                let ref_name = &self.refs[index];
                let enum_name = &self.enum_names[index];
                if enum_name.is_empty() {
                    let real_type = get_real_type(
                        self.export_columns.as_str(),
                        ref_name.as_str(),
                        origin_type.as_str(),
                        false,
                    );
                    let value = cell_to_json(
                        row_value(row, index),
                        real_type.as_str(),
                        file_name,
                        column_name,
                        &self.sheet_name,
                    )?;
                    map.insert(column_name.to_string(), value);
                } else {
                    let value = row_value(row, index).trim();
                    if is_enum_none(value) {
                        map.insert(column_name.to_string(), json!(0));
                    } else {
                        let enum_index = to_enum_index(enum_values, enum_name, value)?;
                        map.insert(column_name.to_string(), json!(enum_index));
                    }
                }
            }

            if !map.is_empty() {
                let key_value = self.export_key(row, enum_values)?;
                result.insert(key_value, json!(map));
            }
        }

        if self.values.is_empty() || file_name.is_empty() {
            return Ok(());
        }
        let obj = json!(result);
        let json = serde_json::to_string_pretty(&obj)
            .map_err(|err| format!("JSON序列化失败 [{}]: {}", file_name, err))?;
        self.write_file(format!("{}/{}.json", out_path, file_name).as_str(), &json)
    }

    fn data_to_lua(&self, out_path: &str, enum_values: &EnumMap) -> ExcelResult<()> {
        let file_name = &self.output_file_name;
        let mut rows = Vec::new();
        let mut keys = Vec::new();
        let mut key_index = 1;
        for index in 1..self.types.len() {
            if !self.types[index].is_empty() {
                keys.push(format!("{} = {}", self.names[index], key_index));
                key_index += 1;
            }
        }

        for row in &self.values {
            let mut columns = Vec::new();
            for index in 1..self.types.len() {
                if self.types[index].is_empty() {
                    continue;
                }
                let column_name = &self.names[index];
                let origin_type = &self.types[index];
                let enum_name = &self.enum_names[index];
                if enum_name.is_empty() {
                    let real_type = get_real_type(
                        self.export_columns.as_str(),
                        self.refs[index].as_str(),
                        origin_type.as_str(),
                        false,
                    );
                    columns.push(cell_to_lua_string(
                        row_value(row, index),
                        real_type.as_str(),
                        file_name,
                        column_name,
                    )?);
                } else {
                    let value = row_value(row, index).trim();
                    if is_enum_none(value) {
                        columns.push("0".to_string());
                    } else {
                        columns.push(to_enum_index(enum_values, enum_name, value)?.to_string());
                    }
                }
            }

            if !columns.is_empty() {
                let key_value = self.export_lua_key(row, enum_values)?;
                rows.push(format!("\t[{}] = {{{}}}", key_value, columns.join(",")));
            }
        }

        if self.values.is_empty() || file_name.is_empty() {
            return Ok(());
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
            keys.join(","),
            rows.join(",\n"),
        );
        self.write_file(format!("{}/{}.lua", out_path, file_name).as_str(), &out)
    }

    fn data_to_ex(&self, dst_path: &str, enum_values: &EnumMap) -> ExcelResult<()> {
        let file_name = &self.output_file_name;
        let mut rows = Vec::new();
        let mut ids = Vec::new();
        for row in &self.values {
            let mut columns = Vec::new();
            for index in 1..self.types.len() {
                if self.types[index].is_empty() {
                    continue;
                }
                let column_name = &self.names[index];
                let origin_type = &self.types[index];
                let enum_name = &self.enum_names[index];
                if enum_name.is_empty() {
                    let real_type = get_real_type(
                        self.export_columns.as_str(),
                        self.refs[index].as_str(),
                        origin_type.as_str(),
                        false,
                    );
                    let value = cell_to_string(
                        row_value(row, index),
                        real_type.as_str(),
                        file_name,
                        column_name,
                    )?;
                    columns.push(format!("\t\t\t{}: {}", column_name.to_lowercase(), value));
                } else {
                    let value = row_value(row, index).trim();
                    if is_enum_none(value) {
                        columns.push(format!("\t\t\t{}: {}", column_name.to_lowercase(), 0));
                    } else {
                        let enum_index = to_enum_index(enum_values, enum_name, value)?;
                        columns.push(format!(
                            "\t\t\t{}: {}",
                            column_name.to_lowercase(),
                            enum_index
                        ));
                    }
                }
            }

            if !columns.is_empty() {
                let key_value = self.export_ex_key(row, enum_values)?;
                rows.push(format!(
                    "  def get({}) do
    %{{
{}
    }}
  end",
                    key_value,
                    columns.join(",\n")
                ));
                ids.push(key_value);
            }
        }

        if self.values.is_empty() || file_name.is_empty() {
            return Ok(());
        }
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
            self.mod_name,
            self.input_file_name,
            self.sheet_name,
            ids.join(", "),
            rows.join("\n\n")
        );
        self.write_file(format!("{}/{}.ex", dst_path, file_name).as_str(), &out)
    }

    fn data_to_pbd(&self, out_path: &str, enum_values: &EnumMap) -> ExcelResult<()> {
        if self.mod_name.is_empty() || self.valid_columns.is_empty() {
            return Ok(());
        }
        let key_name = &self.names[self.valid_columns[0]];
        if self.valid_front_types[0].contains("LIST") || self.valid_front_types[0] == "FLOAT" {
            return Err(format!(
                "主键仅支持整型类型和字符串类型 File: [{}] Sheet: [{}] Mod_name: [{}] Key: [{}] Type: [{}]",
                self.input_file_name, self.sheet_name, self.mod_name, key_name, self.valid_front_types[0]
            ));
        }

        let mut map_key_type = self.valid_front_types[0].to_string();
        if map_key_type == "ENUM" || map_key_type == "STRING_LOC" {
            map_key_type = "UINT32".to_string();
        }

        let msg_name = self.mod_name.replace("Data.", "");
        let mut class_name = msg_name.to_string();
        if self.force_mods.len() > 1 && !self.force_mods[1].is_empty() {
            class_name = self.force_mods[1].to_string();
        }

        let msg_schema = self.pbd_msg_schema(&class_name);
        let out = format!(
            "message Data{}{{\n\
             \tmap<{},{}> data = 1;\n\
            }}\n\
            {}",
            class_name,
            map_key_type.to_lowercase(),
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
        let tmp_dir = format!("{}/.extool_tmp", out_path);
        fs::create_dir_all(&tmp_dir)
            .map_err(|err| format!("创建PBD临时目录失败 [{}]: {}", tmp_dir, err))?;
        let proto_path = format!("{}/tmp_{}.proto", tmp_dir, msg_name.to_lowercase());
        self.write_file(&proto_path, &content)?;
        GLOBAL_PBD.write().insert(class_name.to_string(), out);

        let mut builder = prost_reflect_build::Builder::new();
        let bin_path = format!("{}/tmp_{}.bin", tmp_dir, msg_name.to_lowercase());
        builder.file_descriptor_set_path(&bin_path);
        std::env::set_var("OUT_DIR", &tmp_dir);
        builder
            .compile_protos(&[proto_path.as_str()], &[tmp_dir.as_str()])
            .map_err(|err| {
                format!(
                    "协议文件编译错误 File: [{}] Sheet: [{}] Err: {}",
                    self.input_file_name, self.sheet_name, err
                )
            })?;

        let bytes = fs::read(&bin_path)
            .map_err(|err| format!("读取PBD描述文件失败 [{}]: {}", bin_path, err))?;
        let pool = DescriptorPool::decode(bytes.as_ref())
            .map_err(|err| format!("PBD描述池解析失败 [{}]: {}", bin_path, err))?;
        let data_message_name = format!("pbd.Data{}", class_name);
        let row_message_name = format!("pbd.{}", class_name);
        let info_des = pool
            .get_message_by_name(&data_message_name)
            .ok_or_else(|| format!("找不到PBD消息 [{}]", data_message_name))?;
        let mut info_dm = DynamicMessage::new(info_des);
        let msg_des = pool
            .get_message_by_name(&row_message_name)
            .ok_or_else(|| format!("找不到PBD消息 [{}]", row_message_name))?;
        let mut data: IndexMap<MapKey, PValue> = IndexMap::new();

        for (row_index, row) in self.values.iter().enumerate() {
            let mut dm = DynamicMessage::new(msg_des.clone());
            for (type_index, column_index) in self.valid_columns.iter().enumerate() {
                let field_type = &self.valid_front_types[type_index];
                let field_name = &self.names[*column_index];
                let field_value = row_value(row, *column_index);
                let enum_name = &self.enum_names[*column_index];
                if field_type == "ENUM" {
                    if is_enum_none(field_value) || field_value == "0" {
                        dm.set_field_by_name(field_name, PValue::U32(0));
                    } else {
                        let enum_index = to_enum_index(enum_values, enum_name, field_value)?;
                        dm.set_field_by_name(field_name, PValue::U32(enum_index as u32));
                    }
                } else {
                    let p_val = cell_to_pvalue(
                        field_value,
                        field_type,
                        &self.input_file_name,
                        &self.sheet_name,
                        field_name,
                    )?;
                    dm.set_field_by_name(field_name, p_val);
                }
            }
            let key_val = dm
                .get_field_by_name_mut(key_name)
                .ok_or_else(|| {
                    format!(
                        "PBD主键字段不存在 File: [{}] Sheet: [{}] Row: {} Key: [{}]",
                        self.input_file_name,
                        self.sheet_name,
                        row_index + 1,
                        key_name
                    )
                })?
                .clone();
            let dy_msg = PValue::Message(dm);
            match key_val {
                PValue::I32(value) => data.insert(MapKey::I32(value), dy_msg),
                PValue::I64(value) => data.insert(MapKey::I64(value), dy_msg),
                PValue::U32(value) => data.insert(MapKey::U32(value), dy_msg),
                PValue::U64(value) => data.insert(MapKey::U64(value), dy_msg),
                PValue::String(value) => data.insert(MapKey::String(value), dy_msg),
                _ => {
                    return Err(format!(
                    "键值的数据类型不对 File: [{}] Sheet: [{}] Mod_name: [{}] Row: {} Key: [{}]",
                    self.input_file_name,
                    self.sheet_name,
                    self.mod_name,
                    row_index + 1,
                    key_name
                ))
                }
            };
        }
        info_dm.set_field_by_name("data", PValue::Map(data));
        let mut buf = Vec::new();
        info_dm
            .encode(&mut buf)
            .map_err(|err| format!("PBD编码失败 [{}]: {}", self.mod_name, err))?;
        let out_pbd_path = format!("{}/data_{}.bytes", out_path, msg_name.to_lowercase());
        fs::write(&out_pbd_path, buf)
            .map_err(|err| format!("写入PBD数据失败 [{}]: {}", out_pbd_path, err))?;
        Ok(())
    }

    fn data_to_lang(&self) -> ExcelResult<()> {
        for index in 1..self.types.len() {
            if self.types[index] == "STRING_LOC" {
                let lang_field_name = format!("{}_local", self.names[index].trim());
                let lang_field_index = self
                    .names
                    .iter()
                    .position(|name| name == &lang_field_name)
                    .ok_or_else(|| {
                        format!(
                            "没有找到字段名 [{}]，STRING_LOC字段需要配置本地化文本列 File: [{}] Sheet: [{}]",
                            lang_field_name, self.input_file_name, self.sheet_name
                        )
                    })?;
                for row in &self.values {
                    let lang_key = row_value(row, index).to_string();
                    let lang_val = row_value(row, lang_field_index).to_string();
                    if lang_key.is_empty() {
                        continue;
                    }
                    let hash = to_hash_id(&lang_key);
                    if let Some(existing_key) = GLOBAL_LANG_HASHES.read().get(&hash) {
                        if existing_key != &lang_key {
                            return Err(format!(
                                "STRING_LOC hash碰撞 Hash: [{}] KeyA: [{}] KeyB: [{}]",
                                hash, existing_key, lang_key
                            ));
                        }
                    }
                    GLOBAL_LANG_HASHES
                        .write()
                        .insert(hash, lang_key.to_string());
                    GLOBAL_LANG.write().insert(lang_key.to_string(), lang_val);
                    GLOBAL_LANG_SOURCE
                        .write()
                        .insert(lang_key.to_string(), self.input_file_name.to_string());
                }
            }
        }
        Ok(())
    }

    fn pbd_msg_schema(&self, class_name: &str) -> String {
        let mut field_schemas = Vec::new();
        for (type_index, column_index) in self.valid_columns.iter().enumerate() {
            let field_type = &self.valid_front_types[type_index];
            let field_name = &self.names[*column_index];
            let describe = self
                .describes
                .get(*column_index)
                .map(String::as_str)
                .unwrap_or("");
            let field_schema = match field_type.as_str() {
                "LIST_UINT32" => "repeated uint32".to_string(),
                "LIST_UINT64" => "repeated uint64".to_string(),
                "LIST_INT32" => "repeated int32".to_string(),
                "LIST_INT64" => "repeated int64".to_string(),
                "LIST_FLOAT" => "repeated float".to_string(),
                "LIST_STRING" => "repeated string".to_string(),
                "ENUM" => "uint32".to_string(),
                "STRING_LOC" => "uint32".to_string(),
                _ => field_type.to_lowercase(),
            };
            field_schemas.push(format!(
                "\t{} {} = {}; //{}",
                field_schema,
                field_name,
                type_index + 1,
                describe
            ));
        }

        format!(
            "message {}{{\n\
            {}\n\
            }}",
            class_name,
            field_schemas.join("\n"),
        )
    }

    fn export_key(&self, row: &[String], enum_values: &EnumMap) -> ExcelResult<String> {
        let enum_name = &self.enum_names[1];
        let value = row_value(row, 1).trim();
        if enum_name.is_empty() {
            Ok(value.to_string())
        } else {
            to_enum_index(enum_values, enum_name, value).map(|value| value.to_string())
        }
    }

    fn export_lua_key(&self, row: &[String], enum_values: &EnumMap) -> ExcelResult<String> {
        let enum_name = &self.enum_names[1];
        if enum_name.is_empty() {
            cell_to_lua_string(
                row_value(row, 1),
                self.types[1].as_str(),
                &self.output_file_name,
                &self.names[1],
            )
        } else {
            to_enum_index(enum_values, enum_name, row_value(row, 1)).map(|value| value.to_string())
        }
    }

    fn export_ex_key(&self, row: &[String], enum_values: &EnumMap) -> ExcelResult<String> {
        let enum_name = &self.enum_names[1];
        if enum_name.is_empty() {
            cell_to_string(
                row_value(row, 1),
                self.types[1].as_str(),
                &self.output_file_name,
                &self.names[1],
            )
        } else {
            to_enum_index(enum_values, enum_name, row_value(row, 1)).map(|value| value.to_string())
        }
    }
}

fn row_value(row: &[String], index: usize) -> &str {
    row.get(index).map(String::as_str).unwrap_or("")
}
