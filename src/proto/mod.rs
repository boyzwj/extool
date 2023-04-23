use std::fs::{self, File};
use std::io::{BufRead, BufReader};

use inflector::Inflector;
#[derive(Debug)]

struct Element {
    package: String,
    proto: String,
    id: u128,
}

fn find_element(
    filepath: &str,
    elements: &mut Vec<Element>,
) -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open(filepath)?;
    let reader = BufReader::new(file);

    let mut _find_package = String::new();
    let mut stack: Vec<Element> = vec![];
    for line in reader.lines() {
        let temp = line?.trim().to_string();
        if temp.starts_with("package") {
            if _find_package.is_empty() {
                let v: Vec<&str> = temp.split(&[' ', ';']).collect();
                _find_package = v[1].to_string();
                continue;
            } else {
                panic!("proto file={:?} is invalid!", filepath)
            }
        }

        if temp.starts_with("message") {
            let v: Vec<&str> = temp.split('{').collect();
            let head = v[0].trim();
            if head.ends_with("2C") || head.ends_with("2S") {
                let v2: Vec<&str> = head.split("message").collect();
                let msg_name = v2[1].trim().to_string();
                let temp_proto = format!("{}.{}", &_find_package, msg_name);
                let arrs: Vec<&str> = temp_proto.split(&['.']).collect();
                let mut new_arrs: Vec<String> = vec![];
                for t_str in arrs.into_iter() {
                    new_arrs.push(t_str.to_pascal_case());
                }
                let proto = new_arrs.join(".");
                let digest = md5::compute(&proto);
                let bytes = format!("{:x}", digest);
                let k2 = u128::from_str_radix(&bytes, 16).unwrap();
                let k3: u128 = 65536;
                let msg_id = k2 % k3;
                let elem = Element {
                    package: _find_package.to_string(),
                    proto,
                    id: msg_id,
                };
                stack.push(elem);
            }
        }

        if temp.ends_with("}") {
            let last = stack.pop();
            if stack.is_empty() && last.is_some() {
                match last {
                    Some(value) => elements.push(value),
                    None => continue,
                };
            }
        }
    }

    Ok(())
}

pub fn create(files: Vec<String>, out_path: &String, format: &String, type_input: &String) {
    let mut elements: Vec<Element> = vec![];
    for file in &files {
        find_element(&file, &mut elements).unwrap();
    }
    elements.sort_by(|a, b| a.proto.cmp(&b.proto));
    match format.to_uppercase().as_str() {
        "CS" => data_to_cs(out_path, &elements),
        "LUA" => data_to_lua(out_path, &elements),
        _ => panic!(
            "type input={:?} and format={} is unsupport!",
            type_input, format
        ),
    }
}

fn data_to_cs(out_path: &String, elements: &Vec<Element>) {
    let mut constdef: Vec<String> = vec![];
    let mut dicid: Vec<String> = vec![];
    let mut dicparser: Vec<String> = vec![];
    for elem in elements {
        let const_name = &elem.proto.to_class_case();
        constdef.push(format!(
            "\t\tpublic const ushort {} = {};",
            const_name, elem.id
        ));
        dicid.push(format!(
            "\t\t\t{{typeof({}), {}}},",
            &elem.proto, const_name
        ));
        dicparser.push(format!(
            "\t\t\t{{{},  {}.Parser}},",
            const_name, &elem.proto
        ));
    }

    let out = format!(
        "using System;
using Google.Protobuf;
using System.Collections.Generic;
namespace Script.Network
{{
    public class PB
    {{
{}


        private static Dictionary<Type, ushort> _dic_id = new Dictionary<Type, ushort>()
        {{
{}

        }};

        private static Dictionary<ushort, MessageParser> _dic_parser = new Dictionary<ushort, MessageParser>()
        {{
{}

        }};
        public static ushort GetCmdID(IMessage obj)
        {{
          ushort cmd = 0;
          Type type = obj.GetType();
          if (_dic_id.TryGetValue(type, out cmd))
          {{
            return cmd;
          }}

          return cmd;
        }}

        public static MessageParser GetParser(ushort id)
        {{
          MessageParser parser;
          if (_dic_parser.TryGetValue(id, out parser))
          {{
            return parser;
          }}

          return parser;
        }}
    }}
}}",
        constdef.join("\n"),
        dicid.join("\n"),
        dicparser.join("\n")
    );
    let output_file_name = String::from("PB.cs");
    let path_str = format!("{}/{}", out_path, output_file_name);
    fs::write(path_str, out).unwrap();
}

fn data_to_lua(out_path: &String, elements: &Vec<Element>) {
    let mut pt: Vec<String> = vec![];
    let mut pt_names: Vec<String> = vec![];
    let mut pkgs: Vec<String> = vec![];

    for elem in elements {
        let const_name = &elem.proto.replace(".", "_").to_lowercase();
        pt.push(format!("\t{} = {}", const_name, elem.id));

        let pt_str = format!(
            "\t[{}] = [[{}]]",
            &elem.id,
            lowercase_first_letter(&elem.proto)
        );
        pt_names.push(pt_str);

        let pkgs_str = format!("\t{}", &elem.package);
        if !pkgs.contains(&pkgs_str) {
            pkgs.push(pkgs_str)
        }
    }

    let out = format!(
        "PT = {{
{}
}}
PT_NAMES = {{
{}
}}
PT_PKGS = {{
{}
}}",
        pt.join(",\n"),
        pt_names.join(",\n"),
        pkgs.join(",\n"),
    );
    let output_file_name = String::from("PT.lua");
    let path_str = format!("{}/{}", out_path, output_file_name);
    fs::write(path_str, out).unwrap();
}

fn lowercase_first_letter(s: &str) -> String {
    s[0..1].to_lowercase() + &s[1..]
}
