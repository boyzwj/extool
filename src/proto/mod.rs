use std::fs::{File, self};
use std::io::{BufRead, BufReader};

struct Element {
    package: String,
    proto: String,
    id: u128
}


fn find_element(filepath: &str,elements: &mut Vec<Element>) -> Result<(), Box<dyn std::error::Error>> {
    let file = File::open(filepath)?;
    let reader = BufReader::new(file);

    let mut _find_package  = String::new();
    let mut stack: Vec<Element> = vec![];
    for line in reader.lines() {
        let temp = line?.trim().to_string();
        if temp.starts_with("package") {
            if _find_package.is_empty() {
                let v: Vec<&str> =  temp.split(&[' ', ';']).collect();
                _find_package = v[1].to_string();
                continue;
            }
            else{
                error!("proto file={:?} is invalid!", filepath)
            }
        }

        if temp.starts_with("message"){
            let v: Vec<&str> =  temp.split('{').collect();
            let head = v[0].trim();
            if head.ends_with("2C") || head.ends_with("2S") {
                let v2: Vec<&str> = head.split("message").collect();
                let msg_name = v2[1].trim().to_string();
                let proto = format!("{}.{}",uppercase_first_letter(&_find_package),msg_name);
                let digest = md5::compute(&proto);
                let bytes = format!("{:x}", digest);
                let k2 = u128::from_str_radix(&bytes, 16).unwrap();
                let k3: u128 = 65536;
                let msg_id = k2 % k3;
                let elem = Element{package: _find_package.to_string(),proto,id: msg_id};
                stack.push(elem);
            }
        }

        if temp.ends_with("}"){
            let last = stack.pop();
            if stack.is_empty() && last.is_some() {
                match last {
                    Some(value) => elements.push(value),
                    None => continue
                };
            }
        }
    }

    Ok(())
}


pub fn create(files: Vec<String>,out_path: &String,format: &String,type_input: &String){
    let mut elements: Vec<Element> = vec![];
    for file in &files {
        find_element(&file,&mut elements).unwrap();
    }
    match format.to_uppercase().as_str() {
        "CS" => data_to_cs(out_path,&elements),
        "LUA" => data_to_lua(out_path,&elements),
        _ => error!("type input={:?} and format={} is unsupport!", type_input, format),
    }
    
}

fn data_to_cs(out_path: &String,elements: &Vec<Element>) {
    let mut constdef: Vec<String> = vec![];
    let mut type2id: Vec<String> = vec![];
    let mut id2parser: Vec<String> = vec![];
   
    for elem in elements {
        let const_name = &elem.proto.replace(".","");
        constdef.push(format!("\t\tpublic const ushort {} = {};",const_name,elem.id));

        let const_type2id = format!(
"\t\t\tif (type == typeof({}))
\t\t\t{{
\t\t\t    return {};
\t\t\t}} ",&elem.proto,const_name);
        type2id.push(const_type2id);

        let const_id2parser = format!(
"\t\t\tif (id == {})
\t\t\t{{
\t\t\t    return {}.Parser;
\t\t\t}} ",const_name,&elem.proto);
        id2parser.push(const_id2parser)
    }

    let out = format!(
"using System;
using Google.Protobuf;
namespace Script.Network
{{
    public class PB
    {{
{}


        public static ushort GetCmdID(IMessage obj)
        {{
            Type type = obj.GetType();
{}

            return 0;
        }}

        public static MessageParser GetParser(ushort id)
        {{
{}

            return null;
        }}
    }}
}}",
        constdef.join("\n"),
        type2id.join("\n"),
        id2parser.join("\n"),
    );
    let output_file_name = String::from("PB.cs");
    let path_str = format!("{}/{}", out_path, output_file_name);
    fs::write(path_str, out).unwrap();
}




fn data_to_lua(out_path: &String,elements: &Vec<Element>) {
    let mut pt: Vec<String> = vec![];
    let mut pt_names: Vec<String> = vec![];
    let mut pkgs: Vec<String> = vec![];
   
    for elem in elements {
        let const_name = &elem.proto.replace(".","_").to_lowercase();
        pt.push(format!("\t{} = {}",const_name,elem.id));

        let pt_str = format!("\t[{}] = [[{}]]",&elem.id,lowercase_first_letter(&elem.proto));
        pt_names.push(pt_str);

        let pkgs_str = format!("\t{}",&elem.package);
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

fn uppercase_first_letter(s: &str) -> String {
    s[0..1].to_uppercase() + &s[1..]
}

fn lowercase_first_letter(s: &str) -> String {
    s[0..1].to_lowercase() + &s[1..]
}