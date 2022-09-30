#[macro_use]
extern crate serde_json;
extern crate ahash;
extern crate calamine;
extern crate clap;
extern crate inflector;

#[macro_use]
extern crate log;

extern crate indexmap;
extern crate num_cpus;
extern crate prost;
extern crate prost_reflect;
extern crate regex;
extern crate serde;
extern crate simple_excel_writer;
extern crate static_init;
extern crate threads_pool;
use std::fs;
use std::sync::mpsc::channel;
use std::time::SystemTime;
use threads_pool::prelude::*;
// use clap::{App, Arg, SubCommand};
// use threads_pool::*;
mod excel;
mod proto;
use clap::Parser;
/// Excel表格导出工具
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Args {
    ///源文件类型 EXCEL|PROTO
    #[clap(short, long, value_parser, default_value = "EXCEL")]
    type_input: String,
    ///源文件目录
    #[clap(short, long, value_parser, default_value = "./")]
    input_path: String,
    ///导出目录
    #[clap(short, long, value_parser, default_value = "./")]
    output_path: String,
    ///导出格式 NONE | JSON | LUA | EX | CS | PBD | LANG
    #[clap(short, long, value_parser, default_value = "NONE")]
    format: String,
    ///是否允许多sheets导出 TRUE|FALSE
    #[clap(short, long, value_parser, default_value = "FALSE")]
    multi_sheets: String,
    ///导出的列 FRONT|BACK|BOTH
    #[clap(short, long, value_parser, default_value = "FRONT")]
    export_columns: String,
}

fn main() -> Result<(), usize> {
    env_logger::init();
    let args = Args::parse();
    let type_input = args.type_input.to_uppercase();
    match type_input.as_str() {
        "EXCEL" => match gen_from_excel(args) {
            0 => return Ok(()),
            _ => return Err(1),
        },
        "PROTO" => match gen_from_proto(args) {
            0 => return Ok(()),
            _ => return Err(1),
        },
        _ => {
            error!("type input={:?} is unsupported!", type_input);
            return Err(1);
        }
    }
}

fn all_files(path_str: &str, exts: Vec<String>) -> Vec<String> {
    let mut res: Vec<String> = vec![];

    let objects = fs::read_dir(path_str).unwrap();
    for obj in objects {
        let path = obj.unwrap().path();
        let file_name = path.file_name().unwrap().to_str().unwrap();
        if file_name.starts_with("~$") {
            continue;
        }
        match path.extension() {
            Some(x) if exts.iter().any(|e| e.as_str() == x) => {
                let p_str = format!("{}", path.display());
                res.push(p_str);
            }
            _ => {}
        }
    }
    res
}

fn gen_from_excel(args: Args) -> usize {
    let now = SystemTime::now();
    info!("导出格式: {} ", args.format);
    let xls_files = all_xls(args.input_path.as_str());
    let pool_size = num_cpus::get();

    let mut pool = ThreadPool::new(pool_size);
    let (tx, rc) = channel();

    //******************* BUILD ID   ***************
    info!("开始构建全局索引并进行ID检查 ...");
    for file in &xls_files {
        let file1 = file.clone();
        let tx = tx.clone();
        let multi_sheets = args.multi_sheets.clone().to_uppercase() == "TRUE";
        let export_columns = args.export_columns.clone().to_uppercase();
        pool.execute(move || {
            let code = excel::build_id(file1, multi_sheets, export_columns);
            tx.send(code).unwrap();
        })
        .ok();
    }
    let mut err_flag: usize = 0;
    for _ in 0..xls_files.len() {
        let result = rc.recv().unwrap();
        err_flag = err_flag + result
    }
    if err_flag > 0 {
        pool.clear();
        pool.close();
        return 1;
    }

    //******************* EXPORT FILE  ***************//
    for file in &xls_files {
        let file1 = file.clone();
        let tx = tx.clone();
        let dst_path = args.output_path.clone();
        let format = args.format.clone();
        let multi_sheets = args.multi_sheets.clone().to_uppercase() == "TRUE";
        let export_columns = args.export_columns.clone().to_uppercase();
        pool.execute(move || {
            let code = excel::xls_to_file(file1, dst_path, format, multi_sheets, export_columns);
            tx.send(code).unwrap();
        })
        .ok();
    }
    for _ in 0..xls_files.len() {
        let result = rc.recv().unwrap();
        err_flag = err_flag + result;
    }
    pool.clear();
    pool.close();
    if err_flag > 0 {
        return 1;
    }

    if args.type_input.to_uppercase() == "EXCEL" && args.format.to_uppercase() == "PBD" {
        match excel::create_pbd_file(&args.output_path) {
            0 => (),
            _ => return 1,
        }
    }

    if args.type_input.to_uppercase() == "EXCEL" && args.format.to_uppercase() == "LANG" {
        match excel::create_lang_file(&args.output_path) {
            0 => (),
            _ => return 1,
        }
    }

    match now.elapsed() {
        Ok(elapsed) => {
            info!("任务完成,总共耗时 {} 毫秒!", elapsed.as_millis());
        }
        Err(e) => {
            error!("Error: {:?}", e);
            return 1;
        }
    }
    return 0;
}

fn all_xls(path_str: &str) -> Vec<String> {
    all_files(
        path_str,
        [String::from("xlsx"), String::from("xls")].to_vec(),
    )
}

fn gen_from_proto(args: Args) -> usize {
    let exts = [String::from("proto")].to_vec();
    let files = all_files(args.input_path.as_str(), exts);
    proto::create(files, &args.output_path, &args.format, &args.type_input);
    return 0;
}
