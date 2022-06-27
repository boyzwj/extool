#[macro_use]
extern crate serde_json;
extern crate ahash;
extern crate calamine;
extern crate clap;
extern crate inflector;

#[macro_use]
extern crate log;

extern crate num_cpus;
extern crate serde;
extern crate static_init;
extern crate threads_pool;
use std::fs;
use std::sync::mpsc::channel;
use std::time::SystemTime;
use threads_pool::prelude::*;
// use clap::{App, Arg, SubCommand};
// use threads_pool::*;
mod tool;
use clap::Parser;
/// Excel表格导出工具
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Args {
    ///Excel源文件目录
    #[clap(short, long, value_parser, default_value = "./")]
    input_path: String,
    ///导出目录
    #[clap(short, long, value_parser, default_value = "./")]
    output_path: String,
    ///导出格式 NONE | JSON | LUA | EX
    #[clap(short, long, value_parser, default_value = "NONE")]
    format: String,
}

fn main() {
    env_logger::init();
    let args = Args::parse();
    gen(args);
}

fn gen(args: Args) {
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
        pool.execute(move || {
            tool::build_id(file1);
            tx.send(()).unwrap();
        })
        .ok();
    }
    for _ in 0..xls_files.len() {
        rc.recv().unwrap();
    }

    //******************* EXPORT FILE  ***************//
    for file in &xls_files {
        let file1 = file.clone();
        let tx = tx.clone();
        let dst_path = args.output_path.clone();
        let format = args.format.clone();
        pool.execute(move || {
            tool::xls_to_file(file1, dst_path, format);
            tx.send(()).unwrap();
        })
        .ok();
    }
    for _ in 0..xls_files.len() {
        rc.recv().unwrap();
    }

    pool.clear();
    pool.close();
    match now.elapsed() {
        Ok(elapsed) => {
            info!("任务完成,总共耗时 {} 毫秒!", elapsed.as_millis());
        }
        Err(e) => {
            error!("Error: {:?}", e);
        }
    }
}

fn all_xls(path_str: &str) -> Vec<String> {
    let mut res: Vec<String> = vec![];
    let objects = fs::read_dir(path_str).unwrap();
    for obj in objects {
        let path = obj.unwrap().path();
        match path.extension() {
            Some(x) if x == "xlsx" || x == "xls" => {
                let p_str = format!("{}", path.display());
                res.push(p_str);
            }
            _ => {}
        }
    }
    res
}
