use std::collections::HashMap;

pub(crate) type ExcelResult<T> = Result<T, String>;
pub(crate) type EnumMap = HashMap<String, Vec<EnumValue>>;

pub(crate) struct SheetData {
    pub(crate) input_file_name: String,
    pub(crate) output_file_name: String,
    pub(crate) sheet_name: String,
    pub(crate) mod_name: String,
    pub(crate) names: Vec<String>,
    pub(crate) refs: Vec<String>,
    pub(crate) types: Vec<String>,
    pub(crate) describes: Vec<String>,
    pub(crate) enum_names: Vec<String>,
    pub(crate) values: Vec<Vec<String>>,
    pub(crate) force_mods: Vec<String>,
    pub(crate) export_columns: String,
    pub(crate) valid_columns: Vec<usize>,
    pub(crate) valid_front_types: Vec<String>,
}

pub(crate) struct EnumValue {
    pub(crate) index: i32,
    pub(crate) comment: String,
}

pub(crate) fn log_result(result: ExcelResult<()>) -> usize {
    match result {
        Ok(()) => 0,
        Err(err) => {
            error!("{}", err);
            1
        }
    }
}
