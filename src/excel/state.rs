use ahash::{AHashMap, AHashSet};
use static_init::dynamic;

#[dynamic]
pub(crate) static mut GLOBAL_IDS: AHashSet<String> = AHashSet::new();

#[dynamic]
pub(crate) static mut GLOBAL_FRONT_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_BACK_PRIMARYS: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_PBD: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_LANG: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_LANG_SOURCE: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_LANG_HASHES: AHashMap<u32, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_EXCLUDE_SHEETS: AHashSet<String> = AHashSet::new();

#[dynamic]
pub(crate) static mut GLOBAL_MODS: AHashSet<String> = AHashSet::new();

#[dynamic]
pub(crate) static mut GLOBAL_GROUP_NAMES: AHashMap<String, String> = AHashMap::new();

#[dynamic]
pub(crate) static mut GLOBAL_GROUP_IDS: AHashSet<String> = AHashSet::new();

#[dynamic]
pub(crate) static mut GLOBAL_MOD_IDS: AHashMap<String, Vec<String>> = AHashMap::new();

pub(crate) fn sheet_key(input_file_name: &str, sheet_name: &str) -> String {
    format!("{}.{}", input_file_name, sheet_name)
}

pub(crate) fn exclude_sheet(input_file_name: &str, sheet_name: &str) {
    GLOBAL_EXCLUDE_SHEETS
        .write()
        .insert(sheet_key(input_file_name, sheet_name));
}

pub(crate) fn is_excluded_sheet(input_file_name: &str, sheet_name: &str) -> bool {
    GLOBAL_EXCLUDE_SHEETS
        .read()
        .contains(&sheet_key(input_file_name, sheet_name))
}

pub(crate) fn find_group_mod_names(group_name: &str) -> Vec<String> {
    let group_names = GLOBAL_GROUP_NAMES.read();
    let mut contents = Vec::new();
    for (mod_name, value) in group_names.iter() {
        if value == group_name {
            contents.push(mod_name.to_string());
        }
    }
    contents
}

pub(crate) fn find_group_mod_name_that_contain_id(group_name: &str, id: &str) -> Vec<String> {
    let mod_names = find_group_mod_names(group_name);
    let mut contents = Vec::new();
    for mod_name in mod_names {
        if let Some(ids) = GLOBAL_MOD_IDS.read().get(&mod_name) {
            if ids.contains(&id.to_string()) {
                contents.push(mod_name);
            }
        }
    }
    contents
}

pub(crate) fn get_all_grouped_mods() -> AHashMap<String, Vec<String>> {
    let group_names = GLOBAL_GROUP_NAMES.read();
    let mut grouped: AHashMap<String, Vec<String>> = AHashMap::new();
    for (mod_name, group_name) in group_names.iter() {
        grouped
            .entry(group_name.to_string())
            .or_insert_with(Vec::new)
            .push(mod_name.to_string());
    }
    grouped
}

pub(crate) fn get_ids_set(mod_names: &[String]) -> AHashMap<String, Vec<String>> {
    let mut result = AHashMap::new();
    for mod_name in mod_names {
        if let Some(ids) = GLOBAL_MOD_IDS.read().get(mod_name) {
            result.insert(mod_name.to_string(), ids.to_vec());
        }
    }
    result
}
