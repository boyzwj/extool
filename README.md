# extool

[English](README.md) | [中文](docs/README.zh-CN.md)

`extool` is a configuration export tool for game projects. It focuses on the project workflow of exporting Excel and Proto sources to client and server data assets.

## Features

- Rust-based CLI with batch export and thread-pool concurrency.
- Supports Excel and Proto inputs.
- Excel outputs support `JSON`, `LUA`, `EX`, `PBD`, `LANG`, and `GROUP`.
- Proto outputs support `LUA` and `CS`.
- Excel validation covers duplicate primary keys, duplicate MOD names, duplicate field names, invalid references, invalid enum values, and `string_loc` hash collisions.
- One `.xlsx` file may contain only one real data sheet, while additional documentation sheets are allowed.

## Build

```bash
cargo build --release
```

PBD dependencies build protobuf. If your machine uses CMake 4 and hits an old protobuf CMake compatibility error, use:

```bash
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo check
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo build --release
```

`static_init` is updated to `1.0.4` for compatibility with Rust 1.93 and newer.

## Usage

### Excel Export

Windows:

```bat
SET RUST_LOG=debug&&extool -t EXCEL -i ExcelInputDir -o OutputDir -f LUA -m TRUE -e FRONT
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t EXCEL -i ExcelInputDir -o OutputDir -f LUA -m TRUE -e FRONT
```

### Proto Export

Windows:

```bat
SET RUST_LOG=debug&&extool -t PROTO -i ProtoInputDir -o OutputDir -f CS
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t PROTO -i ProtoInputDir -o OutputDir -f CS
```

## Arguments

| Argument | Default | Description |
| --- | --- | --- |
| `-t, --type-input` | `EXCEL` | Input type: `EXCEL` or `PROTO`. |
| `-i, --input-path` | `./` | Input directory. |
| `-o, --output-path` | `./` | Output directory. |
| `-f, --format` | `NONE` | Output format: `JSON`, `LUA`, `EX`, `CS`, `PBD`, `LANG`, or `GROUP`. |
| `-m, --multi-sheets` | `FALSE` | Scan all sheets to allow documentation sheets. This is not multi data-sheet export. |
| `-e, --export-columns` | `FRONT` | Export column set: `FRONT`, `BACK`, or `BOTH`. |
| `-p, --pbe-file` | `pbe.proto` | Enum proto file name, read from `{input_path}/enum/pbe.proto` by default. |

## Excel File Rules

- Use `.xlsx` files.
- Recommended naming style: Chinese initials plus Chinese name, for example `C常量配置.xlsx`.
- One `.xlsx` file represents one config module.
- One `.xlsx` file can have at most one data sheet with a valid `MOD`.
- Other sheets can contain README content, field notes, images, examples, or drafts. Sheets without `MOD` are ignored.
- With `-m TRUE`, the tool scans all sheets, finds the single data sheet, and ignores documentation sheets.
- With `-m FALSE`, only the first sheet is read.

## Sheet Protocol

A data sheet uses fixed protocol rows:

| Column A | From Column B |
| --- | --- |
| `MOD` | Module name, for example `Data.Const`. |
| `BACK_TYPE` | Server-side field types. |
| `FRONT_TYPE` | Client-side field types. |
| `DES` | Field description. |
| `NAMES` | Field names, with column B as the primary key. |
| `ENUM` | Enum name. |
| `REF` | Referenced module name. |
| `FORCE_MOD` | Optional PBD class-name override. |
| `GROUP` | Optional GROUP aggregation name. |
| `VALUE` | Data row. |

Only rows whose column A is `VALUE` are exported. Other rows can be used for notes, drafts, or layout.

## Types

Common types:

- Scalars: `string`, `string_loc`, `bool`, `int`, `int32`, `int64`, `uint32`, `uint64`, `float`
- Lists: `list`, `list_int`, `list_int32`, `list_int64`, `list_uint32`, `list_uint64`, `list_string`, `list_float`

Type rules:

- `int` is equivalent to `uint32`.
- `list` and `list_int` are equivalent to `list_uint32`.
- Empty numeric cells export as `0`.
- Empty list cells export as empty lists.
- For `bool`, `是` or `TRUE` exports as `true`; other values export as `false`.
- `string_loc` exports the low 32 bits of the md5 hash of the string key. `LANG` export checks for hash collisions.

## Validation

Before export, the tool builds global indexes and validates:

- `MOD` must be unique.
- The primary key field in column B of `NAMES` cannot be empty.
- Field names cannot be duplicated or start with a digit.
- Primary key values cannot be empty or duplicated.
- `REF` values must exist in the target module primary-key set. List references are validated item by item.
- `ENUM` values must be found in enum proto comments.
- One Excel file cannot contain multiple exportable data sheets.

## Outputs

- `JSON`: Generates an object map keyed by primary key.
- `LUA`: Generates read-only tables. Field names are mapped to array indexes through `KT`.
- `EX`: Generates Elixir modules.
- `PBD`: Generates `data_*.bytes` and aggregated `pbd.proto`.
- `LANG`: Collects `string_loc` fields and generates `D多语言简体中文表.xlsx`.
- `GROUP`: Generates mapping tables based on `GROUP` settings.

## Code Layout

Excel export logic is split by responsibility:

- `src/excel/workbook.rs`: Opens workbooks, selects sheets, and organizes the export flow.
- `src/excel/sheet.rs`: Parses protocol rows and builds primary-key, reference, and enum validation.
- `src/excel/export.rs`: Exports one table to JSON, LUA, EX, PBD, or LANG.
- `src/excel/output_files.rs`: Writes aggregated PBD, LANG, and GROUP files.
- `src/excel/value.rs`: Handles type normalization, cell conversion, and enum parsing.
- `src/excel/state.rs`: Stores global index state.

## Example

| A | B | C | D |
| --- | --- | --- | --- |
| `MOD` | `Data.Const` | | |
| `BACK_TYPE` | `uint32` | `string` | |
| `FRONT_TYPE` | `uint32` | `string` | |
| `DES` | `id` | `Description` | |
| `NAMES` | `id` | `desc` | |
| `ENUM` | | | |
| `REF` | | | |
| `VALUE` | `1` | `Example config` | |
