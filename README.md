# extool

游戏项目配置导出工具，面向“Excel/Proto -> 客户端与服务端配置资源”的项目内工作流。

A configuration export tool for game projects, focused on the project workflow of exporting Excel/Proto sources to client and server data assets.

## 特性 / Features

- 基于 Rust CLI，支持批量导出和线程池并发处理。
- Rust-based CLI with batch export and thread-pool concurrency.
- 支持 Excel 与 Proto 两类输入。
- Supports Excel and Proto inputs.
- Excel 输出支持 `JSON`、`LUA`、`EX`、`PBD`、`LANG`、`GROUP`。
- Excel outputs support `JSON`, `LUA`, `EX`, `PBD`, `LANG`, and `GROUP`.
- Proto 输出支持 `LUA`、`CS`。
- Proto outputs support `LUA` and `CS`.
- Excel 校验包含主键重复、MOD 重复、字段名重复、引用合法性、枚举合法性、`string_loc` hash 碰撞。
- Excel validation covers duplicate primary keys, duplicate MOD names, duplicate field names, invalid references, invalid enum values, and `string_loc` hash collisions.
- 一个 `.xlsx` 文件只允许一个真实数据 Sheet；可以包含多个文档/说明 Sheet。
- One `.xlsx` file may contain only one real data sheet, while additional documentation sheets are allowed.

## 构建 / Build

```bash
cargo build --release
```

PBD 相关依赖会构建 protobuf。若本机使用 CMake 4，并遇到旧 protobuf CMake 兼容错误，可以临时这样检查或构建：

PBD dependencies build protobuf. If your machine uses CMake 4 and hits an old protobuf CMake compatibility error, use:

```bash
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo check
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo build --release
```

`static_init` 已更新到 `1.0.4`，以兼容 Rust 1.93 及更新版本。

`static_init` is updated to `1.0.4` for compatibility with Rust 1.93 and newer.

## 使用 / Usage

### Excel 导出 / Excel Export

Windows:

```bat
SET RUST_LOG=debug&&extool -t EXCEL -i 源Excel目录 -o 输出目录 -f LUA -m TRUE -e FRONT
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t EXCEL -i 源Excel目录 -o 输出目录 -f LUA -m TRUE -e FRONT
```

### Proto 导出 / Proto Export

Windows:

```bat
SET RUST_LOG=debug&&extool -t PROTO -i 源Proto目录 -o 输出目录 -f CS
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t PROTO -i 源Proto目录 -o 输出目录 -f CS
```

## 参数 / Arguments

| 参数 / Argument | 默认值 / Default | 说明 / Description |
| --- | --- | --- |
| `-t, --type-input` | `EXCEL` | 输入类型：`EXCEL` 或 `PROTO`。Input type: `EXCEL` or `PROTO`. |
| `-i, --input-path` | `./` | 输入目录。Input directory. |
| `-o, --output-path` | `./` | 输出目录。Output directory. |
| `-f, --format` | `NONE` | 输出格式：`JSON`、`LUA`、`EX`、`CS`、`PBD`、`LANG`、`GROUP`。Output format. |
| `-m, --multi-sheets` | `FALSE` | 是否扫描所有 Sheet。用于允许文档 Sheet；不是多数据 Sheet 导出。Scan all sheets to allow documentation sheets; this is not multi data-sheet export. |
| `-e, --export-columns` | `FRONT` | 导出列：`FRONT`、`BACK`、`BOTH`。Export column set: `FRONT`, `BACK`, or `BOTH`. |
| `-p, --pbe-file` | `pbe.proto` | 枚举 proto 文件名，默认读取 `{input_path}/enum/pbe.proto`。Enum proto file name, read from `{input_path}/enum/pbe.proto` by default. |

## Excel 文件规则 / Excel File Rules

- 文件统一使用 `.xlsx`。
- Use `.xlsx` files.
- 推荐命名为“中文名称首字母 + 中文名称”，例如 `C常量配置.xlsx`。
- Recommended naming style: Chinese initials plus Chinese name, for example `C常量配置.xlsx`.
- 一个 `.xlsx` 文件代表一个配置模块。
- One `.xlsx` file represents one config module.
- 一个 `.xlsx` 文件最多只能有一个带有效 `MOD` 的数据 Sheet。
- One `.xlsx` file can have at most one data sheet with a valid `MOD`.
- 其他 Sheet 可以放 README、字段说明、图片、示例、草稿；没有 `MOD` 的 Sheet 会被忽略。
- Other sheets can contain README content, field notes, images, examples, or drafts. Sheets without `MOD` are ignored.
- 如果使用 `-m TRUE`，工具会扫描所有 Sheet，找到唯一数据 Sheet 并忽略文档 Sheet。
- With `-m TRUE`, the tool scans all sheets, finds the single data sheet, and ignores documentation sheets.
- 如果使用 `-m FALSE`，只读取第一个 Sheet。
- With `-m FALSE`, only the first sheet is read.

## Sheet 协议 / Sheet Protocol

数据 Sheet 使用固定协议行：

A data sheet uses fixed protocol rows:

| A 列 / Column A | B 列开始 / From Column B |
| --- | --- |
| `MOD` | 模块名，例如 `Data.Const`。Module name, for example `Data.Const`. |
| `BACK_TYPE` | 服务端字段类型。Server-side field types. |
| `FRONT_TYPE` | 客户端字段类型。Client-side field types. |
| `DES` | 字段中文描述。Field description. |
| `NAMES` | 字段名，B 列固定为主键。Field names, with column B as the primary key. |
| `ENUM` | 枚举名。Enum name. |
| `REF` | 引用的模块名。Referenced module name. |
| `FORCE_MOD` | 可选，PBD 类名覆盖。Optional PBD class-name override. |
| `GROUP` | 可选，GROUP 聚合名。Optional GROUP aggregation name. |
| `VALUE` | 数据行。Data row. |

只有 A 列为 `VALUE` 的行会导出。其他行可以用作注释、草稿或排版。

Only rows whose column A is `VALUE` are exported. Other rows can be used for notes, drafts, or layout.

## 类型 / Types

常用类型：

Common types:

- 标量 / Scalars: `string`, `string_loc`, `bool`, `int`, `int32`, `int64`, `uint32`, `uint64`, `float`
- 列表 / Lists: `list`, `list_int`, `list_int32`, `list_int64`, `list_uint32`, `list_uint64`, `list_string`, `list_float`

类型规则：

Type rules:

- `int` 等价于 `uint32`。
- `int` is equivalent to `uint32`.
- `list` 和 `list_int` 等价于 `list_uint32`。
- `list` and `list_int` are equivalent to `list_uint32`.
- 空数字导出为 `0`。
- Empty numeric cells export as `0`.
- 空列表导出为空列表。
- Empty list cells export as empty lists.
- `bool` 中 `是` 或 `TRUE` 导出为 `true`，其他值导出为 `false`。
- For `bool`, `是` or `TRUE` exports as `true`; other values export as `false`.
- `string_loc` 导出为字符串 key 的 md5 低 32 位 hash；`LANG` 导出会检查 hash 碰撞。
- `string_loc` exports the low 32 bits of the md5 hash of the string key. `LANG` export checks for hash collisions.

## 校验 / Validation

导出前会先构建全局索引并校验：

Before export, the tool builds global indexes and validates:

- `MOD` 不能重复。
- `MOD` must be unique.
- `NAMES` 的 B 列主键字段不能为空。
- The primary key field in column B of `NAMES` cannot be empty.
- 字段名不能重复，不能以数字开头。
- Field names cannot be duplicated or start with a digit.
- 主键值不能为空，不能重复。
- Primary key values cannot be empty or duplicated.
- `REF` 引用必须能在目标模块主键集合中找到；列表引用会逐项校验。
- `REF` values must exist in the target module primary-key set. List references are validated item by item.
- `ENUM` 值必须能在枚举 proto 的注释中找到。
- `ENUM` values must be found in enum proto comments.
- 一个 Excel 文件中不能出现多个可导出的数据 Sheet。
- One Excel file cannot contain multiple exportable data sheets.

## 输出说明 / Outputs

- `JSON`：按主键生成对象映射。
- `JSON`: Generates an object map keyed by primary key.
- `LUA`：生成只读表，字段名通过 `KT` 映射到数组下标。
- `LUA`: Generates read-only tables. Field names are mapped to array indexes through `KT`.
- `EX`：生成 Elixir 模块。
- `EX`: Generates Elixir modules.
- `PBD`：生成 `data_*.bytes` 与聚合 `pbd.proto`。
- `PBD`: Generates `data_*.bytes` and aggregated `pbd.proto`.
- `LANG`：收集 `string_loc` 字段，生成 `D多语言简体中文表.xlsx`。
- `LANG`: Collects `string_loc` fields and generates `D多语言简体中文表.xlsx`.
- `GROUP`：按 `GROUP` 配置生成映射表。
- `GROUP`: Generates mapping tables based on `GROUP` settings.

## 代码结构 / Code Layout

Excel 导出逻辑按职责拆分：

Excel export logic is split by responsibility:

- `src/excel/workbook.rs`：打开 workbook、选择 Sheet、组织导出流程。
- `src/excel/workbook.rs`: Opens workbooks, selects sheets, and organizes the export flow.
- `src/excel/sheet.rs`：解析协议行，构建主键/ref/enum 校验。
- `src/excel/sheet.rs`: Parses protocol rows and builds primary-key, reference, and enum validation.
- `src/excel/export.rs`：JSON/LUA/EX/PBD/LANG 单表导出。
- `src/excel/export.rs`: Exports one table to JSON, LUA, EX, PBD, or LANG.
- `src/excel/output_files.rs`：PBD/LANG/GROUP 聚合文件。
- `src/excel/output_files.rs`: Writes aggregated PBD, LANG, and GROUP files.
- `src/excel/value.rs`：类型归一化、单元格值转换、枚举解析。
- `src/excel/value.rs`: Handles type normalization, cell conversion, and enum parsing.
- `src/excel/state.rs`：全局索引状态。
- `src/excel/state.rs`: Stores global index state.

## 示例 / Example

| A | B | C | D |
| --- | --- | --- | --- |
| `MOD` | `Data.Const` | | |
| `BACK_TYPE` | `uint32` | `string` | |
| `FRONT_TYPE` | `uint32` | `string` | |
| `DES` | `id` | `描述 / Description` | |
| `NAMES` | `id` | `desc` | |
| `ENUM` | | | |
| `REF` | | | |
| `VALUE` | `1` | `示例配置 / Example config` | |
