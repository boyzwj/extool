# extool

[English](../README.md) | [中文](README.zh-CN.md)

`extool` 是游戏项目配置导出工具，面向“Excel/Proto -> 客户端与服务端配置资源”的项目内工作流。

## 特性

- 基于 Rust CLI，支持批量导出和线程池并发处理。
- 支持 Excel 与 Proto 两类输入。
- Excel 输出支持 `JSON`、`LUA`、`EX`、`PBD`、`LANG`、`GROUP`。
- Proto 输出支持 `LUA`、`CS`。
- Excel 校验包含主键重复、MOD 重复、字段名重复、引用合法性、枚举合法性、`string_loc` hash 碰撞。
- 一个 `.xlsx` 文件只允许一个真实数据 Sheet；可以包含多个文档/说明 Sheet。

## 构建

```bash
cargo build --release
```

PBD 相关依赖会构建 protobuf。若本机使用 CMake 4，并遇到旧 protobuf CMake 兼容错误，可以临时这样检查或构建：

```bash
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo check
CMAKE_POLICY_VERSION_MINIMUM=3.5 cargo build --release
```

`static_init` 已更新到 `1.0.4`，以兼容 Rust 1.93 及更新版本。

## 使用

### Excel 导出

Windows:

```bat
SET RUST_LOG=debug&&extool -t EXCEL -i 源Excel目录 -o 输出目录 -f LUA -m TRUE -e FRONT
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t EXCEL -i 源Excel目录 -o 输出目录 -f LUA -m TRUE -e FRONT
```

### Proto 导出

Windows:

```bat
SET RUST_LOG=debug&&extool -t PROTO -i 源Proto目录 -o 输出目录 -f CS
```

macOS / Linux:

```bash
RUST_LOG=debug extool -t PROTO -i 源Proto目录 -o 输出目录 -f CS
```

## 参数

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `-t, --type-input` | `EXCEL` | 输入类型：`EXCEL` 或 `PROTO`。 |
| `-i, --input-path` | `./` | 输入目录。 |
| `-o, --output-path` | `./` | 输出目录。 |
| `-f, --format` | `NONE` | 输出格式：`JSON`、`LUA`、`EX`、`CS`、`PBD`、`LANG`、`GROUP`。 |
| `-m, --multi-sheets` | `FALSE` | 是否扫描所有 Sheet。用于允许文档 Sheet；不是多数据 Sheet 导出。 |
| `-e, --export-columns` | `FRONT` | 导出列：`FRONT`、`BACK`、`BOTH`。 |
| `-p, --pbe-file` | `pbe.proto` | 枚举 proto 文件名，默认读取 `{input_path}/enum/pbe.proto`。 |

## Excel 文件规则

- 文件统一使用 `.xlsx`。
- 推荐命名为“中文名称首字母 + 中文名称”，例如 `C常量配置.xlsx`。
- 一个 `.xlsx` 文件代表一个配置模块。
- 一个 `.xlsx` 文件最多只能有一个带有效 `MOD` 的数据 Sheet。
- 其他 Sheet 可以放 README、字段说明、图片、示例、草稿；没有 `MOD` 的 Sheet 会被忽略。
- 如果使用 `-m TRUE`，工具会扫描所有 Sheet，找到唯一数据 Sheet 并忽略文档 Sheet。
- 如果使用 `-m FALSE`，只读取第一个 Sheet。

## Sheet 协议

数据 Sheet 使用固定协议行：

| A 列 | B 列开始 |
| --- | --- |
| `MOD` | 模块名，例如 `Data.Const`。 |
| `BACK_TYPE` | 服务端字段类型。 |
| `FRONT_TYPE` | 客户端字段类型。 |
| `DES` | 字段中文描述。 |
| `NAMES` | 字段名，B 列固定为主键。 |
| `ENUM` | 枚举名。 |
| `REF` | 引用的模块名。 |
| `FORCE_MOD` | 可选，PBD 类名覆盖。 |
| `GROUP` | 可选，GROUP 聚合名。 |
| `VALUE` | 数据行。 |

只有 A 列为 `VALUE` 的行会导出。其他行可以用作注释、草稿或排版。

## 类型

常用类型：

- 标量：`string`、`string_loc`、`bool`、`int`、`int32`、`int64`、`uint32`、`uint64`、`float`
- 列表：`list`、`list_int`、`list_int32`、`list_int64`、`list_uint32`、`list_uint64`、`list_string`、`list_float`

类型规则：

- `int` 等价于 `uint32`。
- `list` 和 `list_int` 等价于 `list_uint32`。
- 空数字导出为 `0`。
- 空列表导出为空列表。
- `bool` 中 `是` 或 `TRUE` 导出为 `true`，其他值导出为 `false`。
- `string_loc` 导出为字符串 key 的 md5 低 32 位 hash；`LANG` 导出会检查 hash 碰撞。

## 校验

导出前会先构建全局索引并校验：

- `MOD` 不能重复。
- `NAMES` 的 B 列主键字段不能为空。
- 字段名不能重复，不能以数字开头。
- 主键值不能为空，不能重复。
- `REF` 引用必须能在目标模块主键集合中找到；列表引用会逐项校验。
- `ENUM` 值必须能在枚举 proto 的注释中找到。
- 一个 Excel 文件中不能出现多个可导出的数据 Sheet。

## 输出说明

- `JSON`：按主键生成对象映射。
- `LUA`：生成只读表，字段名通过 `KT` 映射到数组下标。
- `EX`：生成 Elixir 模块。
- `PBD`：生成 `data_*.bytes` 与聚合 `pbd.proto`。
- `LANG`：收集 `string_loc` 字段，生成 `D多语言简体中文表.xlsx`。
- `GROUP`：按 `GROUP` 配置生成映射表。

## 代码结构

Excel 导出逻辑按职责拆分：

- `src/excel/workbook.rs`：打开 workbook、选择 Sheet、组织导出流程。
- `src/excel/sheet.rs`：解析协议行，构建主键/ref/enum 校验。
- `src/excel/export.rs`：JSON/LUA/EX/PBD/LANG 单表导出。
- `src/excel/output_files.rs`：PBD/LANG/GROUP 聚合文件。
- `src/excel/value.rs`：类型归一化、单元格值转换、枚举解析。
- `src/excel/state.rs`：全局索引状态。

## 示例

| A | B | C | D |
| --- | --- | --- | --- |
| `MOD` | `Data.Const` | | |
| `BACK_TYPE` | `uint32` | `string` | |
| `FRONT_TYPE` | `uint32` | `string` | |
| `DES` | `id` | `描述` | |
| `NAMES` | `id` | `desc` | |
| `ENUM` | | | |
| `REF` | | | |
| `VALUE` | `1` | `示例配置` | |
