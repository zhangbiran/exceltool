# ExcelTool

面向人和 AI Agent 的 `.xls/.xlsx` 命令行工具。它提供稳定的局部查看、二维
JSON 批量写入、查找、范围样式、sheet/行/列结构操作，以及单工作簿事务式 patch，调用者
无需为每次修改临时编写脚本。

## 开始使用

依赖、源码安装、升级、卸载和字体配置见 [Linux 安装说明](INSTALL_LINUX.md)或
[Windows 安装说明](INSTALL_WINDOWS.md)。

Linux 在仓库根目录可以直接运行：

```bash
./exceltool --version
./exceltool --help
```

Windows 使用 `exceltool.cmd`，安装到 `PATH` 后使用 `exceltool`。下文统一展示
`./exceltool`；Windows 调用时替换为 `exceltool`，命令参数保持一致。

所有读取或修改工作簿的命令都会取得当前操作系统用户的全局命令锁，并持有到
LibreOffice、UNO、验证和发布全部结束。若已有命令运行，新进程会在 stderr 提示一次
`ExcelTool 正被其他任务使用，等待前一个任务完成……`，随后等待并自动继续；正常退出
和异常终止都会由操作系统释放锁。`--help`、`--version` 和 `font check` 不使用该锁。

工作簿命令默认自动使用当前用户的轻量 daemon：daemon 只管理一个专属
LibreOffice 的启动、保活和关闭，CLI 仍直接连接 UNO 执行原有业务。每条命令会
关闭自己打开的工作簿，健康的 LibreOffice 保留给后续命令；连续空闲 300 秒后
daemon 和它创建的 LibreOffice 一起退出。运行状态和诊断入口：

```bash
./exceltool daemon start
./exceltool daemon status --json
./exceltool daemon stop
./exceltool --no-daemon view --file book.xls --sheet Sheet1
```

`--no-daemon` 保留一次一启、一次一关的隔离路径，用于故障诊断和对照回归。

调用方仍应串行执行命令，不要主动并发或在等待时启动重试；这可以保持资源占用、输入
版本、发布顺序和验证结果明确。同一工作簿的多项连续修改使用一次 `patch`，多个工作簿
按明确清单逐个处理。不要删除、替换或绕过工具的命令锁。

## 查看

查看整个 sheet 的有效区域：

```bash
./exceltool view --file book.xls --sheet Sheet1
```

按行列查看：

```bash
./exceltool view --file book.xls --sheet Sheet1 --rows 1:3 --cols A:F
```

查看有效区域最后 20 行，可继续限制列：

```bash
./exceltool view --file book.xls --sheet Sheet1 --tail 20 --cols A:F
```

`--tail` 按每个 sheet 各自的有效区域计算，数量必须是正整数；当有效行数不足 N
时返回全部有效行。它可以与 `--cols` 组合，但不能与 `--rows`、`--range` 或
`--from/--n` 同时使用。

矩形范围与正方区域：

```bash
./exceltool view --file book.xlsx --sheet Sheet1 --range B3:F20
./exceltool view --file book.xlsx --sheet Sheet1 --from C3 --n 5
```

一次 `view` 只处理一个工作簿；需要查看多个文件时，应先明确文件清单，再逐个调用。

### 可回写的纯 JSON

`--json` 要求明确一个文件和一个 sheet，stdout 只输出二维数组：

```bash
./exceltool view --file book.xls --sheet Sheet1 --range A1:C3 --json
```

```json
[
  [1, "苹果", null],
  [2, "香蕉", true],
  [3, "梨", "=A3*2"]
]
```

编码规则：

- 空单元格为 `null`。
- 数字和布尔值保持 JSON 类型。
- 公式是以 `=` 开头的字符串。
- 以 `=` 开头的普通文本编码为前导单引号形式，例如 `'=普通文本`。
- 该数组可直接传给 `write --stdin`。

### 带元数据的 JSON

```bash
./exceltool view --file book.xls --sheet Sheet1 --range A1:C3 --json-full
```

指定一个 sheet 时，输出包含 `file`、`sheet`、`range`、`begin` 和 `values`；省略
sheet 时输出 `{file, sheets}`，每个 sheet 独立包含范围和 values。

需要独立检查字体和字号时增加 `--include-style`：

```bash
./exceltool view --file book.xls --sheet Sheet1 --range A1:C3 \
  --json-full --include-style
```

输出增加与 `values` 对齐的二维 `styles`，每个单元格分别列出 `western`、
`asian`、`complex` 三个字体槽及字号。字体写后验证按单元格实际文字脚本检查相关
槽：拉丁字母/数字检查 Western，中日韩文字检查 Asian，阿拉伯/希伯来文字检查
Complex；结果仍完整报告三槽，便于识别 LibreOffice 的回退。`--include-style` 必须与 `--json-full`
一起使用，避免破坏可回写的裸二维 JSON。

## Sheet 操作

```bash
./exceltool sheet list --file book.xls
./exceltool sheet list --file book.xls --json
./exceltool sheet info --file book.xls
./exceltool sheet info --file book.xls --sheet Sheet1 --json

./exceltool sheet add --file book.xls --name 新表
./exceltool sheet delete --file book.xls --sheet 旧表
./exceltool sheet rename --file book.xls --sheet Sheet1 --name 数据
./exceltool sheet copy --file book.xls --sheet 模板 --name 新表
```

`sheet list --json` stdout 直接输出 sheet 名称数组。`sheet info` 返回 sheet 的有效范围、
有效行列数和最后行列；省略 `--sheet` 时返回全部 sheet。JSON 字段为 `used_range`、
`used_rows`、`used_cols`、`last_row` 和 `last_col`。这里的“有效”采用 LibreOffice 的
有效区域定义，因此只有格式的尾部空行或空列也可能计入。

## 查找

默认在所有 sheet 的有效区域中，不区分大小写地查找显示值和公式：

```bash
./exceltool find --file book.xls --text "任务" --json
```

也可以限制 sheet、矩形范围和查找来源：

```bash
./exceltool find --file book.xls --text "A10*2" \
  --sheet 任务 --range A1:K100 --look-in formulas \
  --case-sensitive --limit 200 --json
```

`find` 与 `view` 共用范围参数：`--rows`、`--cols`、`--range`、`--tail` 和
`--from/--n`。例如只在 F 列的有效行中查找：

```bash
./exceltool find --file book.xls --sheet 任务 --text "苹果" --cols F --json
```

也可以同时限制行列，或查找有效区域最后 20 行：

```bash
./exceltool find --file book.xls --sheet 任务 --text "苹果" \
  --rows 10:100 --cols F:H --json
./exceltool find --file book.xls --sheet 任务 --text "苹果" \
  --tail 20 --cols A:F --json
```

`--look-in` 可取 `values`、`formulas` 或 `both`，默认 `both`；`--limit` 默认
100。JSON 结果包含 `file`、`query`、`matches` 和 `truncated`，每条匹配包含
`sheet`、`cell`、`display`、`formula` 和 `match_in`。当实际匹配超过上限时
`truncated` 为 `true`，调用者应缩小范围或明确调大上限。

需要按格式查找错误码、ID 等结构化文本时使用 Python 标准正则：

```bash
./exceltool find --file book.xls --sheet 错误码 --cols C:C \
  --text '^ERR_[0-9]{4}$' --regex --json
```

`--regex` 与普通 `find` 使用完全相同的 sheet 和范围规则，可以搜索全部 sheet，
也可以配合 `--sheet`、`--rows`、`--cols`、`--range`、`--tail` 或 `--from/--n`
缩小范围。`--case-sensitive` 同样适用于正则；未指定时正则不区分大小写。非法正则
会在启动 LibreOffice 前返回退出码 3。Python 标准正则没有可靠的单次匹配超时，
因此 AI 应根据工作簿大小自行选择合理范围，并使用简单模式，避免嵌套重复等可能
产生灾难性回溯的写法。

查找 JSON 顶层还包含 `regex` 和 `case_sensitive`，便于调用者确认实际匹配模式。

## 行列定位与结构操作

行使用从 1 开始的数字，列可以使用 Excel 字母。常用定位参数如下：

```text
--rows 1:3       第 1～3 行
--tail 20        有效区域最后 20 行
--cols A:F       A～F 列
--range B2:F10   矩形范围
--from C3 --n 5  从 C3 开始的 5×5 正方区域
--begin AA3      AA 列第 3 行
```

在第 10 行前插入两行：

```bash
./exceltool row insert --file book.xls --sheet Sheet1 \
  --before 10 --count 2
```

删除第 3～5 行：

```bash
./exceltool row delete --file book.xls --sheet Sheet1 \
  --rows 3:5
```

复制第 3～5 行并插入第 20 行前：

```bash
./exceltool row copy --file book.xls --sheet Sheet1 \
  --rows 3:5 --insert-before 20
```

`row copy` 复制值、公式、样式和行高，并由 LibreOffice 按复制目标调整相对公式。
JSON 管道只复制值及公式文本，不复制样式和行高。

在 F 列前插入两列、删除 F～H 列：

```bash
./exceltool col insert --file book.xls --sheet Sheet1 \
  --before F --count 2
./exceltool col delete --file book.xls --sheet Sheet1 \
  --cols F:H
```

复制 B～D 列并插入 F 列前：

```bash
./exceltool col copy --file book.xls --sheet Sheet1 \
  --cols B:D --insert-before F
```

`col copy` 复制值、公式、样式和列宽，并由 LibreOffice 按目标位置调整相对公式。

按内容自适应 A～F 列，并把每列最大宽度限制为 60 mm：

```bash
./exceltool col autofit --file book.xls --sheet Sheet1 \
  --cols A:F --max-width-mm 60
```

`--max-width-mm` 默认 60，必须是有限的正数。自适应是显式操作；`write`、
`row copy` 等命令不会隐式改变列宽，避免意外破坏既有版式。`.xls/.xlsx` 保存列宽
时会按各自格式的单位量化，因此重开后的宽度允许不超过 0.1 mm 的格式舍入误差。

`col autofit --json` 的 `changes.widths_mm` 保持原有简洁数组，同时
`changes.columns` 按列报告计算宽度和保存后重开得到的实际宽度：

```json
{
  "columns": [
    {
      "column": "A",
      "width_mm": 18.4,
      "actual_width_mm": 18.4
    }
  ]
}
```

Patch 中每个 `col.autofit` 操作的结果也包含 `changes.columns`，并增加
`final_column` 表示后续列结构操作完成后的最终列标。如果该列随后被删除，
`final_column` 和 `actual_width_mm` 均为 `null`。

## 修改

### 二维 JSON 批量写入

从 F 列第 3 行开始写入：

```bash
./exceltool write \
  --file book.xls --sheet Sheet1 --begin F3 \
  '[[1,"苹果"],[2,"香蕉"]]'
```

`F3` 拆分为 F 列、第 3 行；`AA3` 拆分为 AA 列、第 3 行：

```bash
./exceltool write \
  --file book.xls --sheet Sheet1 --begin AA3 \
  '[["v1","v2"],["v3","v4"]]' \
  --out result.xls
```

对应位置为 `AA3/AB3` 和 `AA4/AB4`。

### 数组语义

```json
[
  [100, "苹果", true],
  [],
  [200, null, "=A3*2"]
]
```

- 外层数组按连续行定位，内层数组按连续列定位。
- `[]` 跳过该行，不修改任何单元格。
- `null` 清空对应单元格内容并保留样式。
- 未提供的行尾单元格保持不变。
- `=...` 写为公式；`'=...` 写为以 `=` 开头的普通文本。
- JSON 公式文本原样写入，不自动平移引用；需要 Excel 相对引用调整时使用
  `row copy` 或 `col copy`。

### 三种数据输入

位置参数：

```bash
./exceltool write --file book.xls --sheet Sheet1 --begin A1 \
  '[[1,"苹果"],[2,"香蕉"]]'
```

UTF-8 JSON 文件：

```bash
./exceltool write --file book.xls --sheet Sheet1 --begin A1 \
  --values-file values.json
```

stdin 与 view 管道：

```bash
./exceltool view --file source.xls --sheet Sheet1 --range A1:C10 --json |
./exceltool write --file target.xls --sheet Data --begin F3 --stdin
```

三种输入必须且只能选择一种。

### 写入字体和字号

```bash
./exceltool write \
  --file book.xls --sheet Sheet1 --begin A1 \
  '[[1,"苹果"],[2,"香蕉"]]' \
  --font "微软雅黑" --font-size 12
```

字体和字号应用于 JSON 中实际出现的位置，包括值为 `null` 的位置；`[]` 跳过
的行不应用样式。`--font-size` 必须是有限的正数。编辑结果中的 `font_check`
记录请求字体与系统解析字体，`changes.actual_font` 记录保存后实际字体槽；验证
失败时错误会指出具体 sheet、单元格、请求值和实际值。

### 范围样式与清空

只修改字体和字号，不改变内容或其他样式：

```bash
./exceltool style --file book.xls --sheet Sheet1 --range A1:F10 \
  --font "微软雅黑" --font-size 12
```

清空内容并保留样式：

```bash
./exceltool clear --file book.xls --sheet Sheet1 --range B3:F10
```

同时移除该范围的硬样式：

```bash
./exceltool clear --file book.xls --sheet Sheet1 --range B3:F10 \
  --clear-style
```

范围参数也可以是单个单元格。单格写入继续使用统一的二维 JSON：

```bash
./exceltool write --file book.xls --sheet Sheet1 --begin B3 \
  '[["完成"]]'
```

写入数值、布尔值、公式并同时设置字体字号：

```bash
./exceltool write --file book.xls --sheet Sheet1 --begin B3 \
  '[[100,true,"=A1*2"]]' --font "微软雅黑" --font-size 12
```

清空单元格内容并保留样式：

```bash
./exceltool clear --file book.xls --sheet Sheet1 --range B3
```

单格样式也使用 `style --range B3`。工具不再维护另一套 `cell` 命令和类型参数。

### 单元格批注

写入单元格批注；目标已有批注时替换原文本：

```bash
./exceltool comment --file book.xls --sheet Sheet1 --cell B3 \
  --text "由财务复核"
```

批注文本不能为空。操作只修改批注，不改变单元格的值、公式或样式，并与其他编辑
命令一样在保存后重新打开验证。`.xls` 和 `.xlsx` 使用相同接口。

## 事务式 Patch

需要连续修改一个工作簿的多个位置或多个 sheet 时，可以将操作写入一个 UTF-8
JSON 文件。文件路径、另存和覆盖策略只放在命令行中，不写入 patch：

```bash
./exceltool patch --file book.xls --patch patch.json --json
./exceltool patch --file book.xls --patch patch.json \
  --out result.xls --json
```

第二条命令保留输入文件；`result.xls` 已存在时必须增加 `--overwrite`。
Patch JSON 继续只描述工作簿操作；文件路径、输出、覆盖和 SHA256 前置条件都放在
命令行。

### Patch v1 格式

```json
{
  "version": 1,
  "operations": [
    {
      "id": "insert-task",
      "op": "row.insert",
      "sheet": "任务",
      "before": 10,
      "count": 1
    },
    {
      "id": "write-task",
      "op": "write",
      "sheet": "任务",
      "begin": "A10",
      "values": [[null, null, 1001, "示例任务", "示例说明"]],
      "font": "Microsoft YaHei",
      "font_size": 10
    }
  ]
}
```

`id` 可选，用于成功结果和失败定位；`op` 以及相应操作的必填字段不可省略。
v1 对顶层和操作字段执行严格校验，未知字段会失败，以便尽早发现拼写错误。
顶层只允许 `version` 和 `operations`，因此不能在 JSON 中放入 `file`、`out`、
`overwrite` 或其他发布策略。

支持以下操作：

```json
{"op":"write","sheet":"任务","begin":"A10","values":[[1001,"任务一",true],[],[1003,null,"=A3*2","'=普通文本"]],"font":"Microsoft YaHei","font_size":10}
{"op":"style","sheet":"任务","range":"A10:K10","font":"Microsoft YaHei","font_size":10}
{"op":"clear","sheet":"任务","range":"A10:K10","with_style":false}
{"op":"comment","sheet":"任务","cell":"D10","text":"由负责人确认"}
{"op":"row.insert","sheet":"任务","before":10,"count":1}
{"op":"row.delete","sheet":"任务","rows":"10:12"}
{"op":"row.copy","sheet":"任务","rows":"3:3","insert_before":10}
{"op":"col.insert","sheet":"任务","before":"F","count":1}
{"op":"col.delete","sheet":"任务","cols":"F:H"}
{"op":"col.copy","sheet":"任务","cols":"B:D","insert_before":"F"}
{"op":"col.autofit","sheet":"任务","cols":"A:F","max_width_mm":60}
{"op":"sheet.add","name":"新配置"}
{"op":"sheet.delete","sheet":"旧配置"}
{"op":"sheet.rename","sheet":"旧名称","name":"新名称"}
{"op":"sheet.copy","sheet":"模板","name":"新配置"}
```

各操作沿用对应单命令语义：

- `write` 的 `[]` 跳过整行，`null` 清空内容并保留样式，未提供的行尾不修改；
  `=...` 是公式，`'=...` 是以等号开头的普通文本；`font`、`font_size` 可选。
- `style` 至少提供 `font` 或 `font_size`。
- `clear.with_style` 默认 `false`。
- `comment` 写入或替换一个单元格的非空批注文本。
- `row.insert.count` 默认 `1`。
- `row.copy` 复制值、公式、样式和行高，相对公式由 LibreOffice 调整。
- `col.insert.count` 默认 `1`。
- `col.copy` 复制值、公式、样式和列宽，相对公式由 LibreOffice 调整。
- `col.autofit.max_width_mm` 默认 `60`，自适应后应用最大宽度限制。
- Patch 的 `col.autofit` 成功项包含逐列 `width_mm`、`final_column` 和保存后
  `actual_width_mm`，便于调用者进行数值比较。

### 执行与发布规则

- `operations` 按数组顺序执行，坐标基于前面操作完成后的实时工作簿状态。
- 一个 patch 只能操作 `--file` 指定的一个工作簿，但可包含多个 sheet。
- 全部操作在临时副本的同一个 LibreOffice 会话中完成。
- 任一操作失败，临时副本被丢弃，不发布部分结果。
- 全部操作完成后只保存一次，并只重新打开一次。
- 重开后验证最终受影响的内容、公式、明确样式、行高、列宽和 sheet 结构。
- 所有最终公式在保存前建立快照，重开后逐个核对；公式缺失、变成常量或文本、
  公式文本改变时不发布结果。
- 不包含 sheet/行/列结构操作时，还会检查操作声明范围之外的公式在操作前后不变。
  结构操作允许 LibreOffice 按操作语义移动和调整公式，但仍执行保存前后守恒。
- 验证成功后才原子替换输入文件或发布 `--out`。
- 默认不创建 `.bak`；是否在执行前备份仍由调用者明确决定。

成功的 `--json` 输出：

```json
{
  "ok": true,
  "operation": "patch",
  "input": "/path/book.xls",
  "output": "/path/book.xls",
  "verified": true,
  "operations": [
    {"index": 0, "id": "insert-task", "op": "row.insert", "verified": true},
    {"index": 1, "id": "write-task", "op": "write", "verified": true}
  ],
  "summary": {
    "operation_count": 2,
    "sheets": ["任务"],
    "saved_once": true
  }
}
```

操作失败会在 stderr 中指出下标、可选 `id` 和操作名，并明确没有发布：

```json
{
  "ok": false,
  "error": "未找到 sheet: 任务",
  "code": 3,
  "operation": "patch",
  "published": false,
  "failed_operation": {
    "index": 1,
    "id": "write-task",
    "op": "write"
  }
}
```

## 输出与安全

编辑命令默认修改 `--file` 指定的原文件，不需要额外输出选项：

```bash
./exceltool write --file book.xls --sheet Sheet1 --begin F3 \
  '[[1,"苹果"]]'
```

需要确保输入文件仍是调用者检查过的版本时，先计算 SHA256，再将其作为编辑前置
条件：

```bash
sha256sum book.xls
./exceltool patch --file book.xls --patch patch.json \
  --expect-sha256 0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef \
  --json
```

`--expect-sha256` 支持所有编辑命令，必须是 64 位十六进制。哈希不匹配会在
LibreOffice 启动前返回退出码 3。就地修改时，工具还会在发布前重新计算原路径
SHA256；如果编辑期间文件被其他程序或 SVN 更新，则拒绝覆盖。使用 `--out` 时，
输出基于已经通过哈希校验并复制完成的输入快照。

工具不自动创建 `.bak`。是否备份由调用者在修改前明确决定，例如：

```bash
cp -- book.xls book.xls.bak
./exceltool write --file book.xls --sheet Sheet1 --begin F3 \
  '[[1,"苹果"]]'
```

需要保留原文件时，使用 `--out` 另存；目标已存在时还需明确 `--overwrite`：

```bash
--out result.xls
--out result.xls --overwrite
```

`--out` 只支持同格式另存：`.xls` 必须输出为 `.xls`，`.xlsx` 必须输出为
`.xlsx`；ExcelTool 不提供格式转换。

无论修改原文件还是另存，编辑都在临时副本中完成，保存后重新打开验证，成功后
才原子发布。普通编辑命令执行一个操作；`patch` 在一次保存中顺序执行多个操作。
宏不会执行；失败不会把半成品替换为结果。

## AI 与退出码

编辑命令加 `--json` 后输出结果对象，例如：

```json
{
  "ok": true,
  "operation": "write",
  "input_sha256": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
  "output": "/path/book.xls",
  "verified": true,
  "formula_verification": {
    "checked": true,
    "operation_guard": "exact_outside_declared",
    "before_operation_count": 20,
    "before_save_count": 20,
    "after_save_count": 20,
    "normalized_string_results": 2,
    "unexpected_changes": 0
  },
  "changes": {
    "sheet": "Sheet1",
    "begin": "F3",
    "rows": 2,
    "columns": 2,
    "cells": 4
  }
}
```

`operation_guard` 为 `exact_outside_declared` 时，已验证非结构操作声明范围外的
公式不变；为 `structural_operations` 时，结构操作允许调整公式坐标，但已验证
保存前与重开后的全部最终公式一致。报告只陈述实际检查的公式，不会把未执行的
全工作簿值、样式或合并单元格比较伪报为零差异。

保存 `.xls` 前，工具会依据 LibreOffice 的真实公式结果类型，原样重赋所有字符串
结果的普通公式，规避 BIFF 导出把字符串缓存写成数值 `0` 的问题；数值、日期、
布尔和错误结果公式不重赋。`normalized_string_results` 报告实际归一化数量。
字符串结果数组公式不能安全逐格重赋时，工具会拒绝保存且不发布临时结果。

公式验证失败使用退出码 6，错误包含 `formula_verification` 和
`unexpected_formula_changes`，逐项报告阶段、sheet、单元格、保存前公式及
保存后类型和值；临时结果不会发布。

错误写到 stderr：

- `0`：成功
- `2`：命令行参数错误
- `3`：文件、sheet、范围、JSON 数据或输出目标错误
- `4`：格式或环境能力不支持
- `5`：LibreOffice 编辑或保存失败
- `6`：写后验证失败

## 验证与限制

```bash
python3 -m compileall -q src tests
PYTHONPATH=src python3 -m unittest discover -s tests -v
```

当前支持 `.xls/.xlsx` 的查看和本文列出的实用编辑能力。首版不承诺宏、图表、
数据透视表、字符级富文本和跨工作簿结构复制；LibreOffice 保存可能重写文件的
内部二进制结构，因此保证逻辑内容和目标操作，而非字节级一致。

部分受限沙箱会禁止 LibreOffice 启动子进程或创建本地 UNO socket，并返回
`Operation not permitted`。这时需按宿主环境规则授权 `exceltool` 在允许启动
LibreOffice 的环境执行；不要绕过权限或把该错误当作工作簿损坏。

## 项目开发文档

稳定决策见 [项目范围](docs/PROJECT.md)、[架构](docs/ARCHITECTURE.md) 和
[开发规范](docs/DEVELOPMENT.md)。
