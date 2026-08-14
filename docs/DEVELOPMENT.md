# 开发规范

## 环境与边界

- Python 3.8+；保持标准库实现，UNO 由系统 LibreOffice 提供。
- 正式代码在 `src/exceltool/`，测试在 `tests/`。
- 测试只修改临时目录中的合成工作簿。
- `.zhang-dev/` 是外层 Git 忽略的本机工作流目录。

## 实现规则

- 不把 UNO 对象泄露到 CLI JSON。
- 写入类型来自 JSON 类型和明确的公式/字面等号编码，不从显示文本猜测原始值。
- view/write 使用同一 JSON 编解码语义；新增值类型必须同时补读写往返测试。
- `view --tail` 必须按 sheet 的有效区域独立计算，只允许与列选择组合，并覆盖
  `.xls/.xlsx`、数量超过有效行数、非正数及范围冲突测试。
- `sheet info` 必须只读复用有效区域语义，覆盖全部 sheet、单个 sheet、缺失 sheet
  以及 `.xls/.xlsx` 的稳定 JSON 字段。
- 样式操作只修改明确指定的属性。
- 字体写入前必须精确解析 fontconfig；文本样式使用整段 text cursor，验证错误必须
  按内容脚本检查 Western/Asian/Complex 相关槽，并包含具体 sheet、单元格、请求值
  和保存后的三槽实际值。
- 一次命令只描述一个可审计操作，不使用跨进程剪贴板状态。
- 列宽只在显式 `col autofit` 中自适应；最大宽度按 1/100 mm 计算，往返验证允许
  `.xls/.xlsx` 最多 0.1 mm 的格式量化误差。
- 列宽机器输出必须区分操作时 `width_mm` 和重开后 `actual_width_mm`；Patch 后续
  结构操作必须同步维护 `final_column`，被删除时返回 `null`。
- `find` 是只读操作，默认限制结果数量；新增查找模式必须保持稳定 JSON 字段。
- `find` 与 `view` 必须复用范围参数定义、校验和 `resolve_window` 解析；双格式测试
  必须覆盖按列、行列组合、末尾行、正方区域和冲突参数。
- `view` 一次只接受一个工作簿；不得恢复目录扫描、多文件包装或静默覆盖重复
  `--file`。单格写入、清空和样式必须复用 write/clear/style，不恢复 cell 命令。
- patch 是单工作簿事务边界；协议严格拒绝未知字段和 JSON 内的文件/发布策略，
  操作按实时工作簿坐标执行，失败必须返回操作下标、可选 id 且不发布。
- patch 最终验证比较稳定逻辑语义，不依赖 LibreOffice 会在保存时规范化的内部
  属性状态；`clear.with_style=true` 仍须独立验证硬样式均已清除。
- 普通错误不能向用户输出 Python traceback。

## 检查

```bash
python3 -m compileall -q src tests
PYTHONPATH=src python3 -m unittest discover -s tests -v
PYTHONPATH=src python3 -m exceltool --help
```

`Makefile` 默认安装到 `/usr/local/bin/exceltool` 和 `/usr/local/lib/exceltool`；
开发验证可通过临时 `PREFIX` 执行安装、入口和卸载往返，避免写入系统目录。
README 是 Excel Skill 命令参考的正本；通过 `make sync-skill-docs` 生成插件副本，
通过 `make check-skill-docs` 阻止文档漂移。两个目标都必须显式传入插件 Skill
目录，不在仓库中固化本机路径。

集成测试必须覆盖 `.xls/.xlsx`、纯 JSON 与 json-full、三种 write 输入、空行与
null、公式和字面 `=` 文本、单格/范围样式与清空、样式查看、字体预检、sheet 信息、sheet、行、列、
查找、常见失败和未修改哨兵。Patch 测试还必须覆盖全部 v1 操作、多 sheet、顺序
坐标、单次保存结果、双格式写后验证、列宽上限、失败定位和原文件不变。
LibreOffice 进程与临时用户配置必须由测试清理。

未经用户明确要求，不 commit、push、创建远端或发布包。
