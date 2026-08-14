# 开发规范

## 环境与边界

- Python 3.8+；保持标准库实现，UNO 由系统 LibreOffice 提供。
- 正式代码在 `src/exceltool/`，测试在 `tests/`。
- 测试只修改临时目录中的合成工作簿。
- `.zhang-dev/` 是外层 Git 忽略的本机工作流目录。

## 实现规则

- 不把 UNO 对象泄露到 CLI JSON。
- 类型由命令显式指定，不从显示文本猜测原始值。
- view/write 使用同一 JSON 编解码语义；新增值类型必须同时补读写往返测试。
- 样式操作只修改明确指定的属性。
- 字体写入前必须精确解析 fontconfig；文本样式使用整段 text cursor，验证错误必须
  按内容脚本检查 Western/Asian/Complex 相关槽，并包含具体 sheet、单元格、请求值
  和保存后的三槽实际值。
- 一次命令只描述一个可审计操作，不使用跨进程剪贴板状态。
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
null、公式和字面 `=` 文本、范围样式/清空、样式查看、字体预检、sheet、行、
常见失败和未修改哨兵。Patch 测试还必须覆盖全部 v1 操作、多 sheet、顺序坐标、
单次保存结果、双格式写后验证、失败定位和原文件不变。
LibreOffice 进程与临时用户配置必须由测试清理。

未经用户明确要求，不 commit、push、创建远端或发布包。
