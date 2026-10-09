# 开发规范

## 环境与边界

- Python 3.8+；保持标准库实现。Linux 使用系统 Python/UNO，Windows 使用 LibreOffice
  自带 Python/UNO。
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
- 批注操作只接受单个单元格和非空文本；必须验证保存后的批注文本及原单元格内容。
- 字体写入前必须在 Linux 精确解析 fontconfig、在 Windows 精确匹配系统字体注册表；
  文本样式使用整段 text cursor，验证错误必须
  按内容脚本检查 Western/Asian/Complex 相关槽，并包含具体 sheet、单元格、请求值
  和保存后的三槽实际值。
- 一次命令只描述一个可审计操作，不使用跨进程剪贴板状态。
- 除 `--help`、`--version` 和 `font check` 外，所有 CLI 命令必须在输入校验和
  LibreOffice 启动前取得当前用户的全局排他锁，并持有到命令返回与资源清理结束。
  竞争时只提示一次并等待，不对工作簿命令启动重试；Windows 可以为适配 `msvcrt`
  字节锁在锁实现内部做短间隔等待。进程退出依赖操作系统释放锁。
- 列宽只在显式 `col autofit` 中自适应；最大宽度按 1/100 mm 计算，往返验证允许
  `.xls/.xlsx` 最多 0.1 mm 的格式量化误差。
- 列宽机器输出必须区分操作时 `width_mm` 和重开后 `actual_width_mm`；Patch 后续
  结构操作必须同步维护 `final_column`，被删除时返回 `null`。
- `find` 是只读操作，默认限制结果数量；正则模式必须在启动 LibreOffice 前完成
  语法校验，并与普通查找共用范围能力；新增查找模式必须保持稳定 JSON 字段。
- `find` 与 `view` 必须复用范围参数定义、校验和 `resolve_window` 解析；双格式测试
  必须覆盖按列、行列组合、末尾行、正方区域和冲突参数。
- `view` 一次只接受一个工作簿；不得恢复目录扫描、多文件包装或静默覆盖重复
  `--file`。单格写入、清空和样式必须复用 write/clear/style，不恢复 cell 命令。
- patch 是单工作簿事务边界；协议严格拒绝未知字段和 JSON 内的文件/发布策略，
  操作按实时工作簿坐标执行，失败必须返回操作下标、可选 id 且不发布。
- patch 最终验证比较稳定逻辑语义，不依赖 LibreOffice 会在保存时规范化的内部
  属性状态；`clear.with_style=true` 仍须独立验证硬样式均已清除。
- 全部编辑必须通过统一 safety 生命周期采集 original、planned、reopened 公式
  快照；保存边界不得只验证目标范围或抽样哨兵。
- `.xls` 保存前必须使用 `FormulaResultType2` 精确筛选并原样重赋字符串结果公式，
  归一化前后必须再次执行全工作簿公式快照校验；不得根据显示字符串猜测结果类型。
- 非结构内容操作只豁免明确 write/clear 单元格；style、comment 和 col.autofit 不得豁免
  公式。结构操作必须至少保证 planned 到 reopened 的全工作簿公式一致。
- `--expect-sha256` 必须在 LibreOffice 启动前校验；临时副本必须匹配初始输入，
  就地发布前必须再次核对源路径，失败不得发布。
- 公式失败 JSON 必须指出阶段、sheet、单元格、原公式和保存后类型/值；成功报告
  只能声明实际执行的检查，不为值、样式或合并单元格生成虚假零差异。
- 普通错误不能向用户输出 Python traceback。

## 检查

```bash
python3 -m compileall -q src tests
PYTHONPATH=src python3 -m unittest discover -s tests -v
PYTHONPATH=src python3 -m exceltool --help
```

`Makefile` 默认安装到 `/usr/local/bin/exceltool` 和 `/usr/local/lib/exceltool`；
开发验证可通过临时 `PREFIX` 执行安装、入口和卸载往返，避免写入系统目录。

集成测试必须覆盖 `.xls/.xlsx`、纯 JSON 与 json-full、三种 write 输入、空行与
null、公式和字面 `=` 文本、单格/范围样式与清空、样式查看、字体预检、sheet 信息、
sheet、行、列、查找、正则及其共用范围、常见失败和未修改哨兵。Patch 测试还
必须覆盖全部 v1 操作、多 sheet、顺序坐标、单次保存结果、双格式写后验证、列宽
上限、失败定位和原文件不变。
公式安全测试还必须覆盖双格式未声明公式守恒、目标公式允许替换、结构 Patch 保存
边界、公式转常量差异定位、错误/正确 SHA256 和编辑期间源文件变化拒绝。
LibreOffice 进程与临时用户配置必须由测试清理。
命令锁测试必须覆盖同文件并发写入、等待后读取、不同文件共享全局锁、持锁进程异常
退出、快速连续调用和免锁命令；原有外部修改 SHA 防护仍须保持有效。
Windows 还必须覆盖 `msvcrt` 锁、`soffice.com`/UNO accept 参数、打印机环境变量、字体
注册表精确匹配、UTF-8 启动器和短暂 profile 占用清理。

未经用户明确要求，不 commit、push、创建远端或发布包。
