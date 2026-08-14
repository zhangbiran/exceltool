# AGENTS.md instructions

## 开始工作前

依次阅读：

1. `README.md`
2. `docs/PROJECT.md`
3. `docs/ARCHITECTURE.md`
4. `docs/DEVELOPMENT.md`

个人使用 Zhang Dev 时，再显式阅读本地 `.zhang-dev/AGENTS.md`；它不能覆盖本文件。

同时检查当前工作区和最近提交，保留用户已有改动。

## 工程标准

- 使用生产标准实现，不提交演示式简化或静默降级。
- 修改关键设计、协议、数据模型、权限、配置或故障语义时同步更新 `docs/`。
- 实现前明确需求覆盖、影响范围和验收方式。
- 运行与风险相称的测试、静态检查和构建；记录真实验证结果。
- 不在仓库、日志、错误或测试数据中写入秘密。
- 未经用户要求，不提交、推送、发布或操作生产环境。

## 目录边界

- `README.md` 是使用说明。
- `docs/` 是稳定项目知识库。
- `.zhang-dev/` 是外层仓库忽略的 AI 工作区，不得进入业务仓库提交；经用户明确选择时可作为独立私人 Git 仓库管理稳定资产。
- `.zhang-dev/01-docs/` 保存只读原始输入，`02-work/` 按原文档路径保存分析、计划、审查、测试和 `step.txt`。
- `.zhang-dev/context/` 保存项目上下文，`guides/workflow.md` 保存 Step 和门禁，其他 `guides/<name>.md` 保存工作流所用专项 Skill 的项目规范。
- `.zhang-dev/03-mock/` 和 `tools/` 保存本地可复用验证资产与辅助工具，`tmp/` 只放临时文件。
