# ExcelTool 轻量 Daemon 设计

## 文档状态

本文是解决 Windows 和 Linux 上 LibreOffice 重复启动和关闭耗时的生效技术设计。
当前实现以本文与 [架构说明](ARCHITECTURE.md) 为准。

## 目标与边界

目标是在一段连续工作期间复用一个由 ExcelTool 专属管理的 LibreOffice 进程和 UNO
端点，避免每条命令重复承担进程启动、UNO 就绪、进程退出和临时 profile 清理成本。

必须保持：

- 现有命令、参数、stdout、stderr、JSON 和退出码；
- 相对路径相对于当前 CLI 进程工作目录解析；
- 当前用户级全局工作簿命令锁和串行执行语义；
- 每条命令结束后关闭工作簿，不长期锁定 `.xls/.xlsx` 文件；
- 编辑临时副本、输入 SHA256、防并发覆盖、保存后重开验证和原子发布；
- `font check`、帮助和版本不启动 daemon 或 LibreOffice；
- `--no-daemon` 保留当前一次性进程路径，供故障诊断和强隔离回归使用。

本方案不把工作簿命令转发给 daemon。daemon 不解析 ExcelTool 业务参数，不读取工作簿
数据，也不代理命令的标准输入输出。

Windows 10/11 和 Linux 都是正式支持平台，普通工作簿命令在两者上默认使用轻量 daemon；
平台差异只限于运行目录、后台进程创建、LibreOffice 启动和进程树托管。两端必须保持
相同的命令语义、输出、退出码、租约状态机和故障结果。

## 方案选择

比较三个能够复用 LibreOffice 的方向：

| 方向 | 正确性与成本 | 结论 |
| --- | --- | --- |
| 文件中记录 UNO 端点，CLI 直接连接 | 最轻，但没有进程所有者、使用租约和可靠异常清理，idle 关闭可能撞上正在执行的命令 | 不采用 |
| daemon 代理完整 ExcelTool 命令 | 可以集中认证、排队和清理，但需要转发 argv、cwd、stdin、stdout、stderr、取消和业务错误，形成第二套命令执行边界 | 不采用 |
| daemon 管理 LibreOffice 生命周期，CLI 直接连接 UNO | 保留现有 CLI 业务链，只新增发现、握手、租约和进程管理；全局锁继续负责串行化 | 采用 |

实例文件是发现线索，不是存活权威。daemon 控制握手是实例身份和可用性的权威判断；
长连接租约是 LibreOffice 当前是否可关闭的权威状态。

## 总体结构

```text
PowerShell / Agent
        │
        │ 原有 exceltool 命令
        ▼
ExcelTool CLI
  ├── argparse、路径、stdin/stdout/stderr、退出码
  ├── workbook_command_lock
  ├── daemon 发现与 acquire 租约
  └── 直接 UNO 连接和现有业务执行
        │                         │
        │ 控制连接（租约）        │ UNO
        ▼                         ▼
轻量 daemon ───────────────→ 专属 LibreOffice
  ├── 实例文件                    ├── 随机 UNO 端口
  ├── 单实例与启动协调锁          ├── 独立临时 profile
  ├── idle 状态机                 └── 平台进程树托管
  └── 进程、UNO 就绪和异常清理
```

daemon 只拥有 LibreOffice 进程、profile、管理用 UNO 连接和进程树清理责任。CLI 拥有
具体工作簿及命令结果；工作簿对象不得跨命令缓存。

## CLI 执行流程

工作簿命令按以下顺序执行：

1. CLI 完成 argparse 解析；参数错误直接按现有方式返回。
2. CLI 取得现有 `workbook_command_lock`。daemon 客户端不得在锁外执行 UNO 操作。
3. CLI 读取实例文件并执行控制握手；无有效实例时进入自动启动流程。
4. CLI 在一条控制连接上发送 `acquire`。
5. daemon 原子增加活动租约、取消 idle 退出，并确保 LibreOffice 和管理 UNO 连接就绪。
6. daemon 返回本代 LibreOffice 的 UNO endpoint；控制连接保持打开。
7. CLI 建立业务 UNO 连接，打开工作簿并执行现有读取或编辑流程。
8. CLI 在 `finally` 中关闭本命令创建的全部工作簿和业务 UNO 引用。
9. CLI 发送 `release`；daemon 检查没有残留工作簿后返回 `released`。
10. CLI 关闭控制连接并释放 `workbook_command_lock`。

现有全局锁继续提供不同 CLI 进程间的工作簿命令串行化。daemon 不建立业务 FIFO 队列，
也不在整个常驻生命周期持有该锁。`--no-daemon` 和旧客户端继续通过同一把锁与 daemon
模式串行，但操作系统锁不承诺跨不同等待者的严格 FIFO。

## 实例发现文件

### 路径

Windows：

```text
%LOCALAPPDATA%\ExcelTool\runtime\daemon.json
```

Linux 优先使用：

```text
$XDG_RUNTIME_DIR/exceltool/daemon.json
```

`XDG_RUNTIME_DIR` 缺失、不属于当前用户或权限不安全时，回退到：

```text
/tmp/exceltool-<uid>/runtime/daemon.json
```

Linux 回退目录沿用现有命令锁的用户归属与权限校验：目录必须属于当前 uid，且 group 和
other 不可写；目录权限收紧为 `0700`，实例文件、启动锁和生命周期锁为 `0600`。

runtime 目录和实例文件必须限制为当前用户访问。token 不得进入日志、错误文本或
`daemon status` 输出。

### 内容

```json
{
  "schema_version": 1,
  "protocol_version": 1,
  "exceltool_version": "0.1.0",
  "build_id": "当前代码构建标识",
  "installation_id": "规范化安装根目录摘要",
  "instance_id": "随机 UUID",
  "pid": 12345,
  "host": "127.0.0.1",
  "port": 49152,
  "token": "256 位随机令牌",
  "started_at": "2026-10-10T14:00:00+08:00"
}
```

字段语义：

| 字段 | 生产者与消费者 | 语义 |
| --- | --- | --- |
| `schema_version` | daemon 写、CLI 读 | 实例文件结构版本；缺失或不支持即视为无有效实例 |
| `protocol_version` | daemon 写、CLI 读并在握手复核 | 控制协议版本，不代表业务代码版本 |
| `exceltool_version` | daemon 写、CLI 读 | 人可读版本，用于诊断 |
| `build_id` | daemon 写、CLI 比较 | 防止升级后继续连接已加载旧源码的 daemon |
| `installation_id` | daemon 写、CLI 比较 | 防止另一个安装目录的客户端误用该实例 |
| `instance_id` | daemon 每次启动生成 | 区分 PID/端口复用和新旧实例，也是条件删除文件的依据 |
| `pid` | daemon 写 | 仅供诊断，不能证明实例存活，也不能作为直接杀进程依据 |
| `host`、`port` | daemon 监听成功后写 | 控制端点；host 首版必须是 `127.0.0.1` |
| `token` | daemon 每次启动生成 | 控制协议认证材料，使用 32 字节安全随机数编码 |
| `started_at` | daemon 写 | 带时区的启动时间，仅供状态展示 |

UNO endpoint 不写入实例文件。它只在成功 `acquire` 后返回，因为 LibreOffice 可以在同一
daemon 生命周期内因崩溃或污染而重启。

`build_id` 由客户端和 daemon 调用同一个标准库实现计算：按规范化相对路径排序，将当前
安装中 `exceltool` Python 包全部运行时 `.py` 文件的“相对路径、NUL
分隔符、文件字节”依次输入 SHA256。`__pycache__`、测试、文档、Git 状态和修改时间不参与
计算。文件在计算期间变化时本次握手或启动失败，不使用可能混合两个版本的摘要。该规则
同时适用于源码安装和复制安装，`installation_id` 另行区分规范化安装根目录。

### 发布与清理

daemon 先绑定控制 socket 的端口 `0` 并开始监听，再将内容写入同目录临时文件，flush
后使用 `os.replace()` 原子发布。不能使用“先探测空闲端口、关闭 socket、再监听”的
方式。

退出时 daemon 重新读取实例文件，只有文件中的 `instance_id` 仍等于自身实例时才删除，
防止旧进程删除新实例的文件。损坏、缺失或握手失败的文件都只视为失效线索；PID 不用于
替代握手。
运行中的 daemon 会检查实例文件身份；文件丢失、损坏或被陈旧内容替换时，持有
生命周期锁的 daemon 重新原子发布自己的记录，使客户端能恢复发现。

实例文件是短生命周期运行状态，不做迁移、历史保留或恢复。新 daemon 获得单实例锁后
可以原子覆盖失效文件。

## 自动启动与单实例

CLI 握手失败后的启动流程：

1. 获取当前用户的 `daemon-start.lock`。
2. 获取后重新读取实例文件并握手，避免并发客户端重复启动。
3. 仍无有效实例时，以无窗口、脱离调用控制台的用户进程启动 daemon。
4. 等待 daemon 原子发布实例文件并完成合法 `ping/pong`。
5. 释放启动锁，再执行 `acquire`。

daemon 自身还要在整个生命周期持有独立的 `daemon-lifetime.lock`。新 daemon 无法取得
该锁时不得发布实例文件。启动锁解决并发创建，生命周期锁解决旧 daemon 正在退出、状态
文件损坏或手工启动造成的双实例。

等待启动必须有明确超时；失败返回 daemon 启动错误，不自动退回 `--no-daemon`，避免
静默改变隔离和清理语义。用户可显式使用 `--no-daemon` 诊断。

## 控制协议

控制协议只传递小型管理消息，不传 argv、工作簿数据或命令输出。使用仅监听
`127.0.0.1` 的随机 TCP 端口；每帧为 4 字节大端长度加 UTF-8 JSON，首版单帧上限
64 KiB。超长、非法 UTF-8、非法 JSON、未知消息或字段类型错误都关闭连接并返回协议
错误（能够安全返回时）。

所有请求都包含：

```json
{
  "type": "ping",
  "protocol_version": 1,
  "instance_id": "...",
  "token": "..."
}
```

daemon 使用常量时间比较 token，并复核 `instance_id`。token 或实例错误不得泄露合法值。

### `ping`

CLI 在读取实例文件后发送；成功响应必须回显协议、实例和 build 身份：

```json
{
  "type": "pong",
  "protocol_version": 1,
  "instance_id": "...",
  "build_id": "...",
  "installation_id": "...",
  "state": "idle"
}
```

文件内容、连接目标和响应身份全部匹配才是有效 daemon。

### `acquire`、`release` 与连接租约

CLI 在同一控制连接上发送：

```json
{
  "type": "acquire",
  "protocol_version": 1,
  "instance_id": "...",
  "token": "...",
  "client_pid": 45678
}
```

成功响应：

```json
{
  "type": "acquired",
  "instance_id": "...",
  "generation": 3,
  "uno": {
    "transport": "socket",
    "host": "127.0.0.1",
    "port": 49321
  }
}
```

`generation` 从 1 开始，每次 daemon 重建 LibreOffice 都递增。CLI 只使用本次响应里的
endpoint，不缓存到后续命令。

控制连接从 `acquired` 持续到命令结束，代表一个活动租约。正常完成时 CLI 先关闭所有
工作簿，再发送：

```json
{
  "type": "release",
  "instance_id": "...",
  "token": "...",
  "generation": 3,
  "client_cleanup": "clean"
}
```

`client_cleanup` 必填，只允许 `clean` 或 `failed`。CLI 的全部工作簿和业务 UNO 引用均
成功清理时发送 `clean`；任一清理失败时发送 `failed`，daemon 不再尝试复用本代 session。
daemon 复核 generation；对于 `clean`，还要通过管理 UNO 连接检查没有残留工作簿。随后
响应：

```json
{"type":"released","generation":3,"reusable":true}
```

`reusable=false` 表示 daemon 已接受释放但将废弃本代 session。只有收到 `released` 后，
CLI 才认为租约已由 daemon 接管清理；它不能把 `reusable=false` 改写成业务命令失败，
但状态和诊断日志应记录 session 被废弃。

以下情况都将该代 session 标为 `DIRTY`，不得交给下一条命令：

- 租约连接在合法 `release/released` 前断开；
- CLI 报告工作簿或 UNO 清理失败；
- daemon 检查到残留工作簿；
- LibreOffice 进程或管理 UNO 连接异常；
- 请求中的 generation 与当前实例不一致。

`DIRTY` session 由 daemon 结束进程树、清理 profile 并回到 `STOPPED`；下一次
`acquire` 创建新 generation。daemon 不尝试猜测或复用状态不明的 UNO 对象。

### `status` 与 `stop`

`status` 不启动 LibreOffice，返回 daemon、租约和 LibreOffice 的汇总状态，不返回
token、UNO endpoint、工作簿路径或命令内容。

```json
{
  "type": "status_result",
  "instance_id": "...",
  "daemon_state": "idle",
  "uptime_seconds": 420,
  "idle_seconds": 38,
  "idle_timeout_seconds": 300,
  "active_lease": false,
  "libreoffice": {
    "state": "ready",
    "generation": 3,
    "pid": 23456,
    "requests": 7
  }
}
```

时间字段是非负整数秒；LibreOffice 尚未启动时 `state=stopped`、`generation=0`、`pid=null`。
`daemon_state` 只允许 `idle`、`starting`、`leased`、`stopping`。

`stop` 将 daemon 置为 `STOPPING` 并拒绝新租约；已有租约正常完成后再关闭
LibreOffice。普通停止不强杀正在执行的命令。idle 退出与手工 stop 使用同一清理路径。

`stop` 成功受理后返回 `{"type":"stopping","instance_id":"..."}`；客户端随后等待控制
实例握手失效和实例文件撤下，但不根据 PID 强杀进程。已经处于 `STOPPING` 时重复 stop
返回相同结果。

控制层失败统一返回：

```json
{
  "type": "error",
  "code": "incompatible_build",
  "message": "可供用户诊断且不含秘密的说明"
}
```

首版错误码限定为 `bad_request`、`unauthorized`、`incompatible_protocol`、
`incompatible_build`、`busy`、`starting_failed`、`stopping` 和 `internal_error`。认证失败
只返回统一 `unauthorized`，不说明 token、instance_id 或端口哪一项错误。控制错误由 CLI
转换成现有业务命令错误输出；不占用现有工作簿内容和验证错误码的新含义。

协议不提供命令排队、业务取消和任意函数调用。业务中断继续由 CLI 处理；CLI 异常退出由
租约断开触发 session 废弃。

## LibreOffice 生命周期

```text
STOPPED
   │ acquire
   ▼
STARTING ──失败──→ STOPPED（acquire 返回错误）
   │ UNO ready
   ▼
READY
   │ 租约建立
   ▼
LEASED
   │ 干净 release
   ▼
READY
   │ 空闲 300 秒
   ▼
STOPPING → EXITED（daemon 同时退出）

LEASED ──异常断连/残留/进程异常──→ DIRTY → STOPPING → STOPPED
```

规则：

- daemon 启动时不立即启动 LibreOffice，第一次 `acquire` 才延迟启动；
- 使用独立临时 profile、随机 UNO 端口和现有打印机枚举抑制环境变量；
- 同一时刻最多一个活动租约；正常 CLI 已由全局锁保证这一点，daemon 仍拒绝第二租约；
- idle 计时从干净 `release` 或无租约的启动完成后开始；
- 活动租约期间绝不 idle 关闭；
- 默认 idle timeout 为 300 秒，首版不增加永久配置；
- 正常退出先通过管理 UNO 连接调用 `desktop.terminate()`，再有限等待；
- 超时后只结束 daemon 自己创建并托管的进程树；
- profile 清理完成后 daemon 才完成退出。

idle 判断、租约变化和 `STOPPING` 转换必须在同一个状态锁下完成，避免计时器与新
`acquire` 交叉。

## 平台进程所有权

daemon 只能结束自己创建的 LibreOffice 进程树，不能按进程名扫描、连接或结束用户手工
启动的 LibreOffice。Windows 与 Linux 使用不同托管机制，但必须满足相同的正常退出、
daemon 崩溃清理和所有权验收。

### Windows

daemon 启动的 LibreOffice 必须加入专属 Windows Job Object，并设置
`JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`。Job handle 由 daemon 独占且不向子进程继承。
为消除 `soffice.com` 在加入 Job 前派生子进程的窗口，Windows 启动层使用标准库
`ctypes` 调用 `CreateProcessW`，组合 `CREATE_SUSPENDED`、`CREATE_NO_WINDOW` 和
`CREATE_UNICODE_ENVIRONMENT` 创建主进程；`AssignProcessToJobObject` 成功后才
`ResumeThread`。任一步失败都终止尚未恢复的主进程并关闭句柄，本代启动失败。

实现必须验证 `soffice.com` 及其派生的 `soffice.bin` 都属于该 Job。加入 Job 失败时
LibreOffice 启动整体失败，不允许静默退化为按进程名扫描或结束用户进程。还必须验证
daemon 位于宿主 Job Object 时的嵌套行为，以及挂起创建、加入 Job、恢复和失败回收的
完整路径。

daemon 崩溃时 Job handle 关闭，由 Windows 结束专属 LibreOffice 进程树。daemon 永远不
根据 `soffice.exe` 名称结束系统中其他 LibreOffice 实例。

### Linux

仅使用 `start_new_session=True` 或单独进程组不能覆盖 daemon 被强杀后的清理，因为
LibreOffice 可能成为孤儿进程。Linux 使用一个最小进程守护器管理专属进程组：

1. daemon 创建只用于存活监测的匿名 pipe，并启动守护器进程；
2. 守护器以新 session/process group 启动 `soffice`，只管理该组；
3. daemon 生命周期内保持 pipe 写端打开；正常关闭时请求守护器优雅结束 LibreOffice；
4. daemon 崩溃或被强杀时写端由内核关闭，守护器从 EOF 得知所有者消失；
5. 守护器先向专属进程组发送 `SIGTERM` 并有限等待，超时后发送 `SIGKILL`，确认进程组
   消失后自身退出。

守护器不接受工作簿命令，不保存实例状态，也不按名称查找进程。守护器启动失败、pipe
无法建立或 LibreOffice 无法进入专属进程组时，当前 generation 启动失败，不静默退化成
无托管子进程。正常退出仍优先由 daemon 的管理 UNO 连接调用 `desktop.terminate()`；
守护器承担超时升级和 daemon 异常死亡后的兜底清理。

## CLI 与编辑流程改造

现有 `LibreOfficeSession` 拆分为两个职责：

- 进程拥有者：供 daemon 和 `--no-daemon` 使用，负责 profile、进程、就绪、终止和清理；
- UNO 连接：供 CLI 使用，连接既有 endpoint，负责 `load/create`，但不拥有 LibreOffice
  进程。

`view`、`find`、sheet 读取和 `editing.py` 接收外部 UNO 连接，不再在业务函数内部固定
创建 LibreOffice 进程。CLI 保持现有参数、路径和输出逻辑，因此无需跨进程传递 cwd、
stdin、stdout 或 stderr。

daemon 模式下的编辑流程为：

```text
取得租约并直连 UNO
→ 在临时副本上修改和保存
→ 关闭工作簿
→ 在同一 LibreOffice generation 中只读重开临时副本
→ 执行公式和目标验证
→ 关闭工作簿
→ 原子发布
→ 干净 release
```

这保持“保存到磁盘、关闭、重开、验证后发布”的逻辑保证，但不等同于当前实现使用第二个
LibreOffice 进程和新 profile 的冷重开隔离。文档和成功结果不得把两者描述为同一强度。
`--no-daemon` 保留当前冷重开路径，作为故障诊断与回归基准。

任何清理失败都发生在原子发布之前时，临时结果不得发布；已经完成原子发布后发生的
租约连接丢失不能回滚，只将 session 标为 `DIRTY`，其边界与当前发布后输出丢失一致。

## 故障行为

| 故障 | 行为 |
| --- | --- |
| 实例文件缺失、损坏、端口不可达或握手错误 | 取得启动锁后二次检查，仍失效则启动 daemon |
| 协议、build 或 installation 不匹配 | 不执行命令；请求旧实例安全 stop，等待退出后启动当前实例 |
| daemon 启动或 LibreOffice 就绪超时 | 返回明确错误；不静默切到直连模式 |
| LibreOffice 在命令中退出 | 当前命令沿用 LibreOffice 错误码 5；session 废弃，下一条 acquire 重建 |
| CLI 正常中断且完成 finally 清理 | 发送 release；清理成功则可复用，否则 session 废弃 |
| CLI 崩溃、强杀或控制连接消失 | 租约异常结束，session 标记 DIRTY 并重启 LibreOffice |
| daemon 崩溃 | Windows Job Object 或 Linux 进程守护器清理 LibreOffice；当前 CLI 收到控制或 UNO 连接错误 |
| 工作簿关闭后仍被占用或存在残留组件 | release 失败，session 废弃；不能将污染状态交给下一命令 |
| stop 时存在活动租约 | 拒绝新 acquire，等待现有租约结束后退出 |

修改命令不因 LibreOffice 或 daemon 失败自动重试，避免重复写入或不明确发布。

## 安全边界

- 控制端口只绑定 `127.0.0.1`，使用每实例 256 位 token 和常量时间比较；
- 实例文件由当前用户私有目录保护，并以原子替换发布；
- daemon 控制协议不接受 Python 代码、模块名、业务 argv 或任意函数调用；
- daemon 不记录 token、完整工作簿内容、stdin、UNO endpoint 或工作簿路径；
- CLI 业务仍经过现有 argparse、格式、SHA256、覆盖和发布校验；
- 直接 UNO socket 不复用 daemon token，因此随机端口加私有实例文件不能被描述成能够
  抵抗同机恶意进程的强认证边界。若未来要求该边界，必须重新评估由 daemon 代理业务，
  不在本方案中虚构 UNO 鉴权。

## 管理入口

保留以下管理命令：

```text
exceltool daemon start
exceltool daemon status
exceltool daemon stop
exceltool --no-daemon <workbook-command> ...
```

普通工作簿命令自动发现和启动 daemon，不要求用户预热。`start` 只确保控制 daemon 已
就绪；LibreOffice 仍延迟到首次 `acquire`。`status` 和 `stop` 不取得工作簿命令锁，
但 stop 必须尊重现有租约。`--no-daemon` 工作簿命令继续取得现有全局锁。

## 实施顺序

1. 为 Windows 和 Linux 路径增加分阶段耗时观测，固定启动、UNO ready、load、文档关闭、
   进程退出和 profile 清理基线。
2. 拆分 LibreOffice 进程所有权与 UNO 连接，让读取和编辑入口接受外部连接；保持现有
   直连回归全部通过。
3. 实现实例文件、原子发布、启动锁、生命周期锁和 `ping` 身份校验。
4. 实现 daemon 状态机、管理 UNO 连接、`acquire/release` 长连接租约和 idle 退出。
5. 先将 `view/find/sheet list/sheet info` 路由到 daemon，验证复用、解锁和异常 session
   重建。
6. 重构编辑命令使用租约连接，完成同 generation 重开验证与 `--no-daemon` 冷重开回归。
7. 实现并验证 Windows Job Object、Linux 进程守护器以及各平台后台 daemon 启动。
8. 更新 README、架构、开发规范及 Windows/Linux 安装和故障处理文档，再在两端默认
   启用。

## 验证与验收

必须覆盖：

1. 无实例文件时第一条 `view` 自动启动 daemon 和 LibreOffice，并返回正确结果。
2. 实例文件存在且握手成功时复用；文件损坏、错误 token、错误实例、PID/端口复用均不
   会误认。
3. 两个客户端同时启动时只产生一个 daemon；获取锁后均能按现有全局锁串行完成。
4. 连续读取的 daemon PID、LibreOffice PID 和 generation 保持不变，第二次读取不承担
   LibreOffice 启动和关闭成本。
5. 每条命令 release 后可在对应平台独占打开、重命名或删除目标工作簿，证明没有长期
   文件锁。
6. 活动租约超过 300 秒时不触发 idle 退出；干净 release 后满 300 秒，daemon、
   `soffice.com` 和 `soffice.bin` 全部退出。
7. CLI 在执行中被强杀时控制连接断开，当前 generation 被废弃，下一条命令使用新
   generation，且没有残留工作簿锁。
8. daemon 被强杀时，Windows Job Object 和 Linux 进程守护器都能清理完整 LibreOffice
   进程树；不影响用户手工启动的 LibreOffice。
9. LibreOffice 异常退出时当前命令返回代码 5，编辑不自动重试，下一条命令能够重建。
10. `.xls/.xlsx` 的读取、写入、Patch、保存重开验证、公式保护、SHA256 和原子发布在
    daemon 与 `--no-daemon` 两种路径均通过。
11. 保存后同 generation 重开确实从磁盘得到最终内容；其结果与 `--no-daemon` 冷重开
    回归一致。
12. 中文路径、中文 sheet、微软雅黑、UTF-8 JSON 和 PowerShell 调用保持正常。
13. 帮助、版本和 `font check` 不启动 daemon；相对路径和 stdin 管道行为不变。
14. 升级 build 后客户端不继续使用旧 daemon，旧实例安全退出后由当前代码接管。
15. Windows 和 Linux 真实机器分别记录首条和后续 `A1:C5 view` 各阶段耗时，证明后续
    命令已消除重复 LibreOffice 启停成本；不以单元测试计时替代真实测量。Linux 由当前
    开发环境验证，Windows 使用同一验收脚本由目标 Windows 环境执行并保存输出。

测试通过注入短 idle timeout 验证计时器，不在常规测试中实际等待五分钟。

## 排除项

- 不缓存工作簿、sheet、单元格数据或 UNO 工作簿对象；
- 不并行执行多个工作簿命令；
- 不由 daemon 代理业务 argv、stdin、stdout 或 stderr；
- 不连接或终止用户手工启动的 LibreOffice；
- 不新增第三方 Python 依赖；
- 不自动重试修改命令；
- 不把随机 UNO 端口宣传为强认证机制；
- 首版不增加永久 daemon 配置文件，idle timeout 固定为 300 秒。

## 未决问题

无。
