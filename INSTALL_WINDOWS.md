# ExcelTool Windows 安装说明

本文说明 Windows 10/11 原生环境的 ExcelTool 安装、升级、卸载、中文编码、字体和
LibreOffice 故障处理，不需要 WSL。命令使用方式见 [README](README.md)，Linux 安装见
[Linux 安装说明](INSTALL_LINUX.md)。

## 运行依赖

- Git；
- 64 位 LibreOffice Calc；
- LibreOffice 自带 Python 和 PyUNO；
- Windows 10 或 Windows 11。

Windows 使用 LibreOffice 自带 Python，不需要另装 `uno`，也不要使用
`pip install uno` 代替 PyUNO。默认安装目录为：

```text
C:\Program Files\LibreOffice
```

在 PowerShell 中验证 LibreOffice 和 UNO：

```powershell
& "$env:ProgramFiles\LibreOffice\program\soffice.com" --version
& "$env:ProgramFiles\LibreOffice\program\python.exe" -c "import uno; print('UNO OK')"
```

## 用户级安装

默认 Windows 启动器使用以下布局：

```text
%USERPROFILE%\.local\src\exceltool
%USERPROFILE%\.local\bin\exceltool.cmd
```

在 PowerShell 中克隆源码并复制启动器：

```powershell
$sourceRoot = Join-Path $env:USERPROFILE '.local\src'
$excelToolRoot = Join-Path $sourceRoot 'exceltool'
$binRoot = Join-Path $env:USERPROFILE '.local\bin'

New-Item -ItemType Directory -Force -Path $sourceRoot, $binRoot | Out-Null
git clone https://github.com/zhangbiran/exceltool.git $excelToolRoot
Copy-Item -LiteralPath (Join-Path $excelToolRoot 'exceltool.cmd') -Destination $binRoot
```

将命令目录加入当前用户 `PATH`，并避免重复添加：

```powershell
$binRoot = Join-Path $env:USERPROFILE '.local\bin'
$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$entries = @($userPath -split ';' | Where-Object { $_ })
if ($entries -notcontains $binRoot) {
    [Environment]::SetEnvironmentVariable(
        'Path',
        (($entries + $binRoot) -join ';'),
        'User'
    )
}
```

重新打开终端后验证：

```powershell
Get-Command exceltool
exceltool --version
exceltool --help
```

如果源码或 LibreOffice 不在默认目录，复制启动器后修改其中的 `EXCELTOOL_SOURCE` 或
`LIBREOFFICE_PROGRAM`。启动器固定使用 LibreOffice 自带 Python，并只为 ExcelTool
子进程设置：

```text
PYTHONUTF8=1
PYTHONIOENCODING=utf-8
SAL_DISABLE_SYNCHRONOUS_PRINTER_DETECTION=1
SAL_DISABLE_PRINTERLIST=1
SAL_DISABLE_DEFAULTPRINTER=1
```

这些变量不会永久修改其他程序的环境，也不会改变系统默认打印机。

## 升级与卸载

升级源码：

```powershell
$excelToolRoot = Join-Path $env:USERPROFILE '.local\src\exceltool'
git -C $excelToolRoot pull --ff-only
Copy-Item -LiteralPath (Join-Path $excelToolRoot 'exceltool.cmd') `
    -Destination (Join-Path $env:USERPROFILE '.local\bin') -Force
exceltool --version
```

升级前先检查源码工作区是否有本地改动，不要强制覆盖。卸载只删除用户级源码目录、
启动器，并按需从用户 `PATH` 移除 `.local\bin`；不要删除仍被其他程序使用的
LibreOffice 或字体。

## 中文乱码

启动器中的 `PYTHONUTF8=1` 和 `PYTHONIOENCODING=utf-8` 保证 Python 以 UTF-8 输出。
Windows PowerShell 5.1 读取 UTF-8 文件时仍应显式指定编码：

```powershell
Get-Content -LiteralPath '.\result.json' -Raw -Encoding utf8
```

如果交互式控制台仍显示乱码，只对当前 PowerShell 会话设置 UTF-8：

```powershell
[Console]::InputEncoding = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
```

包含中文、引号或公式的二维数据优先保存为无 BOM 的 UTF-8 JSON 文件，再使用
`--values-file`；复杂 Patch 同样使用 UTF-8 文件。不要依赖多层 PowerShell/CMD 引号
转义传递复杂 JSON。

```powershell
$json = @'
[[1,"中文"],[2,"=A2*2"],[3,"'=以等号开头的文本"]]
'@
[IO.File]::WriteAllText(
    (Join-Path $PWD 'values.json'),
    $json,
    [Text.UTF8Encoding]::new($false)
)
exceltool write --file '.\book.xls' --sheet 'Sheet1' --begin A1 `
    --values-file '.\values.json' --json
```

不要通过转码工作簿、修改系统区域或替换中文内容来掩盖终端编码问题。

## 字体

Windows 版本从系统字体注册表精确检查请求字体，不依赖 Linux 的 `fc-match`：

```powershell
exceltool font check --name 'Microsoft YaHei' --json
```

只有 `available` 和 `exact` 都为 `true` 才表示能够精确使用该字体。ExcelTool 不分发
字体；微软雅黑等专有字体只使用 Windows 已合法安装的系统字体，不要把字体文件提交到
ExcelTool 或其他 Git 仓库。

## “正在等待打印机连接”弹窗

该弹窗通常不是 UNO 端口冲突，而是 LibreOffice 启动时枚举默认打印机或不可达的 WSD
网络打印机。ExcelTool 的 Windows 启动路径禁用同步打印机检测、打印机列表和默认打印机
查询，不需要永久修改用户的默认打印机。

每次 LibreOffice 会话使用独立的空闲端口和用户 profile。确认端口空闲可以排除冲突，
但不能解决打印机枚举阻塞。

## 常见问题

### 轻量 daemon 状态与诊断

工作簿命令会自动启动当前用户的 daemon，不安装 Windows 服务，也不需要管理员
权限。实例文件由 daemon 维护在 `%LOCALAPPDATA%\ExcelTool\runtime\daemon.json`，
CLI 仍会使用真实握手判断存活，不只看文件或 PID。

```powershell
exceltool daemon start
exceltool daemon status --json
exceltool daemon stop
exceltool --no-daemon view --file '.\book.xls' --sheet 'Sheet1'
```

默认命令会关闭工作簿但保留专属 LibreOffice，空闲 300 秒后两者退出。
daemon 使用 Job Object 拥有它创建的 LibreOffice 进程树，不会扫描或结束用户手动
打开的 LibreOffice。

### `ModuleNotFoundError: No module named 'fcntl'`

说明运行的是尚未包含 Windows 命令锁的旧版源码。更新 ExcelTool，并确认
`src/exceltool/locking.py` 在 Windows 分支使用 `msvcrt`。

### `Connection refused` 或启动超时

Windows 启动 `soffice.com`，服务端 accept 参数不带对象名：

```text
--accept=socket,host=127.0.0.1,port=<port>;urp;
```

客户端解析地址仍包含 `StarOffice.ComponentContext`。首次启动可能需要十余秒；不要并发
启动重试，ExcelTool 的 Windows 启动等待上限约为 30 秒。

### 清理临时 profile 报 `WinError 32`

`soffice.bin` 退出后可能短暂持有 `extensions.pmap`。Windows 实现会有限重试临时
profile 清理；不要跳过进程退出或改成无界重试。

### 工作簿正被占用

先关闭 Excel 或 LibreOffice 中打开的目标工作簿，再重试。不要强制终止用户正在使用的
应用。编辑共享文件时仍应在修改前计算 SHA-256，并使用 `--expect-sha256`。

## 完整环境验证

```powershell
& "$env:ProgramFiles\LibreOffice\program\python.exe" -c "import uno; print('UNO OK')"
exceltool --version
exceltool font check --name 'Microsoft YaHei' --json
exceltool sheet list --file '.\book.xls' --json
exceltool view --file '.\book.xls' --sheet 'Sheet1' --range 'A1:C5' --json-full
exceltool daemon status --json
```

仓库还提供一次性验收脚本，它会执行两次 `view`、独占打开检查、正常 stop，然后
强制结束测试中的 daemon 以验证 Job Object 清理。它不修改工作簿，但会停止当前
ExcelTool daemon，因此应在没有其他 ExcelTool 命令运行时执行：

```powershell
& '.\scripts\validate_daemon_windows.ps1' `
  -File "$env:USERPROFILE\Desktop\l凛冬降临.xls" `
  -Sheet '常量表' -Range 'A1:C5'
```

成功时输出 `ok: true`、首次/第二次耗时、复用的 PID/generation、文件释放、正常
stop 和崩溃清理结果。请保留这段 JSON 作为 Windows 验收证据。

编辑验证使用副本或明确授权的测试文件，并要求：

- 编辑结果 `verified: true`；
- `formula_verification.checked: true`；
- `formula_verification.unexpected_changes: 0`；
- 修改后用有界 `view --json-full --include-style` 回读目标范围；
- 连续两次 `view` 的 `libreoffice.generation`、`pid` 保持相同；
- 命令结束后目标工作簿可重命名或独占打开；
- `exceltool daemon stop` 后 daemon 及它创建的 `soffice.com`/`soffice.bin` 全部退出；
- 强制结束 daemon 进程后，Job Object 仍清理其专属 LibreOffice 进程树。
