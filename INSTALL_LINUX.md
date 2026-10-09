# ExcelTool Linux 安装说明

本文说明 ExcelTool 在 Linux 上的运行依赖、安装、升级、卸载和可选字体配置。命令使用方式见
[README](README.md)。

## 运行依赖

- Python 3.8 或更高版本；
- LibreOffice Calc；
- 与 LibreOffice 匹配的 Python UNO bridge；
- 使用字体检查或字体编辑时需要 fontconfig，Ubuntu/Debian 通常由系统字体环境提供。

Ubuntu/Debian 安装依赖：

```bash
sudo apt update
sudo apt install libreoffice-calc python3-uno fontconfig
```

`python3-uno` 通常绑定系统 Python，不要用 `pip install uno` 代替。如果虚拟环境中的
Python 无法导入 UNO，请退出虚拟环境，或使用系统 Python 运行仓库入口。

## 获取并直接运行

```bash
git clone https://github.com/zhangbiran/exceltool.git
cd exceltool
python3 --version
soffice --version
python3 -c 'import uno; print("UNO OK")'
./exceltool --version
./exceltool --help
```

仓库入口固定使用 `/usr/bin/python3`，以便加载系统安装的 UNO。完成依赖检查后，无需
安装即可在仓库根目录用 `./exceltool` 工作。

## 安装到命令搜索路径

安装到当前用户的 `~/.local`，不需要管理员权限：

```bash
make install PREFIX="$HOME/.local"
"$HOME/.local/bin/exceltool" --version
```

如果 `~/.local/bin` 尚未进入 `PATH`，将下面一行加入当前 shell 的启动文件后重新打开
终端：

```bash
export PATH="$HOME/.local/bin:$PATH"
```

安装到系统默认的 `/usr/local` 需要管理员权限：

```bash
sudo make install
exceltool --version
```

只有用户明确允许系统级安装时才使用 `sudo`。安装脚本不会安装系统依赖，也不会修改
shell 启动文件。

## 升级与卸载

升级源码并重新安装到原来的前缀：

```bash
git pull --ff-only
make install PREFIX="$HOME/.local"
exceltool --version
```

系统级安装则使用 `sudo make install`。不要在未确认本地改动和目标版本时强制覆盖
源码工作区。

卸载当前用户安装：

```bash
make uninstall PREFIX="$HOME/.local"
```

卸载系统级安装：

```bash
sudo make uninstall
```

卸载只删除 Makefile 安装的 ExcelTool 入口和 Python 模块，不删除 LibreOffice、UNO、
fontconfig 或字体。

## 安装可选字体

ExcelTool 不分发字体。字体文件必须由使用者从有权使用的来源取得，并确认其许可允许在
目标机器上安装。不要把 Microsoft YaHei 等专有字体文件提交到 ExcelTool、业务项目或
其他 Git 仓库；操作系统自带字体也不等于可以重新分发。

Linux 的当前用户字体可以放在 `~/.local/share/fonts/`。子目录名称只是为了便于管理，
可以按字体家族自行命名。例如，已经合法取得微软雅黑字体文件时：

```bash
mkdir -p "$HOME/.local/share/fonts/microsoft-yahei"
install -m 0644 /path/to/MSYH.TTC \
  "$HOME/.local/share/fonts/microsoft-yahei/MSYH.TTC"
install -m 0644 /path/to/MSYHBD.TTC \
  "$HOME/.local/share/fonts/microsoft-yahei/MSYHBD.TTC"
install -m 0644 /path/to/MSYHL.TTC \
  "$HOME/.local/share/fonts/microsoft-yahei/MSYHL.TTC"
fc-cache -f "$HOME/.local/share/fonts"
```

文件名可能因合法来源和字体版本而不同，只安装实际取得的文件，不要为了凑齐示例文件名
从不明来源下载。`.ttf`、`.ttc` 和 `.otf` 都可以由 fontconfig 管理。

先确认系统能够解析字体，再确认 ExcelTool 得到精确匹配：

```bash
fc-match "Microsoft YaHei"
exceltool font check --name "微软雅黑" --json
```

在仓库内直接运行时，将最后一条命令改为：

```bash
./exceltool font check --name "微软雅黑" --json
```

JSON 结果包含 `requested`、系统实际 `resolved`、`available` 和 `exact`。只有
`available` 与 `exact` 都为 `true`，才表示请求字体可以精确使用。ExcelTool 在字体
不匹配时会在修改工作簿前退出，不会静默换成替代字体。

如果 `fc-match` 仍返回其他字体，依次检查字体文件权限、`fc-cache` 输出和请求的字体
家族名；不要通过忽略 `exact=false` 来继续写入。

## 完整环境验证

```bash
python3 --version
soffice --version
python3 -c 'import uno; print("UNO OK")'
exceltool --version
exceltool --help
```

使用字体编辑时，再执行目标字体的 `font check`。安装完成后的命令使用和安全约束见
[README](README.md)。
