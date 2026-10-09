import os
import socket
import subprocess
import tempfile
import time
from pathlib import Path

from .errors import ExcelToolError, TargetError, UnsupportedError

try:
    import uno
    from com.sun.star.beans import PropertyValue
except ImportError:  # pragma: no cover - exercised through dependency failure
    uno = None
    PropertyValue = None


FILTERS = {
    ".xls": "MS Excel 97",
    ".xlsx": "Calc MS Excel 2007 XML",
}


def property_value(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def file_url(path):
    return uno.systemPathToFileUrl(str(Path(path).resolve()))


def free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


class LibreOfficeSession:
    def __init__(self):
        self.process = None
        self.profile = None
        self.desktop = None

    def __enter__(self):
        if uno is None:
            raise UnsupportedError("缺少 Python UNO；请安装与 LibreOffice 匹配的 python3-uno")
        port = free_port()
        self.profile = tempfile.TemporaryDirectory(prefix="exceltool-lo-")
        profile_url = file_url(self.profile.name)
        accept = (
            "--accept=socket,host=127.0.0.1,port=%d;urp;" % port
            if os.name == "nt"
            else "--accept=socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port
        )
        command = [
            "soffice.com" if os.name == "nt" else "soffice",
            "--headless",
            "--nologo",
            "--nodefault",
            "--nofirststartwizard",
            "--norestore",
            "-env:UserInstallation=%s" % profile_url,
            accept,
        ]
        environment = None
        if os.name == "nt":
            environment = os.environ.copy()
            environment["SAL_DISABLE_SYNCHRONOUS_PRINTER_DETECTION"] = "1"
            environment["SAL_DISABLE_PRINTERLIST"] = "1"
            environment["SAL_DISABLE_DEFAULTPRINTER"] = "1"
        try:
            self.process = subprocess.Popen(
                command,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                env=environment,
            )
        except OSError as exc:
            raise UnsupportedError("无法启动 LibreOffice: %s" % exc)

        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        last_error = None
        attempts = 600 if os.name == "nt" else 80
        for _ in range(attempts):
            if self.process.poll() is not None:
                break
            try:
                context = resolver.resolve(
                    "uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext" % port
                )
                self.desktop = context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", context
                )
                return self
            except Exception as exc:  # UNO exposes implementation-specific exceptions
                last_error = exc
                time.sleep(0.05)
        stderr = ""
        if self.process.poll() is not None and self.process.stderr:
            stderr = self.process.stderr.read().decode("utf-8", "replace").strip()
        self.close()
        raise ExcelToolError("LibreOffice 连接失败: %s%s" % (last_error or "进程退出", ": " + stderr if stderr else ""))

    def load(self, path, read_only=False):
        path = Path(path)
        if path.suffix.lower() not in FILTERS:
            raise UnsupportedError("仅支持 .xls 和 .xlsx: %s" % path)
        if not path.is_file():
            raise TargetError("文件不存在: %s" % path)
        options = [
            property_value("Hidden", True),
            property_value("ReadOnly", read_only),
            property_value("MacroExecutionMode", 0),
        ]
        try:
            document = self.desktop.loadComponentFromURL(file_url(path), "_blank", 0, tuple(options))
        except Exception as exc:
            raise TargetError("无法打开工作簿 %s: %s" % (path, exc))
        if document is None:
            raise TargetError("无法打开工作簿: %s" % path)
        return Workbook(document, path)

    def create(self):
        options = (property_value("Hidden", True),)
        document = self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, options)
        if document is None:
            raise ExcelToolError("无法创建工作簿")
        return Workbook(document, None)

    def close(self):
        if self.desktop is not None:
            try:
                self.desktop.terminate()
            except Exception:
                pass
            self.desktop = None
        if self.process is not None:
            try:
                self.process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                self.process.terminate()
                try:
                    self.process.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    self.process.kill()
                    self.process.wait()
            if self.process.stderr:
                self.process.stderr.close()
            self.process = None
        if self.profile is not None:
            attempts = 100 if os.name == "nt" else 1
            for attempt in range(attempts):
                try:
                    self.profile.cleanup()
                    break
                except PermissionError:
                    if attempt + 1 == attempts:
                        raise
                    time.sleep(0.1)
            self.profile = None

    def __exit__(self, exc_type, exc, traceback):
        self.close()


class Workbook:
    def __init__(self, document, path):
        self.document = document
        self.path = Path(path) if path else None

    def close(self):
        if self.document is not None:
            try:
                self.document.close(True)
            except Exception:
                try:
                    self.document.dispose()
                except Exception:
                    pass
            self.document = None

    def sheet_names(self):
        return list(self.document.Sheets.getElementNames())

    def sheet(self, name):
        sheets = self.document.Sheets
        if name not in sheets.getElementNames():
            raise TargetError("未找到 sheet: %s" % name)
        return sheets.getByName(name)

    def save_as(self, path):
        path = Path(path)
        suffix = path.suffix.lower()
        if suffix not in FILTERS:
            raise UnsupportedError("输出格式必须是 .xls 或 .xlsx: %s" % path)
        path.parent.mkdir(parents=True, exist_ok=True)
        options = (
            property_value("FilterName", FILTERS[suffix]),
            property_value("Overwrite", True),
        )
        try:
            self.document.storeAsURL(file_url(path), options)
            self.path = path
        except Exception as exc:
            raise ExcelToolError("保存工作簿失败 %s: %s" % (path, exc))

    def save(self):
        try:
            self.document.store()
        except Exception as exc:
            raise ExcelToolError("保存工作簿失败 %s: %s" % (self.path, exc))

    def used_size(self, sheet):
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(True)
        address = cursor.RangeAddress
        return address.EndRow + 1, address.EndColumn + 1

    def display_value(self, cell):
        return cell.String

    def json_value(self, cell):
        cell_type = cell.Type.value
        if cell_type == "EMPTY":
            return None
        if cell_type == "VALUE":
            value = cell.Value
            return int(value) if value.is_integer() else value
        if cell_type == "TEXT":
            value = cell.String
            return "'" + value if value.startswith("=") else value
        if cell_type == "FORMULA":
            formula = cell.Formula
            normalized = formula.replace(" ", "").upper()
            if normalized in ("=TRUE()", "=TRUE"):
                return True
            if normalized in ("=FALSE()", "=FALSE"):
                return False
            return formula
        return None
