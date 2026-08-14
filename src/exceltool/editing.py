import os
import shutil
import tempfile
from pathlib import Path

from .engine import LibreOfficeSession
from .errors import TargetError, UnsupportedError


def validate_input(path):
    path = Path(path).resolve()
    if not path.is_file():
        raise TargetError("文件不存在: %s" % path)
    if path.suffix.lower() not in (".xls", ".xlsx"):
        raise UnsupportedError("仅支持 .xls 和 .xlsx: %s" % path)
    return path


def publish(source, destination):
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=".%s." % destination.name, dir=str(destination.parent))
    os.close(handle)
    try:
        shutil.copy2(str(source), temporary)
        os.replace(temporary, str(destination))
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def edit_file(input_path, output_path, overwrite, operation):
    source = validate_input(input_path)
    if output_path:
        destination = Path(output_path).resolve()
        if destination == source:
            raise TargetError("--out 不能等于输入文件；省略 --out 即可修改原文件")
        if destination.suffix.lower() != source.suffix.lower():
            raise TargetError("输出扩展名必须与输入格式一致")
        if destination.exists() and not overwrite:
            raise TargetError("输出文件已存在；使用 --overwrite 明确覆盖: %s" % destination)
    else:
        if overwrite:
            raise TargetError("--overwrite 只能与 --out 一起使用")
        destination = source

    with tempfile.TemporaryDirectory(prefix="exceltool-edit-") as temp_dir:
        working = Path(temp_dir) / ("working" + source.suffix.lower())
        shutil.copy2(str(source), str(working))
        with LibreOfficeSession() as session:
            workbook = session.load(working)
            try:
                changes, verifier = operation(workbook)
                workbook.save()
            finally:
                workbook.close()
        with LibreOfficeSession() as session:
            reopened = session.load(working, read_only=True)
            try:
                verifier(reopened)
            finally:
                reopened.close()

        publish(working, destination)

    return {
        "ok": True,
        "input": str(source),
        "output": str(destination),
        "verified": True,
        "changes": changes,
    }
