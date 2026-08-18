import os
import shutil
import tempfile
from pathlib import Path

from .engine import LibreOfficeSession
from .errors import TargetError, UnsupportedError
from .safety import (
    assert_formula_snapshots,
    formula_verification_report,
    normalize_xls_string_formula_results,
    verify_input_sha256,
    verify_source_unchanged,
    workbook_formula_snapshot,
)


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


def edit_file(input_path, output_path, overwrite, operation,
              formula_policy=None, expect_sha256=None):
    source = validate_input(input_path)
    input_sha256 = verify_input_sha256(source, expect_sha256)
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

    policy = formula_policy or {
        "enforce_declared_changes": False,
        "allowed_changes": {},
    }
    with tempfile.TemporaryDirectory(prefix="exceltool-edit-") as temp_dir:
        working = Path(temp_dir) / ("working" + source.suffix.lower())
        shutil.copy2(str(source), str(working))
        verify_source_unchanged(working, input_sha256)
        with LibreOfficeSession() as session:
            workbook = session.load(working)
            try:
                original_formulas = workbook_formula_snapshot(workbook)
                changes, verifier = operation(workbook)
                planned_formulas = workbook_formula_snapshot(workbook)
                if policy.get("enforce_declared_changes", False):
                    assert_formula_snapshots(
                        original_formulas,
                        planned_formulas,
                        workbook,
                        "operation",
                        policy.get("allowed_changes", {}),
                    )
                normalized_string_results = 0
                if source.suffix.lower() == ".xls":
                    normalized_string_results = normalize_xls_string_formula_results(
                        workbook
                    )
                    normalized_formulas = workbook_formula_snapshot(workbook)
                    assert_formula_snapshots(
                        planned_formulas,
                        normalized_formulas,
                        workbook,
                        "normalization",
                    )
                workbook.save()
            finally:
                workbook.close()
        with LibreOfficeSession() as session:
            reopened = session.load(working, read_only=True)
            try:
                reopened_formulas = workbook_formula_snapshot(reopened)
                assert_formula_snapshots(
                    planned_formulas, reopened_formulas, reopened, "save"
                )
                verifier(reopened)
                formula_report = formula_verification_report(
                    original_formulas,
                    planned_formulas,
                    reopened_formulas,
                    policy,
                    normalized_string_results,
                )
            finally:
                reopened.close()

        if destination == source:
            verify_source_unchanged(source, input_sha256)
        publish(working, destination)

    return {
        "ok": True,
        "input": str(source),
        "output": str(destination),
        "verified": True,
        "input_sha256": input_sha256,
        "formula_verification": formula_report,
        "changes": changes,
    }
