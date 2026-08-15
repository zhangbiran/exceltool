import json
import math
from pathlib import Path

from .errors import ExcelToolError, PatchOperationError, TargetError, VerificationError
from .fonts import inspect_font
from .operations import (
    HARD_STYLE_PROPERTIES,
    COLUMN_WIDTH_TOLERANCE,
    clear_range,
    col_autofit,
    col_copy,
    col_delete,
    col_insert,
    row_copy,
    row_delete,
    row_insert,
    sheet_add,
    sheet_copy,
    sheet_delete,
    sheet_rename,
    style_range,
    required_font_slots,
    write_matrix,
)
from .ranges import cell_range, cell_ref, col_range, column_name, column_number, row_range


OPERATION_FIELDS = {
    "write": {"id", "op", "sheet", "begin", "values", "font", "font_size"},
    "style": {"id", "op", "sheet", "range", "font", "font_size"},
    "clear": {"id", "op", "sheet", "range", "with_style"},
    "row.insert": {"id", "op", "sheet", "before", "count"},
    "row.delete": {"id", "op", "sheet", "rows"},
    "row.copy": {"id", "op", "sheet", "rows", "insert_before"},
    "col.insert": {"id", "op", "sheet", "before", "count"},
    "col.delete": {"id", "op", "sheet", "cols"},
    "col.copy": {"id", "op", "sheet", "cols", "insert_before"},
    "col.autofit": {"id", "op", "sheet", "cols", "max_width_mm"},
    "sheet.add": {"id", "op", "name"},
    "sheet.delete": {"id", "op", "sheet"},
    "sheet.rename": {"id", "op", "sheet", "name"},
    "sheet.copy": {"id", "op", "sheet", "name"},
}

REQUIRED_FIELDS = {
    "write": {"op", "sheet", "begin", "values"},
    "style": {"op", "sheet", "range"},
    "clear": {"op", "sheet", "range"},
    "row.insert": {"op", "sheet", "before"},
    "row.delete": {"op", "sheet", "rows"},
    "row.copy": {"op", "sheet", "rows", "insert_before"},
    "col.insert": {"op", "sheet", "before"},
    "col.delete": {"op", "sheet", "cols"},
    "col.copy": {"op", "sheet", "cols", "insert_before"},
    "col.autofit": {"op", "sheet", "cols"},
    "sheet.add": {"op", "name"},
    "sheet.delete": {"op", "sheet"},
    "sheet.rename": {"op", "sheet", "name"},
    "sheet.copy": {"op", "sheet", "name"},
}


def operation_identity(index, operation):
    identity = {"index": index}
    if isinstance(operation.get("id"), str):
        identity["id"] = operation["id"]
    if isinstance(operation.get("op"), str):
        identity["op"] = operation["op"]
    return identity


def operation_error(index, operation, error):
    if isinstance(error, PatchOperationError):
        return error
    if isinstance(error, ExcelToolError):
        message = error.message
        code = error.code
    else:
        message = "操作执行失败: %s" % error
        code = 5
    wrapped = PatchOperationError(message, code, operation_identity(index, operation))
    if hasattr(error, "details"):
        wrapped.details.update(error.details)
    return wrapped


def require_string(operation, field):
    value = operation.get(field)
    if not isinstance(value, str) or not value:
        raise TargetError("%s 必须是非空字符串" % field)


def require_positive_integer(operation, field, default=None):
    value = operation.get(field, default)
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise TargetError("%s 必须是正整数" % field)
    operation[field] = value


def validate_font_size_value(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TargetError("font_size 必须是有限的正数")


def validate_positive_number(value, field):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TargetError("%s 必须是有限的正数" % field)
    if not math.isfinite(float(value)) or float(value) <= 0:
        raise TargetError("%s 必须是有限的正数" % field)
    if not math.isfinite(float(value)) or float(value) <= 0:
        raise TargetError("font_size 必须是有限的正数")


def validate_matrix(matrix):
    if not isinstance(matrix, list) or not matrix:
        raise TargetError("values 必须是非空二维 JSON 数组")
    touched = False
    for row in matrix:
        if not isinstance(row, list):
            raise TargetError("values 的每一行都必须是数组")
        for value in row:
            touched = True
            if value is None or isinstance(value, (str, bool)):
                continue
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                if not math.isfinite(value):
                    raise TargetError("JSON 数字必须是有限值")
                continue
            raise TargetError("JSON 单元格只支持 null、字符串、数字和布尔值")
    if not touched:
        raise TargetError("values 没有包含任何单元格；空行 [] 只表示跳过")


def validate_operation(operation):
    if not isinstance(operation, dict):
        raise TargetError("operation 必须是 JSON 对象")
    op = operation.get("op")
    if not isinstance(op, str) or op not in OPERATION_FIELDS:
        raise TargetError("不支持的 patch 操作: %s" % (op if op is not None else "未提供"))
    unexpected = sorted(set(operation) - OPERATION_FIELDS[op])
    if unexpected:
        raise TargetError("%s 不支持字段: %s" % (op, ", ".join(unexpected)))
    missing = sorted(REQUIRED_FIELDS[op] - set(operation))
    if missing:
        raise TargetError("%s 缺少字段: %s" % (op, ", ".join(missing)))
    normalized = dict(operation)
    if "id" in normalized:
        require_string(normalized, "id")
    if "sheet" in normalized:
        require_string(normalized, "sheet")
    if "name" in normalized:
        require_string(normalized, "name")
    if "font" in normalized:
        require_string(normalized, "font")
        font_check = inspect_font(normalized["font"])
        if not font_check["exact"]:
            raise TargetError(
                "字体不可精确使用: 请求=%s，系统解析=%s"
                % (normalized["font"], font_check["resolved"] or "未找到")
            )
    if "font_size" in normalized:
        validate_font_size_value(normalized["font_size"])

    if op == "write":
        require_string(normalized, "begin")
        cell_ref(normalized["begin"])
        validate_matrix(normalized["values"])
    elif op == "style":
        require_string(normalized, "range")
        cell_range_value(normalized["range"])
        if "font" not in normalized and "font_size" not in normalized:
            raise TargetError("style 至少提供 font 或 font_size")
    elif op == "clear":
        require_string(normalized, "range")
        cell_range_value(normalized["range"])
        with_style = normalized.get("with_style", False)
        if not isinstance(with_style, bool):
            raise TargetError("with_style 必须是布尔值")
        normalized["with_style"] = with_style
    elif op == "row.insert":
        require_positive_integer(normalized, "before")
        require_positive_integer(normalized, "count", 1)
    elif op == "row.delete":
        require_string(normalized, "rows")
        row_range(normalized["rows"])
    elif op == "row.copy":
        require_string(normalized, "rows")
        row_range(normalized["rows"])
        require_positive_integer(normalized, "insert_before")
    elif op == "col.insert":
        require_string(normalized, "before")
        column_number(normalized["before"])
        require_positive_integer(normalized, "count", 1)
    elif op == "col.delete":
        require_string(normalized, "cols")
        col_range(normalized["cols"])
    elif op == "col.copy":
        require_string(normalized, "cols")
        col_range(normalized["cols"])
        require_string(normalized, "insert_before")
        column_number(normalized["insert_before"])
    elif op == "col.autofit":
        require_string(normalized, "cols")
        col_range(normalized["cols"])
        maximum = normalized.get("max_width_mm", 60.0)
        validate_positive_number(maximum, "max_width_mm")
        normalized["max_width_mm"] = float(maximum)
    return normalized


def load_patch(path):
    patch_path = Path(path).resolve()
    if not patch_path.is_file():
        raise TargetError("patch 文件不存在: %s" % patch_path)
    try:
        document = json.loads(patch_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise TargetError("无法读取 patch JSON: %s" % exc)
    if not isinstance(document, dict):
        raise TargetError("patch 顶层必须是 JSON 对象")
    unexpected = sorted(set(document) - {"version", "operations"})
    if unexpected:
        raise TargetError("patch 顶层不支持字段: %s" % ", ".join(unexpected))
    version = document.get("version")
    if isinstance(version, bool) or version != 1:
        raise TargetError("patch version 必须为 1")
    operations = document.get("operations")
    if not isinstance(operations, list) or not operations:
        raise TargetError("patch operations 必须是非空数组")
    normalized = []
    for index, operation in enumerate(operations):
        raw = operation if isinstance(operation, dict) else {}
        try:
            normalized.append(validate_operation(operation))
        except Exception as exc:
            raise operation_error(index, raw, exc)
    return {"version": 1, "operations": normalized}


def cell_range_value(value):
    return cell_range(value if ":" in value else "%s:%s" % (value, value))


STYLE_PROPERTY_GROUPS = (
    ("CharFontName", "CharFontNameAsian", "CharFontNameComplex"),
    ("CharHeight", "CharHeightAsian", "CharHeightComplex"),
    ("CharWeight", "CharWeightAsian", "CharWeightComplex"),
    ("CellBackColor",),
    ("NumberFormat",),
)


class SheetVerificationError(VerificationError):
    def __init__(self, message, row=None, col=None):
        super().__init__(message)
        self.row = row
        self.col = col


def cell_expectation(cell):
    cell_type = cell.Type.value
    if cell_type == "VALUE":
        content = (cell_type, float(cell.Value))
    elif cell_type == "FORMULA":
        content = (cell_type, cell.Formula)
    elif cell_type == "TEXT":
        content = (cell_type, cell.String)
    else:
        content = (cell_type,)
    properties = {}
    for group in STYLE_PROPERTY_GROUPS:
        if any(
            getattr(cell.getPropertyState(name), "value", str(cell.getPropertyState(name)))
            == "DIRECT_VALUE"
            for name in group
        ):
            names = group
            if group[0] == "CharFontName":
                slot_names = {
                    "western": "CharFontName",
                    "asian": "CharFontNameAsian",
                    "complex": "CharFontNameComplex",
                }
                names = tuple(slot_names[slot] for slot in required_font_slots(cell))
            for name in names:
                properties[name] = getattr(cell, name)
    return {
        "content": content,
        "cell_style": cell.CellStyle,
        "properties": properties,
    }


def matches_cell_expectation(cell, expected):
    actual = cell_expectation(cell)
    if actual["content"] != expected["content"]:
        return False
    if actual["cell_style"] != expected["cell_style"]:
        return False
    for name, value in expected["properties"].items():
        actual_value = getattr(cell, name)
        if isinstance(value, float):
            if abs(float(actual_value) - value) > 0.01:
                return False
        elif actual_value != value:
            return False
    return True


class VerificationPlan:
    def __init__(self):
        self.full_sheets = set()
        self.cells = {}
        self.hard_default_cells = {}
        self.width_columns = {}
        self.width_reports = []
        self.last_by_cell = {}
        self.last_by_column = {}
        self.last_by_sheet = {}
        self.last_structure = None

    def _touch(self, sheet, identity):
        self.last_by_sheet[sheet] = identity

    def record(self, operation, index, changes=None):
        identity = operation_identity(index, operation)
        op = operation["op"]
        sheet = operation.get("sheet")
        if sheet:
            self._touch(sheet, identity)
        if op == "write":
            start_row, start_col = cell_ref(operation["begin"])
            targets = self.cells.setdefault(sheet, set())
            for row_offset, values in enumerate(operation["values"]):
                for col_offset in range(len(values)):
                    target = (start_row + row_offset, start_col + col_offset)
                    targets.add(target)
                    self.last_by_cell.setdefault(sheet, {})[target] = identity
                    if "font" in operation or "font_size" in operation:
                        self.hard_default_cells.setdefault(sheet, set()).discard(target)
        elif op in ("style", "clear"):
            start_row, end_row, start_col, end_col = cell_range_value(operation["range"])
            targets = self.cells.setdefault(sheet, set())
            for row in range(start_row, end_row):
                for col in range(start_col, end_col):
                    target = (row, col)
                    targets.add(target)
                    self.last_by_cell.setdefault(sheet, {})[target] = identity
                    if op == "style":
                        self.hard_default_cells.setdefault(sheet, set()).discard(target)
                    elif operation["with_style"]:
                        self.hard_default_cells.setdefault(sheet, set()).add(target)
        elif op.startswith("row."):
            if op == "row.insert":
                self._insert_rows(sheet, operation["before"] - 1, operation["count"])
            elif op == "row.delete":
                start, end = row_range(operation["rows"])
                self._delete_rows(sheet, start, end)
            else:
                start, end = row_range(operation["rows"])
                self._copy_rows(sheet, start, end, operation["insert_before"] - 1)
            self.full_sheets.add(sheet)
            self.cells.pop(sheet, None)
        elif op.startswith("col."):
            if op == "col.insert":
                before = column_number(operation["before"]) - 1
                self._insert_cols(sheet, before, operation["count"])
                affected = range(before, before + operation["count"])
            elif op == "col.delete":
                start, end = col_range(operation["cols"])
                self._delete_cols(sheet, start, end)
                affected = (start,)
            elif op == "col.copy":
                start, end = col_range(operation["cols"])
                before = column_number(operation["insert_before"]) - 1
                self._copy_cols(sheet, start, end, before)
                affected = range(before, before + end - start)
            else:
                start, end = col_range(operation["cols"])
                affected = range(start, end)
                self.width_columns.setdefault(sheet, set()).update(affected)
                for offset, col in enumerate(affected):
                    self.width_reports.append({
                        "sheet": sheet,
                        "col": col,
                        "item": changes["columns"][offset],
                    })
            for col in affected:
                self.last_by_column.setdefault(sheet, {})[col] = identity
            self.full_sheets.add(sheet)
            self.cells.pop(sheet, None)
        elif op == "sheet.add":
            self.last_structure = identity
            self._touch(operation["name"], identity)
        elif op == "sheet.delete":
            self.last_structure = identity
            self.full_sheets.discard(sheet)
            self.cells.pop(sheet, None)
            self.hard_default_cells.pop(sheet, None)
            self.width_columns.pop(sheet, None)
            self.last_by_cell.pop(sheet, None)
            self.last_by_column.pop(sheet, None)
            for report in self.width_reports:
                if report["sheet"] == sheet:
                    report["sheet"] = None
                    report["col"] = None
        elif op == "sheet.rename":
            self.last_structure = identity
            new_name = operation["name"]
            if sheet in self.full_sheets:
                self.full_sheets.remove(sheet)
                self.full_sheets.add(new_name)
            if sheet in self.cells:
                self.cells[new_name] = self.cells.pop(sheet)
            if sheet in self.hard_default_cells:
                self.hard_default_cells[new_name] = self.hard_default_cells.pop(sheet)
            if sheet in self.last_by_cell:
                self.last_by_cell[new_name] = self.last_by_cell.pop(sheet)
            if sheet in self.width_columns:
                self.width_columns[new_name] = self.width_columns.pop(sheet)
            if sheet in self.last_by_column:
                self.last_by_column[new_name] = self.last_by_column.pop(sheet)
            for report in self.width_reports:
                if report["sheet"] == sheet:
                    report["sheet"] = new_name
            self.last_by_sheet.pop(sheet, None)
            self._touch(new_name, identity)
        elif op == "sheet.copy":
            self.last_structure = identity
            self.full_sheets.add(operation["name"])
            if sheet in self.hard_default_cells:
                self.hard_default_cells[operation["name"]] = set(
                    self.hard_default_cells[sheet]
                )
            if sheet in self.width_columns:
                self.width_columns[operation["name"]] = set(self.width_columns[sheet])
            self._touch(operation["name"], identity)

    def _insert_rows(self, sheet, before, count):
        for mapping in (self.cells, self.hard_default_cells):
            if sheet in mapping:
                mapping[sheet] = {
                    (row + count if row >= before else row, col)
                    for row, col in mapping[sheet]
                }
        if sheet in self.last_by_cell:
            self.last_by_cell[sheet] = {
                (row + count if row >= before else row, col): identity
                for (row, col), identity in self.last_by_cell[sheet].items()
            }

    def _delete_rows(self, sheet, start, end):
        count = end - start
        for mapping in (self.cells, self.hard_default_cells):
            if sheet in mapping:
                mapping[sheet] = {
                    (row - count if row >= end else row, col)
                    for row, col in mapping[sheet]
                    if not start <= row < end
                }
        if sheet in self.last_by_cell:
            self.last_by_cell[sheet] = {
                (row - count if row >= end else row, col): identity
                for (row, col), identity in self.last_by_cell[sheet].items()
                if not start <= row < end
            }

    def _copy_rows(self, sheet, start, end, insert_before):
        count = end - start
        for mapping in (self.cells, self.hard_default_cells):
            if sheet not in mapping:
                continue
            original = mapping[sheet]
            copied = {
                (insert_before + row - start, col)
                for row, col in original
                if start <= row < end
            }
            mapping[sheet] = {
                (row + count if row >= insert_before else row, col)
                for row, col in original
            } | copied
        if sheet in self.last_by_cell:
            original = self.last_by_cell[sheet]
            copied = {
                (insert_before + row - start, col): self.last_by_sheet[sheet]
                for row, col in original
                if start <= row < end
            }
            self.last_by_cell[sheet] = {
                (row + count if row >= insert_before else row, col): identity
                for (row, col), identity in original.items()
            }
            self.last_by_cell[sheet].update(copied)

    def _insert_cols(self, sheet, before, count):
        for mapping in (self.cells, self.hard_default_cells):
            if sheet in mapping:
                mapping[sheet] = {
                    (row, col + count if col >= before else col)
                    for row, col in mapping[sheet]
                }
        if sheet in self.width_columns:
            self.width_columns[sheet] = {
                col + count if col >= before else col
                for col in self.width_columns[sheet]
            }
        if sheet in self.last_by_cell:
            self.last_by_cell[sheet] = {
                (row, col + count if col >= before else col): identity
                for (row, col), identity in self.last_by_cell[sheet].items()
            }
        if sheet in self.last_by_column:
            self.last_by_column[sheet] = {
                (col + count if col >= before else col): identity
                for col, identity in self.last_by_column[sheet].items()
            }
        for report in self.width_reports:
            if (
                report["sheet"] == sheet
                and report["col"] is not None
                and report["col"] >= before
            ):
                report["col"] += count

    def _delete_cols(self, sheet, start, end):
        count = end - start
        for mapping in (self.cells, self.hard_default_cells):
            if sheet in mapping:
                mapping[sheet] = {
                    (row, col - count if col >= end else col)
                    for row, col in mapping[sheet]
                    if not start <= col < end
                }
        if sheet in self.width_columns:
            self.width_columns[sheet] = {
                col - count if col >= end else col
                for col in self.width_columns[sheet]
                if not start <= col < end
            }
        if sheet in self.last_by_cell:
            self.last_by_cell[sheet] = {
                (row, col - count if col >= end else col): identity
                for (row, col), identity in self.last_by_cell[sheet].items()
                if not start <= col < end
            }
        if sheet in self.last_by_column:
            self.last_by_column[sheet] = {
                (col - count if col >= end else col): identity
                for col, identity in self.last_by_column[sheet].items()
                if not start <= col < end
            }
        for report in self.width_reports:
            if report["sheet"] != sheet or report["col"] is None:
                continue
            if start <= report["col"] < end:
                report["col"] = None
            elif report["col"] >= end:
                report["col"] -= count

    def _copy_cols(self, sheet, start, end, insert_before):
        count = end - start
        for mapping in (self.cells, self.hard_default_cells):
            if sheet not in mapping:
                continue
            original = mapping[sheet]
            copied = {
                (row, insert_before + col - start)
                for row, col in original
                if start <= col < end
            }
            mapping[sheet] = {
                (row, col + count if col >= insert_before else col)
                for row, col in original
            } | copied
        if sheet in self.width_columns:
            original = self.width_columns[sheet]
            copied = {
                insert_before + col - start for col in original if start <= col < end
            }
            self.width_columns[sheet] = {
                col + count if col >= insert_before else col for col in original
            } | copied
        if sheet in self.last_by_cell:
            original = self.last_by_cell[sheet]
            copied = {
                (row, insert_before + col - start): self.last_by_sheet[sheet]
                for (row, col), identity in original.items()
                if start <= col < end
            }
            self.last_by_cell[sheet] = {
                (row, col + count if col >= insert_before else col): identity
                for (row, col), identity in original.items()
            }
            self.last_by_cell[sheet].update(copied)
        if sheet in self.last_by_column:
            original = self.last_by_column[sheet]
            copied = {
                insert_before + col - start: self.last_by_sheet[sheet]
                for col in original
                if start <= col < end
            }
            self.last_by_column[sheet] = {
                (col + count if col >= insert_before else col): identity
                for col, identity in original.items()
            }
            self.last_by_column[sheet].update(copied)
        for report in self.width_reports:
            if (
                report["sheet"] == sheet
                and report["col"] is not None
                and report["col"] >= insert_before
            ):
                report["col"] += count

    def snapshot(self, workbook):
        names = workbook.sheet_names()
        snapshots = {}
        for sheet_name in self.full_sheets:
            if sheet_name in names:
                snapshots[sheet_name] = full_sheet_snapshot(
                    workbook, sheet_name, self.width_columns.get(sheet_name, set())
                )
        for sheet_name, targets in self.cells.items():
            if sheet_name in names and sheet_name not in self.full_sheets:
                sheet = workbook.sheet(sheet_name)
                snapshots[sheet_name] = {
                    "cells": {
                        target: cell_expectation(sheet.getCellByPosition(target[1], target[0]))
                        for target in targets
                    }
                }
        for sheet_name, targets in self.hard_default_cells.items():
            if sheet_name in names:
                snapshots.setdefault(sheet_name, {"cells": {}})["hard_default_cells"] = set(targets)
        return {"sheet_names": names, "sheets": snapshots}

    def verify(self, reopened, expected):
        if reopened.sheet_names() != expected["sheet_names"]:
            identity = self.last_structure or self.last_operation()
            raise PatchOperationError("patch sheet 结构验证失败", 6, identity)
        for sheet_name, sheet_snapshot in expected["sheets"].items():
            try:
                verify_sheet_snapshot(reopened, sheet_name, sheet_snapshot)
            except SheetVerificationError as exc:
                identity = self.last_by_cell.get(sheet_name, {}).get((exc.row, exc.col))
                if identity is None and exc.col is not None:
                    identity = self.last_by_column.get(sheet_name, {}).get(exc.col)
                if identity is None:
                    identity = self.last_by_sheet.get(sheet_name, self.last_operation())
                raise PatchOperationError(exc.message, exc.code, identity)
        for report in self.width_reports:
            item = report["item"]
            if report["sheet"] is None or report["col"] is None:
                item["final_column"] = None
                item["actual_width_mm"] = None
                continue
            item["final_column"] = column_name(report["col"] + 1)
            actual = reopened.sheet(report["sheet"]).Columns.getByIndex(report["col"]).Width
            item["actual_width_mm"] = round(actual / 100.0, 2)

    def last_operation(self):
        identities = list(self.last_by_sheet.values())
        if self.last_structure is not None:
            identities.append(self.last_structure)
        return max(identities, key=lambda item: item["index"]) if identities else {"index": 0}


def full_sheet_snapshot(workbook, sheet_name, extra_columns=()):
    sheet = workbook.sheet(sheet_name)
    rows, columns = workbook.used_size(sheet)
    return {
        "size": (rows, columns),
        "row_heights": [sheet.Rows.getByIndex(row).Height for row in range(rows)],
        "column_widths": {
            col: sheet.Columns.getByIndex(col).Width
            for col in range(max([columns] + [col + 1 for col in extra_columns]))
        },
        "cells": {
            (row, col): cell_expectation(sheet.getCellByPosition(col, row))
            for row in range(rows)
            for col in range(columns)
        },
    }


def verify_sheet_snapshot(workbook, sheet_name, expected):
    sheet = workbook.sheet(sheet_name)
    if "size" in expected:
        actual_size = workbook.used_size(sheet)
        if tuple(actual_size) != tuple(expected["size"]):
            raise SheetVerificationError("patch sheet 范围验证失败: %s" % sheet_name)
        for row, height in enumerate(expected["row_heights"]):
            if abs(sheet.Rows.getByIndex(row).Height - height) > 2:
                raise SheetVerificationError(
                    "patch 行高验证失败: %s!%d" % (sheet_name, row + 1), row=row
                )
        for col, width in expected["column_widths"].items():
            if abs(sheet.Columns.getByIndex(col).Width - width) > COLUMN_WIDTH_TOLERANCE:
                raise SheetVerificationError(
                    "patch 列宽验证失败: %s!%s"
                    % (sheet_name, column_name(col + 1)), col=col
                )
    for (row, col), signature in expected["cells"].items():
        cell = sheet.getCellByPosition(col, row)
        if not matches_cell_expectation(cell, signature):
            raise SheetVerificationError(
                "patch 单元格验证失败: %s!%s%d"
                % (sheet_name, column_name(col + 1), row + 1),
                row=row,
                col=col,
            )
    for row, col in expected.get("hard_default_cells", set()):
        cell = sheet.getCellByPosition(col, row)
        if any(
            getattr(cell.getPropertyState(name), "value", str(cell.getPropertyState(name)))
            != "DEFAULT_VALUE"
            for name in HARD_STYLE_PROPERTIES
        ):
            raise SheetVerificationError(
                "patch 硬样式清除验证失败: %s!%s%d"
                % (sheet_name, column_name(col + 1), row + 1),
                row=row,
                col=col,
            )


def execute_operation(workbook, operation):
    op = operation["op"]
    if op == "write":
        return write_matrix(
            workbook, operation["sheet"], operation["begin"], operation["values"],
            operation.get("font"), operation.get("font_size"),
        )
    if op == "style":
        return style_range(
            workbook, operation["sheet"], operation["range"], operation.get("font"),
            operation.get("font_size"),
        )
    if op == "clear":
        return clear_range(
            workbook, operation["sheet"], operation["range"], operation["with_style"]
        )
    if op == "row.insert":
        return row_insert(workbook, operation["sheet"], operation["before"], operation["count"])
    if op == "row.delete":
        return row_delete(workbook, operation["sheet"], operation["rows"])
    if op == "row.copy":
        return row_copy(
            workbook, operation["sheet"], operation["rows"], operation["insert_before"]
        )
    if op == "col.insert":
        return col_insert(
            workbook, operation["sheet"], operation["before"], operation["count"]
        )
    if op == "col.delete":
        return col_delete(workbook, operation["sheet"], operation["cols"])
    if op == "col.copy":
        return col_copy(
            workbook, operation["sheet"], operation["cols"], operation["insert_before"]
        )
    if op == "col.autofit":
        return col_autofit(
            workbook, operation["sheet"], operation["cols"],
            operation["max_width_mm"],
        )
    if op == "sheet.add":
        return sheet_add(workbook, operation["name"])
    if op == "sheet.delete":
        return sheet_delete(workbook, operation["sheet"])
    if op == "sheet.rename":
        return sheet_rename(workbook, operation["sheet"], operation["name"])
    if op == "sheet.copy":
        return sheet_copy(workbook, operation["sheet"], operation["name"])
    raise TargetError("不支持的 patch 操作: %s" % op)


def patch_operation(document):
    operations = document["operations"]

    def apply(workbook):
        plan = VerificationPlan()
        results = []
        summary_sheets = []
        for index, operation in enumerate(operations):
            try:
                operation_changes, _ = execute_operation(workbook, operation)
                plan.record(operation, index, operation_changes)
            except Exception as exc:
                raise operation_error(index, operation, exc)
            result = operation_identity(index, operation)
            result["verified"] = True
            if operation["op"] == "col.autofit":
                result["changes"] = operation_changes
            results.append(result)
            for name in operation_sheet_names(operation):
                if name not in summary_sheets:
                    summary_sheets.append(name)
        try:
            expected = plan.snapshot(workbook)
        except Exception as exc:
            last_index = len(operations) - 1
            raise operation_error(last_index, operations[last_index], exc)

        def verify(reopened):
            try:
                plan.verify(reopened, expected)
            except Exception as exc:
                last_index = len(operations) - 1
                raise operation_error(last_index, operations[last_index], exc)

        return {
            "operations": results,
            "summary": {
                "operation_count": len(results),
                "sheets": summary_sheets,
                "saved_once": True,
            },
        }, verify

    return apply


def operation_sheet_names(operation):
    names = []
    if "sheet" in operation:
        names.append(operation["sheet"])
    if operation["op"] in ("sheet.add", "sheet.rename", "sheet.copy"):
        names.append(operation["name"])
    return names
