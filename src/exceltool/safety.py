import hashlib
import re

from .errors import FormulaVerificationError, InputHashMismatchError, TargetError
from .ranges import cell_range, cell_ref, column_name


FORMULA_CELL_FLAG = 16
MAX_REPORTED_FORMULA_CHANGES = 100
STRUCTURAL_OPERATIONS = {
    "sheet.add",
    "sheet.delete",
    "sheet.rename",
    "sheet.copy",
    "row.insert",
    "row.delete",
    "row.copy",
    "col.insert",
    "col.delete",
    "col.copy",
}


def sha256_file(path):
    digest = hashlib.sha256()
    with open(str(path), "rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def normalize_expected_sha256(value):
    if value is None:
        return None
    normalized = str(value).strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise TargetError("--expect-sha256 必须是 64 位十六进制 SHA256")
    return normalized


def verify_input_sha256(path, expected_sha256=None):
    expected = normalize_expected_sha256(expected_sha256)
    actual = sha256_file(path)
    if expected is not None and actual != expected:
        raise InputHashMismatchError(
            "输入文件 SHA256 不匹配", expected, actual
        )
    return actual


def verify_source_unchanged(path, initial_sha256):
    actual = sha256_file(path)
    if actual != initial_sha256:
        raise InputHashMismatchError(
            "输入文件在编辑期间已被修改，拒绝覆盖", initial_sha256, actual
        )


def workbook_formula_snapshot(workbook):
    snapshot = {}
    for sheet_name in workbook.sheet_names():
        sheet = workbook.sheet(sheet_name)
        formulas = {}
        ranges = sheet.queryContentCells(FORMULA_CELL_FLAG).getRangeAddresses()
        for address in ranges:
            for row in range(address.StartRow, address.EndRow + 1):
                for col in range(address.StartColumn, address.EndColumn + 1):
                    cell = sheet.getCellByPosition(col, row)
                    if cell.Type.value == "FORMULA":
                        formulas[(row, col)] = cell.Formula
        snapshot[sheet_name] = formulas
    return snapshot


def formula_count(snapshot):
    return sum(len(formulas) for formulas in snapshot.values())


def _allowed_write_cells(sheet_name, begin, values):
    start_row, start_col = cell_ref(begin)
    return {
        (start_row + row_offset, start_col + col_offset)
        for row_offset, row_values in enumerate(values)
        for col_offset in range(len(row_values))
    }


def _allowed_range_cells(range_value):
    normalized = range_value if ":" in range_value else "%s:%s" % (range_value, range_value)
    start_row, end_row, start_col, end_col = cell_range(normalized)
    return {
        (row, col)
        for row in range(start_row, end_row)
        for col in range(start_col, end_col)
    }


def formula_policy_for_edit(operation_name, sheet_name=None, begin=None,
                            values=None, range_value=None):
    allowed = {}
    if operation_name == "write":
        allowed[sheet_name] = _allowed_write_cells(sheet_name, begin, values)
    elif operation_name == "clear":
        allowed[sheet_name] = _allowed_range_cells(range_value)
    return {
        "enforce_declared_changes": operation_name not in STRUCTURAL_OPERATIONS,
        "allowed_changes": allowed,
    }


def formula_policy_for_patch(operations):
    allowed = {}
    structural = False
    for operation in operations:
        op = operation["op"]
        structural = structural or op in STRUCTURAL_OPERATIONS
        if op == "write":
            allowed.setdefault(operation["sheet"], set()).update(
                _allowed_write_cells(
                    operation["sheet"], operation["begin"], operation["values"]
                )
            )
        elif op == "clear":
            allowed.setdefault(operation["sheet"], set()).update(
                _allowed_range_cells(operation["range"])
            )
    return {
        "enforce_declared_changes": not structural,
        "allowed_changes": allowed,
    }


def _cell_state(workbook, sheet_name, row, col):
    if sheet_name not in workbook.sheet_names():
        return {"after_type": "MISSING_SHEET", "after": None}
    cell = workbook.sheet(sheet_name).getCellByPosition(col, row)
    cell_type = cell.Type.value
    state = {"after_type": cell_type}
    if cell_type == "FORMULA":
        state["after_formula"] = cell.Formula
    elif cell_type == "VALUE":
        value = float(cell.Value)
        state["after"] = int(value) if value.is_integer() else value
    elif cell_type == "TEXT":
        state["after"] = cell.String
    else:
        state["after"] = None
    return state


def formula_differences(expected, actual, actual_workbook, stage,
                        allowed_changes=None):
    allowed_changes = allowed_changes or {}
    differences = []
    for sheet_name in sorted(set(expected) | set(actual)):
        expected_cells = expected.get(sheet_name, {})
        actual_cells = actual.get(sheet_name, {})
        allowed = allowed_changes.get(sheet_name, set())
        for row, col in sorted(set(expected_cells) | set(actual_cells)):
            if (row, col) in allowed:
                continue
            before_formula = expected_cells.get((row, col))
            after_formula = actual_cells.get((row, col))
            if before_formula == after_formula:
                continue
            difference = {
                "stage": stage,
                "sheet": sheet_name,
                "cell": "%s%d" % (column_name(col + 1), row + 1),
                "before_formula": before_formula,
            }
            difference.update(_cell_state(actual_workbook, sheet_name, row, col))
            differences.append(difference)
    return differences


def assert_formula_snapshots(expected, actual, actual_workbook, stage,
                             allowed_changes=None):
    differences = formula_differences(
        expected, actual, actual_workbook, stage, allowed_changes
    )
    if not differences:
        return
    first = differences[0]
    message = "发现未声明的公式变化: %s!%s" % (first["sheet"], first["cell"])
    report = {
        "checked": True,
        "stage": stage,
        "unexpected_changes": len(differences),
        "reported_changes": min(len(differences), MAX_REPORTED_FORMULA_CHANGES),
        "truncated": len(differences) > MAX_REPORTED_FORMULA_CHANGES,
    }
    raise FormulaVerificationError(
        message, report, differences[:MAX_REPORTED_FORMULA_CHANGES]
    )


def formula_verification_report(original, planned, reopened, policy):
    return {
        "checked": True,
        "operation_guard": (
            "exact_outside_declared"
            if policy.get("enforce_declared_changes", False)
            else "structural_operations"
        ),
        "before_operation_count": formula_count(original),
        "before_save_count": formula_count(planned),
        "after_save_count": formula_count(reopened),
        "unexpected_changes": 0,
    }
