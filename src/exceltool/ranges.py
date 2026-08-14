import re

from .errors import TargetError


CELL_RE = re.compile(r"^([A-Za-z]+)([1-9][0-9]*)$")


def column_number(value):
    value = str(value).strip().upper()
    if value.isdigit():
        number = int(value)
        if number < 1:
            raise TargetError("列号必须大于 0")
        return number
    if not value.isalpha():
        raise TargetError("非法列标记: %s" % value)
    number = 0
    for char in value:
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def column_name(number):
    if number < 1:
        raise TargetError("列号必须大于 0")
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def cell_ref(value):
    match = CELL_RE.match(value.strip())
    if not match:
        raise TargetError("非法单元格地址: %s" % value)
    return int(match.group(2)) - 1, column_number(match.group(1)) - 1


def simple_range(value, parser, label):
    parts = value.strip().split(":")
    if len(parts) > 2 or not parts[0]:
        raise TargetError("%s 格式应为 start:end 或单个值" % label)
    start = parser(parts[0])
    end = parser(parts[1]) if len(parts) == 2 else start
    if end < start:
        raise TargetError("%s 结束位置不能小于开始位置" % label)
    return start - 1, end


def row_range(value):
    def parse(raw):
        try:
            parsed = int(raw.strip())
        except ValueError:
            parsed = 0
        if parsed < 1:
            raise TargetError("行号必须是正整数: %s" % raw)
        return parsed

    return simple_range(value, parse, "行范围")


def col_range(value):
    return simple_range(value, column_number, "列范围")


def cell_range(value):
    parts = value.strip().split(":")
    if len(parts) == 1:
        parts.append(parts[0])
    if len(parts) != 2:
        raise TargetError("范围格式应为 A1:D20 或单个单元格")
    start_row, start_col = cell_ref(parts[0])
    end_row, end_col = cell_ref(parts[1])
    if end_row < start_row or end_col < start_col:
        raise TargetError("结束单元格不能小于开始单元格")
    return start_row, end_row + 1, start_col, end_col + 1


def resolve_window(rows, cols, range_value, from_value, size, tail, max_rows, max_cols):
    if tail is not None:
        if isinstance(tail, bool) or not isinstance(tail, int) or tail < 1:
            raise TargetError("--tail 必须是正整数")
        if rows or range_value or from_value or size is not None:
            raise TargetError("--tail 不能与 --rows、--range 或 --from/--n 同时使用")
        col_start, col_end = col_range(cols) if cols else (0, max_cols)
        window = max(0, max_rows - tail), max_rows, col_start, col_end
    elif from_value is not None or size is not None:
        if not from_value or size is None or size < 1:
            raise TargetError("--from 必须和正整数 --n 一起使用")
        row, col = cell_ref(from_value)
        window = row, row + size, col, col + size
    elif range_value:
        window = cell_range(range_value)
    else:
        row_start, row_end = row_range(rows) if rows else (0, max_rows)
        col_start, col_end = col_range(cols) if cols else (0, max_cols)
        window = row_start, row_end, col_start, col_end
    row_start, row_end, col_start, col_end = window
    return (
        min(max(row_start, 0), max_rows),
        min(max(row_end, 0), max_rows),
        min(max(col_start, 0), max_cols),
        min(max(col_end, 0), max_cols),
    )
