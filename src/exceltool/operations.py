import math

from .errors import TargetError, VerificationError
from .ranges import cell_range, cell_ref, column_name, resolve_window, row_range


HARD_STYLE_PROPERTIES = (
    "CharFontName",
    "CharFontNameAsian",
    "CharFontNameComplex",
    "CharHeight",
    "CharHeightAsian",
    "CharHeightComplex",
    "CharWeight",
    "CharWeightAsian",
    "CharWeightComplex",
    "CellBackColor",
    "NumberFormat",
)


def validate_font_size(font_size):
    if font_size is None:
        return
    if not math.isfinite(float(font_size)) or float(font_size) <= 0:
        raise TargetError("字体大小必须是有限的正数")


def is_direct_property(cell, name):
    state = cell.getPropertyState(name)
    return getattr(state, "value", str(state)) == "DIRECT_VALUE"


def copy_style_signature(cell):
    return (
        cell.CellStyle,
        cell.CharFontName,
        cell.CharFontNameAsian,
        cell.CharFontNameComplex,
        float(cell.CharHeight),
        float(cell.CharHeightAsian),
        float(cell.CharHeightComplex),
        float(cell.CharWeight),
        float(cell.CharWeightAsian),
        float(cell.CharWeightComplex),
        int(cell.CellBackColor),
        int(cell.NumberFormat),
    )


def cell_style(cell):
    return {
        "font": {
            "western": cell.CharFontName,
            "asian": cell.CharFontNameAsian,
            "complex": cell.CharFontNameComplex,
        },
        "font_size": {
            "western": float(cell.CharHeight),
            "asian": float(cell.CharHeightAsian),
            "complex": float(cell.CharHeightComplex),
        },
    }


def content_signature(cell):
    return cell.Type.value, cell.String, cell.Formula, float(cell.Value)


def read_sheet(workbook, sheet_name, options):
    sheet = workbook.sheet(sheet_name)
    max_rows, max_cols = workbook.used_size(sheet)
    start_row, end_row, start_col, end_col = resolve_window(
        options.get("rows"),
        options.get("cols"),
        options.get("range"),
        options.get("from"),
        options.get("n"),
        max_rows,
        max_cols,
    )
    json_values = options.get("json_values", False)
    values = []
    styles = [] if options.get("include_style", False) else None
    for row in range(start_row, end_row):
        values.append([
            (
                workbook.json_value(sheet.getCellByPosition(col, row))
                if json_values
                else workbook.display_value(sheet.getCellByPosition(col, row), options.get("value_mode", "display"))
            )
            for col in range(start_col, end_col)
        ])
        if styles is not None:
            styles.append([
                cell_style(sheet.getCellByPosition(col, row))
                for col in range(start_col, end_col)
            ])
    range_name = None
    if start_row < end_row and start_col < end_col:
        range_name = "%s%d:%s%d" % (
            column_name(start_col + 1), start_row + 1,
            column_name(end_col), end_row,
        )
    result = {
        "sheet": sheet_name,
        "range": range_name,
        "row_start": start_row,
        "col_start": start_col,
        "values": values,
    }
    if styles is not None:
        result["styles"] = styles
    return result


def view_workbook(workbook, sheet_name, options):
    names = workbook.sheet_names()
    if sheet_name:
        if sheet_name not in names:
            raise TargetError("未找到 sheet: %s" % sheet_name)
        names = [sheet_name]
    return [read_sheet(workbook, name, options) for name in names]


def set_cell(workbook, sheet_name, address, value=None, value_type="string", font=None, font_size=None):
    validate_font_size(font_size)
    row, col = cell_ref(address)
    sheet = workbook.sheet(sheet_name)
    cell = sheet.getCellByPosition(col, row)
    if value is not None:
        if value_type == "string":
            cell.String = value
        elif value_type == "number":
            try:
                cell.Value = float(value)
            except ValueError:
                raise TargetError("number 类型需要有效数字: %s" % value)
        elif value_type == "bool":
            lowered = value.lower()
            if lowered not in ("true", "false", "1", "0"):
                raise TargetError("bool 类型只接受 true/false/1/0")
            cell.Formula = "=TRUE()" if lowered in ("true", "1") else "=FALSE()"
        elif value_type == "formula":
            cell.Formula = value if value.startswith("=") else "=" + value
        else:
            raise TargetError("未知值类型: %s" % value_type)
    apply_font(cell, font, font_size)
    expected_formula = cell.Formula
    expected_string = cell.String
    def verify(reopened):
        target = reopened.sheet(sheet_name).getCellByPosition(col, row)
        if target.Formula != expected_formula or target.String != expected_string:
            raise VerificationError("单元格写后验证失败: %s!%s" % (sheet_name, address))
        if not verify_font(target, font, font_size):
            raise VerificationError(
                "单元格样式写后验证失败: %s!%s；%s"
                % (sheet_name, address, font_verification_error(target, font, font_size))
            )

    changes = {"sheet": sheet_name, "cell": address, "value": expected_string}

    def verified(reopened):
        verify(reopened)
        if font is not None:
            target = reopened.sheet(sheet_name).getCellByPosition(col, row)
            changes["actual_font"] = cell_style(target)["font"]

    return changes, verified


def apply_font(cell, font=None, font_size=None):
    validate_font_size(font_size)
    target = cell
    if cell.Type.value == "TEXT" and cell.String:
        target = cell.createTextCursor()
        target.gotoEnd(True)
    if font is not None:
        target.CharFontName = font
        target.CharFontNameAsian = font
        target.CharFontNameComplex = font
    if font_size is not None:
        target.CharHeight = float(font_size)
        target.CharHeightAsian = float(font_size)
        target.CharHeightComplex = float(font_size)


def verify_font(cell, font=None, font_size=None):
    if font is not None:
        actual = cell_style(cell)["font"]
        if any(actual[slot] != font for slot in required_font_slots(cell)):
            return False
    if font_size is not None and any(
        abs(float(value) - float(font_size)) > 0.01
        for value in (cell.CharHeight, cell.CharHeightAsian, cell.CharHeightComplex)
    ):
        return False
    return True


def required_font_slots(cell):
    text = cell.Formula if cell.Type.value == "FORMULA" else cell.String
    has_asian = any(
        0x2E80 <= ord(character) <= 0x9FFF
        or 0xAC00 <= ord(character) <= 0xD7AF
        or 0xF900 <= ord(character) <= 0xFAFF
        for character in text
    )
    has_complex = any(
        0x0590 <= ord(character) <= 0x08FF
        for character in text
    )
    has_western = not (has_asian or has_complex) or any(
        character.isascii() and character.isalnum()
        for character in text
    )
    slots = []
    if has_western:
        slots.append("western")
    if has_asian:
        slots.append("asian")
    if has_complex:
        slots.append("complex")
    return slots


def font_verification_error(cell, font=None, font_size=None):
    actual = cell_style(cell)
    return "请求字体=%s 字号=%s；实际字体=%s 字号=%s" % (
        font if font is not None else "未修改",
        font_size if font_size is not None else "未修改",
        actual["font"],
        actual["font_size"],
    )


def assign_json_value(cell, value):
    if value is None:
        cell.clearContents(23)
    elif isinstance(value, bool):
        cell.Formula = "=TRUE()" if value else "=FALSE()"
    elif isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise TargetError("JSON 数字必须是有限值")
        cell.Value = float(value)
    elif isinstance(value, str):
        if value.startswith("'="):
            cell.String = value[1:]
        elif value.startswith("="):
            cell.Formula = value
        else:
            cell.String = value
    else:
        raise TargetError("JSON 单元格只支持 null、字符串、数字和布尔值")


def normalized_json_value(value):
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def write_matrix(workbook, sheet_name, begin, matrix, font=None, font_size=None):
    start_row, start_col = cell_ref(begin)
    if not isinstance(matrix, list):
        raise TargetError("写入数据必须是二维 JSON 数组")
    if not matrix:
        raise TargetError("写入数据不能为空")
    sheet = workbook.sheet(sheet_name)
    touched = []
    max_columns = 0
    for row_offset, row_values in enumerate(matrix):
        if not isinstance(row_values, list):
            raise TargetError("二维 JSON 数组的每一行都必须是数组")
        max_columns = max(max_columns, len(row_values))
        for col_offset, value in enumerate(row_values):
            row = start_row + row_offset
            col = start_col + col_offset
            cell = sheet.getCellByPosition(col, row)
            assign_json_value(cell, value)
            apply_font(cell, font, font_size)
            touched.append((row, col, normalized_json_value(value)))
    if not touched:
        raise TargetError("写入数据没有包含任何单元格；空行 [] 只表示跳过")

    changes = {
        "sheet": sheet_name,
        "begin": "%s%d" % (column_name(start_col + 1), start_row + 1),
        "rows": len(matrix),
        "columns": max_columns,
        "cells": len(touched),
    }

    def verify(reopened):
        target_sheet = reopened.sheet(sheet_name)
        for row, col, expected_value in touched:
            cell = target_sheet.getCellByPosition(col, row)
            if reopened.json_value(cell) != expected_value:
                raise VerificationError("批量写入验证失败: %s%d" % (column_name(col + 1), row + 1))
            if not verify_font(cell, font, font_size):
                address = "%s%d" % (column_name(col + 1), row + 1)
                raise VerificationError(
                    "批量写入样式验证失败: %s!%s；%s"
                    % (sheet_name, address, font_verification_error(cell, font, font_size))
                )
        if font is not None:
            first_row, first_col, _ = touched[0]
            changes["actual_font"] = cell_style(
                target_sheet.getCellByPosition(first_col, first_row)
            )["font"]

    return changes, verify


def style_range(workbook, sheet_name, range_value, font=None, font_size=None):
    if font is None and font_size is None:
        raise TargetError("style 至少指定 --font 或 --font-size")
    start_row, end_row, start_col, end_col = cell_range(range_value if ":" in range_value else "%s:%s" % (range_value, range_value))
    sheet = workbook.sheet(sheet_name)
    contents = {}
    for row in range(start_row, end_row):
        for col in range(start_col, end_col):
            cell = sheet.getCellByPosition(col, row)
            contents[(row, col)] = content_signature(cell)
            apply_font(cell, font, font_size)

    changes = {"sheet": sheet_name, "range": range_value, "font": font, "font_size": font_size}

    def verify(reopened):
        target_sheet = reopened.sheet(sheet_name)
        for (row, col), expected_content in contents.items():
            cell = target_sheet.getCellByPosition(col, row)
            address = "%s%d" % (column_name(col + 1), row + 1)
            if content_signature(cell) != expected_content:
                raise VerificationError("样式操作改变了单元格内容: %s!%s" % (sheet_name, address))
            if not verify_font(cell, font, font_size):
                raise VerificationError(
                    "样式写后验证失败: %s!%s；%s"
                    % (sheet_name, address, font_verification_error(cell, font, font_size))
                )
        if font is not None and contents:
            first_row, first_col = next(iter(contents))
            changes["actual_font"] = cell_style(
                target_sheet.getCellByPosition(first_col, first_row)
            )["font"]

    return changes, verify


def clear_range(workbook, sheet_name, range_value, with_style=False):
    start_row, end_row, start_col, end_col = cell_range(range_value if ":" in range_value else "%s:%s" % (range_value, range_value))
    sheet = workbook.sheet(sheet_name)
    prior_styles = {}
    flags = 23 + (32 + 256 if with_style else 0)
    for row in range(start_row, end_row):
        for col in range(start_col, end_col):
            cell = sheet.getCellByPosition(col, row)
            prior_styles[(row, col)] = (
                cell.CellStyle,
                cell.CharFontNameAsian,
                float(cell.CharHeightAsian),
            )
            cell.clearContents(flags)

    def verify(reopened):
        target_sheet = reopened.sheet(sheet_name)
        for (row, col), style in prior_styles.items():
            cell = target_sheet.getCellByPosition(col, row)
            if cell.Formula or cell.String:
                raise VerificationError("范围清空写后验证失败")
            if cell.CellStyle != style[0]:
                raise VerificationError("范围清空意外改变了单元格样式")
            if not with_style and (
                cell.CharFontNameAsian != style[1] or abs(float(cell.CharHeightAsian) - style[2]) > 0.01
            ):
                raise VerificationError("范围清空意外改变了样式")
            if with_style and any(is_direct_property(cell, name) for name in HARD_STYLE_PROPERTIES):
                raise VerificationError("范围清空未移除硬样式")

    return {"sheet": sheet_name, "range": range_value, "with_style": with_style}, verify


def clear_cell(workbook, sheet_name, address):
    row, col = cell_ref(address)
    # VALUE | DATETIME | STRING | FORMULA; preserve formatting and annotations.
    workbook.sheet(sheet_name).getCellByPosition(col, row).clearContents(23)

    def verify(reopened):
        cell = reopened.sheet(sheet_name).getCellByPosition(col, row)
        if cell.String or cell.Formula:
            raise VerificationError("单元格清空验证失败: %s!%s" % (sheet_name, address))

    return {"sheet": sheet_name, "cell": address}, verify


def sheet_add(workbook, name):
    if not name:
        raise TargetError("sheet 名不能为空")
    if name in workbook.sheet_names():
        raise TargetError("sheet 已存在: %s" % name)
    workbook.document.Sheets.insertNewByName(name, len(workbook.sheet_names()))

    def verify(reopened):
        if name not in reopened.sheet_names():
            raise VerificationError("新增 sheet 写后验证失败: %s" % name)

    return {"name": name}, verify


def sheet_delete(workbook, name):
    names = workbook.sheet_names()
    if name not in names:
        raise TargetError("未找到 sheet: %s" % name)
    if len(names) == 1:
        raise TargetError("不能删除工作簿中的最后一个 sheet")
    workbook.document.Sheets.removeByName(name)

    def verify(reopened):
        if name in reopened.sheet_names():
            raise VerificationError("删除 sheet 写后验证失败: %s" % name)

    return {"sheet": name}, verify


def sheet_rename(workbook, old_name, new_name):
    if not new_name:
        raise TargetError("新 sheet 名不能为空")
    if new_name in workbook.sheet_names():
        raise TargetError("sheet 已存在: %s" % new_name)
    workbook.sheet(old_name).Name = new_name

    def verify(reopened):
        names = reopened.sheet_names()
        if old_name in names or new_name not in names:
            raise VerificationError("重命名 sheet 写后验证失败")

    return {"sheet": old_name, "name": new_name}, verify


def sheet_copy(workbook, source_name, new_name):
    names = workbook.sheet_names()
    if source_name not in names:
        raise TargetError("未找到 sheet: %s" % source_name)
    if not new_name:
        raise TargetError("新 sheet 名不能为空")
    if new_name in names:
        raise TargetError("sheet 已存在: %s" % new_name)
    source_sheet = workbook.sheet(source_name)
    rows, cols = workbook.used_size(source_sheet)
    expected = [
        [source_sheet.getCellByPosition(col, row).Formula for col in range(cols)]
        for row in range(rows)
    ]
    expected_styles = [
        [copy_style_signature(source_sheet.getCellByPosition(col, row)) for col in range(cols)]
        for row in range(rows)
    ]
    expected_heights = [source_sheet.Rows.getByIndex(row).Height for row in range(rows)]
    workbook.document.Sheets.copyByName(source_name, new_name, len(names))

    def verify(reopened):
        if new_name not in reopened.sheet_names():
            raise VerificationError("复制 sheet 写后验证失败")
        copied = reopened.sheet(new_name)
        for row in range(rows):
            for col in range(cols):
                if copied.getCellByPosition(col, row).Formula != expected[row][col]:
                    raise VerificationError("复制 sheet 内容验证失败")
                if copy_style_signature(copied.getCellByPosition(col, row)) != expected_styles[row][col]:
                    raise VerificationError("复制 sheet 样式验证失败")
            if abs(copied.Rows.getByIndex(row).Height - expected_heights[row]) > 2:
                raise VerificationError("复制 sheet 行高验证失败")

    return {"sheet": source_name, "name": new_name}, verify


def row_insert(workbook, sheet_name, before, count):
    if before < 1 or count < 1:
        raise TargetError("行号和数量必须为正整数")
    sheet = workbook.sheet(sheet_name)
    sentinel = sheet.getCellByPosition(0, before - 1).Formula
    sheet.Rows.insertByIndex(before - 1, count)

    def verify(reopened):
        target = reopened.sheet(sheet_name)
        if sentinel and target.getCellByPosition(0, before - 1 + count).Formula != sentinel:
            raise VerificationError("插入行写后验证失败")

    return {"sheet": sheet_name, "before": before, "count": count}, verify


def row_delete(workbook, sheet_name, rows):
    start, end = row_range(rows)
    count = end - start
    sheet = workbook.sheet(sheet_name)
    next_value = sheet.getCellByPosition(0, end).Formula
    sheet.Rows.removeByIndex(start, count)

    def verify(reopened):
        if next_value and reopened.sheet(sheet_name).getCellByPosition(0, start).Formula != next_value:
            raise VerificationError("删除行写后验证失败")

    return {"sheet": sheet_name, "rows": rows}, verify


def row_copy(workbook, sheet_name, rows, insert_before):
    start, end = row_range(rows)
    count = end - start
    if insert_before < 1:
        raise TargetError("目标行必须为正整数")
    sheet = workbook.sheet(sheet_name)
    _, max_cols = workbook.used_size(sheet)
    source_formulas = [
        [sheet.getCellByPosition(col, row).Formula for col in range(max_cols)]
        for row in range(start, end)
    ]
    heights = [sheet.Rows.getByIndex(row).Height for row in range(start, end)]
    target = insert_before - 1
    sheet.Rows.insertByIndex(target, count)
    adjusted_start = start + count if target <= start else start
    source_range = sheet.getCellRangeByPosition(0, adjusted_start, max_cols - 1, adjusted_start + count - 1)
    destination = sheet.getCellByPosition(0, target).CellAddress
    sheet.copyRange(destination, source_range.RangeAddress)
    for offset, height in enumerate(heights):
        sheet.Rows.getByIndex(target + offset).Height = height

    def verify(reopened):
        target_sheet = reopened.sheet(sheet_name)
        copied = [
            [target_sheet.getCellByPosition(col, target + row).Formula for col in range(max_cols)]
            for row in range(count)
        ]
        # Relative formulas may be adjusted by Calc. Literal cells must remain exact.
        for source_row, copied_row in zip(source_formulas, copied):
            for source_value, copied_value in zip(source_row, copied_row):
                if not str(source_value).startswith("=") and source_value != copied_value:
                    raise VerificationError("复制行写后验证失败")
        for offset, expected_height in enumerate(heights):
            actual_height = target_sheet.Rows.getByIndex(target + offset).Height
            if abs(actual_height - expected_height) > 2:
                raise VerificationError("复制行高度写后验证失败")

    return {"sheet": sheet_name, "rows": rows, "insert_before": insert_before}, verify
