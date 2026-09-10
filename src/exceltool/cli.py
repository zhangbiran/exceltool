import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .editing import edit_file, validate_input
from .engine import LibreOfficeSession
from .errors import ExcelToolError, TargetError
from .fonts import inspect_font
from .patching import load_patch, operation_error, patch_operation
from .operations import (
    clear_range,
    col_autofit,
    col_copy,
    col_delete,
    col_insert,
    find_cells,
    prepare_find_query,
    row_copy,
    row_delete,
    row_insert,
    sheet_add,
    sheet_copy,
    sheet_delete,
    sheet_information,
    sheet_rename,
    style_range,
    view_workbook,
    write_comment,
    write_matrix,
)
from .output import render_table
from .ranges import column_name
from .safety import formula_policy_for_edit, formula_policy_for_patch


class ExcelToolArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        if "--json" in sys.argv[1:] or "--json-full" in sys.argv[1:]:
            print(
                json.dumps({"ok": False, "error": "参数错误: %s" % message, "code": 2}, ensure_ascii=False),
                file=sys.stderr,
            )
            raise SystemExit(2)
        super().error(message)


class StoreOnce(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        if getattr(namespace, self.dest, None) is not None:
            parser.error("%s 只能指定一次" % option_string)
        setattr(namespace, self.dest, values)


def add_json(parser):
    parser.add_argument("--json", action="store_true", help="输出机器可读 JSON")


def add_edit_output(parser):
    parser.add_argument("--out", help="另存为新文件；省略时修改输入文件")
    parser.add_argument("--overwrite", action="store_true", help="允许覆盖已存在的 --out")
    parser.add_argument(
        "--expect-sha256",
        action=StoreOnce,
        help="要求输入文件匹配指定 SHA256，并保护就地发布",
    )
    add_json(parser)


def add_file_sheet(parser, sheet_required=True):
    parser.add_argument("--file", required=True, action=StoreOnce, help="输入 .xls/.xlsx 文件")
    parser.add_argument("--sheet", required=sheet_required, help="sheet 名")


def add_file(parser):
    parser.add_argument("--file", required=True, action=StoreOnce, help="输入 .xls/.xlsx 文件")


def add_window_selection(parser):
    parser.add_argument("--rows", help="行范围，如 1:20")
    parser.add_argument("--tail", type=int, help="有效区域最后 N 行")
    parser.add_argument("--cols", help="列范围，如 A:F 或 1:6")
    parser.add_argument("--range", dest="cell_range", help="单元格范围，如 A1:F20")
    parser.add_argument("--from", dest="from_cell", help="正方区域起点")
    parser.add_argument("--n", type=int, help="正方区域大小")


def window_options(args):
    if args.cell_range and (
        args.rows or args.cols or args.from_cell or args.n is not None or args.tail is not None
    ):
        raise TargetError("--range 不能与其他范围参数同时使用")
    if (args.from_cell or args.n is not None) and (args.rows or args.cols):
        raise TargetError("--from/--n 不能与 --rows/--cols 同时使用")
    if args.tail is not None:
        if args.tail < 1:
            raise TargetError("--tail 必须是正整数")
        if args.rows or args.cell_range or args.from_cell or args.n is not None:
            raise TargetError("--tail 不能与 --rows、--range 或 --from/--n 同时使用")
    return {
        "rows": args.rows,
        "tail": args.tail,
        "cols": args.cols,
        "range": args.cell_range,
        "from": args.from_cell,
        "n": args.n,
    }


def build_parser():
    parser = ExcelToolArgumentParser(prog="exceltool", description="查看和局部编辑 .xls/.xlsx")
    parser.add_argument("--version", action="version", version="exceltool %s" % __version__)
    commands = parser.add_subparsers(dest="command", required=True)

    view = commands.add_parser("view", help="查看工作簿局部区域")
    add_file(view)
    view.add_argument("--sheet", help="sheet 名；省略则查看全部")
    add_window_selection(view)
    view.add_argument("--include-style", action="store_true", help="在 --json-full 中包含字体和字号")
    view_json = view.add_mutually_exclusive_group()
    view_json.add_argument("--json", action="store_true", help="输出可回写的纯二维 JSON 数组")
    view_json.add_argument("--json-full", action="store_true", help="输出带文件、sheet 和范围元数据的 JSON")

    write = commands.add_parser("write", help="从起始单元格批量写入二维 JSON")
    add_file_sheet(write)
    write.add_argument("--begin", required=True, help="起始单元格，如 F3 或 AA3")
    write.add_argument("data", nargs="?", help="二维 JSON 数组")
    write.add_argument("--stdin", action="store_true", help="从标准输入读取二维 JSON")
    write.add_argument("--values-file", help="从 UTF-8 JSON 文件读取二维数组")
    write.add_argument("--font", help="写入位置的字体名称")
    write.add_argument("--font-size", type=float, help="写入位置的字号")
    add_edit_output(write)

    patch = commands.add_parser("patch", help="顺序执行一个工作簿的 JSON 操作批次")
    add_file(patch)
    patch.add_argument("--patch", required=True, help="UTF-8 patch JSON 文件")
    add_edit_output(patch)

    style = commands.add_parser("style", help="设置范围字体和字号")
    add_file_sheet(style)
    style.add_argument("--range", required=True, dest="cell_range", help="范围或单元格，如 A1:F10 或 B3")
    style.add_argument("--font", help="字体名称")
    style.add_argument("--font-size", type=float, help="字号")
    add_edit_output(style)

    clear = commands.add_parser("clear", help="清空范围内容")
    add_file_sheet(clear)
    clear.add_argument("--range", required=True, dest="cell_range", help="范围或单元格，如 B3:F10 或 B3")
    clear.add_argument("--clear-style", action="store_true", help="同时清除硬样式")
    add_edit_output(clear)

    comment = commands.add_parser("comment", help="写入或替换单元格批注")
    add_file_sheet(comment)
    comment.add_argument("--cell", required=True, help="单元格，如 B3")
    comment.add_argument("--text", required=True, help="批注文本")
    add_edit_output(comment)

    sheet = commands.add_parser("sheet", help="sheet 查看和编辑")
    sheet_commands = sheet.add_subparsers(dest="sheet_command", required=True)
    sheet_list_parser = sheet_commands.add_parser("list", help="列出 sheet")
    add_file(sheet_list_parser)
    add_json(sheet_list_parser)
    sheet_info_parser = sheet_commands.add_parser("info", help="查看 sheet 有效范围和行列数")
    add_file_sheet(sheet_info_parser, sheet_required=False)
    add_json(sheet_info_parser)
    sheet_add_parser = sheet_commands.add_parser("add", help="新增 sheet")
    add_file(sheet_add_parser)
    sheet_add_parser.add_argument("--name", required=True, help="新 sheet 名")
    add_edit_output(sheet_add_parser)
    sheet_delete_parser = sheet_commands.add_parser("delete", help="删除 sheet")
    add_file_sheet(sheet_delete_parser)
    add_edit_output(sheet_delete_parser)
    sheet_rename_parser = sheet_commands.add_parser("rename", help="重命名 sheet")
    add_file_sheet(sheet_rename_parser)
    sheet_rename_parser.add_argument("--name", required=True, help="新 sheet 名")
    add_edit_output(sheet_rename_parser)
    sheet_copy_parser = sheet_commands.add_parser("copy", help="复制 sheet")
    add_file_sheet(sheet_copy_parser)
    sheet_copy_parser.add_argument("--name", required=True, help="副本 sheet 名")
    add_edit_output(sheet_copy_parser)

    row = commands.add_parser("row", help="行结构编辑")
    row_commands = row.add_subparsers(dest="row_command", required=True)
    row_insert_parser = row_commands.add_parser("insert", help="在指定行前插入空行")
    add_file_sheet(row_insert_parser)
    row_insert_parser.add_argument("--before", type=int, required=True, help="在该行号之前插入")
    row_insert_parser.add_argument("--count", type=int, default=1, help="插入行数，默认 1")
    add_edit_output(row_insert_parser)
    row_delete_parser = row_commands.add_parser("delete", help="删除行范围")
    add_file_sheet(row_delete_parser)
    row_delete_parser.add_argument("--rows", required=True, help="删除行范围，如 3:5")
    add_edit_output(row_delete_parser)
    row_copy_parser = row_commands.add_parser("copy", help="复制行并在目标行前插入")
    add_file_sheet(row_copy_parser)
    row_copy_parser.add_argument("--rows", required=True, help="源行范围，如 3:5")
    row_copy_parser.add_argument("--insert-before", type=int, required=True, help="将副本插入该行之前")
    add_edit_output(row_copy_parser)

    col = commands.add_parser("col", help="列结构和宽度编辑")
    col_commands = col.add_subparsers(dest="col_command", required=True)
    col_insert_parser = col_commands.add_parser("insert", help="在指定列前插入空列")
    add_file_sheet(col_insert_parser)
    col_insert_parser.add_argument("--before", required=True, help="目标列，如 F 或 AA")
    col_insert_parser.add_argument("--count", type=int, default=1, help="插入列数，默认 1")
    add_edit_output(col_insert_parser)
    col_delete_parser = col_commands.add_parser("delete", help="删除列范围")
    add_file_sheet(col_delete_parser)
    col_delete_parser.add_argument("--cols", required=True, help="列范围，如 F:H")
    add_edit_output(col_delete_parser)
    col_copy_parser = col_commands.add_parser("copy", help="复制列并在目标列前插入")
    add_file_sheet(col_copy_parser)
    col_copy_parser.add_argument("--cols", required=True, help="源列范围，如 B:D")
    col_copy_parser.add_argument("--insert-before", required=True, help="将副本插入该列之前，如 F")
    add_edit_output(col_copy_parser)
    col_autofit_parser = col_commands.add_parser("autofit", help="自适应列宽并限制最大宽度")
    add_file_sheet(col_autofit_parser)
    col_autofit_parser.add_argument("--cols", required=True, help="列范围，如 A:F")
    col_autofit_parser.add_argument("--max-width-mm", type=float, default=60.0, help="最大列宽毫米数，默认 60")
    add_edit_output(col_autofit_parser)

    find = commands.add_parser("find", help="在工作簿中查找值或公式")
    add_file(find)
    find.add_argument("--text", required=True, help="查找文本")
    find.add_argument("--sheet", help="sheet 名；省略则查找全部")
    add_window_selection(find)
    find.add_argument("--look-in", choices=("values", "formulas", "both"), default="both", help="查找显示值、公式或两者，默认 both")
    find.add_argument("--case-sensitive", action="store_true", help="区分大小写")
    find.add_argument(
        "--regex",
        action="store_true",
        help="将 --text 作为 Python 正则",
    )
    find.add_argument("--limit", type=int, default=100, help="最大结果数，默认 100")
    add_json(find)

    font = commands.add_parser("font", help="字体环境检查")
    font_commands = font.add_subparsers(dest="font_command", required=True)
    font_check = font_commands.add_parser("check", help="检查字体能否精确匹配")
    font_check.add_argument("--name", required=True, help="字体名称")
    add_json(font_check)

    return parser


def run_view(args):
    if args.include_style and not args.json_full:
        raise TargetError("--include-style 必须与 --json-full 一起使用")
    options = window_options(args)
    options.update({
        "json_values": args.json or args.json_full,
        "include_style": args.include_style,
    })
    path = validate_input(args.file)
    if args.json and not args.sheet:
        raise TargetError("view --json 必须明确指定一个 --sheet；多 sheet 请使用 --json-full")
    with LibreOfficeSession() as session:
        workbook = session.load(path, read_only=True)
        try:
            sheets = view_workbook(workbook, args.sheet, options)
        finally:
            workbook.close()
    if args.json:
        print(json.dumps(sheets[0]["values"], ensure_ascii=False, indent=2))
    elif args.json_full:
        if len(sheets) == 1:
            sheet = sheets[0]
            payload = {
                "file": str(path),
                "sheet": sheet["sheet"],
                "range": sheet["range"],
                "begin": (
                    "%s%d" % (column_name(sheet["col_start"] + 1), sheet["row_start"] + 1)
                    if sheet["range"] else None
                ),
                "values": sheet["values"],
            }
            if "styles" in sheet:
                payload["styles"] = sheet["styles"]
        else:
            payload = {"file": str(path), "sheets": sheets}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print("\n========== %s ==========" % path)
        for sheet in sheets:
            print("\n--- Sheet: %s ---" % sheet["sheet"])
            if sheet["range"] is None:
                print("(所选范围内没有可显示内容)")
                continue
            print("范围: %s" % sheet["range"])
            print(render_table(sheet["values"], sheet["row_start"], sheet["col_start"]))
    return {"ok": True, "file": str(path), "sheets": sheets}


def load_write_data(args):
    sources = int(args.data is not None) + int(args.stdin) + int(args.values_file is not None)
    if sources != 1:
        raise TargetError("write 必须且只能从位置 JSON、--stdin 或 --values-file 读取数据")
    try:
        if args.stdin:
            raw = sys.stdin.read()
        elif args.values_file:
            raw = Path(args.values_file).read_text(encoding="utf-8")
        else:
            raw = args.data
        return json.loads(raw)
    except (OSError, json.JSONDecodeError) as exc:
        raise TargetError("无法读取二维 JSON: %s" % exc)


def run_sheet_list(args):
    path = validate_input(args.file)
    with LibreOfficeSession() as session:
        workbook = session.load(path, read_only=True)
        try:
            names = workbook.sheet_names()
        finally:
            workbook.close()
    if args.json:
        print(json.dumps(names, ensure_ascii=False, indent=2))
    else:
        for name in names:
            print(name)
    return {"ok": True, "file": str(path), "sheets": names}


def run_sheet_info(args):
    path = validate_input(args.file)
    with LibreOfficeSession() as session:
        workbook = session.load(path, read_only=True)
        try:
            sheets = sheet_information(workbook, args.sheet)
        finally:
            workbook.close()
    result = {"file": str(path), "sheets": sheets}
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for info in sheets:
            print("Sheet: %s" % info["sheet"])
            print("有效范围: %s" % info["used_range"])
            print("有效行数: %d" % info["used_rows"])
            print("有效列数: %d" % info["used_cols"])
    return result


def run_font_check(args):
    result = inspect_font(args.name)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("请求字体: %s" % result["requested"])
        print("解析字体: %s" % (result["resolved"] or "未找到"))
        print("精确可用: %s" % ("是" if result["exact"] else "否"))
    return result


def run_find(args):
    options = window_options(args)
    prepared_query = prepare_find_query(args.text, args.case_sensitive, args.regex)
    path = validate_input(args.file)
    with LibreOfficeSession() as session:
        workbook = session.load(path, read_only=True)
        try:
            matches, truncated = find_cells(
                workbook, args.text, args.sheet, options, args.look_in,
                args.case_sensitive, args.limit, args.regex, prepared_query,
            )
        finally:
            workbook.close()
    result = {
        "file": str(path),
        "query": args.text,
        "regex": args.regex,
        "case_sensitive": args.case_sensitive,
        "matches": matches,
        "truncated": truncated,
    }
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for match in matches:
            locations = ",".join(match["match_in"])
            print("%s!%s [%s] %s" % (
                match["sheet"], match["cell"], locations, match["display"]
            ))
        if truncated:
            print("结果已截断；使用 --limit 调整上限")
    return result


# Dispatch one editing command through the shared transactional lifecycle.
def run_edit(args):
    font_check = None
    requested_font = getattr(args, "font", None)
    if requested_font is not None:
        font_check = inspect_font(requested_font)
        if not font_check["exact"]:
            raise TargetError(
                "字体不可精确使用: 请求=%s，系统解析=%s"
                % (requested_font, font_check["resolved"] or "未找到")
            )
    if args.command == "write":
        operation_name = "write"
        matrix = load_write_data(args)
        operation = lambda book: write_matrix(
            book, args.sheet, args.begin, matrix, args.font, args.font_size
        )
        formula_policy = formula_policy_for_edit(
            operation_name, args.sheet, begin=args.begin, values=matrix
        )
    elif args.command == "style":
        operation_name = "style"
        operation = lambda book: style_range(
            book, args.sheet, args.cell_range, args.font, args.font_size
        )
        formula_policy = formula_policy_for_edit(operation_name)
    elif args.command == "clear":
        operation_name = "clear"
        operation = lambda book: clear_range(
            book, args.sheet, args.cell_range, args.clear_style
        )
        formula_policy = formula_policy_for_edit(
            operation_name, args.sheet, range_value=args.cell_range
        )
    elif args.command == "comment":
        operation_name = "comment"
        operation = lambda book: write_comment(
            book, args.sheet, args.cell, args.text
        )
        formula_policy = formula_policy_for_edit(operation_name)
    elif args.command == "sheet":
        if args.sheet_command == "add":
            operation_name = "sheet.add"
            operation = lambda book: sheet_add(book, args.name)
        elif args.sheet_command == "delete":
            operation_name = "sheet.delete"
            operation = lambda book: sheet_delete(book, args.sheet)
        elif args.sheet_command == "rename":
            operation_name = "sheet.rename"
            operation = lambda book: sheet_rename(book, args.sheet, args.name)
        else:
            operation_name = "sheet.copy"
            operation = lambda book: sheet_copy(book, args.sheet, args.name)
        formula_policy = formula_policy_for_edit(operation_name)
    elif args.command == "row":
        if args.row_command == "insert":
            operation_name = "row.insert"
            operation = lambda book: row_insert(book, args.sheet, args.before, args.count)
        elif args.row_command == "delete":
            operation_name = "row.delete"
            operation = lambda book: row_delete(book, args.sheet, args.rows)
        else:
            operation_name = "row.copy"
            operation = lambda book: row_copy(book, args.sheet, args.rows, args.insert_before)
        formula_policy = formula_policy_for_edit(operation_name)
    else:
        if args.col_command == "insert":
            operation_name = "col.insert"
            operation = lambda book: col_insert(book, args.sheet, args.before, args.count)
        elif args.col_command == "delete":
            operation_name = "col.delete"
            operation = lambda book: col_delete(book, args.sheet, args.cols)
        elif args.col_command == "copy":
            operation_name = "col.copy"
            operation = lambda book: col_copy(book, args.sheet, args.cols, args.insert_before)
        else:
            operation_name = "col.autofit"
            operation = lambda book: col_autofit(
                book, args.sheet, args.cols, args.max_width_mm
            )
        formula_policy = formula_policy_for_edit(operation_name)
    result = edit_file(
        args.file, args.out, args.overwrite, operation,
        formula_policy, args.expect_sha256,
    )
    result["operation"] = operation_name
    if font_check is not None:
        result["font_check"] = font_check
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("成功: %s" % operation_name)
        print("输出: %s" % result["output"])
        print("写后验证: 通过")
        formula_report = result["formula_verification"]
        print("公式验证: %d -> %d，意外变化: %d" % (
            formula_report["before_save_count"],
            formula_report["after_save_count"],
            formula_report["unexpected_changes"],
        ))
    return result


def run_patch(args):
    document = load_patch(args.patch)
    try:
        result = edit_file(
            args.file, args.out, args.overwrite, patch_operation(document),
            formula_policy_for_patch(document["operations"]), args.expect_sha256,
        )
    except Exception as exc:
        if hasattr(exc, "details"):
            raise
        last_index = len(document["operations"]) - 1
        raise operation_error(last_index, document["operations"][last_index], exc)
    changes = result.pop("changes")
    result["operation"] = "patch"
    result.update(changes)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("成功: patch")
        print("操作数: %d" % result["summary"]["operation_count"])
        print("输出: %s" % result["output"])
        print("写后验证: 通过")
        formula_report = result["formula_verification"]
        print("公式验证: %d -> %d，意外变化: %d" % (
            formula_report["before_save_count"],
            formula_report["after_save_count"],
            formula_report["unexpected_changes"],
        ))
    return result


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "view":
            run_view(args)
        elif args.command == "patch":
            run_patch(args)
        elif args.command == "sheet" and args.sheet_command == "list":
            run_sheet_list(args)
        elif args.command == "sheet" and args.sheet_command == "info":
            run_sheet_info(args)
        elif args.command == "font" and args.font_command == "check":
            run_font_check(args)
        elif args.command == "find":
            run_find(args)
        else:
            run_edit(args)
        return 0
    except ExcelToolError as exc:
        payload = {"ok": False, "error": exc.message, "code": exc.code}
        if args.command == "patch":
            payload.update({"operation": "patch", "published": False})
        if hasattr(exc, "details"):
            payload.update(exc.details)
        if getattr(args, "json", False) or getattr(args, "json_full", False):
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        else:
            print("错误: %s" % exc.message, file=sys.stderr)
        return exc.code
    except KeyboardInterrupt:
        print("错误: 操作已取消", file=sys.stderr)
        return 130
    except Exception as exc:
        payload = {"ok": False, "error": "操作失败: %s" % exc, "code": 5}
        if args.command == "patch":
            payload.update({"operation": "patch", "published": False})
        if getattr(args, "json", False) or getattr(args, "json_full", False):
            print(json.dumps(payload, ensure_ascii=False), file=sys.stderr)
        else:
            print("错误: %s" % payload["error"], file=sys.stderr)
        return 5


if __name__ == "__main__":
    sys.exit(main())
