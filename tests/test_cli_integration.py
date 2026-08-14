import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from exceltool.engine import LibreOfficeSession


ROOT = Path(__file__).resolve().parents[1]
TEST_FONT = "Microsoft YaHei"


def create_fixture(path, extra_sheet=False):
    with LibreOfficeSession() as session:
        workbook = session.create()
        try:
            sheet = workbook.sheet(workbook.sheet_names()[0])
            sheet.Name = "Data"
            sheet.Rows.getByIndex(1).Height = 800
            sheet.Columns.getByIndex(1).Width = 4200
            values = (
                ("ID", "Name", "Double"),
                (1, "Alpha", "=A2*2"),
                (2, "Beta", "=A3*2"),
                (3, "Sentinel", "=A4*2"),
            )
            for row_index, row in enumerate(values):
                for col_index, value in enumerate(row):
                    cell = sheet.getCellByPosition(col_index, row_index)
                    if isinstance(value, int):
                        cell.Value = float(value)
                    elif value.startswith("="):
                        cell.Formula = value
                    else:
                        cell.String = value
            sheet.getCellRangeByName("B2").CharWeight = 150.0
            sheet.getCellRangeByName("D2").String = "=literal"
            sheet.getCellRangeByName("E2").Formula = "=TRUE()"
            sheet.getCellRangeByName("F4").String = "KEEP"
            if extra_sheet:
                workbook.document.Sheets.insertNewByName("Extra", 1)
                workbook.sheet("Extra").getCellRangeByName("A1").String = "Alpha Extra"
            workbook.save_as(path)
        finally:
            workbook.close()


class CliIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="exceltool-test-")
        self.directory = Path(self.temp.name)
        self.env = os.environ.copy()
        self.env["PYTHONPATH"] = str(ROOT / "src")

    def tearDown(self):
        self.temp.cleanup()

    def run_cli(self, *arguments, expected=0, input_text=None):
        command = [sys.executable, "-m", "exceltool"] + list(arguments)
        result = subprocess.run(
            command,
            env=self.env,
            text=True,
            input=input_text,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.assertEqual(
            result.returncode,
            expected,
            "command: %s\nstdout:\n%s\nstderr:\n%s" % (command, result.stdout, result.stderr),
        )
        return result

    def inspect(self, path, callback):
        with LibreOfficeSession() as session:
            workbook = session.load(path, read_only=True)
            try:
                return callback(workbook)
            finally:
                workbook.close()

    def test_view_compatibility_and_errors(self):
        xlsx = self.directory / "view.xlsx"
        xls = self.directory / "view.xls"
        create_fixture(xlsx, extra_sheet=True)
        create_fixture(xls, extra_sheet=True)
        for path in (xlsx, xls):
            with self.subTest(path=path.suffix):
                result = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--rows", "2:3", "--cols", "A:C", "--json-full"
                )
                payload = json.loads(result.stdout)
                self.assertEqual(payload["range"], "A2:C3")
                self.assertEqual(payload["values"][0][:2], [1, "Alpha"])

                square = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--from", "B2", "--n", "2", "--json-full"
                )
                self.assertEqual(json.loads(square.stdout)["range"], "B2:C3")

                raw = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--range", "A2:A2", "--json"
                )
                self.assertEqual(json.loads(raw.stdout), [[1]])
                formula = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--range", "C2:C2", "--json"
                )
                self.assertTrue(json.loads(formula.stdout)[0][0].startswith("="))
                tail = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--tail", "2",
                    "--cols", "A:C", "--json-full",
                )
                tail_payload = json.loads(tail.stdout)
                self.assertEqual(tail_payload["range"], "A3:C4")
                self.assertEqual(tail_payload["values"][0][:2], [2, "Beta"])
                self.assertEqual(tail_payload["values"][1][:2], [3, "Sentinel"])
                oversized_tail = self.run_cli(
                    "view", "--file", str(path), "--sheet", "Data", "--tail", "10",
                    "--cols", "A:C", "--json-full",
                )
                self.assertEqual(json.loads(oversized_tail.stdout)["range"], "A1:C4")

        multiple = self.run_cli(
            "view", "--file", str(xlsx), "--range", "A1:A1", "--json-full"
        )
        multiple_payload = json.loads(multiple.stdout)
        self.assertEqual(multiple_payload["file"], str(xlsx.resolve()))
        self.assertEqual(
            [sheet["sheet"] for sheet in multiple_payload["sheets"]],
            ["Data", "Extra"],
        )

        missing = self.run_cli(
            "view", "--file", str(xlsx), "--sheet", "Missing", "--json", expected=3
        )
        self.assertFalse(json.loads(missing.stderr)["ok"])
        self.assertNotIn("Traceback", missing.stderr)
        conflict = self.run_cli(
            "view", "--file", str(xlsx), "--sheet", "Data", "--tail", "2",
            "--rows", "1:2", "--json-full", expected=3,
        )
        self.assertIn("不能与", conflict.stderr)
        invalid_tail = self.run_cli(
            "view", "--file", str(xlsx), "--sheet", "Data", "--tail", "0",
            "--json-full", expected=3,
        )
        self.assertIn("正整数", invalid_tail.stderr)
        removed_directory = self.run_cli(
            "view", "--file", str(xlsx), "--dir", str(self.directory),
            "--json-full", expected=2,
        )
        self.assertIn("unrecognized arguments", removed_directory.stderr)
        removed_value_mode = self.run_cli(
            "view", "--file", str(xlsx), "--value-mode", "raw",
            "--json-full", expected=2,
        )
        self.assertIn("unrecognized arguments", removed_value_mode.stderr)

    def test_single_cell_write_style_clear_and_default_original_for_both_formats(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("single-cell" + extension)
                output = self.directory / ("single-cell-edited" + extension)
                create_fixture(source)
                result = self.run_cli(
                    "write", "--file", str(source), "--sheet", "Data", "--begin", "B2",
                    '[["测试"]]', "--font", TEST_FONT, "--font-size", "14",
                    "--out", str(output), "--json",
                )
                self.assertTrue(json.loads(result.stdout)["verified"])

                def assert_edit(workbook):
                    sheet = workbook.sheet("Data")
                    target = sheet.getCellRangeByName("B2")
                    self.assertEqual(target.String, "测试")
                    self.assertAlmostEqual(target.CharHeightAsian, 14.0)
                    self.assertEqual(sheet.getCellRangeByName("B4").String, "Sentinel")

                self.inspect(output, assert_edit)
                styled_output = self.directory / ("single-styled" + extension)
                style_result = self.run_cli(
                    "style", "--file", str(output), "--sheet", "Data", "--range", "D2",
                    "--font-size", "15", "--out", str(styled_output), "--json",
                )
                self.assertTrue(json.loads(style_result.stdout)["verified"])

                def assert_single_style(workbook):
                    target = workbook.sheet("Data").getCellRangeByName("D2")
                    self.assertEqual(target.String, "=literal")
                    self.assertAlmostEqual(target.CharHeightAsian, 15.0)

                self.inspect(styled_output, assert_single_style)
                number_output = self.directory / ("single-number" + extension)
                bool_output = self.directory / ("single-bool" + extension)
                formula_output = self.directory / ("single-formula" + extension)
                self.run_cli(
                    "write", "--file", str(styled_output), "--sheet", "Data", "--begin", "A2",
                    "[[42.5]]", "--out", str(number_output), "--json",
                )
                self.run_cli(
                    "write", "--file", str(number_output), "--sheet", "Data", "--begin", "B2",
                    "[[true]]", "--out", str(bool_output), "--json",
                )
                self.run_cli(
                    "write", "--file", str(bool_output), "--sheet", "Data", "--begin", "C2",
                    '[["=A2*3"]]', "--out", str(formula_output), "--json",
                )

                def assert_types(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("A2").Value, 42.5)
                    self.assertIn(sheet.getCellRangeByName("B2").String.upper(), ("TRUE", "1"))
                    self.assertTrue(sheet.getCellRangeByName("B2").Formula.startswith("="))
                    self.assertTrue(sheet.getCellRangeByName("C2").Formula.startswith("="))

                self.inspect(formula_output, assert_types)
                self.run_cli(
                    "clear", "--file", str(formula_output), "--sheet", "Data", "--range", "B2", "--json"
                )
                self.assertFalse(Path(str(formula_output) + ".bak").exists())

                def assert_clear(workbook):
                    target = workbook.sheet("Data").getCellRangeByName("B2")
                    self.assertEqual(target.String, "")
                    self.assertAlmostEqual(target.CharHeightAsian, 14.0)

                self.inspect(formula_output, assert_clear)
                exists = self.run_cli(
                    "write", "--file", str(source), "--sheet", "Data", "--begin", "A2", "[[9]]",
                    "--out", str(output), "--json", expected=3,
                )
                self.assertNotIn("Traceback", exists.stderr)

    def test_sheet_info_for_both_formats(self):
        sources = []
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("sheet-info" + extension)
                sources.append(source)
                create_fixture(source, extra_sheet=True)
                result = self.run_cli(
                    "sheet", "info", "--file", str(source), "--json"
                )
                payload = json.loads(result.stdout)
                self.assertEqual(
                    [item["sheet"] for item in payload["sheets"]],
                    ["Data", "Extra"],
                )
                data = payload["sheets"][0]
                self.assertEqual(data["used_range"], "A1:F4")
                self.assertEqual(data["used_rows"], 4)
                self.assertEqual(data["used_cols"], 6)
                self.assertEqual(data["last_row"], 4)
                self.assertEqual(data["last_col"], "F")
                selected = self.run_cli(
                    "sheet", "info", "--file", str(source), "--sheet", "Extra",
                    "--json",
                )
                self.assertEqual(json.loads(selected.stdout)["sheets"], [{
                    "sheet": "Extra",
                    "used_range": "A1:A1",
                    "used_rows": 1,
                    "used_cols": 1,
                    "last_row": 1,
                    "last_col": "A",
                }])

        missing = self.run_cli(
            "sheet", "info", "--file", str(sources[0]), "--sheet", "Missing",
            "--json", expected=3,
        )
        self.assertIn("未找到 sheet", missing.stderr)

    def test_sheet_and_row_operations_for_both_formats(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("struct" + extension)
                added = self.directory / ("added" + extension)
                renamed = self.directory / ("renamed" + extension)
                deleted_sheet = self.directory / ("deleted-sheet" + extension)
                inserted = self.directory / ("inserted" + extension)
                deleted_row = self.directory / ("deleted-row" + extension)
                copied = self.directory / ("copied" + extension)
                create_fixture(source)

                self.run_cli("sheet", "add", "--file", str(source), "--name", "Extra", "--out", str(added), "--json")
                self.run_cli("sheet", "rename", "--file", str(added), "--sheet", "Extra", "--name", "Renamed", "--out", str(renamed), "--json")
                self.run_cli("sheet", "delete", "--file", str(renamed), "--sheet", "Renamed", "--out", str(deleted_sheet), "--json")
                self.inspect(deleted_sheet, lambda book: self.assertEqual(book.sheet_names(), ["Data"]))

                self.run_cli("row", "insert", "--file", str(deleted_sheet), "--sheet", "Data", "--before", "3", "--count", "1", "--out", str(inserted), "--json")
                self.inspect(inserted, lambda book: self.assertEqual(book.sheet("Data").getCellRangeByName("B4").String, "Beta"))
                self.run_cli("row", "delete", "--file", str(inserted), "--sheet", "Data", "--rows", "3", "--out", str(deleted_row), "--json")
                self.inspect(deleted_row, lambda book: self.assertEqual(book.sheet("Data").getCellRangeByName("B3").String, "Beta"))
                self.run_cli("row", "copy", "--file", str(deleted_row), "--sheet", "Data", "--rows", "2", "--insert-before", "4", "--out", str(copied), "--json")

                def assert_copy(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("B4").String, "Alpha")
                    self.assertEqual(sheet.getCellRangeByName("B5").String, "Sentinel")
                    self.assertEqual(sheet.getCellRangeByName("C4").Formula, "=A4*2")
                    self.assertLessEqual(abs(sheet.Rows.getByIndex(3).Height - 800), 2)
                    self.assertEqual(sheet.getCellRangeByName("B4").CharWeight, 150.0)

                self.inspect(copied, assert_copy)

    def test_column_operations_and_autofit_for_both_formats(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("columns" + extension)
                inserted = self.directory / ("columns-inserted" + extension)
                deleted = self.directory / ("columns-deleted" + extension)
                copied = self.directory / ("columns-copied" + extension)
                fitted = self.directory / ("columns-fitted" + extension)
                create_fixture(source)

                self.run_cli(
                    "col", "insert", "--file", str(source), "--sheet", "Data",
                    "--before", "B", "--count", "1", "--out", str(inserted), "--json",
                )
                self.inspect(
                    inserted,
                    lambda book: self.assertEqual(
                        book.sheet("Data").getCellRangeByName("C2").String, "Alpha"
                    ),
                )
                self.run_cli(
                    "col", "delete", "--file", str(inserted), "--sheet", "Data",
                    "--cols", "B:B", "--out", str(deleted), "--json",
                )
                self.run_cli(
                    "col", "copy", "--file", str(deleted), "--sheet", "Data",
                    "--cols", "B:C", "--insert-before", "F", "--out", str(copied),
                    "--json",
                )

                def assert_copy(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("F2").String, "Alpha")
                    self.assertEqual(sheet.getCellRangeByName("G2").Formula, "=E2*2")
                    self.assertEqual(sheet.getCellRangeByName("F2").CharWeight, 150.0)
                    self.assertLessEqual(
                        abs(
                            sheet.Columns.getByIndex(5).Width
                            - sheet.Columns.getByIndex(1).Width
                        ),
                        10,
                    )

                self.inspect(copied, assert_copy)
                result = self.run_cli(
                    "col", "autofit", "--file", str(copied), "--sheet", "Data",
                    "--cols", "F:G", "--max-width-mm", "10", "--out", str(fitted),
                    "--json",
                )
                payload = json.loads(result.stdout)
                self.assertEqual(payload["changes"]["max_width_mm"], 10.0)
                self.assertEqual(
                    [item["column"] for item in payload["changes"]["columns"]],
                    ["F", "G"],
                )
                self.assertEqual(
                    payload["changes"]["widths_mm"],
                    [item["width_mm"] for item in payload["changes"]["columns"]],
                )
                self.assertTrue(all(
                    "actual_width_mm" in item
                    and abs(item["actual_width_mm"] - item["width_mm"]) <= 0.1
                    for item in payload["changes"]["columns"]
                ))
                self.inspect(
                    fitted,
                    lambda book: self.assertTrue(
                        all(
                            book.sheet("Data").Columns.getByIndex(col).Width <= 1010
                            for col in (5, 6)
                        )
                    ),
                )

        invalid = self.directory / "invalid-autofit.xlsx"
        invalid_output = self.directory / "invalid-autofit-out.xlsx"
        create_fixture(invalid)
        result = self.run_cli(
            "col", "autofit", "--file", str(invalid), "--sheet", "Data",
            "--cols", "A", "--max-width-mm", "0", "--out", str(invalid_output),
            "--json", expected=3,
        )
        self.assertIn("有限的正数", result.stderr)
        self.assertFalse(invalid_output.exists())

    def test_find_values_formulas_ranges_and_limits(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("find" + extension)
                create_fixture(source, extra_sheet=True)

                values = self.run_cli(
                    "find", "--file", str(source), "--text", "alpha", "--json"
                )
                value_payload = json.loads(values.stdout)
                self.assertEqual(value_payload["matches"][0]["cell"], "B2")
                self.assertEqual(value_payload["matches"][0]["match_in"], ["values"])
                self.assertEqual(
                    [(match["sheet"], match["cell"]) for match in value_payload["matches"]],
                    [("Data", "B2"), ("Extra", "A1")],
                )
                self.assertFalse(value_payload["truncated"])

                formulas = self.run_cli(
                    "find", "--file", str(source), "--text", "A2*2", "--sheet", "Data",
                    "--range", "C2:C2", "--look-in", "formulas", "--json",
                )
                formula_match = json.loads(formulas.stdout)["matches"][0]
                self.assertEqual(formula_match["formula"], "=A2*2")
                self.assertEqual(formula_match["match_in"], ["formulas"])

                column_only = self.run_cli(
                    "find", "--file", str(source), "--text", "a", "--sheet", "Data",
                    "--cols", "B", "--json",
                )
                self.assertEqual(
                    [match["cell"] for match in json.loads(column_only.stdout)["matches"]],
                    ["B1", "B2", "B3"],
                )
                row_and_column = self.run_cli(
                    "find", "--file", str(source), "--text", "sentinel",
                    "--sheet", "Data", "--rows", "4", "--cols", "B", "--json",
                )
                self.assertEqual(
                    [match["cell"] for match in json.loads(row_and_column.stdout)["matches"]],
                    ["B4"],
                )
                tail = self.run_cli(
                    "find", "--file", str(source), "--text", "keep", "--sheet", "Data",
                    "--tail", "1", "--cols", "F", "--json",
                )
                self.assertEqual(
                    [match["cell"] for match in json.loads(tail.stdout)["matches"]],
                    ["F4"],
                )
                square = self.run_cli(
                    "find", "--file", str(source), "--text", "alpha", "--sheet", "Data",
                    "--from", "B2", "--n", "1", "--json",
                )
                self.assertEqual(
                    [match["cell"] for match in json.loads(square.stdout)["matches"]],
                    ["B2"],
                )

                sensitive = self.run_cli(
                    "find", "--file", str(source), "--text", "alpha",
                    "--case-sensitive", "--json",
                )
                self.assertEqual(json.loads(sensitive.stdout)["matches"], [])
                limited = self.run_cli(
                    "find", "--file", str(source), "--text", "=", "--look-in",
                    "formulas", "--limit", "1", "--json",
                )
                limited_payload = json.loads(limited.stdout)
                self.assertEqual(len(limited_payload["matches"]), 1)
                self.assertTrue(limited_payload["truncated"])

                regex_values = self.run_cli(
                    "find", "--file", str(source), "--sheet", "Data",
                    "--cols", "B:B", "--text", "^a[a-z]+$", "--regex", "--json",
                )
                regex_payload = json.loads(regex_values.stdout)
                self.assertEqual(
                    [match["cell"] for match in regex_payload["matches"]],
                    ["B2"],
                )
                self.assertTrue(regex_payload["regex"])
                self.assertFalse(regex_payload["case_sensitive"])

                regex_formulas = self.run_cli(
                    "find", "--file", str(source), "--sheet", "Data",
                    "--cols", "C", "--text", r"^=A[23]\*2$", "--regex",
                    "--look-in", "formulas", "--case-sensitive", "--json",
                )
                regex_formula_payload = json.loads(regex_formulas.stdout)
                self.assertEqual(
                    [match["cell"] for match in regex_formula_payload["matches"]],
                    ["C2", "C3"],
                )
                self.assertTrue(regex_formula_payload["case_sensitive"])

                regex_unbounded = self.run_cli(
                    "find", "--file", str(source), "--text", "^Alpha", "--regex",
                    "--json",
                )
                self.assertEqual(
                    [(match["sheet"], match["cell"])
                     for match in json.loads(regex_unbounded.stdout)["matches"]],
                    [("Data", "B2"), ("Extra", "A1")],
                )

        source = self.directory / "find-invalid.xlsx"
        create_fixture(source)
        invalid = self.run_cli(
            "find", "--file", str(source), "--text", "Alpha", "--limit", "0",
            "--json", expected=3,
        )
        self.assertIn("正整数", invalid.stderr)
        conflict = self.run_cli(
            "find", "--file", str(source), "--text", "Alpha", "--tail", "1",
            "--rows", "1", "--json", expected=3,
        )
        self.assertIn("不能与", conflict.stderr)
        ambiguous = self.run_cli(
            "find", "--file", str(source), "--text", "Alpha", "--range", "B2",
            "--cols", "B", "--json", expected=3,
        )
        self.assertIn("不能与其他范围参数", ambiguous.stderr)
        invalid_regex = self.run_cli(
            "find", "--file", str(self.directory / "missing.xlsx"),
            "--text", "(", "--sheet", "Data", "--cols", "B", "--regex",
            "--json", expected=3,
        )
        self.assertIn("非法正则表达式", invalid_regex.stderr)
        self.assertNotIn("文件不存在", invalid_regex.stderr)

    def test_unexpected_uno_error_has_stable_boundary(self):
        source = self.directory / "invalid-name.xlsx"
        output = self.directory / "invalid-name-out.xlsx"
        create_fixture(source)
        result = self.run_cli(
            "sheet", "add", "--file", str(source), "--name", "bad/name", "--out", str(output), "--json", expected=5
        )
        self.assertFalse(json.loads(result.stderr)["ok"])
        self.assertNotIn("Traceback", result.stderr)

    def test_argument_error_is_json_when_requested(self):
        result = self.run_cli("write", "--json", expected=2)
        payload = json.loads(result.stderr)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["code"], 2)

        removed_cell = self.run_cli("cell", "set", "--json", expected=2)
        self.assertIn("invalid choice", json.loads(removed_cell.stderr)["error"])

    def test_practical_json_write_style_clear_and_sheet_copy(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("practical" + extension)
                piped = self.directory / ("piped" + extension)
                positional = self.directory / ("positional" + extension)
                from_file = self.directory / ("from-file" + extension)
                styled = self.directory / ("styled-range" + extension)
                cleared = self.directory / ("cleared-range" + extension)
                cleared_style = self.directory / ("cleared-style" + extension)
                copied_sheet = self.directory / ("copied-sheet" + extension)
                create_fixture(source)

                viewed = self.run_cli(
                    "view", "--file", str(source), "--sheet", "Data", "--range", "A2:F2", "--json"
                )
                values = json.loads(viewed.stdout)
                self.assertEqual(values, [[1, "Alpha", "=A2*2", "'=literal", True, None]])
                self.run_cli(
                    "write", "--file", str(source), "--sheet", "Data", "--begin", "AA3", "--stdin",
                    "--font", TEST_FONT, "--font-size", "13", "--out", str(piped), "--json",
                    input_text=viewed.stdout,
                )

                def assert_piped(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("AA3").Value, 1)
                    self.assertEqual(sheet.getCellRangeByName("AB3").String, "Alpha")
                    self.assertEqual(sheet.getCellRangeByName("AC3").Formula, "=A2*2")
                    self.assertEqual(sheet.getCellRangeByName("AD3").String, "=literal")
                    self.assertIn(sheet.getCellRangeByName("AE3").String.upper(), ("TRUE", "1"))
                    self.assertAlmostEqual(sheet.getCellRangeByName("AB3").CharHeightAsian, 13.0)
                    self.assertAlmostEqual(sheet.getCellRangeByName("AF3").CharHeightAsian, 13.0)

                self.inspect(piped, assert_piped)

                positional_data = '[[9,"row3"],[],[10,null]]'
                self.run_cli(
                    "write", "--file", str(piped), "--sheet", "Data", "--begin", "F3", positional_data,
                    "--out", str(positional), "--json",
                )

                def assert_positional(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("F3").Value, 9)
                    self.assertEqual(sheet.getCellRangeByName("G3").String, "row3")
                    self.assertEqual(sheet.getCellRangeByName("F4").String, "KEEP")
                    self.assertEqual(sheet.getCellRangeByName("F5").Value, 10)
                    self.assertEqual(sheet.getCellRangeByName("G5").String, "")

                self.inspect(positional, assert_positional)

                values_file = self.directory / ("values-" + extension[1:] + ".json")
                values_file.write_text('[["file",5]]', encoding="utf-8")
                self.run_cli(
                    "write", "--file", str(positional), "--sheet", "Data", "--begin", "H2",
                    "--values-file", str(values_file), "--out", str(from_file), "--json",
                )
                style_result = self.run_cli(
                    "style", "--file", str(from_file), "--sheet", "Data", "--range", "A1:D2",
                    "--font", TEST_FONT, "--font-size", "16", "--out", str(styled), "--json",
                )
                style_payload = json.loads(style_result.stdout)
                self.assertTrue(style_payload["font_check"]["exact"])
                self.assertEqual(style_payload["changes"]["actual_font"]["western"], TEST_FONT)
                self.assertEqual(style_payload["changes"]["actual_font"]["asian"], TEST_FONT)
                styled_view = self.run_cli(
                    "view", "--file", str(styled), "--sheet", "Data", "--range", "D2:D2",
                    "--json-full", "--include-style",
                )
                styled_payload = json.loads(styled_view.stdout)
                self.assertEqual(styled_payload["values"], [["'=literal"]])
                literal_style = styled_payload["styles"][0][0]
                self.assertEqual(literal_style["font"]["western"], TEST_FONT)
                self.assertEqual(literal_style["font"]["asian"], TEST_FONT)
                self.assertEqual(set(literal_style["font_size"].values()), {16.0})
                self.run_cli(
                    "clear", "--file", str(styled), "--sheet", "Data", "--range", "A2:B2",
                    "--out", str(cleared), "--json",
                )

                def assert_clear_preserves_style(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("A2").String, "")
                    self.assertAlmostEqual(sheet.getCellRangeByName("A2").CharHeightAsian, 16.0)

                self.inspect(cleared, assert_clear_preserves_style)
                self.run_cli(
                    "clear", "--file", str(styled), "--sheet", "Data", "--range", "A2:B2", "--clear-style",
                    "--out", str(cleared_style), "--json",
                )
                self.inspect(
                    cleared_style,
                    lambda book: self.assertEqual(
                        book.sheet("Data").getCellRangeByName("A2")
                        .getPropertyState("CharHeightAsian")
                        .value,
                        "DEFAULT_VALUE",
                    ),
                )

                self.run_cli(
                    "sheet", "copy", "--file", str(cleared), "--sheet", "Data", "--name", "Copy",
                    "--out", str(copied_sheet), "--json",
                )
                names = self.run_cli("sheet", "list", "--file", str(copied_sheet), "--json")
                self.assertEqual(json.loads(names.stdout), ["Data", "Copy"])

                def assert_sheet_copy(workbook):
                    source_sheet = workbook.sheet("Data")
                    copied = workbook.sheet("Copy")
                    self.assertEqual(copied.getCellRangeByName("B2").CharWeight, 150.0)
                    self.assertLessEqual(
                        abs(copied.Rows.getByIndex(1).Height - source_sheet.Rows.getByIndex(1).Height), 2
                    )

                self.inspect(copied_sheet, assert_sheet_copy)
                full = self.run_cli(
                    "view", "--file", str(copied_sheet), "--sheet", "Copy", "--range", "H2:I2", "--json-full"
                )
                self.assertEqual(json.loads(full.stdout)["begin"], "H2")

    def test_practical_interface_rejects_ambiguous_inputs(self):
        source = self.directory / "ambiguous.xlsx"
        output = self.directory / "ambiguous-out.xlsx"
        values_file = self.directory / "values.json"
        create_fixture(source)
        values_file.write_text("[[1]]", encoding="utf-8")
        result = self.run_cli(
            "write", "--file", str(source), "--sheet", "Data", "--begin", "A1", "[[1]]",
            "--values-file", str(values_file), "--out", str(output), "--json", expected=3,
        )
        self.assertIn("必须且只能", result.stderr)
        direct = self.run_cli("view", "--file", str(source), "--range", "A1:A1", "--json", expected=3)
        self.assertIn("明确指定", direct.stderr)
        directory = self.run_cli(
            "view", "--dir", str(self.directory), "--sheet", "Data", "--range", "A1:A1", "--json", expected=2
        )
        self.assertIn("required", directory.stderr)
        repeated_file = self.run_cli(
            "view", "--file", str(source), "--file", str(source), "--sheet", "Data",
            "--json", expected=2,
        )
        self.assertIn("只能指定一次", repeated_file.stderr)
        obsolete_clear_style = self.run_cli(
            "clear", "--file", str(source), "--sheet", "Data", "--range", "A1",
            "--with-style", "--json", expected=2,
        )
        self.assertIn("unrecognized arguments", obsolete_clear_style.stderr)
        obsolete_sheet = self.run_cli(
            "sheet", "add", "--file", str(source), "--sheet", "Data", "--name", "New",
            "--json", expected=2,
        )
        self.assertIn("unrecognized arguments", obsolete_sheet.stderr)
        invalid_size = self.run_cli(
            "style", "--file", str(source), "--sheet", "Data", "--range", "A1", "--font-size", "nan",
            "--out", str(output), "--json", expected=3,
        )
        self.assertIn("有限的正数", invalid_size.stderr)
        self.assertFalse(output.exists())
        legacy = self.run_cli(
            "clear", "--file", str(source), "--sheet", "Data", "--range", "A1", "--in-place", "--json",
            expected=2,
        )
        self.assertIn("unrecognized arguments: --in-place", legacy.stderr)
        overwrite = self.run_cli(
            "clear", "--file", str(source), "--sheet", "Data", "--range", "A1", "--overwrite", "--json",
            expected=3,
        )
        self.assertIn("只能与 --out", overwrite.stderr)
        font = self.run_cli("font", "check", "--name", TEST_FONT, "--json")
        font_payload = json.loads(font.stdout)
        self.assertTrue(font_payload["exact"])
        self.assertEqual(font_payload["requested"], TEST_FONT)
        missing_font = self.run_cli(
            "style", "--file", str(source), "--sheet", "Data", "--range", "A1",
            "--font", "ExcelTool Definitely Missing Font", "--out", str(output), "--json", expected=3,
        )
        self.assertIn("字体不可精确使用", missing_font.stderr)
        include_without_full = self.run_cli(
            "view", "--file", str(source), "--sheet", "Data", "--range", "A1",
            "--include-style", expected=3,
        )
        self.assertIn("必须与 --json-full", include_without_full.stderr)

    def test_patch_v1_executes_all_operations_for_both_formats(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("patch-source" + extension)
                output = self.directory / ("patch-output" + extension)
                patch_file = self.directory / ("patch-" + extension[1:] + ".json")
                create_fixture(source)
                patch_file.write_text(
                    json.dumps(
                        {
                            "version": 1,
                            "operations": [
                                {"id": "add", "op": "sheet.add", "name": "Extra"},
                                {
                                    "id": "write-extra", "op": "write", "sheet": "Extra",
                                    "begin": "A1", "values": [["meta", 7]],
                                },
                                {
                                    "id": "rename", "op": "sheet.rename", "sheet": "Extra",
                                    "name": "Config",
                                },
                                {
                                    "id": "copy-sheet", "op": "sheet.copy", "sheet": "Config",
                                    "name": "ConfigCopy",
                                },
                                {"id": "delete-sheet", "op": "sheet.delete", "sheet": "Config"},
                                {
                                    "id": "insert", "op": "row.insert", "sheet": "Data",
                                    "before": 3, "count": 1,
                                },
                                {
                                    "id": "copy-row", "op": "row.copy", "sheet": "Data",
                                    "rows": "2:2", "insert_before": 4,
                                },
                                {
                                    "id": "delete-row", "op": "row.delete", "sheet": "Data",
                                    "rows": "5:5",
                                },
                                {
                                    "id": "write", "op": "write", "sheet": "Data", "begin": "A3",
                                    "values": [[9999, "任务", "=A3*2", "'=普通文本"]],
                                    "font": TEST_FONT, "font_size": 10,
                                },
                                {
                                    "id": "style", "op": "style", "sheet": "Data",
                                    "range": "A3:D3", "font_size": 11,
                                },
                                {
                                    "id": "clear", "op": "clear", "sheet": "Data",
                                    "range": "E4:E4",
                                },
                                {
                                    "op": "clear", "sheet": "Data",
                                    "range": "F1:F1", "with_style": True,
                                },
                            ],
                        },
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )
                result = self.run_cli(
                    "patch", "--file", str(source), "--patch", str(patch_file),
                    "--out", str(output), "--json",
                )
                payload = json.loads(result.stdout)
                self.assertTrue(payload["verified"])
                self.assertEqual(payload["operation"], "patch")
                self.assertEqual(payload["summary"]["operation_count"], 12)
                self.assertTrue(payload["summary"]["saved_once"])
                self.assertEqual(payload["operations"][8]["id"], "write")
                self.assertNotIn("id", payload["operations"][11])
                self.assertTrue(all(item["verified"] for item in payload["operations"]))

                def assert_patch(workbook):
                    self.assertEqual(workbook.sheet_names(), ["Data", "ConfigCopy"])
                    data = workbook.sheet("Data")
                    self.assertEqual(data.getCellRangeByName("A3").Value, 9999)
                    self.assertEqual(data.getCellRangeByName("B3").String, "任务")
                    self.assertEqual(data.getCellRangeByName("D3").String, "=普通文本")
                    self.assertEqual(data.getCellRangeByName("B3").CharFontNameAsian, TEST_FONT)
                    self.assertAlmostEqual(data.getCellRangeByName("A3").CharHeightAsian, 11.0)
                    self.assertEqual(data.getCellRangeByName("A4").Value, 1)
                    self.assertEqual(data.getCellRangeByName("B4").String, "Alpha")
                    self.assertEqual(data.getCellRangeByName("C4").Formula, "=A4*2")
                    self.assertEqual(data.getCellRangeByName("E4").Formula, "")
                    self.assertEqual(data.getCellRangeByName("B5").String, "Sentinel")
                    self.assertEqual(
                        data.getCellRangeByName("F1").getPropertyState("CharHeightAsian").value,
                        "DEFAULT_VALUE",
                    )
                    copied = workbook.sheet("ConfigCopy")
                    self.assertEqual(copied.getCellRangeByName("A1").String, "meta")
                    self.assertEqual(copied.getCellRangeByName("B1").Value, 7)

                self.inspect(output, assert_patch)
                self.inspect(
                    source,
                    lambda workbook: self.assertEqual(workbook.sheet_names(), ["Data"]),
                )

    def test_patch_v1_failure_is_atomic_and_reports_operation(self):
        source = self.directory / "patch-atomic.xlsx"
        patch_file = self.directory / "patch-failure.json"
        invalid_top = self.directory / "patch-invalid-top.json"
        existing_output = self.directory / "patch-existing.xlsx"
        output_patch = self.directory / "patch-output-policy.json"
        font_patch = self.directory / "patch-font-failure.json"
        font_output = self.directory / "patch-font-failure-out.xlsx"
        create_fixture(source)
        original = source.read_bytes()
        patch_file.write_text(
            json.dumps(
                {
                    "version": 1,
                    "operations": [
                        {
                            "id": "insert-first", "op": "row.insert", "sheet": "Data",
                            "before": 2, "count": 1,
                        },
                        {
                            "id": "write-fails", "op": "write", "sheet": "Missing",
                            "begin": "A1", "values": [[1]],
                        },
                    ],
                }
            ),
            encoding="utf-8",
        )
        failure = self.run_cli(
            "patch", "--file", str(source), "--patch", str(patch_file), "--json",
            expected=3,
        )
        payload = json.loads(failure.stderr)
        self.assertEqual(payload["operation"], "patch")
        self.assertEqual(payload["failed_operation"], {
            "index": 1, "id": "write-fails", "op": "write",
        })
        self.assertFalse(payload["published"])
        self.assertEqual(source.read_bytes(), original)

        invalid_top.write_text(
            json.dumps({"version": 1, "file": "forbidden.xlsx", "operations": []}),
            encoding="utf-8",
        )
        invalid = self.run_cli(
            "patch", "--file", str(source), "--patch", str(invalid_top), "--json",
            expected=3,
        )
        invalid_payload = json.loads(invalid.stderr)
        self.assertIn("顶层不支持字段: file", invalid_payload["error"])
        self.assertFalse(invalid_payload["published"])
        self.assertNotIn("failed_operation", invalid_payload)

        create_fixture(existing_output)
        existing_bytes = existing_output.read_bytes()
        output_patch.write_text(
            json.dumps({
                "version": 1,
                "operations": [
                    {"id": "clear", "op": "clear", "sheet": "Data", "range": "A1"}
                ],
            }),
            encoding="utf-8",
        )
        exists = self.run_cli(
            "patch", "--file", str(source), "--patch", str(output_patch),
            "--out", str(existing_output), "--json", expected=3,
        )
        exists_payload = json.loads(exists.stderr)
        self.assertIn("--overwrite", exists_payload["error"])
        self.assertFalse(exists_payload["published"])
        self.assertEqual(existing_output.read_bytes(), existing_bytes)

        font_patch.write_text(
            json.dumps({
                "version": 1,
                "operations": [
                    {
                        "id": "font-fails", "op": "write", "sheet": "Data",
                        "begin": "B2", "values": [["مهمة"]], "font": TEST_FONT,
                    },
                    {
                        "id": "later-clear", "op": "clear", "sheet": "Data",
                        "range": "F1:F1",
                    },
                ],
            }, ensure_ascii=False),
            encoding="utf-8",
        )
        font_failure = self.run_cli(
            "patch", "--file", str(source), "--patch", str(font_patch),
            "--out", str(font_output), "--json", expected=6,
        )
        font_payload = json.loads(font_failure.stderr)
        self.assertEqual(font_payload["failed_operation"], {
            "index": 0, "id": "font-fails", "op": "write",
        })
        self.assertFalse(font_payload["published"])
        self.assertFalse(font_output.exists())

    def test_patch_v1_column_operations_use_live_coordinates(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension):
                source = self.directory / ("patch-columns-source" + extension)
                output = self.directory / ("patch-columns-output" + extension)
                patch_file = self.directory / ("patch-columns-" + extension[1:] + ".json")
                create_fixture(source)
                patch_file.write_text(
                    json.dumps({
                        "version": 1,
                        "operations": [
                            {
                                "id": "insert-col", "op": "col.insert", "sheet": "Data",
                                "before": "B", "count": 1,
                            },
                            {
                                "id": "write-live", "op": "write", "sheet": "Data",
                                "begin": "B2", "values": [["Inserted"]],
                            },
                            {
                                "id": "copy-col", "op": "col.copy", "sheet": "Data",
                                "cols": "B:B", "insert_before": "F",
                            },
                            {
                                "id": "delete-col", "op": "col.delete", "sheet": "Data",
                                "cols": "B:B",
                            },
                            {
                                "id": "fit-col", "op": "col.autofit", "sheet": "Data",
                                "cols": "E:E", "max_width_mm": 10,
                            },
                            {
                                "id": "shift-fitted-col", "op": "col.insert",
                                "sheet": "Data", "before": "A", "count": 1,
                            },
                        ],
                    }),
                    encoding="utf-8",
                )
                result = self.run_cli(
                    "patch", "--file", str(source), "--patch", str(patch_file),
                    "--out", str(output), "--json",
                )
                payload = json.loads(result.stdout)
                self.assertEqual(payload["summary"]["operation_count"], 6)
                self.assertTrue(payload["summary"]["saved_once"])
                fit_changes = payload["operations"][4]["changes"]
                self.assertEqual(fit_changes["columns"][0]["column"], "E")
                self.assertEqual(fit_changes["columns"][0]["final_column"], "F")
                self.assertLessEqual(fit_changes["columns"][0]["actual_width_mm"], 10.1)

                def assert_patch(workbook):
                    sheet = workbook.sheet("Data")
                    self.assertEqual(sheet.getCellRangeByName("C2").String, "Alpha")
                    self.assertEqual(sheet.getCellRangeByName("D2").Formula, "=B2*2")
                    self.assertEqual(sheet.getCellRangeByName("F2").String, "Inserted")
                    self.assertLessEqual(sheet.Columns.getByIndex(5).Width, 1010)

                self.inspect(output, assert_patch)


if __name__ == "__main__":
    unittest.main()
