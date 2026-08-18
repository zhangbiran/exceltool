import hashlib
import tempfile
import unittest
from pathlib import Path

from exceltool.errors import (
    FormulaVerificationError,
    InputHashMismatchError,
    TargetError,
    UnsupportedError,
)
from exceltool.safety import (
    assert_formula_snapshots,
    formula_differences,
    formula_policy_for_edit,
    formula_policy_for_patch,
    normalize_xls_string_formula_results,
    normalize_expected_sha256,
    sha256_file,
    verify_input_sha256,
    verify_source_unchanged,
)


class FakeType:
    def __init__(self, value):
        self.value = value


class FakeCell:
    def __init__(self, cell_type, value=0.0, text="", formula=""):
        self.Type = FakeType(cell_type)
        self.Value = value
        self.String = text
        self.Formula = formula


class FakeSheet:
    def __init__(self, cells):
        self.cells = cells

    def getCellByPosition(self, col, row):
        return self.cells[(row, col)]


class FakeWorkbook:
    def __init__(self, sheets):
        self.sheets = sheets

    def sheet_names(self):
        return list(self.sheets)

    def sheet(self, name):
        return self.sheets[name]


class FakeAddress:
    def __init__(self, row, col):
        self.StartRow = self.EndRow = row
        self.StartColumn = self.EndColumn = col


class FakeFormulaRanges:
    def __init__(self, cells):
        self.cells = cells

    def getRangeAddresses(self):
        return [FakeAddress(row, col) for row, col in self.cells]


class FakeFormulaSheet(FakeSheet):
    def queryContentCells(self, unused_flag):
        return FakeFormulaRanges(self.cells)


class FakeDocument:
    def __init__(self):
        self.calculate_all_calls = 0

    def calculateAll(self):
        self.calculate_all_calls += 1


class FakeFormulaCell(FakeCell):
    def __init__(self, result_type, formula, array_formula=""):
        self.assignments = 0
        self._formula = formula
        super().__init__("FORMULA", formula=formula)
        self.assignments = 0
        self.FormulaResultType2 = result_type
        self.array_formula = array_formula

    @property
    def Formula(self):
        return self._formula

    @Formula.setter
    def Formula(self, value):
        self._formula = value
        self.assignments += 1

    def getArrayFormula(self):
        return self.array_formula


class SafetyTests(unittest.TestCase):
    def test_normalize_xls_reassigns_only_string_result_formulas(self):
        text_formula = FakeFormulaCell(2, '=A1&"_small"')
        value_formula = FakeFormulaCell(1, "=A1*2")
        workbook = FakeWorkbook({
            "Data": FakeFormulaSheet({
                (0, 0): text_formula,
                (0, 1): value_formula,
            }),
        })
        workbook.document = FakeDocument()

        self.assertEqual(normalize_xls_string_formula_results(workbook), 1)
        self.assertEqual(text_formula.assignments, 1)
        self.assertEqual(value_formula.assignments, 0)
        self.assertEqual(workbook.document.calculate_all_calls, 1)

    def test_normalize_xls_rejects_string_result_array_formula(self):
        array_formula = FakeFormulaCell(2, "=TRANSPOSE(A1:A2)", "=TRANSPOSE(A1:A2)")
        workbook = FakeWorkbook({
            "Data": FakeFormulaSheet({(0, 0): array_formula}),
        })
        workbook.document = FakeDocument()

        with self.assertRaises(UnsupportedError) as raised:
            normalize_xls_string_formula_results(workbook)
        self.assertIn("Data!A1", str(raised.exception))
        self.assertEqual(array_formula.assignments, 0)
        self.assertEqual(workbook.document.calculate_all_calls, 0)

    def test_sha256_validation_and_change_detection(self):
        with tempfile.TemporaryDirectory(prefix="exceltool-safety-") as directory:
            path = Path(directory) / "book.xls"
            path.write_bytes(b"before")
            expected = hashlib.sha256(b"before").hexdigest()
            self.assertEqual(sha256_file(path), expected)
            self.assertEqual(verify_input_sha256(path, expected.upper()), expected)
            with self.assertRaises(InputHashMismatchError):
                verify_input_sha256(path, "0" * 64)
            path.write_bytes(b"after")
            with self.assertRaises(InputHashMismatchError):
                verify_source_unchanged(path, expected)

        with self.assertRaises(TargetError):
            normalize_expected_sha256("not-a-hash")

    def test_formula_policies_distinguish_content_and_structure(self):
        write = formula_policy_for_edit(
            "write", "Data", begin="B2", values=[[1, 2], [], [3]]
        )
        self.assertTrue(write["enforce_declared_changes"])
        self.assertEqual(write["allowed_changes"]["Data"], {
            (1, 1), (1, 2), (3, 1),
        })
        style = formula_policy_for_edit("style")
        self.assertTrue(style["enforce_declared_changes"])
        self.assertEqual(style["allowed_changes"], {})
        structural = formula_policy_for_edit("row.insert")
        self.assertFalse(structural["enforce_declared_changes"])

        patch = formula_policy_for_patch([
            {
                "op": "write", "sheet": "Data", "begin": "C3",
                "values": [["=A1"]],
            },
            {
                "op": "clear", "sheet": "Data", "range": "D4:E4",
                "with_style": False,
            },
        ])
        self.assertTrue(patch["enforce_declared_changes"])
        self.assertEqual(patch["allowed_changes"]["Data"], {
            (2, 2), (3, 3), (3, 4),
        })
        patch_with_rows = formula_policy_for_patch([
            {"op": "row.insert", "sheet": "Data", "before": 2, "count": 1},
        ])
        self.assertFalse(patch_with_rows["enforce_declared_changes"])

    def test_formula_difference_reports_formula_to_value(self):
        before = {"Data": {(1, 2): "=A2*2"}}
        after = {"Data": {}}
        workbook = FakeWorkbook({
            "Data": FakeSheet({(1, 2): FakeCell("VALUE", value=4.0)}),
        })
        differences = formula_differences(
            before, after, workbook, "save"
        )
        self.assertEqual(differences, [{
            "stage": "save",
            "sheet": "Data",
            "cell": "C2",
            "before_formula": "=A2*2",
            "after_type": "VALUE",
            "after": 4,
        }])
        with self.assertRaises(FormulaVerificationError) as raised:
            assert_formula_snapshots(before, after, workbook, "save")
        self.assertEqual(
            raised.exception.details["formula_verification"]["unexpected_changes"], 1
        )
        self.assertFalse(raised.exception.details["published"])


if __name__ == "__main__":
    unittest.main()
