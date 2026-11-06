import tempfile
import unittest
import zipfile
from pathlib import Path

from exceltool.ooxml import normalize_xlsx_default_grid_color


def write_xlsx(path, worksheets, comment=b""):
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as archive:
        archive.comment = comment
        archive.writestr("[Content_Types].xml", b"content-types")
        for name, data in worksheets.items():
            archive.writestr("xl/worksheets/%s.xml" % name, data)


class OoxmlTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="exceltool-ooxml-")
        self.directory = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_removes_redundant_default_grid_color_attributes_independently(self):
        path = self.directory / "book.xlsx"
        write_xlsx(path, {
            "sheet1": (
                b'<worksheet><sheetViews><sheetView workbookViewId="0" '
                b'defaultGridColor="true" colorId="64" zoomScale="100"/>'
                b'</sheetViews></worksheet>'
            ),
            "sheet2": (
                b"<worksheet><sheetViews>"
                b"<sheetView colorId='64' defaultGridColor='1' workbookViewId='0'/>"
                b"</sheetViews></worksheet>"
            ),
            "sheet3": (
                b'<worksheet><sheetView colorId="64"/></worksheet>'
            ),
            "sheet4": (
                b'<worksheet><sheetView defaultGridColor="true"/></worksheet>'
            ),
        }, comment=b"archive-comment")

        self.assertEqual(normalize_xlsx_default_grid_color(path), 4)

        with zipfile.ZipFile(str(path), "r") as archive:
            self.assertEqual(archive.comment, b"archive-comment")
            self.assertEqual(archive.read("[Content_Types].xml"), b"content-types")
            for name in ("sheet1", "sheet2", "sheet3", "sheet4"):
                data = archive.read("xl/worksheets/%s.xml" % name)
                self.assertNotIn(b"defaultGridColor", data)
                self.assertNotIn(b"colorId", data)

    def test_removes_defaults_and_preserves_non_default_grid_color_settings(self):
        path = self.directory / "book.xlsx"
        worksheets = {
            "false-default": (
                b'<worksheet><sheetView defaultGridColor="false" colorId="64"/>'
                b"</worksheet>"
            ),
            "custom-color": (
                b'<worksheet><sheetView defaultGridColor="true" colorId="12"/>'
                b"</worksheet>"
            ),
        }
        write_xlsx(path, worksheets)

        self.assertEqual(normalize_xlsx_default_grid_color(path), 2)
        with zipfile.ZipFile(str(path), "r") as archive:
            false_default = archive.read(
                "xl/worksheets/false-default.xml"
            )
            custom_color = archive.read("xl/worksheets/custom-color.xml")
        self.assertIn(b'defaultGridColor="false"', false_default)
        self.assertNotIn(b"colorId", false_default)
        self.assertNotIn(b"defaultGridColor", custom_color)
        self.assertIn(b'colorId="12"', custom_color)

    def test_ignores_xls_files(self):
        path = self.directory / "book.xls"
        path.write_bytes(b"not-a-zip")

        self.assertEqual(normalize_xlsx_default_grid_color(path), 0)
        self.assertEqual(path.read_bytes(), b"not-a-zip")


if __name__ == "__main__":
    unittest.main()
