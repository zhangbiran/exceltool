import unittest
from unittest import mock

from exceltool import fonts


class WindowsFontInspectionTests(unittest.TestCase):
    @mock.patch.object(fonts, "_windows_font_names")
    def test_exact_family_is_available(self, font_names):
        font_names.return_value = {"Microsoft YaHei", "Microsoft YaHei UI"}

        result = fonts._inspect_windows_font("Microsoft YaHei")

        self.assertEqual(result["resolved"], "Microsoft YaHei")
        self.assertTrue(result["available"])
        self.assertTrue(result["exact"])

    @mock.patch.object(fonts, "_windows_font_names")
    def test_missing_family_is_not_available(self, font_names):
        font_names.return_value = {"Microsoft YaHei"}

        result = fonts._inspect_windows_font("Missing Font")

        self.assertIsNone(result["resolved"])
        self.assertFalse(result["available"])
        self.assertFalse(result["exact"])


if __name__ == "__main__":
    unittest.main()
