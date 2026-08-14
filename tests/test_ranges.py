import unittest

from exceltool.errors import TargetError
from exceltool.ranges import (
    cell_range,
    cell_ref,
    col_range,
    column_name,
    column_number,
    resolve_window,
    row_range,
)


class RangeTests(unittest.TestCase):
    def test_columns_round_trip(self):
        for number, name in ((1, "A"), (26, "Z"), (27, "AA"), (703, "AAA")):
            self.assertEqual(column_number(name), number)
            self.assertEqual(column_name(number), name)

    def test_ranges_are_zero_based_half_open(self):
        self.assertEqual(cell_ref("B3"), (2, 1))
        self.assertEqual(row_range("2:5"), (1, 5))
        self.assertEqual(col_range("B:D"), (1, 4))
        self.assertEqual(cell_range("B2:D5"), (1, 5, 1, 4))

    def test_reverse_range_fails(self):
        with self.assertRaises(TargetError):
            row_range("5:2")

    def test_zero_square_size_fails(self):
        with self.assertRaises(TargetError):
            resolve_window(None, None, None, "A1", 0, 5, 5)


if __name__ == "__main__":
    unittest.main()
