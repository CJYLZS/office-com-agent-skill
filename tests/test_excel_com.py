import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from excel_com import MAX_CELLS_PER_COM_READ, matrix, normalized_hwnd, parse_address, range_tiles, matching_workbook, read_range
from types import SimpleNamespace
from unittest.mock import Mock, patch


class AddressTests(unittest.TestCase):
    def test_single_cell_and_absolute_rectangle(self):
        self.assertEqual(parse_address('b3'), {'address': 'B3:B3', 'rows': 1, 'columns': 1})
        self.assertEqual(parse_address('$A$1:$D$20'), {'address': 'A1:D20', 'rows': 20, 'columns': 4})

    def test_large_columns(self):
        self.assertEqual(parse_address('XFD1048576')['columns'], 1)

    def test_rejects_non_rectangular_or_oversized_inputs(self):
        for address in ('Sheet1!A1', 'A:A', 'A1,B2', 'A1#', 'A2:A1', 'XFE1', 'A0', 'A1:ZZ1000'):
            with self.subTest(address=address), self.assertRaises(ValueError):
                parse_address(address)

    def test_rejects_excessive_cell_budget(self):
        with self.assertRaises(ValueError):
            parse_address('A1:B2', max_cells=3)

    def test_tiles_cover_rectangle_with_bounded_calls(self):
        tiles = list(range_tiles('A1:O20'))
        self.assertTrue(all(h*w <= MAX_CELLS_PER_COM_READ for _, _, h, w, _ in tiles))
        self.assertEqual(sum(h*w for _, _, h, w, _ in tiles), 300)
        self.assertEqual(tiles[0][4], 'A1:O6')

    def test_high_bit_handle_binding_and_ambiguity(self):
        self.assertEqual(normalized_hwnd(-2147483648), 2147483648)
        windows = SimpleNamespace(Count=1, Item=lambda i: SimpleNamespace(Hwnd=-2147483648))
        wb = SimpleNamespace(Windows=windows)
        app = SimpleNamespace(Workbooks=SimpleNamespace(Count=1, Item=lambda i: wb))
        self.assertIs(matching_workbook(app, 2147483648), wb)
        with self.assertRaises(ValueError):
            matching_workbook(app, 42)
        app.Workbooks.Count = 2
        with self.assertRaises(ValueError):
            matching_workbook(app, 2147483648)

    def test_invalid_range_rejected_before_com(self):
        wb = Mock()
        with self.assertRaises(ValueError):
            read_range(wb, 'Data', 'A:A')
        self.assertEqual(wb.mock_calls, [])

    def test_property_error_and_text_budget_fail_explicitly(self):
        tile = SimpleNamespace(Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=1),
                               Areas=SimpleNamespace(Count=1), Value2='long text')
        sheet = SimpleNamespace(Name='Data', CodeName='Sheet1', Range=lambda address: tile)
        wb = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda name: sheet))
        with patch('excel_com.MAX_TEXT_CHARS', 2), self.assertRaisesRegex(ValueError, 'Text budget'):
            read_range(wb, 'Data', 'A1', mode='values')
        class ErrorRange:
            @property
            def Value2(self):
                raise NotImplementedError('typecode 10')
        # Shape checks use the first range, then property reads use the next tile.
        calls = iter([tile, ErrorRange()])
        sheet.Range = lambda address: next(calls)
        with self.assertRaisesRegex(ValueError, 'variant cannot be decoded'):
            read_range(wb, 'Data', 'A1', mode='values')

    def test_matrix_scalar_and_rectangular_shape(self):
        self.assertEqual(matrix('value', 1, 1), [['value']])
        with self.assertRaisesRegex(ValueError, 'surrogate'):
            matrix(chr(0xd83d), 1, 1)
        self.assertEqual(matrix(((1, 2), (3, 4)), 2, 2), [[1, 2], [3, 4]])
        with self.assertRaises(ValueError):
            matrix(((1,),), 1, 2)


if __name__ == '__main__':
    unittest.main()
