import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from excel_com import MAX_CELLS_PER_COM_READ, matrix, normalized_hwnd, parse_address, range_tiles, matching_workbook, read_range, decode_scalar, inspect_range
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
        with patch('excel_com.read_property', return_value=([['long text']], [[{'kind': 'string', 'vartype': 8}]])), \
                patch('excel_com.MAX_TEXT_CHARS', 2), self.assertRaisesRegex(ValueError, 'Text budget'):
            read_range(wb, 'Data', 'A1', mode='values')
        with patch('excel_com.read_property', side_effect=ValueError('Unsupported COM VARTYPE 9')), \
                self.assertRaisesRegex(ValueError, 'Unsupported COM VARTYPE'):
            read_range(wb, 'Data', 'A1', mode='values')


    def test_raw_error_and_number_vartypes_remain_distinct(self):
        from comtypes.automation import VARIANT, VT_ERROR
        error = VARIANT()
        error.vt = VT_ERROR
        error._.VT_I4 = 2007
        value, info = decode_scalar(error)
        self.assertEqual(value['name'], '#DIV/0!')
        self.assertEqual(info, {'kind': 'error', 'vartype': VT_ERROR})
        number = VARIANT(2007)
        self.assertEqual(decode_scalar(number), (2007, {'kind': 'number', 'vartype': 3}))
        self.assertEqual(decode_scalar(VARIANT(True))[1]['kind'], 'boolean')

    def test_raw_hresult_unknown_error_and_unsupported_types(self):
        from comtypes.automation import VARIANT
        error = VARIANT()
        error.vt, error._.VT_I4 = 10, -2146826281
        value, _ = decode_scalar(error)
        self.assertEqual(value['scode_hex'], '0x800A07D7')
        self.assertEqual(value['name'], '#DIV/0!')
        error._.VT_I4 = 12345
        self.assertIsNone(decode_scalar(error)[0]['name'])
        self.assertEqual(decode_scalar(error)[0]['scode'], 12345)
        with self.assertRaisesRegex(ValueError, 'Non-finite'):
            decode_scalar(VARIANT(float('nan')))
        raw = VARIANT()
        raw.vt = 0x4005  # BYREF is deliberately unsupported; do not dereference arbitrary pointers.
        try:
            with self.assertRaisesRegex(ValueError, 'Unsupported'):
                decode_scalar(raw)
        finally:
            raw.vt = 0

    def test_utf16_bstr_uses_code_units_and_embedded_nul(self):
        import ctypes
        from excel_com import automation_api
        from comtypes.automation import VARIANT
        api = automation_api()
        allocate = api.SysAllocStringLen
        allocate.argtypes, allocate.restype = [ctypes.c_void_p, ctypes.c_uint], ctypes.c_void_p
        text = '中😀\0tail'
        data = text.encode('utf-16-le')
        source = ctypes.create_string_buffer(data)
        raw = VARIANT()
        raw.vt, raw._.c_void_p = 8, allocate(source, len(data)//2)
        try:
            self.assertEqual(decode_scalar(raw), (text, {'kind': 'string', 'vartype': 8}))
        finally:
            api.VariantClear(ctypes.byref(raw))

    def test_inspect_budget_rejected_before_com(self):
        wb = Mock()
        with self.assertRaises(ValueError):
            inspect_range(wb, 'Data', 'A1:A21')
        self.assertEqual(wb.mock_calls, [])

    def test_inspect_unknown_structure_fails_closed(self):
        class ItemCollection:
            Count = 0
            def Item(self, *args):
                return self
        cell = SimpleNamespace(MergeCells=False, HasArray=False, HasSpill=False, HasFormula=False)
        rng = SimpleNamespace(Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=1),
                              Areas=SimpleNamespace(Count=1), Cells=SimpleNamespace(Item=lambda *args: cell))
        sheet = SimpleNamespace(Name='Data', CodeName='Sheet1', ProtectContents=False,
                                ProtectDrawingObjects=False, ProtectScenarios=False,
                                ListObjects=ItemCollection(), PivotTables=lambda: ItemCollection(),
                                QueryTables=ItemCollection(), Range=lambda address: rng)
        wb = SimpleNamespace(Worksheets=SimpleNamespace(Item=lambda name: sheet), ReadOnly=False,
                             ProtectStructure=False, HasVBProject=False, MultiUserEditing=False)
        with patch('excel_com.read_range', return_value={'data': {}, 'sheet': {'name': 'Data', 'code_name': 'Sheet1'}}):
            result = inspect_range(wb, 'Data', 'A1')
        self.assertTrue(result['checked_structure_clear'])
        self.assertFalse(result['write_authorized'])
        self.assertEqual(result['blockers'], [])
        cell.HasSpill = None
        with patch('excel_com.read_range', return_value={'data': {}, 'sheet': {'name': 'Data', 'code_name': 'Sheet1'}}):
            result = inspect_range(wb, 'Data', 'A1')
        self.assertFalse(result['checked_structure_clear'])
        self.assertTrue(any('HasSpill' in reason for reason in result['blockers']))

    def test_matrix_scalar_and_rectangular_shape(self):
        self.assertEqual(matrix('value', 1, 1), [['value']])
        with self.assertRaisesRegex(ValueError, 'surrogate'):
            matrix(chr(0xd83d), 1, 1)
        self.assertEqual(matrix(((1, 2), (3, 4)), 2, 2), [[1, 2], [3, 4]])
        with self.assertRaises(ValueError):
            matrix(((1,),), 1, 2)


if __name__ == '__main__':
    unittest.main()
