"""Pure mocked safety regressions: never create a bridge or contact live Office."""
import re
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace as N
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import excel_tools as x


class Tracked:
    def __init__(self, journal, **fields):
        object.__setattr__(self, '_journal', journal)
        self.__dict__.update(fields)
    def __setattr__(self, name, value):
        self._journal.append((self, name, value))
        object.__setattr__(self, name, value)


class Cell(Tracked):
    def __init__(self, journal, value=None):
        super().__init__(journal, Value2=value, Formula=value, HasFormula=False,
                         MergeCells=False, HasArray=False, HasSpill=False)
        self.__dict__.update(fail_after=False, ignore_write=False, NumberFormat='General',
                             HorizontalAlignment=1, ColumnWidth=8.43,
                             Font=Tracked(journal, Bold=False, Size=11, Name='Calibri', Color=0),
                             Interior=Tracked(journal, Color=16777215))
    def __setattr__(self, name, value):
        if name not in ('Value2', 'Formula'):
            return super().__setattr__(name, value)
        self._journal.append((self, name, value))
        if not self.ignore_write:
            self.__dict__[name] = value
            self.__dict__['HasFormula'] = name == 'Formula' and isinstance(value, str) and value.startswith('=')
            if name == 'Value2': self.__dict__['Formula'] = value
            elif value == '=1+2': self.__dict__['Value2'] = 3
        if self.fail_after:
            raise RuntimeError('COM disconnected after assignment')


class Rectangle(Tracked):
    def __init__(self, owner, row, col, height, width):
        super().__init__(owner.journal, Rows=N(Count=height), Columns=N(Count=width),
                         Areas=N(Count=1), NumberFormat='General', HorizontalAlignment=1)
        self.__dict__.update(owner=owner, row=row, col=col)
        def name(number):
            text = ''
            while number:
                number, remainder = divmod(number - 1, 26)
                text = chr(65 + remainder) + text
            return text
        self.__dict__['Address'] = (f'${name(col)}${row}:'
                                    f'${name(col + width - 1)}${row + height - 1}')
        self.__dict__['Cells'] = N(Item=self.item)
        self.__dict__['Font'] = self.item(1, 1).Font
        self.__dict__['Interior'] = self.item(1, 1).Interior
        self.__dict__['EntireColumn'] = self.item(1, 1)
    def __setattr__(self, name, value):
        super().__setattr__(name, value)
        if name in ('NumberFormat', 'HorizontalAlignment') and 'owner' in self.__dict__:
            for r in range(1, self.Rows.Count + 1):
                for c in range(1, self.Columns.Count + 1):
                    self.item(r, c).__dict__[name] = value
    def item(self, row, col):
        if not (1 <= row <= self.Rows.Count and 1 <= col <= self.Columns.Count):
            raise AssertionError('Cell index outside requested rectangle')
        return self.owner.cells[self.row + row - 1, self.col + col - 1]
    def Calculate(self):
        self._journal.append((self, 'Calculate', None))


class Sheet:
    Name = 'Data'
    CodeName = 'Sheet1'
    ProtectContents = ProtectDrawingObjects = ProtectScenarios = False
    def __init__(self, journal):
        self.journal = journal
        self.cells = {(r, c): Cell(journal) for r in range(1, 6) for c in range(1, 7)}
        self.ranges = {}
        self.ListObjects = self.QueryTables = N(Count=0)
        self.PivotTables = lambda: N(Count=0)
        self.used_address = 'A1:F5'
    @property
    def UsedRange(self):
        return self.Range(self.used_address)
    def Range(self, address):
        bounds = x.parse_address(address)
        address = bounds['address']
        match = re.match(r'([A-Z]+)([0-9]+)', address)
        col = 0
        for char in match[1]: col = col * 26 + ord(char) - 64
        if address not in self.ranges:
            self.ranges[address] = Rectangle(self, int(match[2]), col, bounds['rows'], bounds['columns'])
        return self.ranges[address]


class Bridge:
    def __init__(self, wb, identity):
        self.wb, self.identity = wb, identity
        self.calls = 0
        self.on_select = None
    def select(self, pid, hwnd, full_name):
        self.calls += 1
        if self.on_select: self.on_select(self.calls)
        return self.wb, dict(self.identity)


class Tests(unittest.TestCase):
    def setUp(self):
        self.journal = []
        self.sheet = Sheet(self.journal)
        self.spec = {'name': 'Data', 'code_name': 'Sheet1'}
        self.target = dict(application='excel', pid=17, hwnd=42, name='owned.xlsx', full_name='owned.xlsx')
        self.wb = N(Worksheets=N(Item=lambda name: self.sheet), Application=N(Version='16.0'),
                    ReadOnly=False, ProtectStructure=False, HasVBProject=False,
                    MultiUserEditing=False, Saved=False, FileFormat=51)
        self.bridge = Bridge(self.wb, self.target)
        self.saves = []
        def save_copy(filename):
            self.saves.append(filename)
            Path(filename).write_bytes(b'owned mock workbook')
        self.wb.SaveCopyAs = save_copy
        self.reader = patch.object(x, 'read_range', side_effect=self.snapshot)
        self.reader.start()
        self.addCleanup(self.reader.stop)
    def snapshot(self, wb, sheet, address, **kwargs):
        rng = self.sheet.Range(address)
        values, formulas, types = [], [], []
        for r in range(1, rng.Rows.Count + 1):
            cells = [rng.item(r, c) for c in range(1, rng.Columns.Count + 1)]
            values.append([c.Value2 for c in cells])
            formulas.append([c.Formula for c in cells])
            types.append([{'kind': 'boolean' if type(c.Value2) is bool else 'number' if type(c.Value2) in (int, float) else 'string' if isinstance(c.Value2, str) else 'empty'} for c in cells])
        return dict(sheet=self.spec, address=x.parse_address(address)['address'],
                    data={'values': values, 'formulas_or_constants': formulas}, types={'values': types})
    def evidence(self, address='B2:C2'):
        return x.read_edit_range(self.wb, self.spec, address)
    def request(self, kind='set_values', address='B2:C2', **fields):
        op = dict(kind=kind, sheet=self.spec, address=address,
                  expected_sha256=self.evidence(address)['sha256'])
        if kind == 'set_values': op['values'] = [[12, '世界 😀']]
        op.update(fields)
        return dict(target=dict(self.target), operation=op)
    def rejected(self, req, error=Exception):
        self.journal.clear()
        with self.assertRaises(error): x.apply_request(self.bridge, req, commit=True)
        self.assertEqual(self.journal, [])
        self.assertEqual(self.saves, [])
    def test_default_preview_zero_writes(self):
        result = x.apply_request(self.bridge, self.request())
        self.assertFalse(result['committed'])
        self.assertFalse(result['saved_in_place'])
        self.assertEqual(self.journal, [])
    def test_stale_evidence_zero_writes(self):
        req = self.request()
        self.sheet.cells[2, 2].__dict__['Value2'] = 99
        self.rejected(req)
    def test_identity_mismatch_and_rebind_change_zero_writes(self):
        req = self.request()
        req['target']['hwnd'] += 1
        self.rejected(req)
        req = self.request()
        self.bridge.calls = 0
        self.bridge.on_select = lambda n: self.bridge.identity.update(name='replacement.xlsx') if n == 2 else None
        self.rejected(req)
    def test_evidence_changes_between_checks_zero_writes(self):
        req = self.request()
        def change(n):
            if n == 2: self.sheet.cells[2, 2].__dict__['Value2'] = 99
        self.bridge.on_select = change
        self.rejected(req)
    def test_spill_unknown_major16_and_unknown_version_fail_closed(self):
        req = self.request()
        for version in ('16.0', 'unknown', ''):
            with self.subTest(version=version):
                self.wb.Application.Version = version
                self.sheet.cells[2, 2].__dict__.pop('HasSpill', None)
                self.rejected(req)
    def test_only_confirmed_legacy_version_allows_missing_spill(self):
        self.wb.Application.Version = '15.0'
        del self.sheet.cells[2, 2].__dict__['HasSpill']
        result = x.apply_request(self.bridge, self.request(), commit=True)
        self.assertTrue(result['committed'])
    def test_protection_merge_array_spill_and_mixed_unknown_zero_writes(self):
        for obj, field, value in [(self.wb, 'ReadOnly', True), (self.wb, 'ProtectStructure', True),
                (self.sheet, 'ProtectContents', True), (self.sheet, 'ProtectDrawingObjects', True),
                (self.sheet.cells[2, 3], 'MergeCells', True), (self.sheet.cells[2, 3], 'HasArray', True),
                (self.sheet.cells[2, 3], 'HasSpill', True), (self.sheet.cells[2, 3], 'HasSpill', None)]:
            with self.subTest(field=field, value=value):
                req = self.request()
                old = getattr(obj, field)
                obj.__dict__[field] = value
                self.rejected(req)
                obj.__dict__[field] = old
    def test_literal_formula_injection_and_bad_unicode_zero_writes(self):
        for value in ('=1+2', '+SUM(A1:A2)', '-1+2', '@SUM(A1:A2)', chr(0xD800), float('nan')):
            with self.subTest(value=repr(value)):
                self.rejected(self.request(values=[[value, 'safe']]))
    def test_typed_values_unicode_and_target_bounds(self):
        for matrix in ([[12, '中文 😀']], [[True, None]], [[3.25, '00123']]):
            req = self.request(values=matrix)
            self.journal.clear()
            result = x.apply_request(self.bridge, req, commit=True)
            self.assertTrue(result['committed'])
            self.assertEqual(self.snapshot(self.wb, 'Data', 'B2:C2')['data']['values'], matrix)
            self.assertEqual(len(self.journal), 2)
            self.assertIsNone(self.sheet.cells[2, 1].Value2)
            self.assertIsNone(self.sheet.cells[2, 4].Value2)
    def test_formula_write_read_and_calculate(self):
        req = self.request('set_formulas', address='B2', formulas=[['=1+2']])
        x.apply_request(self.bridge, req, commit=True)
        evidence = self.evidence('B2')
        self.assertEqual(evidence['formulas'], [['=1+2']])
        x.apply_request(self.bridge, self.request('calculate_range', address='B2'), commit=True)
        self.assertEqual(sum(name == 'Calculate' for _, name, _ in self.journal), 1)
    def test_formula_external_and_dynamic_grammar_rejected(self):
        for formula in ('=WEBSERVICE("https://example.com")', '=SEQUENCE(3)',
                        "='[other.xlsx]Data'!A1", '=HYPERLINK("x")',
                        '=B2C2', '=-B2', '=B2:'):
            self.rejected(self.request('set_formulas', address='B2', formulas=[[formula]]))
    def test_styles_preview_commit_and_invalid_style(self):
        req = self.request('format_range', address='B2', number_format='0.00', bold=True,
                           font_size=14, font_name='等线', color=255, fill_color=65535,
                           horizontal_alignment='center', column_width=18)
        x.apply_request(self.bridge, req)
        self.assertEqual(self.journal, [])
        x.apply_request(self.bridge, req, commit=True)
        rng = self.sheet.Range('B2')
        self.assertEqual((rng.NumberFormat, rng.Font.Bold, rng.Font.Name, rng.Interior.Color,
                          rng.HorizontalAlignment, rng.EntireColumn.ColumnWidth),
                         ('0.00', True, '等线', 65535, -4108, 18))
        self.rejected(self.request('format_range', address='B2', bold='yes'))
    def test_save_copy_no_clobber_and_preserves_source_saved(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'copy.xlsx'
            req = {'target': self.target, 'operation': {'kind': 'save_copy', 'path': str(path)}}
            x.apply_request(self.bridge, req)
            self.assertEqual(self.saves, [])
            self.assertFalse(path.exists())
            for saved in (False, True):
                self.wb.Saved = saved
                path = Path(directory) / ('saved.xlsx' if saved else 'dirty.xlsx')
                req['operation']['path'] = str(path)
                x.apply_request(self.bridge, req, commit=True)
                self.assertEqual(path.read_bytes(), b'owned mock workbook')
                self.assertEqual(self.wb.Saved, saved)
                count = len(self.saves)
                with self.assertRaises(Exception): x.apply_request(self.bridge, req, commit=True)
                self.assertEqual(len(self.saves), count)
                self.assertEqual(path.read_bytes(), b'owned mock workbook')
    def test_partial_write_unknown_no_retry_or_rollback(self):
        req = self.request()
        self.sheet.cells[2, 3].__dict__['fail_after'] = True
        with self.assertRaises(x.WriteOutcomeUnknown): x.apply_request(self.bridge, req, commit=True)
        self.assertEqual(len(self.journal), 2)
        self.assertEqual(self.sheet.cells[2, 2].Value2, 12)
        self.assertEqual(self.sheet.cells[2, 3].Value2, '世界 😀')
    def test_postcondition_mismatch_unknown(self):
        req = self.request()
        self.sheet.cells[2, 3].__dict__['ignore_write'] = True
        with self.assertRaisesRegex(x.WriteOutcomeUnknown, '(?i)postcondition|mismatch'):
            x.apply_request(self.bridge, req, commit=True)
        self.assertEqual(len(self.journal), 2)

    def test_bounded_scalar_sheet_allows_missing_spill_only_with_clean_inventory(self):
        self.wb.Application.Version = '16.0'
        del self.sheet.cells[2, 2].__dict__['HasSpill']
        evidence = x.read_edit_range(
            self.wb, self.spec, 'B2:C2', spill_policy='bounded_scalar_sheet'
        )
        proof = evidence['structural_checks']['spill_proof']
        self.assertEqual(proof['used_address'], 'A1:F5')
        self.assertEqual(proof['formula_cells'], [])
        req = {'target': dict(self.target), 'operation': {
            'kind': 'set_values', 'sheet': self.spec, 'address': 'B2:C2',
            'expected_sha256': evidence['sha256'], 'spill_policy': 'bounded_scalar_sheet',
            'values': [[12, '世界 😀']],
        }}
        self.journal.clear()
        result = x.apply_request(self.bridge, req, commit=True)
        self.assertTrue(result['committed'])
        self.assertEqual(len(self.journal), 2)
        self.assertEqual(self.sheet.cells[2, 2].Value2, 12)
        self.assertEqual(self.sheet.cells[2, 3].Value2, '世界 😀')

    def test_bounded_scalar_sheet_rejects_any_dynamic_formula_or_array(self):
        self.wb.Application.Version = '16.0'
        del self.sheet.cells[2, 2].__dict__['HasSpill']
        self.sheet.cells[4, 4].__dict__.update(Formula='=SEQUENCE(2)', HasFormula=True)
        with self.assertRaises(ValueError):
            x.read_edit_range(
                self.wb, self.spec, 'B2:C2', spill_policy='bounded_scalar_sheet'
            )
        self.sheet.cells[4, 4].__dict__.update(Formula=None, HasFormula=False)
        self.sheet.cells[4, 5].__dict__['HasArray'] = True
        with self.assertRaises(ValueError):
            x.read_edit_range(
                self.wb, self.spec, 'B2:C2', spill_policy='bounded_scalar_sheet'
            )


if __name__ == '__main__':
    unittest.main()
