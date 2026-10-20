"""Capability diagnostics are read-only evidence, never automatic fallback."""
import contextlib
import io
import json
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace as N
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import excel_capabilities as c
import excel_tools as x
import test_excel_tools as fixtures


class Tests(unittest.TestCase):
    def setUp(self):
        self.journal = []
        self.sheet = fixtures.Sheet(self.journal)
        self.spec = {'name': 'Data', 'code_name': 'Sheet1'}
        self.wb = N(Worksheets=N(Item=lambda name: self.sheet),
                    Application=N(Version='16.0', Build='14026.0'), Saved=False,
                    ReadOnly=False, ProtectStructure=False, HasVBProject=False,
                    MultiUserEditing=False)

    def probe(self, policy='native'):
        result = c.probe_capabilities(self.wb, self.spec, 'B2:C2', policy)
        self.assertEqual(self.journal, [])
        self.assertFalse(result['write_authorized'])
        self.assertFalse(result['safety']['write_authorized'])
        self.assertFalse(result['dynamic_arrays_validated'])
        self.assertTrue(result['saved_state_unchanged'])
        json.dumps(result, allow_nan=False)
        return result

    def test_readable_not_setter_or_authorization(self):
        result = self.probe()
        p = result['sample']['properties']['HasSpill']
        self.assertTrue(p['readable'])
        self.assertFalse(p['boolean_value'])
        self.assertIsNone(p['setter_supported'])
        self.assertFalse(p['write_tested'])
        self.assertTrue(result['safety']['checked_structure_clear'])

    def test_name_resolved_does_not_prove_getter(self):
        class Fails:
            @property
            def Formula2(self):
                raise RuntimeError('getter unavailable')
        with patch.object(c, '_resolve_name', return_value=True):
            result = c._property(Fails(), 'Formula2')
        self.assertTrue(result['name_resolved'])
        self.assertFalse(result['readable'])
        self.assertEqual(result['read_error']['type'], 'RuntimeError')
        self.assertIsNone(result['setter_supported'])

    def test_missing_spill_blocks_no_policy_fallback(self):
        for cell in self.sheet.cells.values():
            del cell.__dict__['HasSpill']
        result = self.probe()
        self.assertEqual(result['spill_policy'], 'native')
        self.assertFalse(result['safety']['checked_structure_clear'])
        self.assertTrue(any('HasSpill' in b for b in result['safety']['blockers']))
        self.assertNotIn('spill_proof', result['safety']['checks'])
        explicit = self.probe('bounded_scalar_sheet')
        self.assertTrue(explicit['safety']['checked_structure_clear'])
        self.assertIn('spill_proof', explicit['safety']['checks'])

    def test_each_target_cell_is_checked_not_just_sample(self):
        self.sheet.cells[2, 3].__dict__['HasSpill'] = True
        result = self.probe()
        self.assertFalse(result['sample']['properties']['HasSpill']['boolean_value'])
        self.assertFalse(result['safety']['checked_structure_clear'])
        self.assertIn('cell[1,2].HasSpill', result['safety']['blockers'])

    def test_mixed_unknown_and_protection_remain_blockers(self):
        self.wb.ReadOnly = True
        self.sheet.cells[2, 3].__dict__['HasArray'] = None
        result = self.probe()
        self.assertIn('workbook.ReadOnly', result['safety']['blockers'])
        self.assertIn('cell[1,2].HasArray', result['safety']['blockers'])

    def test_large_used_range_refuses_inventory_without_scanning(self):
        self.sheet.used_address = 'A1:Z100000'
        result = self.probe('bounded_scalar_sheet')
        self.assertFalse(result['safety']['checked_structure_clear'])
        self.assertIn('bounded_scalar_inventory: ValueError', result['safety']['blockers'])
        with self.assertRaises(ValueError):
            x._structure(self.wb, self.sheet, self.sheet.Range('B2:C2'),
                         x.parse_address('B2:C2'), 'bounded_scalar_sheet')

    def test_getter_probe_omits_formula_and_value_contents(self):
        cell = self.sheet.cells[2, 2]
        cell.__dict__.update(Value2='private content', Formula='private content')
        result = self.probe()
        self.assertNotIn('private content', json.dumps(result))
        self.assertNotIn('sha256', result)

    def test_invalid_target_policy_and_sheet_identity_refused(self):
        for address in ('A1:Z1000', 'A:A', 'A1,B2'):
            with self.assertRaises(ValueError):
                c.probe_capabilities(self.wb, self.spec, address)
        with self.assertRaises(ValueError):
            self.probe('automatic')
        with self.assertRaises(ValueError):
            c.probe_capabilities(self.wb, {'name': 'Data', 'code_name': 'other'}, 'B2')
        self.assertEqual(self.journal, [])

    def test_bounded_inventory_redacts_formulas_outside_target(self):
        cell = self.sheet.cells[1, 1]
        cell.__dict__.update(HasFormula=True, Formula='=123456+654321')
        result = self.probe('bounded_scalar_sheet')
        proof = result['safety']['checks']['spill_proof']
        self.assertEqual(proof['formula_cell_count'], 1)
        self.assertNotIn('formula_cells', proof)
        self.assertNotIn('123456', json.dumps(result))

    def test_legacy_version_assumption_not_native_property_evidence(self):
        self.wb.Application.Version = '15.0'
        for cell in self.sheet.cells.values():
            del cell.__dict__['HasSpill']
        result = self.probe()
        self.assertFalse(result['safety']['checked_structure_clear'])
        self.assertIn('legacy_version_assumption_not_native_spill_evidence',
                      result['safety']['blockers'])

    def test_cli_commit_is_rejected_without_bridge(self):
        with patch.object(x, 'ExcelBridge') as bridge, contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                x.main(['capabilities', '--commit'])
        bridge.assert_not_called()

    def test_cli_dispatch_and_rebind_without_read_edit_evidence(self):
        identity = {'application': 'excel', 'pid': 1, 'hwnd': 2,
                    'name': 'owned.xlsx', 'full_name': 'owned.xlsx'}
        bridge = fixtures.Bridge(self.wb, identity)
        output = io.StringIO()
        with patch.object(x, 'ExcelBridge', return_value=bridge), \
                patch.object(c, 'probe_capabilities', return_value={'write_authorized': False}) as probe, \
                patch.object(x, 'read_edit_range') as edit_reader, \
                contextlib.redirect_stdout(output):
            code = x.main(['capabilities', '--worker', '--pid', '1', '--hwnd', '2',
                           '--sheet', 'Data', '--range', 'B2:C2'])
        self.assertEqual(code, 0)
        self.assertEqual(bridge.calls, 2)
        probe.assert_called_once_with(self.wb, self.spec, 'B2:C2', 'native')
        edit_reader.assert_not_called()
        self.assertFalse(json.loads(output.getvalue())['data']['write_authorized'])

    def test_cli_failure_redacts_unhandled_exception_text(self):
        output = io.StringIO()
        with patch.object(x, 'ExcelBridge', side_effect=RuntimeError('private document text')), \
                contextlib.redirect_stderr(output):
            code = x.main(['capabilities', '--worker'])
        self.assertEqual(code, 1)
        report = json.loads(output.getvalue())
        self.assertEqual(report['type'], 'RuntimeError')
        self.assertNotIn('private document text', output.getvalue())

    def test_cli_watchdog_does_not_retry(self):
        output = io.StringIO()
        with patch.object(x.subprocess, 'run', side_effect=subprocess.TimeoutExpired('worker', 9)) as run, \
                patch.object(x, 'ExcelBridge') as bridge, contextlib.redirect_stderr(output):
            code = x.main(['capabilities'])
        self.assertEqual(code, 2)
        run.assert_called_once()
        self.assertEqual(run.call_args.kwargs['timeout'], 9)
        bridge.assert_not_called()
        self.assertEqual(json.loads(output.getvalue())['status'], 'error')

    def test_runtime_metadata_unavailable_is_not_feature_inference(self):
        self.wb.Application = N()
        result = self.probe()
        self.assertIsNone(result['runtime']['version'])
        self.assertIsNone(result['runtime']['build'])
        self.assertFalse(result['dynamic_arrays_validated'])


if __name__ == '__main__':
    unittest.main()
