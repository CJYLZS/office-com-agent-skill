"""Opt-in live regression. Only create, edit, save and close workbooks owned here.

Run: .venv/Scripts/python.exe tests/live_excel_tools.py --run
No user workbook is selected, and Application.Quit is never called.
"""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', action='store_true', help='Create owned live Excel workbooks')
    args = parser.parse_args(argv)
    if not args.run:
        parser.error('Explicit --run required; no Excel connection was made')
    from comtypes.client import CreateObject
    from excel_com import ExcelBridge, normalized_hwnd, read_range
    import excel_tools as x
    owned = []
    passed, timings = [], {}
    started_all = time.perf_counter()
    root = ROOT / 'tmp'
    root.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='live-excel-tools-', dir=root) as folder:
        folder = Path(folder)
        app = CreateObject('Excel.Application', dynamic=True)
        try:
            app.Visible = True
            wb = app.Workbooks.Add()
            owned.append(wb)
            sheet = wb.Worksheets.Item(1)
            sheet.Name = 'Data'
            # Every path is unique under the runner's new temporary directory.
            source = folder / 'source.xlsx'
            assert not source.exists()
            wb.SaveAs(str(source), 51)
            sheet.Range('A2').Value2 = 'left sentinel'
            sheet.Range('F2').Value2 = 'right sentinel'
            spec = {'name': str(sheet.Name), 'code_name': str(sheet.CodeName)}
            bridge = ExcelBridge()
            hwnd = normalized_hwnd(wb.Windows.Item(1).Hwnd)
            pid = bridge.pid(hwnd)
            other = app.Workbooks.Add()
            owned.append(other)
            other.Worksheets.Item(1).Range('B2').Value2 = 'other workbook sentinel'
            attached, target = bridge.select(pid, hwnd, str(wb.FullName))
            assert str(attached.FullName) == str(wb.FullName)
            assert str(app.ActiveWorkbook.Name) == str(other.Name)
            passed.append('explicit_target_with_another_workbook_active')
            policy = 'native'
            try:
                x.read_edit_range(attached, spec, 'B2:E2')
            except ValueError as exc:
                assert 'HasSpill' in str(exc), str(exc)
                assert sheet.Range('B2').Value2 is None
                policy = 'bounded_scalar_sheet'
                passed.append('native_unknown_spill_fails_closed')

            def evidence(address):
                return x.read_edit_range(attached, spec, address, spill_policy=policy)

            def request(kind, address='B2:E2', **fields):
                op = dict(kind=kind, sheet=spec, address=address,
                          expected_sha256=evidence(address)['sha256'], spill_policy=policy)
                op.update(fields)
                return {'target': target, 'operation': op}

            def apply(req, commit=True):
                before = time.perf_counter()
                result = x.apply_request(bridge, req, commit=commit)
                timings[req['operation']['kind']] = round(time.perf_counter() - before, 3)
                return result

            def refused(req):
                before = read_range(attached, 'Data', 'A1:F5')['data']
                saved = bool(wb.Saved)
                try:
                    x.apply_request(bridge, req, commit=True)
                except ValueError:
                    pass
                else:
                    raise AssertionError('Expected a pre-write rejection')
                assert read_range(attached, 'Data', 'A1:F5')['data'] == before
                assert bool(wb.Saved) == saved

            values = [[12.5, True, '中文 😀', '00123']]
            req = request('set_values', values=values)
            before = read_range(attached, 'Data', 'B2:E2')['data']
            assert apply(req, commit=False)['committed'] is False
            assert read_range(attached, 'Data', 'B2:E2')['data'] == before
            apply(req)
            assert read_range(attached, 'Data', 'B2:E2')['data']['values'] == values
            assert sheet.Range('A2').Value2 == 'left sentinel'
            assert sheet.Range('F2').Value2 == 'right sentinel'
            assert other.Worksheets.Item(1).Range('B2').Value2 == 'other workbook sentinel'
            passed.append('preview_zero_writes_typed_unicode_values_and_surrounding_sentinels')

            stale = request('set_values', values=[[0, False, 'new', 'new']])
            sheet.Range('B2').Value2 = 99
            refused(stale)
            passed.append('stale_local_evidence_zero_writes')
            for text in ('=1+2', '  +1+2', '@SUM(B2:B3)'):
                refused(request('set_values', address='B3', values=[[text]]))
            passed.append('literal_formula_injection_zero_writes')

            apply(request('set_formulas', address='B3', formulas=[['=SUM(B2:B2)']]))
            assert evidence('B3')['formulas'] == [['=SUM(B2:B2)']]
            apply(request('calculate_range', address='B3'))
            assert read_range(attached, 'Data', 'B3')['data']['values'] == [[99.0]]
            passed.append('formula_roundtrip_range_calculation')
            apply(request('format_range', address='B2', number_format='0.00', bold=True,
                          font_size=14, font_name='Arial', color=255, fill_color=65535,
                          horizontal_alignment='center', column_width=18))
            assert bool(sheet.Range('B2').Font.Bold)
            assert str(sheet.Range('B2').NumberFormat) == '0.00'
            assert int(sheet.Range('B2').Interior.Color) == 65535
            assert abs(float(sheet.Range('B2').ColumnWidth) - 18) < .1
            passed.append('format_fields_verified_without_content_change')

            protection_req = request('set_values', address='B2', values=[[1]])
            sheet.Protect()
            try:
                refused(protection_req)
            finally:
                sheet.Unprotect()
            # A mixed merged/unmerged rectangle must be rejected per cell.
            merge_req = request('set_values', address='B4:D4', values=[[1, 2, 3]])
            sheet.Range('C4:D4').Merge()
            try:
                refused(merge_req)
            finally:
                sheet.Range('C4:D4').UnMerge()
            array_req = request('set_values', address='B5', values=[[1]])
            sheet.Range('B5:C5').FormulaArray = '=1+2'
            try:
                refused(array_req)
            finally:
                sheet.Range('B5:C5').ClearContents()
            passed.append('protected_merged_and_array_cells_zero_writes')

            # SaveCopyAs must preserve both the on-disk source and its dirty flag.
            original_bytes = source.read_bytes()
            for saved in (False, True):
                if saved:
                    wb.Save()  # Only this newly-created owned source is saved.
                    original_bytes = source.read_bytes()
                else:
                    sheet.Range('E3').Value2 = 'owned dirty change'
                state = bool(wb.Saved)
                destination = folder / ('clean-copy.xlsx' if saved else 'dirty-copy.xlsx')
                copy_req = {'target': target, 'operation': {'kind': 'save_copy', 'path': str(destination)}}
                apply(copy_req, commit=False)
                assert not destination.exists()
                apply(copy_req)
                assert destination.exists()
                assert bool(wb.Saved) == state
                assert source.read_bytes() == original_bytes
                digest = hashlib.sha256(destination.read_bytes()).digest()
                try:
                    x.apply_request(bridge, copy_req, commit=True)
                except (ValueError, FileExistsError):
                    pass
                else:
                    raise AssertionError('Existing destination was not rejected')
                assert hashlib.sha256(destination.read_bytes()).digest() == digest
            passed.append('save_copy_no_overwrite_preserves_source_bytes_and_saved_state')

            before_cli = bool(wb.Saved)
            command = [sys.executable, str(ROOT / 'scripts' / 'excel_tools.py'), 'read',
                       '--pid', str(pid), '--hwnd', str(hwnd), '--document', str(wb.FullName),
                       '--sheet', 'Data', '--range', 'B2', '--spill-policy', policy]
            started = time.perf_counter()
            result = subprocess.run(command, capture_output=True, text=True, encoding='utf-8', timeout=12)
            timings['supervised_cli_read'] = round(time.perf_counter() - started, 3)
            assert result.returncode == 0, result.stderr or result.stdout
            assert timings['supervised_cli_read'] < 9, timings
            report = json.loads(result.stdout)
            assert report['target'] == target
            assert report['data']['snapshot']['data']['values'] == [[99.0]]
            assert bool(wb.Saved) == before_cli
            passed.append('supervised_cli_read_under_9_seconds')
            timings['total'] = round(time.perf_counter() - started_all, 3)
            print(json.dumps({'passed': passed, 'timings_seconds': timings,
                              'spill_policy': policy, 'user_workbooks_modified': False,
                              'application_quit_called': False}, ensure_ascii=False, indent=2))
        finally:
            primary = sys.exc_info()[1]
            failures = []
            for workbook in reversed(owned):
                try:
                    workbook.Close(False)
                except Exception as exc:
                    failures.append(str(exc))
            # A COM server can be shared with a user. Never call Application.Quit.
            if failures:
                message = 'Owned workbook cleanup failed: ' + '; '.join(failures)
                if primary is None:
                    raise RuntimeError(message)
                print(message, file=sys.stderr)


if __name__ == '__main__':
    main()
