"""Opt-in Excel integration test. Only writes/closes an owned, unsaved instance."""
import json
import sys
from pathlib import Path
import time
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from excel_com import ExcelBridge, normalized_hwnd, probe, read_range, inspect_range


def main():
    if '--run' not in sys.argv:
        raise SystemExit('Explicit --run required; creates own temporary Excel instance')
    from comtypes.client import CreateObject
    app = CreateObject('Excel.Application', dynamic=True)
    owned = []
    passed = []
    try:
        app.Visible = True
        wb = app.Workbooks.Add()
        owned.append(wb)
        sheet = wb.Worksheets.Item(1)
        sheet.Name = 'Data'
        # Assign through individual cells; the automation SAFEARRAY tuple shape is backend-specific.
        for address, value in [('A1', 12), ('B1', '中文'), ('C1', True),
                               ('A2', 3.5), ('B2', 'text'), ('C2', False)]:
            sheet.Range(address).Value2 = value
        # Construct supplementary Unicode inside Excel, independently of BSTR setter marshalling.
        sheet.Range('B1').Formula = '=UNICHAR(128512)'
        sheet.Range('D1').Formula = '=1+2'
        bridge = ExcelBridge()
        hwnd = normalized_hwnd(wb.Windows.Item(1).Hwnd)
        pid = bridge.pid(hwnd)
        other = app.Workbooks.Add()
        owned.append(other)  # Active workbook now differs from our target.
        attached, target = bridge.select(pid, hwnd)
        assert str(attached.Name) == str(wb.Name)
        passed.append('bind_target_with_other_workbook_active')
        before = bool(wb.Saved)
        data = read_range(attached, 'Data', 'A1:C2')['data']['values']
        assert data == [[12.0, chr(128512), True], [3.5, 'text', False]], data
        passed.append('unsaved_unicode_numeric_boolean_rectangle')
        assert read_range(attached, 'Data', 'A1')['data']['values'] == [[12.0]]
        assert len(read_range(attached, 'Data', 'A1:C1')['data']['values'][0]) == 3
        assert len(read_range(attached, 'Data', 'A1:A2')['data']['values']) == 2
        passed.append('scalar_row_column_shapes')
        formula = read_range(attached, 'Data', 'D1', mode='formulas')['data']['formulas_or_constants']
        assert formula == [['=1+2']], formula
        assert read_range(attached, 'Data', 'D1', mode='values')['data']['values'] == [[3.0]]
        passed.append('formula_and_cached_result')
        sheet.Range('E1').Formula = '=1/0'
        errors = read_range(attached, 'Data', 'E1', mode='values')
        assert errors['data']['values'][0][0]['name'] == '#DIV/0!', errors
        assert errors['types']['values'] == [[{'kind': 'error', 'vartype': 10}]]
        sheet.Range('F1').Value2 = 2007
        sheet.Range('E2').Formula = '=NA()'
        sheet.Range('F2').Value2 = -2146826281
        mixed = read_range(attached, 'Data', 'E1:F2')
        assert mixed['data']['values'][1][0]['name'] == '#N/A', mixed
        assert mixed['types']['values'][0][1]['kind'] == 'number'
        assert mixed['types']['values'][1][1]['kind'] == 'number'
        assert mixed['data']['formulas_or_constants'][0][0] == '=1/0'
        passed.append('scalar_and_array_errors_not_confused_with_numbers')
        risk = inspect_range(attached, 'Data', 'A1')
        assert not risk['write_authorized'], risk
        assert all('HasSpill' in x and 'unavailable' in x for x in risk['blockers']), risk
        if risk['blockers']:
            passed.append('legacy_HasSpill_unavailable_fails_closed')
        assert any('HasFormula' in x for x in inspect_range(attached, 'Data', 'D1')['blockers'])
        sheet.Range('H1:I1').Merge()
        assert any('MergeCells' in x for x in inspect_range(attached, 'Data', 'H1')['blockers'])
        sheet.Range('H2:I2').FormulaArray = '=1+2'
        assert any('HasArray' in x for x in inspect_range(attached, 'Data', 'H2')['blockers'])
        if not risk['blockers']:
            sheet.Range('K1').Formula2 = '=SEQUENCE(2)'
            assert any('HasSpill' in x for x in inspect_range(attached, 'Data', 'K2')['blockers'])
            passed.append('dynamic_array_spill_detected')
        sheet.Range('M1').Value2 = 'Header'
        sheet.Range('M2').Value2 = 1
        table = sheet.ListObjects.Add(1, sheet.Range('M1:M2'), None, 1)
        assert any('ListObjects' in x for x in inspect_range(attached, 'Data', 'A1')['blockers'])
        table.Delete()
        sheet.Protect()
        assert any('ProtectContents' in x for x in inspect_range(attached, 'Data', 'A1')['blockers'])
        sheet.Unprotect()
        passed.append('inspect_detects_formula_merge_array_table_protection')
        started = time.perf_counter()
        inspect_range(attached, 'Data', 'A1:D5')
        inspect_seconds = time.perf_counter() - started
        assert inspect_seconds < 9, inspect_seconds
        passed.append('inspect_20_cells_under_9_seconds')
        sheet.Range('A4:J23').Formula = '=ROW()*100+COLUMN()'
        tiled = read_range(attached, 'Data', 'A4:J23', mode='values')['data']['values']
        assert tiled[0][0] == 401 and tiled[-1][-1] == 2310
        passed.append('multiple_bounded_tiles')
        assert probe(attached)['worksheet_count'] >= 1
        assert bool(wb.Saved) == before == False
        passed.append('reads_preserve_unsaved_state')
        print(json.dumps({'passed': passed, 'inspect_20_cells_seconds': round(inspect_seconds, 3), 'user_workbooks_modified': False, 'saved_files': False}, indent=2))
    finally:
        primary = sys.exc_info()[1]
        cleanup_errors = []
        for owned_wb in reversed(owned):
            try:
                owned_wb.Close(False)
            except Exception as exc:
                cleanup_errors.append(f'Workbook.Close: {exc}')
        try:
            app.Quit()
        except Exception as exc:
            cleanup_errors.append(f'Application.Quit: {exc}')
        if cleanup_errors:
            detail = '; '.join(cleanup_errors)
            if primary is not None:
                print(f'Cleanup errors (original failure preserved): {detail}', file=sys.stderr)
            else:
                raise RuntimeError(f'Owned Excel cleanup failed: {detail}')


if __name__ == '__main__':
    main()
