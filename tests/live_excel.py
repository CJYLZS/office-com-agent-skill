"""Opt-in Excel integration test. Only writes/closes an owned, unsaved instance."""
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from excel_com import ExcelBridge, normalized_hwnd, probe, read_range


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
        try:
            read_range(attached, 'Data', 'E1', mode='values')
        except ValueError as exc:
            assert 'variant cannot be decoded' in str(exc), str(exc)
        else:
            raise AssertionError('Error cell unexpectedly decoded; review backend policy')
        passed.append('error_cell_explicit_failure')
        sheet.Range('A4:J23').Formula = '=ROW()*100+COLUMN()'
        tiled = read_range(attached, 'Data', 'A4:J23', mode='values')['data']['values']
        assert tiled[0][0] == 401 and tiled[-1][-1] == 2310
        passed.append('multiple_bounded_tiles')
        assert probe(attached)['worksheet_count'] >= 1
        assert bool(wb.Saved) == before == False
        passed.append('reads_preserve_unsaved_state')
        print(json.dumps({'passed': passed, 'user_workbooks_modified': False, 'saved_files': False}, indent=2))
    finally:
        try:
            for wb in reversed(owned):
                wb.Close(False)
        finally:
            app.Quit()


if __name__ == '__main__':
    main()
