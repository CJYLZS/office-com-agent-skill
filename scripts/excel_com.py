"""Read-only Excel COM adapter. Attach to open desktop windows; never Open/Save/Quit."""
from __future__ import annotations
import argparse
import ctypes
import hashlib
import json
import math
from functools import lru_cache
import os
import tempfile
from pathlib import Path
import re
import subprocess
import sys
import time
from office_com import OfficeBridge, normalized_hwnd

MAX_CELLS = 10000
MAX_CELLS_PER_COM_READ = 100
MAX_TEXT_CHARS = 1000000


def parse_address(address, max_cells=MAX_CELLS):
    """Only a single explicit A1 cell/rectangle; reject names, formulas and whole columns."""
    if not isinstance(address, str) or len(address) > 64:
        raise ValueError('Expected a bounded A1 rectangle')
    match = re.fullmatch(r'\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?', address)
    if not match:
        raise ValueError('Use A1 or A1:D20, without sheet/workbook/name/union references')
    a, row1, b, row2 = match.groups()
    b, row2 = b or a, row2 or row1
    def column(label):
        n = 0
        for char in label.upper():
            n = n * 26 + ord(char) - ord('A') + 1
        return n
    c1, c2, r1, r2 = column(a), column(b), int(row1), int(row2)
    if not 1 <= c1 <= c2 <= 16384 or not 1 <= r1 <= r2 <= 1048576:
        raise ValueError('Rectangle is reversed or outside Excel limits')
    rows, cols = r2-r1+1, c2-c1+1
    if rows * cols > max_cells:
        raise ValueError(f'Range exceeds {max_cells} cell read budget')
    return {'address': f'{a.upper()}{r1}:{b.upper()}{r2}', 'rows': rows, 'columns': cols}


def range_tiles(address):
    """Yield rectangles <=100 cells; bound materialization before text accounting."""
    bounds = parse_address(address)
    start = bounds['address'].split(':')[0]
    label = ''.join(char for char in start if char.isalpha())
    first_row = int(start[len(label):])
    first_col = 0
    for char in label:
        first_col = first_col * 26 + ord(char) - 64
    def col_name(number):
        result = ''
        while number:
            number, rem = divmod(number-1, 26)
            result = chr(65+rem) + result
        return result
    width = min(bounds['columns'], MAX_CELLS_PER_COM_READ)
    height = max(1, MAX_CELLS_PER_COM_READ // width)
    for r in range(0, bounds['rows'], height):
        h = min(height, bounds['rows']-r)
        for c in range(0, bounds['columns'], width):
            w = min(width, bounds['columns']-c)
            a = f'{col_name(first_col+c)}{first_row+r}'
            b = f'{col_name(first_col+c+w-1)}{first_row+r+h-1}'
            yield r, c, h, w, f'{a}:{b}'


def matrix(value, rows, columns):
    """Normalize the scalar single-cell and 2D SAFEARRAY results without guessing."""
    if rows == columns == 1:
        result = [[value]]
    else:
        if not isinstance(value, (tuple, list)) or len(value) != rows:
            raise ValueError('Unexpected COM row shape')
        result = [list(row) if isinstance(row, (tuple, list)) else [] for row in value]
        if any(len(row) != columns for row in result):
            raise ValueError('Unexpected COM column shape')
    for row in result:
        for item in row:
            if item is not None and type(item) not in (str, int, float, bool):
                raise ValueError('Unsupported COM value type; no implicit string conversion')
            if isinstance(item, str):
                try:
                    item.encode('utf-8')
                except UnicodeEncodeError as exc:
                    raise ValueError('Unpaired Unicode surrogate; refusing to publish damaged text') from exc
    return result


ERROR_NAMES = {2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!',
               2029: '#NAME?', 2036: '#NUM!', 2042: '#N/A', 2043: '#GETTING_DATA',
               2045: '#SPILL!', 2050: '#CALC!'}


@lru_cache(maxsize=1)
def automation_api():
    """Local ABI bindings; never replace comtypes' global VARIANT decoder."""
    dll = ctypes.WinDLL('oleaut32')
    signatures = {
        'SafeArrayGetDim': (ctypes.c_uint, [ctypes.c_void_p]),
        'SafeArrayGetLBound': (ctypes.c_long, [ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_long)]),
        'SafeArrayGetUBound': (ctypes.c_long, [ctypes.c_void_p, ctypes.c_uint, ctypes.POINTER(ctypes.c_long)]),
        'SafeArrayGetElement': (ctypes.c_long, [ctypes.c_void_p, ctypes.POINTER(ctypes.c_long), ctypes.c_void_p]),
        'SysStringLen': (ctypes.c_uint, [ctypes.c_void_p]),
        'SysFreeString': (None, [ctypes.c_void_p]),
        'VariantClear': (ctypes.c_long, [ctypes.c_void_p]),
    }
    for name, (restype, argtypes) in signatures.items():
        function = getattr(dll, name)
        function.restype, function.argtypes = restype, argtypes
    return dll


def check_hresult(hr):
    if hr < 0:
        raise ValueError(f'Automation API failed: HRESULT={hr}')


def decode_scalar(raw):
    """Decode only explicitly supported scalar VARTYPEs, with no numeric error guessing."""
    vt = int(raw.vt)
    if vt in (0, 1):
        value, kind = None, ('empty' if vt == 0 else 'null')
    elif vt == 10:
        scode = int(raw._.VT_I4)
        code = scode & 0xffffffff
        excel_code = code & 0xffff if (code & 0xffff0000) == 0x800a0000 else code
        value = {'type': 'excel_error', 'scode': scode, 'scode_hex': f'0x{code:08X}',
                 'excel_code': excel_code, 'name': ERROR_NAMES.get(excel_code)}
        kind = 'error'
    elif vt == 8:
        ptr = raw._.c_void_p
        size = automation_api().SysStringLen(ptr) if ptr else 0
        # Read length-counted UTF-16, not NUL-terminated wchar or Python codepoint length.
        value = ctypes.string_at(ptr, size * 2).decode('utf-16-le') if ptr else ''
        value.encode('utf-8')
        kind = 'string'
    elif vt == 11:
        value, kind = bool(raw._.VT_BOOL), 'boolean'
    elif vt in (2, 3, 4, 5, 16, 17, 18, 19, 20, 21, 22, 23):
        value, kind = raw.value, 'number'
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError('Non-finite COM number')
    else:
        raise ValueError(f'Unsupported COM VARTYPE {vt}; refusing implicit conversion')
    return value, {'kind': kind, 'vartype': vt}


def decode_variant(raw, rows, columns):
    """Copy a scalar/2D SAFEARRAY(VARIANT), retaining each element's original vt."""
    if int(raw.vt) != 0x200c:
        if rows != 1 or columns != 1:
            raise ValueError('Expected two-dimensional SAFEARRAY(VARIANT)')
        value, info = decode_scalar(raw)
        return [[value]], [[info]]
    from comtypes.automation import VARIANT
    api, ptr = automation_api(), raw._.c_void_p
    if not ptr or api.SafeArrayGetDim(ptr) != 2:
        raise ValueError('Expected two-dimensional SAFEARRAY(VARIANT)')
    lower = []
    for dimension, expected in ((1, rows), (2, columns)):
        lo, hi = ctypes.c_long(), ctypes.c_long()
        check_hresult(api.SafeArrayGetLBound(ptr, dimension, ctypes.byref(lo)))
        check_hresult(api.SafeArrayGetUBound(ptr, dimension, ctypes.byref(hi)))
        if hi.value - lo.value + 1 != expected:
            raise ValueError('SAFEARRAY shape differs from requested range')
        lower.append(lo.value)
    values, types = [], []
    for r in range(rows):
        value_row, type_row = [], []
        for c in range(columns):
            element = VARIANT()
            try:
                indices = (ctypes.c_long * 2)(lower[0]+r, lower[1]+c)
                check_hresult(api.SafeArrayGetElement(ptr, indices, ctypes.byref(element)))
                value, info = decode_scalar(element)
                value_row.append(value)
                type_row.append(info)
            finally:
                check_hresult(api.VariantClear(ctypes.byref(element)))
        values.append(value_row)
        types.append(type_row)
    return values, types


def clear_exception_strings(exception):
    """Free caller-owned EXCEPINFO BSTRs without creating temporary BSTR owners."""
    for name in ('bstrSource', 'bstrDescription', 'bstrHelpFile'):
        slot = ctypes.c_void_p.from_address(ctypes.addressof(exception) + getattr(type(exception), name).offset)
        pointer = slot.value
        if pointer:
            automation_api().SysFreeString(pointer)
            slot.value = None


def read_property(tile, prop, rows, columns):
    """Property-get with an owned raw result; no global monkeypatch or fallback guessing."""
    from comtypes.automation import IDispatch, VARIANT, DISPPARAMS, EXCEPINFO, IID_NULL
    dispatch = tile.QueryInterface(IDispatch)
    result, params, exception, argument = VARIANT(), DISPPARAMS(), EXCEPINFO(), ctypes.c_uint()
    try:
        dispatch._IDispatch__com_Invoke(dispatch.GetIDsOfNames(prop)[0], ctypes.byref(IID_NULL),
                                       0, 2, ctypes.byref(params), ctypes.byref(result),
                                       ctypes.byref(exception), ctypes.byref(argument))
        return decode_variant(result, rows, columns)
    finally:
        primary_error = sys.exc_info()[1]
        try:
            try:
                check_hresult(automation_api().VariantClear(ctypes.byref(result)))
            finally:
                clear_exception_strings(exception)
        except Exception as cleanup_error:
            if primary_error is None:
                raise
            if hasattr(primary_error, 'add_note'):
                primary_error.add_note(f'COM result cleanup failed: {cleanup_error}')


def matching_workbook(app, hwnd, limit=100):
    """Bind by workbook window handle, never ActiveWorkbook or caption guessing."""
    matches = []
    count = int(app.Workbooks.Count)
    if count > limit:
        raise ValueError('Too many workbooks for bounded binding')
    for i in range(1, count+1):
        wb = app.Workbooks.Item(i)
        windows = wb.Windows
        if int(windows.Count) > limit:
            raise ValueError('Too many workbook windows')
        for j in range(1, int(windows.Count)+1):
            if normalized_hwnd(windows.Item(j).Hwnd) == normalized_hwnd(hwnd):
                matches.append(wb)
                break
    if len(matches) != 1:
        raise ValueError('Workbook window binding is missing or ambiguous')
    return matches[0]


class ExcelBridge(OfficeBridge):
    # Reuse Win32 signatures/label/pid only; Word selection/attachment is not used.
    def windows(self):
        result = []
        @self.callback
        def collect(hwnd, _):
            if self.label(hwnd, True) == 'XLMAIN':
                result.append({'pid': self.pid(hwnd), 'hwnd': int(hwnd), 'caption': self.label(hwnd)})
            return True
        self.user32.EnumWindows(collect, 0)
        return result

    def attach(self, hwnd, expected_pid=None):
        if self.label(hwnd, True) != 'XLMAIN':
            raise ValueError('Target is no longer an Excel top-level window')
        pid = self.pid(hwnd)
        if expected_pid is not None and pid != expected_pid:
            raise ValueError('PID mismatch')
        children = []
        @self.callback
        def collect(child, _):
            if self.label(child, True) == 'EXCEL7':
                children.append(int(child))
            return True
        self.user32.EnumChildWindows(hwnd, collect, 0)
        errors = []
        for child in children:
            try:
                ptr = ctypes.POINTER(self.IDispatch)()
                iid = self.IDispatch._iid_
                hr = self.oleacc.AccessibleObjectFromWindow(child, 0xfffffff0, ctypes.byref(iid), ctypes.byref(ptr))
                if hr != 0:
                    raise RuntimeError(f'Native COM interface failed: HRESULT={hr}')
                if not ptr:
                    raise ValueError('Native object pointer is null')
                window = self.Dispatch(ptr)
                if normalized_hwnd(window.Hwnd) != normalized_hwnd(hwnd):
                    raise ValueError('Native window handle mismatch')
                app = window.Application
                if str(app.Name) != 'Microsoft Excel':
                    raise ValueError('Native application is not Microsoft Excel')
                wb = matching_workbook(app, hwnd)
                identity = {'application': 'excel', 'pid': pid, 'hwnd': hwnd,
                            'name': str(wb.Name), 'full_name': str(wb.FullName)}
                if self.pid(hwnd) != pid or self.label(hwnd, True) != 'XLMAIN':
                    raise ValueError('Window changed during attachment')
                return wb, identity
            except Exception as exc:
                errors.append(str(exc))
        raise RuntimeError('Cannot bind Excel workbook; auxiliary/protected/busy windows are not targets. ' + '; '.join(errors))

    def select(self, pid=None, hwnd=None, document=None):
        if type(pid) is not int or type(hwnd) is not int or pid <= 0 or hwnd <= 0:
            raise ValueError('Explicit positive PID and HWND required; run list first')
        wb, identity = self.attach(hwnd, pid)
        if document is not None and identity['full_name'] != document:
            raise ValueError('Workbook full name mismatch; list again')
        return wb, identity

    def inventory(self):
        items = []
        for window in self.windows():
            try:
                wb, target = self.attach(window['hwnd'], window['pid'])
                items.append({'target': target, 'caption': window['caption'],
                              'saved': bool(wb.Saved), 'read_only': bool(wb.ReadOnly)})
            except Exception as exc:
                items.append({**window, 'error': str(exc)})
        return {'windows': items, 'scope': 'desktop Excel windows only; hidden/protected views may not bind'}


def probe(wb, offset=1, limit=20):
    if type(offset) is not int or offset < 1 or type(limit) is not int or not 1 <= limit <= 100:
        raise ValueError('Worksheet offset >= 1, limit 1..100 required')
    count = int(wb.Worksheets.Count)
    sheets = []
    for i in range(offset, min(count+1, offset+limit)):
        sheet = wb.Worksheets.Item(i)
        sheets.append({'index': i, 'name': str(sheet.Name), 'code_name': str(sheet.CodeName),
                       'visible': int(sheet.Visible), 'protect_contents': bool(sheet.ProtectContents)})
    return {'saved': bool(wb.Saved), 'read_only': bool(wb.ReadOnly), 'path': str(wb.Path),
            'date1904': bool(wb.Date1904), 'protect_structure': bool(wb.ProtectStructure),
            'worksheet_count': count, 'sheets': sheets,
            'next_offset': offset+limit if offset+limit <= count else None,
            'scope': 'worksheets only; chart sheets omitted; no UsedRange scan'}


def read_range(wb, sheet_name, address, mode='both', formula_api='Formula'):
    bounds = parse_address(address)  # Validate before any COM access.
    if not isinstance(sheet_name, str) or not sheet_name:
        raise ValueError('Explicit worksheet name required')
    if mode not in ('values', 'formulas', 'both') or formula_api not in ('Formula', 'Formula2'):
        raise ValueError('Unsupported read mode/formula API')
    sheet = wb.Worksheets.Item(sheet_name)
    if str(sheet.Name) != sheet_name:
        raise ValueError('Worksheet name must match exactly')
    rng = sheet.Range(bounds['address'])
    rows, cols = bounds['rows'], bounds['columns']
    if int(rng.Rows.Count) != rows or int(rng.Columns.Count) != cols or int(rng.Areas.Count) != 1:
        raise ValueError('COM range differs from requested rectangle')
    properties = []
    if mode in ('values', 'both'):
        properties.append(('values', 'Value2'))
    if mode in ('formulas', 'both'):
        properties.append(('formulas_or_constants', formula_api))
    data = {key: [[None] * cols for _ in range(rows)] for key, _ in properties}
    types = {key: [[None] * cols for _ in range(rows)] for key, _ in properties}
    text_chars = 0
    for r, c, height, width, tile_address in range_tiles(bounds['address']):
        tile = sheet.Range(tile_address)
        for key, prop in properties:
            values, value_types = read_property(tile, prop, height, width)
            text_chars += sum(len(item) for row in values for item in row if isinstance(item, str))
            if text_chars > MAX_TEXT_CHARS:
                raise ValueError('Text budget exceeded after bounded COM transfer; narrow the rectangle')
            for i, row in enumerate(values):
                data[key][r+i][c:c+width] = row
                types[key][r+i][c:c+width] = value_types[i]
    # Sequential reads are not atomic. This digest is evidence, not a revision token/edit permit.
    payload = json.dumps({'data': data, 'types': types}, ensure_ascii=False, allow_nan=False, sort_keys=True, separators=(',', ':'))
    return {'sheet': {'name': str(sheet.Name), 'code_name': str(sheet.CodeName)},
            **bounds, 'mode': mode, 'formula_api': formula_api if mode != 'values' else None,
            'data': data, 'types': types, 'hash_scope': 'canonical_json_data_and_types_v2', 'data_sha256': hashlib.sha256(payload.encode('utf-8', errors='surrogatepass')).hexdigest(),
            'truncated': False,
            'warnings': ['Sequential live reads, not an atomic snapshot or edit authorization.',
                         'Value2 dates/currency are numeric; errors carry original VT_ERROR/SCODE, not numeric guesses.',
                         'Formula output includes constants; types describe returned VARIANTs, not a formula mask.',
                         'No calculation/refresh or formatted display text; no structural edit safety check.']}


def inspect_range(wb, sheet_name, address):
    """Read-only, conservative <=20-cell risk evidence. Never an edit authorization."""
    bounds = parse_address(address, max_cells=20)
    if not isinstance(sheet_name, str) or not sheet_name:
        raise ValueError('Explicit worksheet name required')
    sheet = wb.Worksheets.Item(sheet_name)
    if str(sheet.Name) != sheet_name:
        raise ValueError('Worksheet name must match exactly')
    rng = sheet.Range(bounds['address'])
    if (int(rng.Rows.Count) != bounds['rows'] or int(rng.Columns.Count) != bounds['columns']
            or int(rng.Areas.Count) != 1):
        raise ValueError('COM range differs from requested rectangle')
    blockers, checks = [], {}
    def require_false(label, get):
        try:
            value = get()
            checks[label] = value if type(value) is bool else None
            if type(value) is not bool or value is not False:
                blockers.append(label + ': true/mixed/unknown')
        except Exception:
            checks[label] = None
            blockers.append(label + ': unavailable')
    for name in ('ReadOnly', 'ProtectStructure', 'HasVBProject', 'MultiUserEditing'):
        require_false('workbook.'+name, lambda name=name: getattr(wb, name))
    for name in ('ProtectContents', 'ProtectDrawingObjects', 'ProtectScenarios'):
        require_false('sheet.'+name, lambda name=name: getattr(sheet, name))
    # Deliberately reject whole-sheet table/pivot/query contexts instead of guessing
    # overlap from incomplete aggregate Range.ListObject semantics.
    for name in ('ListObjects', 'PivotTables', 'QueryTables'):
        try:
            collection = getattr(sheet, name)
            if name == 'PivotTables':
                collection = collection()
            count = collection.Count
            checks['sheet.'+name] = count if type(count) is int else None
            if type(count) is not int or count != 0:
                blockers.append('sheet.'+name+': nonempty/unknown')
        except Exception:
            checks['sheet.'+name] = None
            blockers.append('sheet.'+name+': unavailable')
    for r in range(1, bounds['rows']+1):
        for c in range(1, bounds['columns']+1):
            cell = rng.Cells.Item(r, c)
            for name in ('MergeCells', 'HasArray', 'HasSpill', 'HasFormula'):
                require_false(f'cell[{r},{c}].{name}', lambda name=name, cell=cell: getattr(cell, name))
    snapshot = read_range(wb, sheet_name, bounds['address'])
    return {'sheet': snapshot['sheet'], **bounds, 'checks': checks, 'blockers': blockers,
            'checked_structure_clear': not blockers, 'write_authorized': False,
            'snapshot': snapshot,
            'warnings': ['Read-only risk evidence, not a complete structural audit or write permit.',
                         'Unknown/unavailable/mixed properties fail closed; old Excel without HasSpill is blocked.',
                         'Events/add-ins/recalculation and concurrent changes cannot be ruled out; hash is not a lock.']}


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('command', choices=['list', 'probe', 'read', 'inspect'])
    p.add_argument('--pid', type=int)
    p.add_argument('--hwnd', type=int)
    p.add_argument('--document', help='Expected FullName of already-open workbook; never opens it')
    p.add_argument('--sheet')
    p.add_argument('--range', dest='address')
    p.add_argument('--mode', choices=['values', 'formulas', 'both'], default='both')
    p.add_argument('--formula-api', choices=['Formula', 'Formula2'], default='Formula')
    p.add_argument('--offset', type=int, default=1)
    p.add_argument('--limit', type=int, default=20)
    p.add_argument('--output', type=Path)
    p.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    a = p.parse_args()
    if not a.worker:
        try:
            return subprocess.run([sys.executable, str(Path(__file__).resolve()), *sys.argv[1:], '--worker'], timeout=9).returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({'status': 'error', 'error': 'Read-only worker exceeded 9 seconds; no workbook writes'}), file=sys.stderr)
            return 2
    try:
        t = time.perf_counter()
        if a.command in ('read', 'inspect'):
            parse_address(a.address, max_cells=20 if a.command == 'inspect' else MAX_CELLS)
            if not a.sheet:
                raise ValueError('--sheet required')
        bridge = ExcelBridge()
        target = None
        if a.command == 'list':
            data = bridge.inventory()
        else:
            wb, target = bridge.select(a.pid, a.hwnd, a.document)
            before_saved = bool(wb.Saved)
            if a.command == 'probe':
                data = probe(wb, a.offset, a.limit)
            elif a.command == 'inspect':
                data = inspect_range(wb, a.sheet, a.address)
            else:
                data = read_range(wb, a.sheet, a.address, a.mode, a.formula_api)
            # Rebind for end-of-read identity check, not a concurrency lock.
            _, current = bridge.select(a.pid, a.hwnd, target['full_name'])
            if current != target:
                raise ValueError('Workbook identity changed during read')
            data['saved_before'] = before_saved
            data['saved_after'] = bool(wb.Saved)
        report = {'schema_version': 2, 'status': 'ok', 'operation': a.command, 'target': target,
                  'elapsed_seconds': round(time.perf_counter()-t, 3),
                  'limits': {'max_cells': MAX_CELLS, 'max_inspect_cells': 20, 'max_text_chars': MAX_TEXT_CHARS, 'supervised_cli_timeout_seconds': 9, 'max_cells_per_com_read': MAX_CELLS_PER_COM_READ}, 'data': data}
        output = json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False)
        if a.output:
            a.output.parent.mkdir(parents=True, exist_ok=True)
            # Atomic publication: consumers must still check exit status and target.
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=a.output.parent,
                                             prefix=a.output.name+'.', suffix='.tmp', delete=False) as handle:
                handle.write(output + chr(10))
                temp_path = handle.name
            try:
                os.replace(temp_path, a.output)
            finally:
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
        print(output)
        return 0
    except Exception as exc:
        print(json.dumps({'schema_version': 2, 'status': 'error', 'operation': a.command,
                          'error': str(exc), 'type': type(exc).__name__}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
