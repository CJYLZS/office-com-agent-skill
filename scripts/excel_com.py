"""Read-only Excel COM adapter. Attach to open desktop windows; never Open/Save/Quit."""
from __future__ import annotations
import argparse
import ctypes
import hashlib
import json
import os
import tempfile
from pathlib import Path
import re
import subprocess
import sys
import time
from word_com import WordBridge

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


def normalized_hwnd(value):
    """Excel exposes HWND as signed Long; normalize its 32-bit representation."""
    return int(value) & 0xffffffff


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


class ExcelBridge(WordBridge):
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
    text_chars = 0
    for r, c, height, width, tile_address in range_tiles(bounds['address']):
        tile = sheet.Range(tile_address)
        for key, prop in properties:
            try:
                values = matrix(getattr(tile, prop), height, width)
            except NotImplementedError as exc:
                raise ValueError('COM variant cannot be decoded (Excel error cells/VT_ERROR unsupported); narrow scope or use formulas mode') from exc
            text_chars += sum(len(item) for row in values for item in row if isinstance(item, str))
            if text_chars > MAX_TEXT_CHARS:
                raise ValueError('Text budget exceeded after bounded COM transfer; narrow the rectangle')
            for i, row in enumerate(values):
                data[key][r+i][c:c+width] = row
    # Sequential reads are not atomic. This digest is evidence, not a revision token/edit permit.
    payload = json.dumps(data, ensure_ascii=False, allow_nan=False, sort_keys=True, separators=(',', ':'))
    return {'sheet': {'name': str(sheet.Name), 'code_name': str(sheet.CodeName)},
            **bounds, 'mode': mode, 'formula_api': formula_api if mode != 'values' else None,
            'data': data, 'data_sha256': hashlib.sha256(payload.encode('utf-8', errors='surrogatepass')).hexdigest(),
            'truncated': False,
            'warnings': ['Sequential live reads, not an atomic snapshot or edit authorization.',
                         'Value2 dates/currency are numeric; undecodable Excel error cells fail the entire read.',
                         'Formula output also includes constants; not a per-cell formula/type mask.',
                         'No calculation/refresh or formatted display text; no structural edit safety check.']}


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('command', choices=['list', 'probe', 'read'])
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
        if a.command == 'read':
            parse_address(a.address)
            if not a.sheet:
                raise ValueError('--sheet required')
        bridge = ExcelBridge()
        target = None
        if a.command == 'list':
            data = bridge.inventory()
        else:
            wb, target = bridge.select(a.pid, a.hwnd, a.document)
            before_saved = bool(wb.Saved)
            data = probe(wb, a.offset, a.limit) if a.command == 'probe' else read_range(wb, a.sheet, a.address, a.mode, a.formula_api)
            # Rebind for end-of-read identity check, not a concurrency lock.
            _, current = bridge.select(a.pid, a.hwnd, target['full_name'])
            if current != target:
                raise ValueError('Workbook identity changed during read')
            data['saved_before'] = before_saved
            data['saved_after'] = bool(wb.Saved)
        report = {'schema_version': 1, 'status': 'ok', 'operation': a.command, 'target': target,
                  'elapsed_seconds': round(time.perf_counter()-t, 3),
                  'limits': {'max_cells': MAX_CELLS, 'max_text_chars': MAX_TEXT_CHARS, 'supervised_cli_timeout_seconds': 9, 'max_cells_per_com_read': MAX_CELLS_PER_COM_READ}, 'data': data}
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
        print(json.dumps({'schema_version': 1, 'status': 'error', 'operation': a.command,
                          'error': str(exc), 'type': type(exc).__name__}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
