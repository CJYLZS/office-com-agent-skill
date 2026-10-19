"""Attach-only, read-only PowerPoint COM adapter; no file opening or application creation."""
from __future__ import annotations
import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
from word_com import WordBridge

MAX_SLIDES = 50
MAX_SHAPES = 500
MAX_CELLS = 500
MAX_TEXT_CHARS = 100000
MAX_TEXT_PER_READ = 32000
MAX_BIND_WINDOWS = 100


def normalized_hwnd(value):
    return int(value) & 0xffffffff


def page(offset, limit, maximum):
    if type(offset) is not int or offset < 1 or type(limit) is not int or not 1 <= limit <= maximum:
        raise ValueError(f'offset >= 1 and limit 1..{maximum} required')


def window_presentation(window):
    """NativeOM's documented paneClassDC return is DocumentWindow, not Application."""
    return window.Presentation


def matching_presentations(app, hwnd, expected_pid, root_hwnd, window_pid, early_bind):
    presentations = app.Presentations
    if int(presentations.Count) > MAX_BIND_WINDOWS:
        raise ValueError('Too many presentations for bounded ROT binding')
    matches = []
    windows_seen = 0
    for i in range(1, int(presentations.Count) + 1):
        presentation = presentations.Item(i)
        windows = presentation.Windows
        for j in range(1, int(windows.Count) + 1):
            windows_seen += 1
            if windows_seen > MAX_BIND_WINDOWS:
                raise ValueError('Too many document windows for bounded ROT binding')
            window = early_bind(windows.Item(j))
            child_hwnd = normalized_hwnd(window.HWND)
            if root_hwnd(child_hwnd) == hwnd and window_pid(child_hwnd) == expected_pid:
                matches.append((presentation, child_hwnd))
    if len(matches) != 1:
        raise ValueError('ROT document window binding missing or ambiguous')
    return matches[0]


class PowerPointBridge(WordBridge):
    # Only reuse Win32 signatures, label(), pid(); not Word attachment/selection.
    def windows(self):
        result = []
        @self.callback
        def collect(hwnd, _):
            if self.label(hwnd, True) == 'PPTFrameClass':
                result.append({'pid': self.pid(hwnd), 'hwnd': int(hwnd), 'caption': self.label(hwnd)})
            return True
        self.user32.EnumWindows(collect, 0)
        if len(result) > MAX_BIND_WINDOWS:
            raise ValueError('Too many PowerPoint windows')
        return result

    def attach(self, hwnd, expected_pid=None):
        if self.label(hwnd, True) != 'PPTFrameClass':
            raise ValueError('Target is not a PowerPoint document frame')
        pid = self.pid(hwnd)
        if not pid or (expected_pid is not None and pid != expected_pid):
            raise ValueError('PID mismatch')
        children = []
        @self.callback
        def collect(child, _):
            if self.label(child, True) == 'paneClassDC' and len(children) <= MAX_BIND_WINDOWS:
                children.append(int(child))
            return True
        self.user32.EnumChildWindows(hwnd, collect, 0)
        if len(children) > MAX_BIND_WINDOWS:
            raise ValueError('Too many NativeOM panes')
        errors = []
        native_matches = []
        for child in children:
            try:
                if self.pid(child) != pid:
                    raise ValueError('Native pane PID mismatch')
                ptr = ctypes.POINTER(self.IDispatch)()
                iid = self.IDispatch._iid_
                hr = self.oleacc.AccessibleObjectFromWindow(child, 0xfffffff0, ctypes.byref(iid), ctypes.byref(ptr))
                if hr != 0 or not ptr:
                    raise RuntimeError(f'NativeOM unavailable: HRESULT={hr}')
                window = self.Dispatch(ptr)
                # Verify the exact pane ancestry with Win32, not PowerPoint's broken
                # hidden DocumentWindow.HWND automation property (DISP_E_MEMBERNOTFOUND).
                from ctypes import wintypes
                self.user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
                self.user32.GetAncestor.restype = wintypes.HWND
                if int(self.user32.GetAncestor(child, 2) or 0) != hwnd:
                    raise ValueError('Native pane root HWND mismatch')
                presentation = window_presentation(window)
                if str(window.Application.Name) != 'Microsoft PowerPoint':
                    raise ValueError('Native application mismatch')
                target = {'application': 'powerpoint', 'pid': pid, 'hwnd': hwnd,
                          'name': str(presentation.Name), 'full_name': str(presentation.FullName),
                          'binding': 'NativeOM:paneClassDC/root-HWND/PID/DocumentWindow.Presentation'}
                if self.pid(hwnd) != pid or self.label(hwnd, True) != 'PPTFrameClass':
                    raise ValueError('Window changed during attachment')
                native_matches.append((presentation, target))
            except Exception as exc:
                errors.append(str(exc))
        if native_matches:
            # Multiple panes may expose the same DocumentWindow. Compare COM identity,
            # not captions, and reject disagreement before returning any candidate.
            keys = set()
            for presentation, _ in native_matches:
                raw = getattr(presentation, '_comobj', presentation)
                unknown = raw.QueryInterface(self.comtypes.IUnknown)
                keys.add(ctypes.cast(unknown, ctypes.c_void_p).value)
            if len(keys) != 1:
                raise ValueError('NativeOM presentation binding is ambiguous')
            return native_matches[0]
        # ROT obtains only a running application. GetActiveObject never creates one.
        # The hidden HWND member fails via IDispatch but works via generated vtable.
        try:
            from comtypes.client import GetActiveObject, GetBestInterface
            from ctypes import wintypes
            self.user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
            self.user32.GetAncestor.restype = wintypes.HWND
            app = GetActiveObject('PowerPoint.Application')
            if str(app.Name) != 'Microsoft PowerPoint':
                raise ValueError('ROT application mismatch')
            presentation, document_hwnd = matching_presentations(
                app, hwnd, pid, lambda h: int(self.user32.GetAncestor(h, 2) or 0), self.pid,
                lambda raw: GetBestInterface(raw._comobj) if hasattr(raw, '_comobj') else raw)
            if self.pid(hwnd) != pid or self.label(hwnd, True) != 'PPTFrameClass':
                raise ValueError('Window changed during ROT attachment')
            return presentation, {'application': 'powerpoint', 'pid': pid, 'hwnd': hwnd,
                                  'document_hwnd': document_hwnd,
                                  'name': str(presentation.Name), 'full_name': str(presentation.FullName),
                                  'binding': 'ROT:Presentation.Windows/vtable-HWND/root-HWND/PID'}
        except Exception as exc:
            errors.append(str(exc))
        raise RuntimeError('Cannot prove PowerPoint Presentation binding; no fallback to active document. ' + '; '.join(errors))

    def select(self, pid=None, hwnd=None, document=None):
        if type(pid) is not int or type(hwnd) is not int or pid <= 0 or hwnd <= 0:
            raise ValueError('Explicit positive PID and HWND required; run list first')
        presentation, target = self.attach(hwnd, pid)
        if document is not None and target['full_name'] != document:
            raise ValueError('Presentation full name mismatch')
        return presentation, target

    def inventory(self):
        items = []
        for window in self.windows():
            try:
                presentation, target = self.attach(window['hwnd'], window['pid'])
                items.append({'target': target, 'caption': window['caption'], 'saved': bool(presentation.Saved)})
            except Exception as exc:
                items.append({**window, 'error': str(exc)})
        return {'windows': items, 'scope': 'desktop document frames only; slideshow/protected/hidden windows unsupported'}


class Budget:
    def __init__(self, text=MAX_TEXT_CHARS):
        if type(text) is not int or not 1 <= text <= MAX_TEXT_CHARS:
            raise ValueError(f'text budget must be 1..{MAX_TEXT_CHARS}')
        self.remaining = text
        self.shapes = 0
        self.cells = 0

    def text(self, shape):
        if not bool(shape.HasTextFrame):
            return None
        text_range = shape.TextFrame.TextRange
        length = int(text_range.Length)
        if length < 0 or length > min(self.remaining, MAX_TEXT_PER_READ):
            raise ValueError('Text exceeds remaining/per-COM-read budget; no partial result published')
        # Never retrieve the entire Text property before checking native Length.
        text = str(text_range.Characters(1, length).Text) if length else ''
        text.encode('utf-8')  # Reject invalid surrogate output, not silently corrupt evidence.
        utf16_length = len(text.encode('utf-16-le')) // 2
        if utf16_length > length or utf16_length > self.remaining:
            raise ValueError('Text changed or exceeded UTF-16 budget during read')
        self.remaining -= length  # Conservative native UTF-16 length budget.
        return {'text': text, 'text_sha256': hashlib.sha256(text.encode('utf-8')).hexdigest(),
                'native_length': length}


def slide_evidence(slide, index):
    return {'index': index, 'slide_id': int(slide.SlideID), 'name': str(slide.Name),
            'shape_count': int(slide.Shapes.Count)}


def probe(presentation, offset=1, limit=20):
    page(offset, limit, MAX_SLIDES)
    count = int(presentation.Slides.Count)
    slides = [slide_evidence(presentation.Slides.Item(i), i)
              for i in range(offset, min(count + 1, offset + limit))]
    return {'slide_count': count, 'slides': slides, 'offset': offset,
            'next_offset': offset + len(slides) if offset + len(slides) <= count else None}


def read_shapes(shapes, offset, limit, budget, tables=False):
    count = int(shapes.Count)
    records = []
    for i in range(offset, min(count + 1, offset + limit)):
        budget.shapes += 1
        if budget.shapes > MAX_SHAPES:
            raise ValueError('Shape structure budget exceeded')
        shape = shapes.Item(i)
        record = {'index': i, 'shape_id': int(shape.Id), 'name': str(shape.Name), 'type': int(shape.Type)}
        text = budget.text(shape)
        if text is not None:
            record.update(text)
        if bool(shape.HasTable):
            record['has_table'] = True
            if tables:
                table = shape.Table
                rows, columns = int(table.Rows.Count), int(table.Columns.Count)
                if rows < 0 or columns < 0 or budget.cells + rows * columns > MAX_CELLS:
                    raise ValueError('Table cell structure budget exceeded')
                budget.cells += rows * columns
                record['table'] = {'rows': rows, 'columns': columns, 'cells': [
                    {'row': r, 'column': c, 'table_shape_id': record['shape_id'],
                     **(budget.text(table.Cell(r, c).Shape) or {})}
                    for r in range(1, rows + 1) for c in range(1, columns + 1)]}
        if int(shape.Type) in (3, 6, 7, 10, 24):
            record['omitted'] = 'child content: chart/group/OLE/SmartArt not expanded'
        # Groups/chart/OLE/SmartArt are recorded, never activated or recursively expanded.
        records.append(record)
    return {'shape_count': count, 'shapes': records, 'offset': offset,
            'next_offset': offset + len(records) if offset + len(records) <= count else None}


def read_slides(presentation, offset=1, limit=5, shape_offset=1, shape_limit=50,
                tables=False, notes=False, text_budget=MAX_TEXT_CHARS):
    page(offset, limit, MAX_SLIDES)
    page(shape_offset, shape_limit, MAX_SHAPES)
    budget = Budget(text_budget)
    count = int(presentation.Slides.Count)
    slides = []
    for i in range(offset, min(count + 1, offset + limit)):
        slide = presentation.Slides.Item(i)
        record = slide_evidence(slide, i)
        record.update(read_shapes(slide.Shapes, shape_offset, shape_limit, budget, tables))
        if notes:
            record['notes'] = read_shapes(slide.NotesPage.Shapes, shape_offset, shape_limit, budget, tables)
        if int(slide.SlideID) != record['slide_id'] or int(slide.Shapes.Count) != record['shape_count']:
            raise ValueError('Slide structure changed during read')
        slides.append(record)
    if int(presentation.Slides.Count) != count:
        raise ValueError('Presentation slide count changed during read')
    return {'slide_count': count, 'slides': slides, 'offset': offset,
            'next_offset': offset + len(slides) if offset + len(slides) <= count else None,
            'text_chars': text_budget - budget.remaining, 'shapes_read': budget.shapes, 'cells_read': budget.cells,
            'tables_included': tables, 'notes_included': notes,
            'consistency': 'best effort live read, not an atomic snapshot'}


def publish(report, output=None):
    text = json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False)
    if output:
        output = Path(output)
        output.parent.mkdir(parents=True, exist_ok=True)
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=output.parent,
                                             prefix=output.name + '.', suffix='.tmp', delete=False) as handle:
                temp_path = handle.name
                handle.write(text + chr(10))
            os.replace(temp_path, output)
        finally:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
    print(text)


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['list', 'probe', 'read'])
    parser.add_argument('--pid', type=int)
    parser.add_argument('--hwnd', type=int)
    parser.add_argument('--document')
    parser.add_argument('--offset', type=int, default=1)
    parser.add_argument('--limit', type=int, default=5)
    parser.add_argument('--shape-offset', type=int, default=1)
    parser.add_argument('--shape-limit', type=int, default=50)
    parser.add_argument('--text-budget', type=int, default=MAX_TEXT_CHARS)
    parser.add_argument('--tables', action='store_true')
    parser.add_argument('--notes', action='store_true')
    parser.add_argument('--output', type=Path)
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    if not args.worker:
        try:
            return subprocess.run([sys.executable, str(Path(__file__).resolve()), *sys.argv[1:], '--worker'], timeout=9).returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({'schema_version': 1, 'status': 'error', 'operation': args.command,
                              'error': 'Read-only worker exceeded 9 seconds; no presentation writes',
                              'type': 'TimeoutExpired'}), file=sys.stderr)
            return 2
    start = time.perf_counter()
    try:
        page(args.offset, args.limit, MAX_SLIDES)
        page(args.shape_offset, args.shape_limit, MAX_SHAPES)
        Budget(args.text_budget)
        bridge = PowerPointBridge()
        target = None
        if args.command == 'list':
            data = bridge.inventory()
        else:
            presentation, target = bridge.select(args.pid, args.hwnd, args.document)
            saved = bool(presentation.Saved)
            data = (probe(presentation, args.offset, args.limit) if args.command == 'probe' else
                    read_slides(presentation, args.offset, args.limit, args.shape_offset, args.shape_limit,
                                args.tables, args.notes, args.text_budget))
            _, current = bridge.select(args.pid, args.hwnd, target['full_name'])
            if current != target:
                raise ValueError('Presentation identity changed during read')
            data.update(saved_before=saved, saved_after=bool(presentation.Saved))
        publish({'schema_version': 1, 'status': 'ok', 'operation': args.command, 'target': target,
                 'elapsed_seconds': round(time.perf_counter() - start, 3),
                 'limits': {'max_slides': MAX_SLIDES, 'max_shapes': MAX_SHAPES, 'max_cells': MAX_CELLS,
                            'max_text_chars': MAX_TEXT_CHARS, 'max_text_per_com_read': MAX_TEXT_PER_READ,
                            'supervised_cli_timeout_seconds': 9}, 'data': data}, args.output)
        return 0
    except Exception as exc:
        print(json.dumps({'schema_version': 1, 'status': 'error', 'operation': args.command,
                          'error': str(exc), 'type': type(exc).__name__}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
