"""Attach-only Microsoft Word bridge. Never opens, closes or quits a user document."""
from __future__ import annotations
import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
from pathlib import Path
import sys


def text_hash(text):
    return hashlib.sha256(text.encode('utf-8', errors='surrogatepass')).hexdigest()


def validate_request(req, identity, body_text):
    if req.get('target') != identity:
        raise ValueError('目标身份已改变；重新读取文档，不得自动重试写入')
    if req.get('text_sha256') != text_hash(body_text):
        raise ValueError('正文已变化；重新读取并重新规划，未执行修改')
    op = req.get('operation', {})
    if op.get('kind') not in ('replace', 'insert', 'comment', 'format'):
        raise ValueError('operation.kind 必须为 replace/insert/comment/format')
    start, end = op.get('start'), op.get('end')
    if type(start) is not int or type(end) is not int or not 0 <= start <= end:
        raise ValueError('需要有效的 Word 原生 start/end')
    if not isinstance(op.get('expected'), str):
        raise ValueError('必须给出范围原文 expected（插入为空字符串）')
    if op['kind'] == 'insert' and (start != end or op['expected'] != ''):
        raise ValueError('insert 必须为零长度范围')
    if op['kind'] == 'replace' and start == end:
        raise ValueError('零长度替换请使用 insert')
    if op['kind'] in ('replace', 'insert', 'comment'):
        if not isinstance(op.get('text'), str):
            raise ValueError('需要 text 字符串')
        if any(c in op['text'] for c in ('\x00', '\x07', '\x13', '\x14', '\x15')):
            raise ValueError('text 不允许含表格/域控制字符')
    if op['kind'] == 'replace' and any(c in op['expected'] for c in ('\r', '\n', '\x07', '\x13', '\x14', '\x15')):
        raise ValueError('拒绝跨段落、表格边界或域的替换；请缩小范围')
    if op['kind'] == 'format':
        style = op.get('font', {})
        if not style or set(style) - {'bold', 'italic', 'name', 'size', 'color'}:
            raise ValueError('font 只支持 bold/italic/name/size/color')
        for key in ('bold', 'italic'):
            if key in style and type(style[key]) is not bool:
                raise ValueError(key + ' 必须是 bool')
        if 'name' in style and (not isinstance(style['name'], str) or not style['name']):
            raise ValueError('name 必须是非空字符串')
        if 'size' in style and (type(style['size']) not in (int, float) or not 1 <= style['size'] <= 1638):
            raise ValueError('size 超出范围')
        if 'color' in style and (type(style['color']) is not int or not 0 <= style['color'] <= 0xffffff):
            raise ValueError('color 应为 0..0xffffff 的 Word BGR 整数')
    return op


class WordBridge:
    def __init__(self):
        if os.name != 'nt':
            raise RuntimeError('仅支持 Windows 桌面交互会话')
        import comtypes
        from comtypes.automation import IDispatch
        from comtypes.client.dynamic import Dispatch
        self.comtypes, self.IDispatch, self.Dispatch = comtypes, IDispatch, Dispatch
        self.user32 = ctypes.WinDLL('user32', use_last_error=True)
        self.oleacc = ctypes.OleDLL('oleacc')
        self.callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        u = self.user32
        u.EnumWindows.argtypes = [self.callback, wintypes.LPARAM]
        u.EnumChildWindows.argtypes = [wintypes.HWND, self.callback, wintypes.LPARAM]
        u.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        u.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        u.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
        self.oleacc.AccessibleObjectFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.POINTER(comtypes.GUID), ctypes.POINTER(ctypes.POINTER(IDispatch))]
        self.oleacc.AccessibleObjectFromWindow.restype = ctypes.HRESULT

    def label(self, hwnd, class_name=False):
        buf = ctypes.create_unicode_buffer(2048)
        (self.user32.GetClassNameW if class_name else self.user32.GetWindowTextW)(hwnd, buf, len(buf))
        return buf.value

    def pid(self, hwnd):
        pid = wintypes.DWORD()
        self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value

    def windows(self):
        result = []
        @self.callback
        def collect(hwnd, _):
            if self.label(hwnd, True) == 'OpusApp':
                result.append({'pid': self.pid(hwnd), 'hwnd': int(hwnd), 'caption': self.label(hwnd)})
            return True
        self.user32.EnumWindows(collect, 0)
        return result

    def attach(self, hwnd, expected_pid=None):
        if self.label(hwnd, True) != 'OpusApp':
            raise ValueError('目标已不是 Word 顶层窗口')
        pid = self.pid(hwnd)
        if expected_pid is not None and pid != expected_pid:
            raise ValueError('PID 不匹配')
        children = []
        @self.callback
        def collect(child, _):
            if self.label(child, True) == '_WwG':
                children.append(child)
            return True
        self.user32.EnumChildWindows(hwnd, collect, 0)
        failures = []
        for child in children:
            try:
                ptr = ctypes.POINTER(self.IDispatch)()
                iid = self.IDispatch._iid_
                self.oleacc.AccessibleObjectFromWindow(child, 0xfffffff0, ctypes.byref(iid), ctypes.byref(ptr))
                window = self.Dispatch(ptr)
                doc = window.Document
                # Do not use Application.ActiveDocument: focus may belong to another document.
                identity = {'pid': pid, 'hwnd': int(hwnd), 'name': str(doc.Name), 'full_name': str(doc.FullName)}
                return doc, identity
            except Exception as exc:
                failures.append(str(exc))
        raise RuntimeError('无法获取 Word 原生对象；退出受保护视图/关闭模态对话框并确认同一权限级别。' + '; '.join(failures))

    def select(self, pid=None, hwnd=None, document=None):
        candidates, errors = [], []
        for item in self.windows():
            if pid is not None and item['pid'] != pid or hwnd is not None and item['hwnd'] != hwnd:
                continue
            try:
                doc, identity = self.attach(item['hwnd'], item['pid'])
                if document and os.path.normcase(os.path.abspath(identity['full_name'])) != os.path.normcase(os.path.abspath(document)):
                    continue
                candidates.append((doc, identity))
            except Exception as exc:
                errors.append({'window': item, 'error': str(exc)})
        if len(candidates) != 1:
            raise RuntimeError(json.dumps({'message': '必须唯一指定文档；使用 list 后按 pid + hwnd 选择', 'matches': [i for _, i in candidates], 'errors': errors}, ensure_ascii=False))
        return candidates[0]

    def inventory(self):
        items = []
        for window in self.windows():
            try:
                doc, identity = self.attach(window['hwnd'], window['pid'])
                items.append({**identity, 'caption': window['caption'], 'saved': bool(doc.Saved), 'read_only': bool(doc.ReadOnly)})
            except Exception as exc:
                items.append({**window, 'error': str(exc)})
        return {'windows': items, 'note': '仅枚举有桌面窗口的 Word 文档；受保护视图/隐藏实例不保证可连接'}


def snapshot(doc, identity, offset=1, limit=50):
    content = doc.Content
    text = str(content.Text)
    count = int(doc.Paragraphs.Count)
    paragraphs = []
    for index in range(offset, min(count + 1, offset + limit)):
        rng = doc.Paragraphs.Item(index).Range
        paragraphs.append({'index': index, 'start': int(rng.Start), 'end': int(rng.End), 'text': str(rng.Text or '')})
    return {'target': identity, 'scope': 'main_story', 'text_sha256': text_hash(text), 'body_end': int(content.End),
            'saved': bool(doc.Saved), 'read_only': bool(doc.ReadOnly), 'track_revisions': bool(doc.TrackRevisions),
            'paragraph_count': count, 'table_count': int(doc.Tables.Count), 'paragraphs': paragraphs,
            'next_offset': offset + limit if offset + limit <= count else None,
            'concurrency': 'text-only optimistic check, NOT an atomic lock; pause human typing during commit'}


def find_text(doc, text, limit=50):
    if not text or len(text) > 250 or any(c in text for c in ('^', '\r', '\n')):
        raise ValueError('find 仅支持 1..250 字符单行字面文本，不允许 ^ 转义')
    hits, start, stop = [], int(doc.Content.Start), int(doc.Content.End)
    while start < stop and len(hits) < limit:
        rng = doc.Range(start, stop)
        finder = rng.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        finder.MatchPrefix = False
        finder.MatchSuffix = False
        finder.MatchPhrase = False
        finder.IgnoreSpace = False
        finder.IgnorePunct = False
        # Text, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        # MatchAllWordForms, Forward, Wrap=wdFindStop, Format
        found = finder.Execute(text, True, False, False, False, False, True, 0, False)
        if not found:
            break
        if str(rng.Text or '') == text:
            hits.append({'start': int(rng.Start), 'end': int(rng.End), 'text': str(rng.Text or '')})
        start = max(int(rng.End), start + 1)
    return {'matches': hits, 'limit_reached': len(hits) == limit}


def apply_request(doc, identity, req, commit=False):
    op = validate_request(req, identity, str(doc.Content.Text))
    if bool(doc.ReadOnly) or int(doc.ProtectionType) != -1:
        raise ValueError('只读或受保护的文档，拒绝编辑')
    if not 0 <= op['start'] <= op['end'] <= int(doc.Content.End) - 1:
        raise ValueError('范围超出正文或触及最终段落标记')
    rng = doc.Range(op['start'], op['end'])
    if str(rng.Text or '') != op['expected']:
        raise ValueError('范围原文 expected 不匹配')
    if op['kind'] in ('replace', 'insert', 'format'):
        # Conservative: complex objects require a task-specific script, not generic text edits.
        if int(doc.ContentControls.Count) or int(doc.Fields.Count):
            raise ValueError('正文含域或内容控件，通用修改拒绝执行；需专用结构感知脚本')
        if bool(rng.Information(12)) or int(rng.Tables.Count):  # wdWithInTable
            raise ValueError('表格内修改需要专用单元格接口')
        if op['start'] != op['end'] and int(rng.InlineShapes.Count):
            raise ValueError('范围含嵌入对象')
    result = {'target': identity, 'operation': op, 'committed': False, 'saved_by_script': False}
    if not commit:
        return {**result, 'message': '预览完成；--commit 才会修改内存文档'}
    app = doc.Application
    undo = app.UndoRecord
    if bool(undo.IsRecordingCustomRecord):
        raise ValueError('Word 已在自定义撤销记录中，拒绝嵌套')
    previous_tracking = bool(doc.TrackRevisions)
    started, modified = False, False
    try:
        undo.StartCustomRecord('AI Word collaboration')
        started = True
        if op['kind'] in ('replace', 'insert'):
            doc.TrackRevisions = True
        # Last-moment guard. COM calls still do not provide an atomic transaction.
        validate_request(req, identity, str(doc.Content.Text))
        if str(rng.Text or '') != op['expected']:
            raise ValueError('提交前范围已改变')
        if op['kind'] in ('replace', 'insert'):
            rng.Text = op['text'].replace('\r\n', '\r').replace('\n', '\r')
            modified = True
        elif op['kind'] == 'comment':
            doc.Comments.Add(rng, op['text'])
            modified = True
        else:
            for key, value in op['font'].items():
                setattr(rng.Font, {'bold': 'Bold', 'italic': 'Italic', 'name': 'Name', 'size': 'Size', 'color': 'Color'}[key], value)
                modified = True
        result.update(committed=True, text_sha256_after=text_hash(str(doc.Content.Text)),
                      message='已更新 Word 内存文档；脚本未调用保存，AutoSave 可能自行保存；可在 Word 中撤销')
        return result
    except Exception as exc:
        raise RuntimeError(f'提交失败，可能有部分修改（已确认修改={modified}）；不自动重试或撤销，请在 Word 检查：{exc}') from exc
    finally:
        try:
            if bool(doc.TrackRevisions) != previous_tracking:
                doc.TrackRevisions = previous_tracking
        finally:
            if started:
                undo.EndCustomRecord()


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['list', 'read', 'find', 'apply', 'save'])
    parser.add_argument('--pid', type=int)
    parser.add_argument('--hwnd', type=int)
    parser.add_argument('--document', help='已打开文档完整路径；不会打开文件')
    parser.add_argument('--offset', type=int, default=1)
    parser.add_argument('--limit', type=int, default=50)
    parser.add_argument('--text')
    parser.add_argument('--request', type=Path)
    parser.add_argument('--commit', action='store_true')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    try:
        if args.offset < 1 or not 1 <= args.limit <= 500:
            raise ValueError('offset >= 1，limit 必须为 1..500')
        bridge = WordBridge()
        if args.command == 'list':
            result = bridge.inventory()
        elif args.command == 'apply':
            if args.request is None:
                raise ValueError('需要 --request JSON 文件')
            req = json.loads(args.request.read_text(encoding='utf-8-sig'))
            target = req['target']
            doc, identity = bridge.select(target['pid'], target['hwnd'], target['full_name'])
            result = apply_request(doc, identity, req, args.commit)
        else:
            doc, identity = bridge.select(args.pid, args.hwnd, args.document)
            if args.command == 'read':
                result = snapshot(doc, identity, args.offset, args.limit)
            elif args.command == 'find':
                result = {**snapshot(doc, identity, 1, 1), **find_text(doc, args.text, args.limit)}
            else:
                if not args.commit or args.pid is None or args.hwnd is None:
                    raise ValueError('保存必须显式指定 --pid --hwnd --commit，且事先获得用户授权')
                if not str(doc.Path) or bool(doc.ReadOnly):
                    raise ValueError('未命名/只读文档不支持 save；请在 Word 另存为')
                doc.Save()
                result = {'target': identity, 'saved': bool(doc.Saved)}
        output = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(output + '\n', encoding='utf-8')
        print(output)
        return 0
    except Exception as exc:
        print(json.dumps({'error': str(exc), 'type': type(exc).__name__}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
