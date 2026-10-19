"""Composable helpers for temporary scripts; CLI is read-only and time-bounded.
No implicit save, selection changes, field updates, retry, or rollback.
"""
from __future__ import annotations
import argparse
from contextlib import contextmanager
import json
from pathlib import Path
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from collections import Counter
from word_com import WordBridge, text_hash

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': W[1:-1]}


def local_range(doc, start, end, max_units=20000):
    """Validate Word main-story coordinates; never derive these from str indices."""
    if type(start) is not int or type(end) is not int or not 0 <= start <= end:
        raise ValueError('Expected ordered integer Word coordinates')
    if end - start > max_units or end > int(doc.Content.End):
        raise ValueError('Range exceeds document or local read budget')
    return doc.Range(start, end)


def read_range(doc, start, end, max_units=20000):
    rng = local_range(doc, start, end, max_units)
    text = str(rng.Text or '')
    return {'start': int(rng.Start), 'end': int(rng.End), 'text': text,
            'text_sha256': text_hash(text), 'scope': 'main_story'}


def xml_summary(xml):
    """Local structural hints, NOT proof of safe editing or Word coordinates."""
    if len(xml.encode('utf-8')) > 8_000_000:
        raise ValueError('XML exceeds parser byte budget')
    if '<!DOCTYPE' in xml.upper() or '<!ENTITY' in xml.upper():
        raise ValueError('DTD/entities are not supported')
    root = ET.fromstring(xml)
    stack, nodes = [(root, 0)], 0
    while stack:
        element, depth = stack.pop()
        nodes += 1
        if nodes > 100000 or depth > 128:
            raise ValueError('XML exceeds structural budget')
        stack.extend((child, depth + 1) for child in element)
    body = root.find('.//w:document/w:body', NS)
    if body is None:
        body = root if root.tag == W + 'body' else None
    if body is None:
        raise ValueError('No Word document body')
    tags = Counter(e.tag[len(W):] for e in body.iter() if e.tag.startswith(W))
    field_types = Counter()
    for e in body.iter(W + 'instrText'):
        parts = (e.text or '').strip().split()
        if parts:
            field_types[parts[0]] += 1
    return {'elements': {key: tags[key] for key in
            ('p', 'tbl', 'tr', 'tc', 'gridSpan', 'vMerge', 'fldChar', 'fldSimple',
             'sdt', 'drawing', 'object', 'bookmarkStart', 'sectPr', 'ins', 'del')},
            'field_instruction_tokens': dict(field_types),
            'warning': 'Clipped XML can omit enclosing fields/controls; not a safety certificate'}


def inspect_range(doc, start, end, max_units=20000, max_xml_bytes=8_000_000):
    rng = local_range(doc, start, end, max_units)
    result = read_range(doc, start, end, max_units)
    xml = str(rng.WordOpenXML)
    size = len(xml.encode('utf-8'))
    if size > max_xml_bytes:
        raise ValueError('XML payload exceeds local budget (already transferred); narrow range')
    result.update(xml_bytes=size, structure=xml_summary(xml), risks=range_risks(doc, start, end))
    return result


def find_literal(doc, text, start=0, end=None, style=None, limit=5,
                 context_units=80, budget_seconds=3):
    """Find in bounded scope; optional built-in style ID/name. No Selection use.
    Budget is cooperative between calls; external worker bounds blocked COM.
    """
    if not text or len(text) > 250 or any(c in text for c in ('^', '\r', '\n')):
        raise ValueError('Use 1..250 single-line literal characters without ^')
    if not 1 <= limit <= 50 or not 0 <= context_units <= 1000 or budget_seconds <= 0:
        raise ValueError('Invalid search budget')
    stop = int(doc.Content.End) if end is None else end
    if type(start) is not int or type(stop) is not int or not 0 <= start <= stop <= int(doc.Content.End):
        raise ValueError('Invalid main-story search bounds')
    deadline = time.perf_counter() + budget_seconds
    hits, cursor, reason = [], start, 'exhausted'
    while cursor < stop:
        if time.perf_counter() >= deadline:
            reason = 'time_budget'
            break
        rng = doc.Range(cursor, stop)
        finder = rng.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        for name in ('MatchPrefix', 'MatchSuffix', 'MatchPhrase', 'IgnoreSpace', 'IgnorePunct'):
            setattr(finder, name, False)
        if style is not None:
            finder.Style = doc.Styles.Item(style)
        if not finder.Execute(text, True, False, False, False, False, True, 0, style is not None):
            break
        a, b = int(rng.Start), int(rng.End)
        if a < cursor or b > stop or b <= a:
            reason = 'word_returned_out_of_scope'
            break  # Fields can make Word Find revisit an enclosing result; never fabricate progress
        if str(rng.Text or '') == text:
            hits.append({'start': a, 'end': b, 'text': text,
                         'context': read_range(doc, max(start, a-context_units),
                                               min(stop, b+context_units), 3000)})
        cursor = max(b, cursor + 1)
        if len(hits) >= limit:
            reason = 'limit'
            break
    return {'matches': hits, 'stop_reason': reason, 'next_start': cursor,
            'search_end': stop, 'complete': reason == 'exhausted'}


def range_risks(doc, start, end):
    """Local diagnostics, not authorization. Probe both ends including insertions.
    Fields.Count alone misses ranges wholly inside a field result.
    """
    rng = local_range(doc, start, end)
    points = [doc.Range(start, start), doc.Range(end, end)]
    return {'fields_in_range': int(rng.Fields.Count),
            'controls_in_range': int(rng.ContentControls.Count),
            'inline_shapes': int(rng.InlineShapes.Count),
            'in_table': bool(rng.Information(12)),
            'endpoint_in_field_code': [bool(p.Information(44)) for p in points],
            'endpoint_in_field_result': [bool(p.Information(45)) for p in points],
            'warning': 'Also inspect enclosing paragraph/table, controls, bookmarks and revisions; not a safe-to-edit flag'}


def cell_text_range(cell):
    """Return duplicate Range excluding cell marker AND final paragraph mark.
    Do not assume row/column indexing works for merged tables.
    """
    rng = cell.Range.Duplicate
    if not str(rng.Text or '').endswith('\r\x07'):
        raise ValueError('Unexpected cell terminator; inspect table before editing')
    expected = str(rng.Text or '')[:-2]
    rng.End = int(rng.End) - 1  # Word represents the two-character terminator in one native unit
    if str(rng.Text or '') != expected:
        raise ValueError('Cell range does not map as expected; refuse to guess coordinates')
    return rng


def assert_expected(rng, expected):
    if str(rng.Text or '') != expected:
        raise ValueError('Range changed; re-read and re-plan, do not auto-retry')


@contextmanager
def tracked_edit(doc, label='Word task'):
    """For explicitly authorized writes only. Grouping is NOT a transaction.
    Caller must validate target/anchors/risks immediately before each write.
    Exceptions may leave partial writes. Never auto-retry or Undo.
    """
    if bool(doc.ReadOnly) or int(doc.ProtectionType) != -1:
        raise ValueError('Document is read-only or protected')
    undo = doc.Application.UndoRecord
    if bool(undo.IsRecordingCustomRecord):
        raise ValueError('Existing UndoRecord; refusing nested edit')
    previous = bool(doc.TrackRevisions)
    undo.StartCustomRecord(label)
    primary = None
    try:
        doc.TrackRevisions = True
        yield
    except BaseException as exc:
        primary = exc
        raise
    finally:
        errors = []
        for label, cleanup in [('restore tracking', lambda: setattr(doc, 'TrackRevisions', previous)),
                               ('end undo record', undo.EndCustomRecord)]:
            try:
                cleanup()
            except Exception as exc:
                errors.append(f'{label}: {exc}')
        if errors:
            message = 'Cleanup failed; document state may be partial/unknown: ' + '; '.join(errors)
            if primary is not None:
                primary.add_note(message)
            else:
                raise RuntimeError(message)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('command', choices=['probe', 'read', 'inspect', 'find'])
    p.add_argument('--pid', type=int, required=True)
    p.add_argument('--hwnd', type=int, required=True)
    p.add_argument('--start', type=int, default=0)
    p.add_argument('--end', type=int)
    p.add_argument('--text')
    p.add_argument('--style', help='Style name or built-in integer ID, e.g. -4')
    p.add_argument('--limit', type=int, default=5)
    p.add_argument('--output', type=Path)
    p.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    a = p.parse_args()
    if not a.worker:
        try:
            return subprocess.run([sys.executable, str(Path(__file__).resolve()),
                                   *sys.argv[1:], '--worker'], timeout=9).returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({'error': 'Read-only worker exceeded 9 seconds; no document writes'}))
            return 2
    t = time.perf_counter()
    doc, identity = WordBridge().select(a.pid, a.hwnd)
    if a.command == 'probe':
        result = {'body_end': int(doc.Content.End), 'tables': int(doc.Tables.Count),
                  'sections': int(doc.Sections.Count), 'fields_main': int(doc.Fields.Count),
                  'content_controls': int(doc.ContentControls.Count),
                  'read_only': bool(doc.ReadOnly), 'protection_type': int(doc.ProtectionType),
                  'track_revisions': bool(doc.TrackRevisions)}
    elif a.command == 'find':
        style = int(a.style) if a.style and a.style.lstrip('-').isdigit() else a.style
        result = find_literal(doc, a.text, a.start, a.end, style, a.limit)
    else:
        if a.end is None:
            p.error('--end required for local read/inspect')
        result = (read_range if a.command == 'read' else inspect_range)(doc, a.start, a.end)
    report = {'target': identity, 'elapsed_seconds': round(time.perf_counter()-t, 3), **result}
    output = json.dumps(report, ensure_ascii=False, indent=2)
    if a.output:
        a.output.parent.mkdir(parents=True, exist_ok=True)
        a.output.write_text(output, encoding='utf-8')
    else:
        print(output)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
