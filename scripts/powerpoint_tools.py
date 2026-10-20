"""Guarded, attach-only PowerPoint editing primitives (points; Office BGR colors).

No operation opens, saves in place, closes, quits, or changes application settings.
Each request targets a fresh PowerPointBridge identity and local optimistic evidence.
Preview is the default. COM is not transactional: a failed/timed-out commit may have
partially completed, including after its caller stops waiting. Never retry or roll
back automatically; inspect the deck and output paths before making a new request.
Only top-level shapes are supported; groups, tables and embedded objects are refused.
Text and expected values are compared byte-for-character without normalization. In
PowerPoint text, CR is a paragraph break and VT (\x0b) is a soft line break. Some
placeholder TextRange writes translate CR to VT; a mismatch is reported as an
unknown write outcome. Inspect native readback before preparing any next request.
"""
from __future__ import annotations

import argparse
import json
import math
import os
from pathlib import Path
import subprocess
import sys
import tempfile

from office_com import WriteOutcomeUnknown, rebind, rebind_after_write
from office_com import text_hash as _text_digest
from powerpoint_com import PowerPointBridge, Budget, MAX_SHAPES, MAX_TEXT_PER_READ

MAX_SLIDE_SCAN = 500
MAX_RUNS = 256
FONT_FIELDS = {'name': 'Name', 'name_far_east': 'NameFarEast', 'size': 'Size',
               'bold': 'Bold', 'italic': 'Italic', 'underline': 'Underline'}


def text_hash(text):
    """SHA256 of exact UTF-8 text; no newline or Unicode normalization.

    Strict encoding rejects unpaired surrogates: PowerPoint text is either exact
    or damaged, and damaged text must not be hashed into an editable evidence.
    """
    return _text_digest(text, strict=True)


def _positive(value, label):
    if type(value) is not int or value <= 0:
        raise ValueError(label + ' must be a positive integer')
    return value


def _ids(collection, attr, maximum):
    count = int(collection.Count)
    if not 0 <= count <= maximum:
        raise ValueError('Collection exceeds bounded resolution limit')
    values = [int(getattr(collection.Item(i), attr)) for i in range(1, count + 1)]
    if int(collection.Count) != count or len(set(values)) != count:
        raise ValueError('Collection changed or IDs are ambiguous')
    return values


def resolve_slide(presentation, slide_id):
    """Resolve a stable SlideID, never a slide index; scan at most 500 slides."""
    _positive(slide_id, 'slide_id')
    ids = _ids(presentation.Slides, 'SlideID', MAX_SLIDE_SCAN)
    if slide_id not in ids:
        raise ValueError('SlideID not found')
    slide = presentation.Slides.Item(ids.index(slide_id) + 1)
    if int(slide.SlideID) != slide_id:
        raise ValueError('Slide changed during resolution')
    return slide


def resolve_shape(slide, shape_id):
    """Resolve Shape.Id within one slide, bounded to 500 top-level shapes."""
    _positive(shape_id, 'shape_id')
    ids = _ids(slide.Shapes, 'Id', MAX_SHAPES)
    if shape_id not in ids:
        raise ValueError('Shape.Id not found')
    shape = slide.Shapes.Item(ids.index(shape_id) + 1)
    if int(shape.Id) != shape_id:
        raise ValueError('Shape changed during resolution')
    return shape


def _font(rng):
    font = rng.Font
    return {**{key: getattr(font, attr) for key, attr in FONT_FIELDS.items()},
            'color': int(font.Color.RGB)}


def _geometry(shape):
    return {key: float(getattr(shape, key.title()))
            for key in ('left', 'top', 'width', 'height', 'rotation')}


def _text_shape(presentation, slide_id, shape_id):
    shape = resolve_shape(resolve_slide(presentation, slide_id), shape_id)
    if int(shape.Type) in (3, 6, 7, 10, 24) or bool(shape.HasTable) or not bool(shape.HasTextFrame):
        raise ValueError('Only simple top-level text shapes supported')
    return shape


def read_shape(presentation, slide_id, shape_id):
    """Read bounded text, native coordinates, font runs and geometry evidence.

    PowerPoint Characters uses 1-based native coordinates, often UTF-16 units;
    Python string offsets must never be substituted for those coordinates.
    Font values of -2/empty names may mean mixed or inherited formatting.
    """
    shape = _text_shape(presentation, slide_id, shape_id)
    evidence = Budget(MAX_TEXT_PER_READ).text(shape)
    rng = shape.TextFrame.TextRange
    runs = rng.Runs()
    count = int(runs.Count)
    if not 0 <= count <= MAX_RUNS:
        raise ValueError('Text formatting run budget exceeded')
    styles = []
    for i in range(1, count + 1):
        run = rng.Runs(i, 1)
        styles.append({'native_start': int(run.Start), 'native_length': int(run.Length),
                       'font': _font(run)})
    result = {'slide_id': slide_id, 'shape_id': shape_id, **evidence,
              'native_start': int(rng.Start), 'geometry': _geometry(shape),
              'font': _font(rng), 'runs': styles,
              'paragraph_count': int(rng.Paragraphs().Count),
              'paragraph_style': {'alignment': int(rng.ParagraphFormat.Alignment),
                                  'bullet_visible': int(rng.ParagraphFormat.Bullet.Visible)}}
    if Budget(MAX_TEXT_PER_READ).text(shape) != evidence:
        raise ValueError('Text changed while reading evidence')
    return result


REFUSED = ('Target cannot be rebound natively ({stage}); no write attempted '
           '(未尝试写入). Reread and replan: {error}')
IDENTITY_CHANGED = ('Native identity changed ({stage}); no write attempted '
                    '(未尝试写入). Reread and replan.')
WRITE_LOST = ('Write attempted, but the target cannot be rebound afterwards, so the '
              'outcome may be partial. Do not retry or undo automatically; inspect '
              '{label} first: {error}')


def _target(bridge, expected, stage='读取前'):
    if not isinstance(expected, dict) or expected.get('application') != 'powerpoint':
        raise ValueError('A full PowerPointBridge target identity is required')
    return rebind(bridge, expected, stage, refusal=REFUSED, mismatch=IDENTITY_CHANGED)


def _expect_ids(collection, attr, maximum, expected):
    actual = _ids(collection, attr, maximum)
    if not isinstance(expected, list) or any(type(v) is not int for v in expected) or actual != expected:
        raise ValueError('Ordered ID/count evidence changed')
    return actual


def _text(value):
    if not isinstance(value, str) or len(value.encode('utf-16-le')) // 2 > MAX_TEXT_PER_READ or '\x00' in value:
        raise ValueError('Bounded text string required (no NUL)')
    return value


def _path(value, suffixes):
    if not isinstance(value, str) or not value:
        raise ValueError('Explicit absolute output path required')
    path = Path(value)
    if not path.is_absolute() or path.suffix.lower() not in suffixes or not path.parent.is_dir():
        raise ValueError('Output needs absolute path, supported suffix and existing parent')
    if path.exists():
        raise ValueError('Output already exists; overwrite is forbidden')
    return path


def _number(value, label, low, high):
    if type(value) not in (int, float) or not math.isfinite(value) or not low <= value <= high:
        raise ValueError(label + ' out of bounds')
    return value


def _validate_font(font):
    for key in ('name', 'name_far_east'):
        if key in font and (not isinstance(font[key], str) or not font[key].strip()):
            raise ValueError('Font names must be nonempty strings')
    for key in ('bold', 'italic'):
        if key in font and type(font[key]) is not bool:
            raise ValueError(key + ' must be bool')
    if 'size' in font:
        _number(font['size'], 'font size', 1, 400)
    if 'color' in font and (type(font['color']) is not int or not 0 <= font['color'] <= 0xffffff):
        raise ValueError('Color must be Office BGR integer 0..0xffffff')


def _set_font(rng, font):
    for key, value in font.items():
        if key == 'color':
            rng.Font.Color.RGB = value
        else:
            setattr(rng.Font, FONT_FIELDS[key], (-1 if value else 0) if type(value) is bool else value)


def _plan(presentation, op):
    """Validate and resolve one local operation without writes; return evidence."""
    kind = op.get('kind')
    if kind in ('replace_text', 'format_text', 'duplicate_slide', 'delete_slide', 'add_textbox'):
        try:
            read_only = presentation.ReadOnly
        except Exception as exc:
            raise ValueError('Cannot establish Presentation.ReadOnly; mutation refused') from exc
        if type(read_only) not in (bool, int) or read_only not in (0, -1, False, True):
            raise ValueError('Unknown Presentation.ReadOnly value; mutation refused')
        if read_only != 0:
            raise ValueError('Read-only presentation; mutation refused')
    if kind in ('replace_text', 'format_text'):
        shape = _text_shape(presentation, op['slide_id'], op['shape_id'])
        evidence = read_shape(presentation, op['slide_id'], op['shape_id'])
        if op.get('expected') != evidence['text'] or op.get('text_sha256') != evidence['text_sha256']:
            raise ValueError('Expected text/hash changed')
        if kind == 'format_text':
            geometry, font = op.get('geometry', {}), op.get('font', {})
            if not geometry and not font and 'alignment' not in op:
                raise ValueError('Explicit formatting change required')
            if set(geometry) - {'left', 'top', 'width', 'height'} or set(font) - {'name', 'name_far_east', 'size', 'bold', 'italic', 'color'}:
                raise ValueError('Unsupported geometry/font field')
            for key, value in geometry.items():
                _number(value, key, 0 if key in ('left', 'top') else .1, 10000)
            _validate_font(font)
            if 'alignment' in op and (type(op['alignment']) is not int or op['alignment'] not in (1, 2, 3, 4)):
                raise ValueError('alignment must be PowerPoint 1=left,2=center,3=right,4=justify')
            return {'evidence': evidence}, shape
        _text(op.get('text'))
        rng = shape.TextFrame.TextRange
        strategy = op.get('strategy', 'native_replace')
        if 'preserve_strategy' in op and strategy != 'whole':
            raise ValueError('preserve_strategy is allowed only for whole replacement')
        mixed = (len(evidence['runs']) != 1 or evidence['paragraph_count'] != 1 or
                 any(v in (-2, '') for v in evidence['font'].values()))
        covers_whole = (strategy == 'native_replace' and op.get('find') == evidence['text']) or (
            strategy == 'characters' and op.get('native_start') == 1 and
            op.get('native_length') == evidence['native_length'])
        if mixed and covers_whole:
            raise ValueError('Whole-range mixed formatting requires whole strategy and explicit preserve_strategy')
        if strategy == 'native_replace':
            find = _text(op.get('find'))
            # Count overlapping exact matches as ambiguous too.
            if not find or sum(evidence['text'].startswith(find, i) for i in range(len(evidence['text']))) != 1:
                raise ValueError('Native Replace requires one exact unique match')
            selected = None
        elif strategy == 'characters':
            start = _positive(op.get('native_start'), 'native_start')
            length = _positive(op.get('native_length'), 'native_length')
            if start + length - 1 > evidence['native_length']:
                raise ValueError('Characters span out of bounds')
            selected = rng.Characters(start, length)
            if str(selected.Text) != op.get('span_expected'):
                raise ValueError('Native span text changed')
        elif strategy == 'whole':
            mixed = (len(evidence['runs']) != 1 or evidence['paragraph_count'] != 1 or
                     any(v in (-2, '') for v in evidence['font'].values()))
            preserve = op.get('preserve_strategy')
            if preserve not in (None, 'first_character'):
                raise ValueError('Unknown preserve_strategy')
            if mixed and preserve != 'first_character':
                raise ValueError('Whole replacement rejects mixed formatting; explicit first_character strategy required')
            if preserve == 'first_character' and not evidence['native_length']:
                raise ValueError('No first character to preserve')
            selected = rng
        else:
            raise ValueError('Unknown replacement strategy')
        return {'evidence': evidence}, (shape, rng, selected)
    if kind in ('duplicate_slide', 'delete_slide', 'export_slide', 'save_copy'):
        ids = _expect_ids(presentation.Slides, 'SlideID', MAX_SLIDE_SCAN, op.get('expected_slide_ids'))
        slide = resolve_slide(presentation, op['slide_id']) if kind != 'save_copy' else None
        if kind == 'duplicate_slide':
            destination = _positive(op.get('move_to'), 'move_to')
            if destination > len(ids) + 1 or len(ids) >= MAX_SLIDE_SCAN:
                raise ValueError('Duplicate destination/slide count out of bounds')
        if kind == 'export_slide':
            _path(op.get('path'), {'.png', '.jpg', '.jpeg'})
            for key in ('width', 'height'):
                if type(op.get(key)) is not int or not 1 <= op[key] <= 8192:
                    raise ValueError('Export dimensions must be integers 1..8192')
        if kind == 'save_copy':
            _path(op.get('path'), {'.pptx'})
            if bool(presentation.HasVBProject):
                raise ValueError('Macro deck cannot be saved as a .pptx copy')
        return {'slide_ids': ids, 'slide_count': len(ids)}, slide
    if kind == 'add_textbox':
        slide = resolve_slide(presentation, op['slide_id'])
        ids = _expect_ids(slide.Shapes, 'Id', MAX_SHAPES, op.get('expected_shape_ids'))
        if len(ids) >= MAX_SHAPES:
            raise ValueError('Shape budget exceeded')
        _text(op.get('text'))
        geometry, font = op.get('geometry', {}), op.get('font', {})
        if set(geometry) != {'left', 'top', 'width', 'height'}:
            raise ValueError('Explicit left/top/width/height in points required')
        for key, value in geometry.items():
            _number(value, key, 0 if key in ('left', 'top') else .1, 10000)
        if not {'name', 'name_far_east', 'size', 'color'} <= set(font) or set(font) - {'name', 'name_far_east', 'size', 'color', 'bold', 'italic'}:
            raise ValueError('Explicit font name/name_far_east/size/color required')
        if any(not isinstance(font[k], str) or not font[k].strip() for k in ('name', 'name_far_east')):
            raise ValueError('Both Latin and CJK font names required')
        _validate_font(font)
        return {'shape_ids': ids, 'shape_count': len(ids)}, slide
    raise ValueError('Unsupported operation kind')


def _file_write(presentation, slide, op):
    # Native COM writes to a private sibling staging directory. Atomic hard-link
    # publication refuses even a destination created concurrently after preview.
    path = _path(op['path'], {'.png', '.jpg', '.jpeg'} if slide is not None else {'.pptx'})
    stage_dir = Path(tempfile.mkdtemp(prefix='.powerpoint-tools-', dir=path.parent))
    staged = stage_dir / path.name
    try:
        if slide is not None:
            slide.Export(str(staged), 'PNG' if path.suffix.lower() == '.png' else 'JPG', op['width'], op['height'])
        else:
            presentation.SaveCopyAs(str(staged), 24)  # ppSaveAsOpenXMLPresentation
        os.link(staged, path)  # Atomic no-clobber; never os.replace.
        staged.unlink()
        stage_dir.rmdir()
    except Exception as exc:
        raise WriteOutcomeUnknown(
            f'File write/publication failed; inspect destination and staging {staged}. '
            f'No retry or rollback: {exc}') from exc
    return {'path': str(path)}


def _verify_format(after, op):
    for key, value in op.get('geometry', {}).items():
        if abs(after['geometry'][key] - value) > .05:
            raise RuntimeError('Geometry postcondition failed: ' + key)
    for key, value in op.get('font', {}).items():
        expected = (-1 if value else 0) if type(value) is bool else value
        actual = after['font'][key]
        if (abs(actual - expected) > .05 if key == 'size' else actual != expected):
            raise RuntimeError('Font postcondition failed: ' + key)
    if 'alignment' in op and after['paragraph_style']['alignment'] != op['alignment']:
        raise RuntimeError('Alignment postcondition failed')


def _write(presentation, op, resolved):
    kind = op['kind']
    if kind == 'format_text':
        for key, value in op.get('geometry', {}).items():
            setattr(resolved, key.title(), value)
        rng = resolved.TextFrame.TextRange
        _set_font(rng, op.get('font', {}))
        if 'alignment' in op:
            rng.ParagraphFormat.Alignment = op['alignment']
        after = read_shape(presentation, op['slide_id'], op['shape_id'])
        if after['text'] != op['expected']:
            raise RuntimeError('Formatting unexpectedly changed text')
        _verify_format(after, op)
        return {'after': after}
    if kind == 'replace_text':
        shape, rng, selected = resolved
        strategy = op.get('strategy', 'native_replace')
        if strategy == 'native_replace':
            intended = op['expected'].replace(op['find'], op['text'], 1)
        elif strategy == 'whole':
            intended = op['text']
        else:
            start, length, total = op['native_start'], op['native_length'], int(rng.Length)
            prefix = str(rng.Characters(1, start - 1).Text) if start > 1 else ''
            end = start + length
            suffix = str(rng.Characters(end, total - end + 1).Text) if end <= total else ''
            intended = prefix + op['text'] + suffix
        if strategy == 'native_replace':
            replaced = rng.Replace(op['find'], op['text'], 0, -1, 0)  # After=0, MatchCase=True, WholeWords=False
            if replaced is None:
                raise RuntimeError('Native Replace returned no match')
        else:
            first_font = _font(rng.Characters(1, 1)) if strategy == 'whole' and op.get('preserve_strategy') == 'first_character' else None
            selected.Text = op['text']
            if first_font:
                for key, attr in FONT_FIELDS.items():
                    setattr(rng.Font, attr, first_font[key])
                rng.Font.Color.RGB = first_font['color']
        after = read_shape(presentation, op['slide_id'], op['shape_id'])
        if after['text'] != intended:
            raise RuntimeError('Replacement text postcondition failed')
        return {'after': after}
    if kind == 'duplicate_slide':
        duplicated = resolved.Duplicate()
        if int(duplicated.Count) != 1:
            raise RuntimeError('Unexpected duplicate result')
        new_slide = duplicated.Item(1)
        new_id = int(new_slide.SlideID)
        new_slide.MoveTo(op['move_to'])
        actual = _ids(presentation.Slides, 'SlideID', MAX_SLIDE_SCAN)
        expected = list(op['expected_slide_ids'])
        if new_id in expected:
            raise RuntimeError('Duplicate reused an existing SlideID')
        expected.insert(op['move_to'] - 1, new_id)
        if actual != expected:
            raise RuntimeError('Duplicate order/count postcondition failed')
        return {'slide_id': new_id, 'slide_ids': actual}
    if kind == 'delete_slide':
        resolved.Delete()
        actual = _ids(presentation.Slides, 'SlideID', MAX_SLIDE_SCAN)
        if actual != [i for i in op['expected_slide_ids'] if i != op['slide_id']]:
            raise RuntimeError('Delete order/count postcondition failed')
        return {'slide_ids': actual}
    if kind == 'add_textbox':
        g, f = op['geometry'], op['font']
        shape = resolved.Shapes.AddTextbox(1, g['left'], g['top'], g['width'], g['height'])
        shape.TextFrame.AutoSize = 0  # ppAutoSizeNone: honor caller geometry.
        shape.TextFrame.WordWrap = -1
        rng = shape.TextFrame.TextRange
        rng.Text = op['text']
        _set_font(rng, f)
        rng.ParagraphFormat.Bullet.Visible = 0
        shape_id = int(shape.Id)
        after = read_shape(presentation, op['slide_id'], shape_id)
        if after['text'] != op['text'] or after['paragraph_style']['bullet_visible'] != 0:
            raise RuntimeError('Textbox text/bullet postcondition failed')
        _verify_format(after, op)
        return {'shape_id': shape_id, 'after': after}
    return _file_write(presentation, resolved, op)


def apply_request(bridge, request, commit=False):
    """Preview or explicitly commit one guarded JSON operation.

    request = {target: full PowerPointBridge identity, operation: {kind: ...}}.
    Text strings are exact: CR denotes paragraph breaks and VT (\x0b) denotes
    soft line breaks. No CR/LF conversion or normalization is performed. Native
    placeholder writes may translate CR to VT; exact postconditions then report
    an unknown outcome. Inspect read_shape before planning another mutation.
    replace_text requires slide_id, shape_id, expected (whole current text),
    text_sha256, text, strategy. native_replace also needs find; characters needs
    native_start/native_length and span_expected. Whole replacements require
    uniform formatting or preserve_strategy='first_character' (font inherited
    explicitly; paragraph formatting is native PowerPoint inheritance).
    format_text requires expected/hash and accepts geometry/font subsets plus
    alignment=1/2/3/4. Font bold/italic are booleans.
    duplicate_slide/delete_slide/export_slide/save_copy require ordered
    expected_slide_ids, proving count and order. duplicate_slide uses move_to
    (final 1-based index). add_textbox requires expected_shape_ids, text, geometry
    in points, font={name,name_far_east,size,color}; textbox orientation is horizontal.
    export_slide requires explicit path and pixel width/height; save_copy writes
    .pptx only. File operations are separate and never overwrite an output.

    Identity and local evidence are checked again immediately before commit;
    this is optimistic concurrency, not an atomic lock. A request returning
    WriteOutcomeUnknown must never be retried automatically.
    """
    if type(commit) is not bool:
        raise ValueError('commit must be explicit bool')
    # Freeze caller-owned mutable JSON before checks.
    request = json.loads(json.dumps(request, allow_nan=False))
    expected, op = request['target'], request['operation']
    presentation, identity = _target(bridge, expected)
    evidence, resolved = _plan(presentation, op)
    result = {'target': identity, 'operation': op, 'committed': False,
              'saved_in_place': False, 'evidence': evidence}
    if not commit:
        return result
    presentation, identity = _target(bridge, expected, '提交前')
    fresh, resolved = _plan(presentation, op)
    if fresh != evidence:
        raise ValueError('Local evidence changed before commit; reread and replan')
    try:
        data = _write(presentation, op, resolved)
        rebind_after_write(bridge, expected, label='PowerPoint', message=WRITE_LOST)
    except WriteOutcomeUnknown:
        raise
    except Exception as exc:
        raise WriteOutcomeUnknown('Write attempted; partial/unknown outcome. No retry or rollback. ' + str(exc)) from exc
    return {**result, 'committed': True, 'data': data}


def main(argv=None):
    """JSON CLI: apply --request FILE [--commit]; read --target FILE --slide-id ID --shape-id ID.

    Worker deadline bounds caller waiting, not PowerPoint execution. A commit
    timeout reports unknown outcome, even if no write was observed. No retry.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=('read', 'apply'))
    parser.add_argument('--request', type=Path)
    parser.add_argument('--target', type=Path)
    parser.add_argument('--slide-id', type=int)
    parser.add_argument('--shape-id', type=int)
    parser.add_argument('--commit', action='store_true')
    parser.add_argument('--timeout', type=int, default=9)
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if not 1 <= args.timeout <= 120:
        parser.error('--timeout must be 1..120 seconds')
    if args.commit and args.command != 'apply':
        parser.error('--commit applies only to apply')
    if not args.worker:
        forwarded = list(sys.argv[1:] if argv is None else argv)
        try:
            return subprocess.run([sys.executable, str(Path(__file__).resolve()), *forwarded, '--worker'],
                                  timeout=args.timeout).returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({'status': 'unknown' if args.commit else 'error',
                              'error': 'Worker timeout; PowerPoint may continue processing. Inspect before any new request; no retry/rollback.',
                              'write_outcome': 'unknown' if args.commit else 'no writes requested'}), file=sys.stderr)
            return 2
    try:
        bridge = PowerPointBridge()
        if args.command == 'apply':
            if args.request is None:
                raise ValueError('--request required')
            result = apply_request(bridge, json.loads(args.request.read_text(encoding='utf-8-sig')), args.commit)
        else:
            if args.target is None:
                raise ValueError('--target required')
            expected = json.loads(args.target.read_text(encoding='utf-8-sig'))
            presentation, identity = _target(bridge, expected)
            result = {'target': identity, 'data': read_shape(presentation, args.slide_id, args.shape_id)}
            _target(bridge, expected)
        print(json.dumps({'status': 'ok', **result}, ensure_ascii=False, allow_nan=False))
        return 0
    except Exception as exc:
        print(json.dumps({'status': 'unknown' if isinstance(exc, WriteOutcomeUnknown) else 'error',
                          'type': type(exc).__name__, 'error': str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
