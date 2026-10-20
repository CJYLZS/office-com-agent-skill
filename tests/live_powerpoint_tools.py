"""Opt-in authentic PowerPoint COM tests: owned decks only; NEVER Application.Quit.

Run with the COM-enabled interpreter: python tests/live_powerpoint_tools.py --run
Imports and COM attachment occur only after the explicit opt-in. Output artifacts
are retained in a unique workspace tmp directory for inspection after a failure.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import time

ROOT = Path(__file__).resolve().parents[1]


def ids(collection, attribute):
    return [int(getattr(collection.Item(i), attribute))
            for i in range(1, int(collection.Count) + 1)]


def refused(action):
    try:
        action()
    except ValueError:
        return
    raise AssertionError('Expected pre-write ValueError refusal')


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', action='store_true')
    args = parser.parse_args(argv)
    if not args.run:
        parser.error('Explicit --run required; creates/closes only owned temporary presentations')

    sys.path.insert(0, str(ROOT / 'scripts'))
    from comtypes.client import CreateObject, PumpEvents
    from powerpoint_com import PowerPointBridge
    from powerpoint_tools import apply_request, read_shape, resolve_shape, resolve_slide

    bridge = PowerPointBridge()
    before_windows = bridge.windows()
    # PowerPoint is a single-instance application: an existing frame is reused, so
    # no owned top-level window appears and every later assertion loses its target.
    # This suite never quits the shared instance, so close PowerPoint and re-run.
    if before_windows:
        raise SystemExit(json.dumps({
            'status': 'error',
            'error': 'A PowerPoint frame already exists; close PowerPoint and re-run',
            'windows': before_windows,
        }, ensure_ascii=False))
    # PowerPoint is a singleton; this does NOT imply ownership of Application.
    app = CreateObject('PowerPoint.Application', dynamic=True)
    owned, passed = [], []
    (ROOT / 'tmp').mkdir(exist_ok=True)
    artifacts = Path(tempfile.mkdtemp(prefix='powerpoint-tools-live-', dir=ROOT / 'tmp'))
    cli_elapsed = None
    try:
        deck = app.Presentations.Add(True)
        owned.append(deck)
        assert type(deck.ReadOnly) is int and deck.ReadOnly == 0, repr(deck.ReadOnly)
        slide = deck.Slides.Add(1, 12)
        sid = int(slide.SlideID)
        shape = slide.Shapes.AddTextbox(1, 20, 30, 360, 60)
        shape.TextFrame.TextRange.Text = 'Uniform text alpha'
        font = shape.TextFrame.TextRange.Font
        font.Name, font.NameFarEast, font.Size = 'Arial', 'Microsoft YaHei', 24
        font.Bold, font.Italic, font.Underline = 0, 0, 0
        font.Color.RGB = 0x123456
        shid = int(shape.Id)
        mixed = slide.Shapes.AddTextbox(1, 20, 110, 360, 60)
        mixed.TextFrame.TextRange.Text = 'Mixed styles'
        mixed.TextFrame.TextRange.Font.Name = 'Arial'
        mixed.TextFrame.TextRange.Font.Size = 22
        mixed.TextFrame.TextRange.Characters(1, 5).Font.Bold = -1
        mixed.TextFrame.TextRange.Characters(6, 7).Font.Bold = 0
        mid = int(mixed.Id)
        deck.Slides.Add(2, 12)
        PumpEvents(0.5)
        old_hwnds = {w['hwnd'] for w in before_windows}
        created = [w for w in bridge.windows() if w['hwnd'] not in old_hwnds]
        assert len(created) == 1, created
        hwnd = created[0]['hwnd']
        pid = bridge.pid(hwnd)
        # A second owned deck stays active throughout targeting tests.
        other = app.Presentations.Add(True)
        owned.append(other)
        other_slide = other.Slides.Add(1, 12)
        sentinel = other_slide.Shapes.AddTextbox(1, 10, 10, 200, 30)
        sentinel.TextFrame.TextRange.Text = 'OTHER DECK MUST STAY UNCHANGED'
        other_before = (str(other.Name), bool(other.Saved), ids(other.Slides, 'SlideID'),
                        str(sentinel.TextFrame.TextRange.Text), ids(other_slide.Shapes, 'Id'))
        _, target = bridge.select(pid, hwnd)
        assert target['name'] == str(deck.Name)
        assert str(app.ActivePresentation.Name) == str(other.Name)

        def state():
            slides = []
            for slide_id in ids(deck.Slides, 'SlideID'):
                current = resolve_slide(deck, slide_id)
                slides.append((slide_id, [read_shape(deck, slide_id, shape_id)
                                         for shape_id in ids(current.Shapes, 'Id')]))
            return (str(deck.Name), str(deck.FullName), bool(deck.Saved), slides)

        def request(op):
            return {'target': target, 'operation': op}

        def preview(op):
            before = state()
            result = apply_request(bridge, request(op))
            assert result['committed'] is False
            assert result['saved_in_place'] is False
            assert state() == before, 'Preview mutated presentation'
            if 'path' in op:
                assert not Path(op['path']).exists(), 'Preview wrote output'
            return result

        def commit(op):
            preview(op)
            result = apply_request(bridge, request(op), commit=True)
            assert result['committed'] is True and result['saved_in_place'] is False
            return result['data']

        def text_op(kind, shape_id=shid, **extra):
            evidence = read_shape(deck, sid, shape_id)
            assert evidence['text_sha256'] == hashlib.sha256(evidence['text'].encode('utf-8')).hexdigest()
            return dict(kind=kind, slide_id=sid, shape_id=shape_id,
                        expected=evidence['text'], text_sha256=evidence['text_sha256'], **extra)

        assert int(resolve_slide(deck, sid).SlideID) == sid
        assert int(resolve_shape(slide, shid).Id) == shid
        refused(lambda: resolve_slide(deck, max(ids(deck.Slides, 'SlideID')) + 1000))
        refused(lambda: resolve_shape(slide, max(ids(slide.Shapes, 'Id')) + 1000))
        passed.append('stable_id_resolution_and_missing_id_refusal')

        original = read_shape(deck, sid, shid)
        op = text_op('replace_text', text='Replaced text beta', strategy='whole')
        commit(op)
        after = read_shape(deck, sid, shid)
        assert after['text'] == 'Replaced text beta'
        assert after['font'] == original['font'], (original['font'], after['font'])
        assert str(app.ActivePresentation.Name) == str(other.Name)
        passed.extend(['preview_unchanged', 'exact_target_with_other_active', 'whole_replacement_preserves_uniform_font'])

        op = text_op('replace_text', text='gamma', strategy='native_replace', find='beta')
        commit(op)
        assert read_shape(deck, sid, shid)['text'] == 'Replaced text gamma'
        op = text_op('replace_text', text='Native', strategy='characters',
                     native_start=1, native_length=8, span_expected='Replaced')
        commit(op)
        assert read_shape(deck, sid, shid)['text'] == 'Native text gamma'
        passed.append('native_replace_and_native_characters')

        before = state()
        mixed_op = text_op('replace_text', shape_id=mid, text='Flattened', strategy='whole')
        refused(lambda: apply_request(bridge, request(mixed_op), commit=True))
        for changes in ({'expected': 'stale'}, {'text_sha256': '0' * 64}):
            stale = text_op('replace_text', text='wrong', strategy='whole')
            stale.update(changes)
            refused(lambda: apply_request(bridge, request(stale), commit=True))
        stale_target = copy.deepcopy(target)
        stale_target['name'] += '-stale'
        refused(lambda: apply_request(bridge, {'target': stale_target, 'operation': mixed_op}, commit=True))
        assert state() == before
        passed.append('mixed_whole_stale_expected_hash_and_identity_refused')

        old_ids = ids(deck.Slides, 'SlideID')
        duplicate = commit(dict(kind='duplicate_slide', slide_id=sid,
                                expected_slide_ids=old_ids, move_to=len(old_ids) + 1))
        duplicate_id = duplicate['slide_id']
        assert duplicate_id not in old_ids
        assert ids(deck.Slides, 'SlideID') == old_ids + [duplicate_id]
        assert duplicate['slide_ids'] == ids(deck.Slides, 'SlideID')
        before = state()
        refused(lambda: apply_request(bridge, request(dict(kind='delete_slide', slide_id=duplicate_id,
                                                          expected_slide_ids=old_ids)), commit=True))
        assert state() == before
        passed.append('duplicate_new_id_move_and_stale_slide_ids_refused')

        geometry = dict(left=40, top=190, width=320, height=70)
        font_spec = dict(name='Arial', name_far_east='Microsoft YaHei', size=20,
                         color=0x654321, bold=True, italic=False)
        add = dict(kind='add_textbox', slide_id=sid, expected_shape_ids=ids(slide.Shapes, 'Id'),
                   text='Editable 文本', geometry=geometry, font=font_spec)
        new_shape_id = commit(add)['shape_id']
        new_shape = resolve_shape(slide, new_shape_id)
        evidence = read_shape(deck, sid, new_shape_id)
        assert int(new_shape.Type) == 17 and bool(new_shape.HasTextFrame)
        assert evidence['text'] == 'Editable 文本'
        for key, value in geometry.items():
            assert abs(evidence['geometry'][key] - value) < .01
        for key, value in font_spec.items():
            expected = (-1 if value else 0) if type(value) is bool else value
            assert evidence['font'][key] == expected
        before = state()
        refused(lambda: apply_request(bridge, request(add), commit=True))
        assert state() == before
        changed_geometry = dict(left=55, top=200, width=350, height=80)
        changed_font = dict(name='Arial', name_far_east='Microsoft YaHei', size=26,
                            color=0x112233, bold=False, italic=True)
        formatted = commit(text_op('format_text', shape_id=new_shape_id, font=changed_font,
                                   geometry=changed_geometry, alignment=2))['after']
        assert formatted['text'] == 'Editable 文本'
        for key, value in changed_geometry.items():
            assert abs(formatted['geometry'][key] - value) < .01
        for key, value in changed_font.items():
            expected = (-1 if value else 0) if type(value) is bool else value
            assert formatted['font'][key] == expected
        assert int(new_shape.TextFrame.TextRange.ParagraphFormat.Alignment) == 2
        passed.append('editable_textbox_font_geometry_alignment_and_stale_shape_ids')

        png = artifacts / 'slide.png'
        export = dict(kind='export_slide', slide_id=sid, expected_slide_ids=ids(deck.Slides, 'SlideID'),
                      path=str(png), width=800, height=450)
        source_state = state()
        commit(export)
        header = png.read_bytes()
        assert header[:8] == b'\x89PNG\r\n\x1a\n'
        assert struct.unpack('>II', header[16:24]) == (800, 450)
        assert state() == source_state
        refused(lambda: apply_request(bridge, request(export), commit=True))
        assert png.read_bytes() == header
        passed.append('native_png_export_dimensions_no_overwrite')

        for saved in (False, True):
            # Saved changes only our returned owned object; test both source states.
            deck.Saved = saved
            source_state = state()
            output = artifacts / ('copy-saved.pptx' if saved else 'copy-unsaved.pptx')
            save = dict(kind='save_copy', expected_slide_ids=ids(deck.Slides, 'SlideID'), path=str(output))
            commit(save)
            assert output.read_bytes()[:2] == b'PK'
            assert state() == source_state, 'SaveCopy changed source name, Saved, or content'
            digest = hashlib.sha256(output.read_bytes()).hexdigest()
            refused(lambda: apply_request(bridge, request(save), commit=True))
            assert hashlib.sha256(output.read_bytes()).hexdigest() == digest
            assert state() == source_state
        passed.append('save_copy_preserves_source_name_saved_both_states_no_overwrite')

        expected_ids = ids(deck.Slides, 'SlideID')
        commit(dict(kind='delete_slide', slide_id=duplicate_id, expected_slide_ids=expected_ids))
        assert ids(deck.Slides, 'SlideID') == old_ids
        refused(lambda: resolve_slide(deck, duplicate_id))
        passed.append('delete_exact_slide_id')

        op = text_op('replace_text', text='CLI MUST ONLY PREVIEW', strategy='whole')
        request_path = artifacts / 'cli-request.json'
        request_path.write_text(json.dumps(request(op), ensure_ascii=False), encoding='utf-8')
        before = state()
        started = time.perf_counter()
        # File redirection avoids named-pipe capture and exercises the real supervisor.
        with (artifacts / 'cli-stdout.json').open('w', encoding='utf-8') as stdout, \
                (artifacts / 'cli-stderr.txt').open('w', encoding='utf-8') as stderr:
            completed = subprocess.run([sys.executable, str(ROOT / 'scripts' / 'powerpoint_tools.py'),
                                        'apply', '--request', str(request_path), '--timeout', '8'],
                                       stdout=stdout, stderr=stderr, timeout=10)
        cli_elapsed = time.perf_counter() - started
        assert completed.returncode == 0, (artifacts / 'cli-stderr.txt').read_text(encoding='utf-8')
        result = json.loads((artifacts / 'cli-stdout.json').read_text(encoding='utf-8'))
        assert result['status'] == 'ok' and result['committed'] is False
        assert result['target'] == target and cli_elapsed < 10
        assert state() == before
        passed.append('supervised_cli_preview_under_10s')

        assert (str(other.Name), bool(other.Saved), ids(other.Slides, 'SlideID'),
                str(sentinel.TextFrame.TextRange.Text), ids(other_slide.Shapes, 'Id')) == other_before
        passed.append('other_owned_active_deck_untouched')
        print(json.dumps(dict(status='ok', passed=passed, artifacts=str(artifacts),
                              cli_elapsed_seconds=round(cli_elapsed, 3),
                              existing_windows_before=before_windows,
                              application_quit=False, user_presentations_modified=False),
                         ensure_ascii=False, indent=2))
    except Exception:
        print(json.dumps(dict(status='failed', passed=passed, artifacts=str(artifacts)),
                         ensure_ascii=False), file=sys.stderr)
        raise
    finally:
        primary_error = sys.exc_info()[1]
        cleanup_errors = []
        for presentation in reversed(owned):
            try:
                presentation.Saved = True
                presentation.Close()
            except Exception as exc:
                cleanup_errors.append(str(exc))
        if cleanup_errors:
            print(json.dumps({'cleanup_errors': cleanup_errors}), file=sys.stderr)
            if primary_error is None:
                raise RuntimeError('Could not close some owned presentations')
        # NEVER Quit: other user presentations may share this Application.
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
