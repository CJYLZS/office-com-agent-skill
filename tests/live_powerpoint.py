"""Opt-in independent COM test; create/close only owned presentations, NEVER Quit."""
import json
from pathlib import Path
import subprocess
import sys
import time
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from powerpoint_com import PowerPointBridge, normalized_hwnd, probe, read_slides


def main():
    if '--run' not in sys.argv:
        raise SystemExit('Explicit --run required; creates only owned temporary presentations')
    from comtypes.client import CreateObject
    bridge = PowerPointBridge()
    existing_windows = bridge.windows()
    # PowerPoint may share a singleton! Never assume ownership of Application.
    app = CreateObject('PowerPoint.Application', dynamic=True)
    owned = []
    passed = []
    try:
        deck = app.Presentations.Add(True)
        owned.append(deck)
        assert probe(deck)['slide_count'] == 0
        assert read_slides(deck)['slides'] == []
        passed.append('empty_deck')
        slide = deck.Slides.Add(1, 12)
        shape = slide.Shapes.AddTextbox(1, 10, 10, 300, 40)
        shape.TextFrame.TextRange.Text = 'COM live 中文'
        table = slide.Shapes.AddTable(2, 2, 10, 70, 300, 100).Table
        table.Cell(1, 1).Shape.TextFrame.TextRange.Text = 'cell evidence'
        slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'owned notes'
        deck.Slides.Add(2, 12)
        from comtypes.client import PumpEvents
        PumpEvents(0.5)
        new_windows = [w for w in bridge.windows() if w['hwnd'] not in {x['hwnd'] for x in existing_windows}]
        assert len(new_windows) == 1, new_windows
        hwnd = new_windows[0]['hwnd']
        pid = bridge.pid(hwnd)
        other = app.Presentations.Add(True)
        owned.append(other)  # Now a different owned presentation is active.
        start = time.perf_counter()
        attach_error = None
        try:
            attached, target = bridge.select(pid, hwnd)
        except Exception as exc:
            attach_error = str(exc)
            children = []
            @bridge.callback
            def collect(child, _):
                children.append([int(child), bridge.label(child, True)])
                return True
            bridge.user32.EnumChildWindows(hwnd, collect, 0)
            print(json.dumps({'diagnostic_children': children}))
            attached, target = deck, {'name': str(deck.Name)}
        assert target['name'] == str(deck.Name)
        if attach_error is None:
            passed.append('exact_hwnd_target_with_other_presentation_active')
        saved = bool(deck.Saved)
        data = read_slides(attached, limit=1, tables=True, notes=True)
        assert data['next_offset'] == 2
        assert data['slides'][0]['slide_id'] == int(slide.SlideID)
        assert data['slides'][0]['shapes'][0]['shape_id'] == int(shape.Id)
        assert data['slides'][0]['shapes'][0]['text'] == 'COM live 中文'
        assert data['slides'][0]['shapes'][1]['table']['cells'][0]['text'] == 'cell evidence'
        assert any(s.get('text') == 'owned notes' for s in data['slides'][0]['notes']['shapes'])
        assert bool(deck.Saved) == saved
        passed.extend(['text_table_notes_ids_pagination', 'preserve_unsaved_state'])
        try:
            bridge.select(pid + 1, hwnd)
        except ValueError:
            passed.append('wrong_pid_rejected')
        else:
            raise AssertionError('Wrong PID accepted')
        try:
            read_slides(attached, text_budget=2)
        except ValueError:
            passed.append('text_budget_fail_closed')
        else:
            raise AssertionError('Text budget ignored')
        if attach_error is not None:
            print(json.dumps({'status': 'partial', 'passed': passed, 'attach_error': attach_error,
                              'direct_owned_read_seconds': round(time.perf_counter() - start, 3),
                              'existing_windows_before': existing_windows,
                              'application_quit': False, 'user_presentations_modified': False}, ensure_ascii=False, indent=2))
            return 3
        # Exercise the actual supervised CLI without pipe capture requirements.
        result_path = Path(__file__).resolve().parents[1] / 'tmp' / 'powerpoint-live-read.json'
        cli_start = time.perf_counter()
        result = subprocess.run([sys.executable, str(Path(__file__).resolve().parents[1] / 'scripts' / 'powerpoint_com.py'),
                                 'read', '--pid', str(pid), '--hwnd', str(hwnd), '--limit', '1',
                                 '--notes', '--tables', '--output', str(result_path)], timeout=10)
        cli_elapsed = time.perf_counter() - cli_start
        assert result.returncode == 0
        assert cli_elapsed < 10
        passed.append('supervised_cli_under_10s')
        print(json.dumps({'passed': passed, 'cli_elapsed_seconds': round(cli_elapsed, 3),
                          'total_read_seconds': round(time.perf_counter() - start, 3),
                          'existing_windows_before': existing_windows,
                          'application_quit': False, 'user_presentations_modified': False}, ensure_ascii=False, indent=2))
    finally:
        primary_error = sys.exc_info()[1]
        cleanup_errors = []
        for presentation in reversed(owned):
            # Only exact returned objects are owned; continue after one cleanup failure.
            try:
                presentation.Saved = True
                presentation.Close()
            except Exception as exc:
                cleanup_errors.append(str(exc))
        # Never Quit a possibly shared Application or mask a primary test failure.
        if cleanup_errors:
            print(json.dumps({'cleanup_errors': cleanup_errors}), file=sys.stderr)
            if primary_error is None:
                raise RuntimeError('Some owned test presentations could not be closed')


if __name__ == '__main__':
    raise SystemExit(main())
