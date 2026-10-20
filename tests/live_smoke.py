"""Opt-in integration test. Creates a private Word instance, never edits user docs."""
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from word_com import WordBridge, snapshot, find_text, apply_request
from comtypes.client import CreateObject

if '--run' not in sys.argv:
    raise SystemExit('Explicit --run required (creates temporary visible Word instance)')
app = CreateObject('Word.Application', dynamic=True)
app.Visible = True
doc = None
results = []
try:
    doc = app.Documents.Add()
    doc.Content.Text = 'Alpha 中文😀 Beta\rSecond paragraph\r'
    bridge = WordBridge()
    hwnd = int(doc.Windows.Item(1).Hwnd)
    pid = bridge.pid(hwnd)
    attached, identity = bridge.select(pid, hwnd)
    assert attached.Name == doc.Name
    results.append('attach_pid_hwnd')
    # A target that no longer resolves must be refused before any write is attempted.
    missing = dict(identity, hwnd=identity['hwnd'] + 1)
    unresolved = {'target': missing, 'text_sha256': snapshot(attached, identity)['text_sha256'],
                  'operation': {'kind': 'insert', 'start': 0, 'end': 0, 'expected': '', 'text': 'x'}}
    before_unresolved = attached.Content.Text
    try:
        apply_request(attached, missing, unresolved, True, bridge=bridge)
    except ValueError as exc:
        assert '未尝试写入' in str(exc), exc
    else:
        raise AssertionError('unresolvable native target was not refused')
    assert attached.Content.Text == before_unresolved
    results.append('unresolvable_native_target_refused_before_write')
    def request(op):
        snap = snapshot(attached, identity)
        return {'target': identity, 'text_sha256': snap['text_sha256'], 'operation': op}
    hit = find_text(attached, 'Beta')['matches'][0]
    assert attached.Range(hit['start'], hit['end']).Text == 'Beta'
    results.append('find_after_non_bmp_unicode')
    req = request({'kind': 'replace', 'start': hit['start'], 'end': hit['end'], 'expected': 'Beta', 'text': 'Gamma'})
    before = attached.Content.Text
    assert not apply_request(attached, identity, req, bridge=bridge)['committed']
    assert attached.Content.Text == before
    results.append('dry_run_no_mutation')
    assert apply_request(attached, identity, req, True, bridge=bridge)['committed']
    assert find_text(attached, 'Gamma')['matches']
    assert not bool(attached.TrackRevisions)
    assert attached.Revisions.Count > 0
    results.append('tracked_replace_restore_setting')
    attached.Undo()
    assert attached.Content.Text == before
    results.append('custom_undo')
    stale = request({'kind': 'insert', 'start': 0, 'end': 0, 'expected': '', 'text': 'prefix '})
    attached.Range(0, 0).Text = 'X'
    try:
        apply_request(attached, identity, stale, True, bridge=bridge)
        raise AssertionError('stale request accepted')
    except ValueError:
        results.append('stale_rejected')
    hit = find_text(attached, 'Alpha')['matches'][0]
    apply_request(attached, identity, request({'kind': 'comment', 'start': hit['start'], 'end': hit['end'], 'expected': 'Alpha', 'text': '测试批注'}), True, bridge=bridge)
    assert attached.Comments.Count == 1
    results.append('comment')
    apply_request(attached, identity, request({'kind': 'format', 'start': hit['start'], 'end': hit['end'], 'expected': 'Alpha', 'font': {'bold': True, 'size': 14}}), True, bridge=bridge)
    assert attached.Range(hit['start'], hit['end']).Font.Bold != 0
    results.append('font_format')
    apply_request(attached, identity, request({'kind': 'insert', 'start': 0, 'end': 0, 'expected': '', 'text': 'PREFIX '}), True, bridge=bridge)
    assert find_text(attached, 'PREFIX')['matches']
    results.append('tracked_insert')
    # Supervised CLI path: proves the CLI passes a bridge, not only the import API.
    prefix = find_text(attached, 'PREFIX')['matches'][0]
    (ROOT / 'tmp').mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='live-smoke-', dir=ROOT / 'tmp') as folder:
        request_file = Path(folder) / 'request.json'
        request_file.write_text(json.dumps(request({
            'kind': 'format', 'start': prefix['start'], 'end': prefix['end'],
            'expected': 'PREFIX', 'font': {'italic': True}}), ensure_ascii=False), encoding='utf-8')
        command = [sys.executable, str(ROOT / 'scripts' / 'word_com.py'),
                   'apply', '--request', str(request_file)]
        started = time.perf_counter()
        preview = subprocess.run(command, capture_output=True, text=True, encoding='utf-8', timeout=15)
        preview_seconds = time.perf_counter() - started
        assert preview.returncode == 0, preview.stderr or preview.stdout
        assert json.loads(preview.stdout)['committed'] is False
        started = time.perf_counter()
        commit = subprocess.run(command + ['--commit'], capture_output=True, text=True,
                                encoding='utf-8', timeout=15)
        commit_seconds = time.perf_counter() - started
        assert commit.returncode == 0, commit.stderr or commit.stdout
        assert json.loads(commit.stdout)['committed'] is True
        assert attached.Range(prefix['start'], prefix['end']).Font.Italic != 0
        assert preview_seconds < 10 and commit_seconds < 10, (preview_seconds, commit_seconds)
    results.append('supervised_cli_apply_preview_and_commit_under_10_seconds')
    report = {'passed': results, 'cli_preview_seconds': round(preview_seconds, 3),
              'cli_commit_seconds': round(commit_seconds, 3),
              'private_test_target': identity, 'user_documents_modified': False, 'saved_files': False}
    out = Path(__file__).resolve().parents[1] / 'artifacts' / 'live-smoke.json'
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
finally:
    if doc is not None:
        doc.Close(0)
    # Application ownership is not implied by CreateObject; never Quit Office.
