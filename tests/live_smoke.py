"""Opt-in integration test. Creates a private Word instance, never edits user docs."""
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
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
    def request(op):
        snap = snapshot(attached, identity)
        return {'target': identity, 'text_sha256': snap['text_sha256'], 'operation': op}
    hit = find_text(attached, 'Beta')['matches'][0]
    assert attached.Range(hit['start'], hit['end']).Text == 'Beta'
    results.append('find_after_non_bmp_unicode')
    req = request({'kind': 'replace', 'start': hit['start'], 'end': hit['end'], 'expected': 'Beta', 'text': 'Gamma'})
    before = attached.Content.Text
    assert not apply_request(attached, identity, req)['committed']
    assert attached.Content.Text == before
    results.append('dry_run_no_mutation')
    assert apply_request(attached, identity, req, True)['committed']
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
        apply_request(attached, identity, stale, True)
        raise AssertionError('stale request accepted')
    except ValueError:
        results.append('stale_rejected')
    hit = find_text(attached, 'Alpha')['matches'][0]
    apply_request(attached, identity, request({'kind': 'comment', 'start': hit['start'], 'end': hit['end'], 'expected': 'Alpha', 'text': '测试批注'}), True)
    assert attached.Comments.Count == 1
    results.append('comment')
    apply_request(attached, identity, request({'kind': 'format', 'start': hit['start'], 'end': hit['end'], 'expected': 'Alpha', 'font': {'bold': True, 'size': 14}}), True)
    assert attached.Range(hit['start'], hit['end']).Font.Bold != 0
    results.append('font_format')
    apply_request(attached, identity, request({'kind': 'insert', 'start': 0, 'end': 0, 'expected': '', 'text': 'PREFIX '}), True)
    assert find_text(attached, 'PREFIX')['matches']
    results.append('tracked_insert')
    report = {'passed': results, 'private_test_target': identity, 'user_documents_modified': False, 'saved_files': False}
    out = Path(__file__).resolve().parents[1] / 'artifacts' / 'live-smoke.json'
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
finally:
    if doc is not None:
        doc.Close(0)
    # Application ownership is not implied by CreateObject; never Quit Office.
