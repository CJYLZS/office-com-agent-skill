"""Opt-in helper tests in a private unsaved Word instance."""
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from word_tools import *
from comtypes.client import CreateObject

if '--run' not in sys.argv:
    raise SystemExit('Requires --run; creates private Word instance')
app = CreateObject('Word.Application', dynamic=True)
app.Visible = True
doc = None
passed = []
try:
    doc = app.Documents.Add()
    doc.Content.Text = 'Repeated title\rEmoji 😀 target\rRepeated title\rEnd\r'
    doc.Paragraphs.Item(3).Style = doc.Styles.Item(-4)
    hits = find_literal(doc, 'Repeated title', style=-4)['matches']
    assert len(hits) == 1 and hits[0]['start'] == int(doc.Paragraphs.Item(3).Range.Start)
    passed.append('style_filter_excludes_plain_duplicate')
    hit = find_literal(doc, 'target')['matches'][0]
    assert read_range(doc, hit['start'], hit['end'])['text'] == 'target'
    assert not find_literal(doc, 'target', end=hit['start'])['matches']
    passed.append('native_unicode_coordinates_and_bounds')
    field = doc.Fields.Add(doc.Range(doc.Content.End-1, doc.Content.End-1), -1, 'QUOTE "FIELD_VALUE"', False)
    a, b = int(field.Result.Start), int(field.Result.End)
    risk = range_risks(doc, a+1, b-1)
    assert any(risk['endpoint_in_field_result']), risk
    risk = range_risks(doc, a+1, a+1)
    assert any(risk['endpoint_in_field_result']), risk
    passed.append('field_result_and_insertion_endpoint_detection')
    table = doc.Tables.Add(doc.Range(doc.Content.End-1, doc.Content.End-1), 2, 2)
    table.Cell(1, 1).Merge(table.Cell(1, 2))
    cell = table.Cell(1, 1)
    assert cell_text_range(cell).Text in (None, '')
    with tracked_edit(doc, 'Private helper test'):
        cell_text_range(cell).Text = 'Merged cell'
    assert not bool(doc.TrackRevisions)
    assert str(cell.Range.Text).endswith('\r\x07')
    assert 'Merged cell' in str(cell.Range.Text), repr((cell.Range.Text, doc.Content.Text))
    assert int(table.Rows.Count) == 2
    assert_expected(doc.Range(hit['start'], hit['end']), 'target')
    passed.append('merged_empty_cell_tracking_and_terminator_preservation')
    report = inspect_range(doc, int(table.Range.Start), int(table.Range.End))
    assert report['structure']['elements']['tbl'] >= 1
    passed.append('local_xml_structure_and_risks')
    print(json.dumps({'passed': passed, 'user_documents_modified': False}, indent=2))
finally:
    if doc is not None:
        doc.Close(0)
    app.Quit(0)
