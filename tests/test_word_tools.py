import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from word_tools import local_range, xml_summary, tracked_edit, find_literal


class HelperTests(unittest.TestCase):
    def test_find_rejects_backward_word_result(self):
        finder = SimpleNamespace(ClearFormatting=lambda: None,
                                 Replacement=SimpleNamespace(ClearFormatting=lambda: None),
                                 Execute=lambda *args: True)
        doc = SimpleNamespace(Content=SimpleNamespace(End=100),
                              Range=lambda a, b: SimpleNamespace(Find=finder, Start=3, End=5, Text='ab'))
        result = find_literal(doc, 'ab', start=10)
        self.assertEqual(result['matches'], [])
        self.assertEqual(result['stop_reason'], 'word_returned_out_of_scope')
        self.assertFalse(result['complete'])

    def test_parser_rejects_entities_and_excessive_depth(self):
        with self.assertRaises(ValueError):
            xml_summary('<!DOCTYPE body><body/>')
        with self.assertRaises(ValueError):
            xml_summary('<a>' * 130 + '</a>' * 130)

    def test_cleanup_does_not_mask_primary_error(self):
        def fail():
            raise RuntimeError('cleanup')
        undo = SimpleNamespace(IsRecordingCustomRecord=False,
                               StartCustomRecord=lambda label: None, EndCustomRecord=fail)
        doc = SimpleNamespace(ReadOnly=False, ProtectionType=-1, TrackRevisions=False,
                              Application=SimpleNamespace(UndoRecord=undo))
        with self.assertRaisesRegex(ValueError, 'primary') as error:
            with tracked_edit(doc):
                raise ValueError('primary')
        self.assertIn('cleanup', error.exception.__notes__[0])

    def test_native_bounds_reject_before_com(self):
        doc = SimpleNamespace(Content=SimpleNamespace(End=100),
                              Range=lambda *args: self.fail('COM range should not be called'))
        for start, end in [(-1, 1), (3, 2), (True, 4), (0, 101), (0, 21000)]:
            with self.assertRaises(ValueError):
                local_range(doc, start, end)

    def test_structural_hints_include_nested_and_simple_fields(self):
        xml = '''<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p><w:fldSimple w:instr="PAGE"/><w:r><w:instrText> REF target </w:instrText></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
        <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:body>'''
        report = xml_summary(xml)
        self.assertEqual(report['elements']['tbl'], 2)
        self.assertEqual(report['elements']['fldSimple'], 1)
        self.assertEqual(report['elements']['gridSpan'], 1)
        self.assertEqual(report['field_instruction_tokens'], {'REF': 1})

    def test_tracking_restored_on_partial_failure_without_undo(self):
        calls = []
        undo = SimpleNamespace(IsRecordingCustomRecord=False,
                               StartCustomRecord=lambda label: calls.append('start'),
                               EndCustomRecord=lambda: calls.append('end'))
        doc = SimpleNamespace(ReadOnly=False, ProtectionType=-1, TrackRevisions=False,
                              Application=SimpleNamespace(UndoRecord=undo))
        with self.assertRaisesRegex(RuntimeError, 'partial'):
            with tracked_edit(doc):
                self.assertTrue(doc.TrackRevisions)
                raise RuntimeError('partial')
        self.assertFalse(doc.TrackRevisions)
        self.assertEqual(calls, ['start', 'end'])


if __name__ == '__main__':
    unittest.main()
