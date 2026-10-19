import hashlib
import sys
from pathlib import Path
import unittest
from types import SimpleNamespace as N
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from powerpoint_com import Budget, PowerPointBridge, page, probe, read_shapes, read_slides, publish, matching_presentations

class Collection:
    def __init__(self, items): self.items=items; self.Count=len(items)
    def Item(self,i): return self.items[i-1]

def shape(text='', typ=17, table=None):
    tr=N(Length=len(text), Characters=lambda a,b:N(Text=text))
    return N(Id=7,Name='Shape',Type=typ,HasTextFrame=bool(text),
             TextFrame=N(TextRange=tr),HasTable=table is not None,Table=table)

def deck(shapes=()):
    return N(Slides=Collection([N(SlideID=42,Name='Slide',Shapes=Collection(shapes))]))

class Tests(unittest.TestCase):
    def test_rot_unique_window_binding(self):
        a=N(Windows=Collection([N(HWND=101)]))
        b=N(Windows=Collection([N(HWND=102)]))
        app=N(Presentations=Collection([a,b]))
        result=matching_presentations(app,200,9,lambda h:200 if h==101 else 300,lambda h:9,lambda w:w)
        self.assertIs(result[0],a)
        with self.assertRaises(ValueError):
            matching_presentations(app,200,9,lambda h:200,lambda h:9,lambda w:w)
        with self.assertRaises(ValueError):
            matching_presentations(app,200,8,lambda h:200,lambda h:9,lambda w:w)
    def test_rot_enumeration_budget(self):
        app=N(Presentations=Collection([N(Windows=Collection([N(HWND=i) for i in range(101)]))]))
        with self.assertRaises(ValueError):
            matching_presentations(app,200,9,lambda h:200,lambda h:9,lambda w:w)
    def test_bounds(self):
        for args in [(0,1,50),(1,51,50),(True,1,50)]:
            with self.assertRaises(ValueError): page(*args)
    def test_empty(self):
        d=N(Slides=Collection([]))
        self.assertEqual(probe(d)['slides'],[])
        self.assertEqual(read_slides(d)['slides'],[])
    def test_ids_hash(self):
        r=read_slides(deck([shape('中文')]))['slides'][0]
        self.assertEqual(r['slide_id'],42)
        self.assertEqual(r['shapes'][0]['shape_id'],7)
        self.assertEqual(r['shapes'][0]['text_sha256'],hashlib.sha256('中文'.encode()).hexdigest())
    def test_text_limit_before_materialization(self):
        s=shape('abc')
        s.TextFrame.TextRange.Characters=lambda a,b:self.fail('must not retrieve')
        with self.assertRaises(ValueError): Budget(2).text(s)
    def test_utf16_budget_rechecks_supplementary_characters(self):
        s=shape('placeholder')
        s.TextFrame.TextRange.Length=2
        s.TextFrame.TextRange.Characters=lambda a,b:N(Text='😀😀')
        with self.assertRaisesRegex(ValueError, 'UTF-16 budget'):
            Budget(2).text(s)
    def test_shape_pages(self):
        r=read_shapes(Collection([shape(),shape(),shape()]),2,1,Budget())
        self.assertEqual(r['next_offset'],3)
        self.assertEqual(r['shapes'][0]['index'],2)
    def test_table_budget(self):
        t=N(Rows=N(Count=1000),Columns=N(Count=1000))
        with self.assertRaises(ValueError): read_slides(deck([shape(table=t)]),tables=True)
        self.assertNotIn('table',read_slides(deck([shape(table=t)]))['slides'][0]['shapes'][0])
    def test_notes_budget_and_opt_in(self):
        d=deck([]); slide=d.Slides.Item(1)
        slide.NotesPage=N(Shapes=Collection([shape('too long')]))
        read_slides(d,text_budget=1)
        with self.assertRaises(ValueError): read_slides(d,notes=True,text_budget=1)
    def test_groups_no_recursion(self):
        r=read_slides(deck([shape(typ=6)]))['slides'][0]['shapes'][0]
        self.assertEqual(r['type'],6)
        self.assertNotIn('children',r)
    def test_shape_budget(self):
        b=Budget(); b.shapes=500
        with self.assertRaises(ValueError): read_shapes(Collection([shape()]),1,1,b)
    def test_explicit_target(self):
        b=PowerPointBridge.__new__(PowerPointBridge)
        for p,h in [(None,None),(1,0),(True,1)]:
            with self.assertRaises(ValueError): b.select(p,h)
    def test_window_class_pid_mismatch(self):
        b=PowerPointBridge.__new__(PowerPointBridge)
        b.label=lambda *a:'not PowerPoint'
        with self.assertRaises(ValueError): b.attach(123,5)
        b.label=lambda *a:'PPTFrameClass'; b.pid=lambda *a:6
        with self.assertRaises(ValueError): b.attach(123,5)
    def test_atomic_publication_failure_keeps_previous(self):
        import tempfile
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as td:
            out=Path(td)/'result.json'; out.write_text('previous')
            with patch('powerpoint_com.os.replace',side_effect=OSError('blocked')):
                with self.assertRaises(OSError): publish({'status':'ok'},out)
            self.assertEqual(out.read_text(),'previous')
            self.assertEqual(len(list(Path(td).iterdir())),1)

if __name__ == '__main__': unittest.main()
