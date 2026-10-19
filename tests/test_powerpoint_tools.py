"""Pure mocks: never instantiate a COM bridge or contact live Office."""
import copy
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace as N
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import powerpoint_tools as p


def units(text):
    return len(text.encode('utf-16-le')) // 2


class Collection:
    def __init__(self, items): self.items = items
    @property
    def Count(self): return len(self.items)
    def Item(self, index): return self.items[index - 1]


class Range:
    def __init__(self, text='', owner=None, start=1, length=None):
        self.owner = owner
        self._text = text
        self.Start = start
        self.span_length = length
        self.Font = N(Name='Arial', NameFarEast='等线', Size=20, Bold=0, Italic=0,
                      Underline=0, Color=N(RGB=123)) if owner is None else owner.Font
        self.ParagraphFormat = N(Alignment=1, Bullet=N(Visible=0))
        self.run_count = 1
        self.writes = 0
        self.replace_calls = []
        self.fail_after_write = False
    @property
    def Text(self):
        if self.owner is None: return self._text
        data = self.owner.Text.encode('utf-16-le')
        return data[(self.Start-1)*2:(self.Start-1+self.span_length)*2].decode('utf-16-le')
    @Text.setter
    def Text(self, value):
        root = self.owner or self
        root.writes += 1
        if self.owner is None:
            self._text = value
        else:
            data = root.Text.encode('utf-16-le')
            root._text = (data[:(self.Start-1)*2] + value.encode('utf-16-le') +
                          data[(self.Start-1+self.span_length)*2:]).decode('utf-16-le')
        if root.fail_after_write: raise RuntimeError('COM disconnected after mutation')
    @property
    def Length(self): return units(self.Text)
    def Characters(self, start, length): return Range(owner=self, start=start, length=length)
    def Runs(self, start=None, length=None): return N(Count=self.run_count) if start is None else self
    def Paragraphs(self): return N(Count=self.Text.count('\r')+1)
    def Replace(self, find, replacement, after, match_case, whole_words):
        self.replace_calls.append((find, replacement, after, match_case, whole_words))
        self.Text = self.Text.replace(find, replacement, 1)
        return self


def shape(text='Hello 世界', shape_id=7):
    return N(Id=shape_id, Type=17, HasTextFrame=-1, HasTable=0,
             Left=1, Top=2, Width=300, Height=70, Rotation=0,
             TextFrame=N(TextRange=Range(text)))


class Bridge:
    def __init__(self, presentation, target):
        self.presentation, self.target = presentation, target
        self.calls = 0
        self.on_select = None
    def select(self, pid, hwnd, name):
        self.calls += 1
        if self.on_select: self.on_select(self.calls)
        return self.presentation, dict(self.target)


class Tests(unittest.TestCase):
    def setUp(self):
        self.s = shape()
        self.slide = N(SlideID=42, Shapes=Collection([self.s]))
        self.deck = N(Slides=Collection([self.slide]), HasVBProject=False, Saved=False, ReadOnly=0)
        self.target = dict(application='powerpoint', pid=1, hwnd=2, full_name='deck.pptx',
                           name='deck.pptx', document_hwnd=3, binding='mock')
        self.bridge = Bridge(self.deck, self.target)
    def request(self, **op): return {'target': self.target, 'operation': op}
    def replacement(self, **extra):
        op = dict(kind='replace_text', slide_id=42, shape_id=7, expected=self.s.TextFrame.TextRange.Text,
                  text_sha256=p.text_hash(self.s.TextFrame.TextRange.Text), text='New', find='Hello')
        op.update(extra)
        return self.request(**op)
    def test_preview_default_and_native_replace(self):
        req = self.replacement()
        self.assertFalse(p.apply_request(self.bridge, req)['committed'])
        self.assertEqual(self.s.TextFrame.TextRange.writes, 0)
        result = p.apply_request(self.bridge, req, commit=True)
        self.assertTrue(result['committed'])
        self.assertEqual(self.s.TextFrame.TextRange.Text, 'New 世界')
        self.assertEqual(self.s.TextFrame.TextRange.replace_calls, [('Hello','New',0,-1,0)])
        self.assertFalse(self.deck.Saved)
    def test_native_span_after_supplementary_character(self):
        self.s.TextFrame.TextRange = Range('😀中文')
        req = self.replacement(strategy='characters', native_start=3, native_length=1, span_expected='中', text='好')
        p.apply_request(self.bridge, req, True)
        self.assertEqual(self.s.TextFrame.TextRange.Text, '😀好文')
    def test_bad_native_span_refused(self):
        for changes in [dict(native_start=0, native_length=1),dict(native_start=1,native_length=500),
                        dict(native_start=1,native_length=1,span_expected='wrong')]:
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                p.apply_request(self.bridge, self.replacement(strategy='characters', **changes), True)
        self.assertEqual(self.s.TextFrame.TextRange.writes, 0)
    def test_identity_changed_at_commit(self):
        self.bridge.on_select = lambda n: self.bridge.target.update(document_hwnd=99) if n == 2 else None
        with self.assertRaisesRegex(ValueError, 'identity changed'):
            p.apply_request(self.bridge, self.replacement(), True)
        self.assertEqual(self.s.TextFrame.TextRange.writes, 0)
    def test_local_change_at_commit(self):
        self.bridge.on_select = lambda n: setattr(self.s, 'Width', 301) if n == 2 else None
        with self.assertRaisesRegex(ValueError, 'Local evidence changed'):
            p.apply_request(self.bridge, self.replacement(), True)
        self.assertEqual(self.s.TextFrame.TextRange.writes, 0)
    def test_expected_and_hash_both_required(self):
        for key in ('expected','text_sha256'):
            req=self.replacement(); req['operation'][key]='stale'
            with self.assertRaises(ValueError): p.apply_request(self.bridge, req, True)
    def test_nonunique_including_overlapping_replace_refused(self):
        self.s.TextFrame.TextRange=Range('aaa')
        with self.assertRaisesRegex(ValueError,'unique'):
            p.apply_request(self.bridge,self.replacement(find='aa'),True)
    def test_mixed_whole_requires_explicit_strategy(self):
        self.s.TextFrame.TextRange.run_count=2
        with self.assertRaisesRegex(ValueError,'mixed formatting'):
            p.apply_request(self.bridge,self.replacement(strategy='whole'),True)
        p.apply_request(self.bridge,self.replacement(strategy='whole',preserve_strategy='first_character'),True)
        self.assertEqual(self.s.TextFrame.TextRange.Font.NameFarEast,'等线')
    def test_partial_write_never_retried_or_undone(self):
        rng=self.s.TextFrame.TextRange; rng.fail_after_write=True
        with self.assertRaisesRegex(p.WriteOutcomeUnknown,'No retry or rollback'):
            p.apply_request(self.bridge,self.replacement(),True)
        self.assertEqual(rng.writes,1)
        self.assertEqual(rng.Text,'New 世界')
    def test_stable_ids_and_bounded_resolution(self):
        self.deck.Slides.items.insert(0,N(SlideID=10))
        self.assertIs(p.resolve_slide(self.deck,42),self.slide)
        for value in (1, True, 0):
            with self.assertRaises(ValueError): p.resolve_slide(self.deck,value)
        self.deck.Slides=N(Count=501, Item=Mock(side_effect=AssertionError('must not enumerate')))
        with self.assertRaisesRegex(ValueError,'bounded'): p.resolve_slide(self.deck,42)
    def test_format_explicit_subset(self):
        req=self.replacement(kind='format_text',font={'italic':False,'bold':True,'name_far_east':'微软雅黑'},
                             geometry={'width':250},alignment=2)
        p.apply_request(self.bridge,req,True)
        self.assertEqual(self.s.Width,250)
        self.assertEqual(self.s.TextFrame.TextRange.Font.Bold,-1)
        self.assertEqual(self.s.TextFrame.TextRange.Font.Italic,0)
        self.assertEqual(self.s.TextFrame.TextRange.ParagraphFormat.Alignment,2)
    def test_add_editable_textbox_preview_and_commit(self):
        added=shape('',8)
        def add(orientation, left, top, width, height):
            added.Left, added.Top, added.Width, added.Height = left, top, width, height
            self.slide.Shapes.items.append(added)
            return added
        self.slide.Shapes.AddTextbox=Mock(side_effect=add)
        req=self.request(kind='add_textbox',slide_id=42,expected_shape_ids=[7],text='可编辑',
                         geometry=dict(left=20,top=30,width=100,height=50),
                         font=dict(name='Arial',name_far_east='等线',size=18,color=0x123456,bold=True))
        p.apply_request(self.bridge,req)
        self.slide.Shapes.AddTextbox.assert_not_called()
        p.apply_request(self.bridge,req,True)
        self.slide.Shapes.AddTextbox.assert_called_once_with(1,20,30,100,50)
        self.assertEqual(added.TextFrame.TextRange.Text,'可编辑')
        self.assertEqual(added.TextFrame.TextRange.ParagraphFormat.Bullet.Visible,0)
        self.assertEqual(added.TextFrame.AutoSize,0)
        self.assertEqual(added.TextFrame.WordWrap,-1)
    def test_delete_only_id_and_count_evidence(self):
        self.slide.Delete=Mock(side_effect=lambda:self.deck.Slides.items.clear())
        req=self.request(kind='delete_slide',slide_id=42,expected_slide_ids=[42])
        bad=copy.deepcopy(req); bad['operation']['expected_slide_ids']=[]
        with self.assertRaises(ValueError): p.apply_request(self.bridge,bad,True)
        self.slide.Delete.assert_not_called()
        self.assertEqual(p.apply_request(self.bridge,req,True)['data']['slide_ids'],[])
    def test_duplicate_uses_native_template_and_move(self):
        new=N(SlideID=99)
        def move(index):
            self.deck.Slides.items.remove(new)
            self.deck.Slides.items.insert(index - 1, new)
        new.MoveTo=Mock(side_effect=move)
        def duplicate():
            self.deck.Slides.items.append(new)
            return Collection([new])
        self.slide.Duplicate=Mock(side_effect=duplicate)
        req=self.request(kind='duplicate_slide',slide_id=42,expected_slide_ids=[42],move_to=1)
        self.assertEqual(p.apply_request(self.bridge,req,True)['data']['slide_id'],99)
        self.slide.Duplicate.assert_called_once_with()
        new.MoveTo.assert_called_once_with(1)
    def test_save_copy_unsaved_no_overwrite(self):
        with tempfile.TemporaryDirectory() as td:
            dest=Path(td)/'copy.pptx'
            self.deck.SaveCopyAs=Mock(side_effect=lambda path,fmt:Path(path).write_bytes(b'copy'))
            req=self.request(kind='save_copy',expected_slide_ids=[42],path=str(dest))
            p.apply_request(self.bridge,req)
            self.deck.SaveCopyAs.assert_not_called()
            p.apply_request(self.bridge,req,True)
            self.assertEqual(dest.read_bytes(),b'copy')
            self.assertFalse(self.deck.Saved)
            with self.assertRaisesRegex(ValueError,'overwrite'): p.apply_request(self.bridge,req,True)
            self.assertEqual(self.deck.SaveCopyAs.call_count,1)
    def test_export_and_concurrent_destination_no_clobber(self):
        with tempfile.TemporaryDirectory() as td:
            dest=Path(td)/'slide.png'
            def export(path,fmt,width,height):
                Path(path).write_bytes(b'new'); dest.write_bytes(b'concurrent')
            self.slide.Export=Mock(side_effect=export)
            req=self.request(kind='export_slide',slide_id=42,expected_slide_ids=[42],path=str(dest),width=640,height=360)
            with self.assertRaises(p.WriteOutcomeUnknown): p.apply_request(self.bridge,req,True)
            self.assertEqual(dest.read_bytes(),b'concurrent')
            self.assertEqual(self.slide.Export.call_count,1)
    def test_cli_commit_timeout_is_unknown_without_retry(self):
        with patch.object(p.subprocess,'run',side_effect=subprocess.TimeoutExpired('worker',1)) as run, patch('sys.stderr',new_callable=io.StringIO) as err:
            self.assertEqual(p.main(['apply','--request','unused.json','--commit']),2)
            self.assertEqual(json.loads(err.getvalue())['write_outcome'],'unknown')
            self.assertEqual(run.call_count,1)
    def test_paragraph_break_conversion_is_unknown(self):
        class PlaceholderRange(Range):
            @Range.Text.setter
            def Text(self, value):
                Range.Text.fset(self, value.replace('\r', '\x0b'))
        rng = PlaceholderRange('Hello 世界')
        self.s.TextFrame.TextRange = rng
        req = self.replacement(strategy='whole', preserve_strategy='first_character', text='Heading\rDetail')
        with self.assertRaisesRegex(p.WriteOutcomeUnknown, 'postcondition failed'):
            p.apply_request(self.bridge, req, True)
        self.assertEqual(rng.Text, 'Heading\x0bDetail')
        self.assertEqual(rng.writes, 1)
    def test_preserve_strategy_rejected_outside_whole_without_writes(self):
        rng = self.s.TextFrame.TextRange
        rng.run_count = 2
        before = copy.deepcopy(vars(rng.Font))
        for strategy in ('characters', 'native_replace'):
            for preserve in ('first_character', None):
                req = self.replacement(strategy=strategy, preserve_strategy=preserve,
                                       native_start=1, native_length=1, span_expected='H')
                with self.subTest(strategy=strategy,preserve=preserve), self.assertRaisesRegex(ValueError,'only for whole'):
                    p.apply_request(self.bridge, req, True)
        self.assertEqual(rng.writes, 0)
        self.assertEqual(rng.replace_calls, [])
        self.assertEqual(vars(rng.Font), before)
        self.assertEqual(rng.Text, 'Hello 世界')
    def test_readonly_and_unknown_state_fail_closed_before_plan(self):
        for value in (-1, True, -2, 'false', None):
            self.deck.ReadOnly=value
            with self.subTest(value=value), patch.object(p,'read_shape',side_effect=AssertionError('no plan')):
                with self.assertRaisesRegex(ValueError,'mutation refused'):
                    p.apply_request(self.bridge,self.replacement(),True)
        del self.deck.ReadOnly
        with self.assertRaisesRegex(ValueError,'Cannot establish'):
            p.apply_request(self.bridge,self.replacement(),True)
        self.assertEqual(self.s.TextFrame.TextRange.writes,0)
    def test_readonly_change_at_commit_is_refused(self):
        self.bridge.on_select=lambda n:setattr(self.deck,'ReadOnly',-1) if n == 2 else None
        with self.assertRaisesRegex(ValueError,'Read-only'):
            p.apply_request(self.bridge,self.replacement(),True)
        self.assertEqual(self.s.TextFrame.TextRange.writes,0)
    def test_readonly_allows_separate_copy_and_export(self):
        self.deck.ReadOnly=-1
        with tempfile.TemporaryDirectory() as td:
            for op in (dict(kind='save_copy',path=str(Path(td)/'copy.pptx')),
                       dict(kind='export_slide',slide_id=42,path=str(Path(td)/'slide.png'),width=640,height=360)):
                self.assertFalse(p.apply_request(self.bridge,self.request(expected_slide_ids=[42],**op))['committed'])
    def test_cli_reconfigures_utf8_and_default_timeout(self):
        out, err = Mock(), Mock()
        with patch('sys.stdout',out), patch('sys.stderr',err), patch.object(p.subprocess,'run',return_value=N(returncode=0)) as run:
            self.assertEqual(p.main(['apply','--request','unused.json']),0)
        out.reconfigure.assert_called_once_with(encoding='utf-8')
        err.reconfigure.assert_called_once_with(encoding='utf-8')
        self.assertEqual(run.call_args.kwargs['timeout'],9)
    def test_native_replace_silent_noop_detected(self):
        self.s.TextFrame.TextRange.Replace=Mock(return_value=self.s.TextFrame.TextRange)
        with self.assertRaisesRegex(p.WriteOutcomeUnknown,'postcondition failed'):
            p.apply_request(self.bridge,self.replacement(),True)
        self.s.TextFrame.TextRange.Replace.assert_called_once()
    def test_postwrite_identity_change_is_unknown(self):
        self.bridge.on_select=lambda n:self.bridge.target.update(hwnd=77) if n == 3 else None
        with self.assertRaises(p.WriteOutcomeUnknown):
            p.apply_request(self.bridge,self.replacement(),True)
        self.assertEqual(self.s.TextFrame.TextRange.writes,1)
    def test_mixed_whole_range_cannot_bypass_guard(self):
        self.s.TextFrame.TextRange.run_count=2
        for extra in (dict(find='Hello 世界'), dict(strategy='characters',native_start=1,
                      native_length=units('Hello 世界'),span_expected='Hello 世界')):
            with self.subTest(extra=extra), self.assertRaisesRegex(ValueError,'Whole-range mixed'):
                p.apply_request(self.bridge,self.replacement(**extra),True)
    def test_duplicate_silent_move_failure_is_unknown(self):
        new=N(SlideID=99,MoveTo=Mock())
        def duplicate():
            self.deck.Slides.items.append(new)
            return Collection([new])
        self.slide.Duplicate=Mock(side_effect=duplicate)
        req=self.request(kind='duplicate_slide',slide_id=42,expected_slide_ids=[42],move_to=1)
        with self.assertRaisesRegex(p.WriteOutcomeUnknown,'order/count'):
            p.apply_request(self.bridge,req,True)
        self.slide.Duplicate.assert_called_once()
        self.assertEqual(self.deck.Slides.Count,2)
    def test_unsupported_complex_shape_refused(self):
        self.s.Type=6
        with self.assertRaisesRegex(ValueError,'simple top-level'): p.read_shape(self.deck,42,7)
    def test_read_evidence_native_coordinates_and_geometry(self):
        self.s.TextFrame.TextRange=Range('😀中')
        data=p.read_shape(self.deck,42,7)
        self.assertEqual(data['native_length'],3)
        self.assertEqual(data['runs'][0]['native_start'],1)
        self.assertEqual(data['geometry']['width'],300)
        self.assertEqual(data['font']['name_far_east'],'等线')


if __name__ == '__main__': unittest.main()
