"""Shared apply_request regressions; pure mocks, never connect to live Office.

Compose existing mock objects, not TestCases (no duplicated inherited suites).
Only Excel's native VARIANT transport is patched; actual planning, evidence,
structural guards, writes and verification execute. Word resolves the target
through a bridge stub like the other apps and reports partial writes through
WriteOutcomeUnknown. Native window ambiguity, Protected View, dialogs, real
timeouts and same-name multi-window binding still need live tests.
"""
import copy
from pathlib import Path
import sys
from types import SimpleNamespace as N
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import word_com as w
import excel_tools as x
import powerpoint_tools as p
import test_excel_tools as excel_mocks
import test_powerpoint_tools as powerpoint_mocks


class WordRange:
    def __init__(self, doc, start, end):
        self.doc, self.Start, self.End = doc, start, end
        self.Tables = self.InlineShapes = N(Count=0)

    def Information(self, unused):
        return False

    @property
    def Text(self):
        return self.doc.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        self.doc.writes.append(value)
        self.doc.text = self.doc.text[:self.Start] + value + self.doc.text[self.End:]
        if self.doc.fail_after_write:
            raise RuntimeError("disconnected after mutation")


class WordDocument:
    def __init__(self):
        self.text = "Hello world\r"
        self.writes = []
        self.fail_after_write = False
        self.ReadOnly = False
        self.ProtectionType = -1
        self.TrackRevisions = False
        self.ContentControls = self.Fields = N(Count=0)
        self.on_start = None
        self.undo = N(IsRecordingCustomRecord=False,
                      StartCustomRecord=Mock(side_effect=self.start_record),
                      EndCustomRecord=Mock(), Undo=Mock())
        self.Application = N(UndoRecord=self.undo, Quit=Mock(), Undo=Mock())
        self.Save, self.Close, self.Undo = Mock(), Mock(), Mock()

    def start_record(self, label):
        if self.on_start:
            self.on_start()

    @property
    def Content(self):
        return N(Text=self.text, End=len(self.text))

    def Range(self, start, end):
        return WordRange(self, start, end)


class WordBridgeMock:
    """Native window resolution stub; select() must be called again per rewrite."""

    def __init__(self, doc, identity):
        self.doc, self.identity = doc, dict(identity)
        self.calls = 0
        self.on_select = None
        self.fail_on = None

    def select(self, pid=None, hwnd=None, document=None):
        self.calls += 1
        if self.fail_on == self.calls:
            raise RuntimeError("no native window for this target")
        if self.on_select:
            self.on_select(self.calls)
        return self.doc, dict(self.identity)


class WordFixture:
    name = "word"
    unknown = w.WriteOutcomeUnknown
    unknown_message = "可能有部分修改"

    def __init__(self):
        self.doc = WordDocument()
        self.target = dict(pid=11, hwnd=22, name="owned.docx", full_name="owned.docx")
        self.identity = dict(self.target)
        self.bridge = WordBridgeMock(self.doc, self.identity)
        self.req = dict(target=dict(self.target), text_sha256=w.text_hash(self.doc.text),
                        operation=dict(kind="replace", start=0, end=5,
                                       expected="Hello", text="New"))

    def apply(self, commit=False):
        return w.apply_request(self.doc, self.identity, self.req, commit=commit,
                               bridge=None if self.bridge is None else self.bridge)

    def change_evidence(self):
        self.doc.text = "Other world\r"

    def at_commit(self, callback):
        self.doc.on_start = callback

    def at_rebind(self, callback):
        self.bridge.on_select = lambda n: callback() if n == 2 else None

    def native_identity(self):
        return self.bridge.identity

    def fail_write(self):
        self.doc.fail_after_write = True

    def writes(self):
        return self.doc.writes

    def content(self):
        return self.doc.text

    def assert_no_lifecycle(self, case):
        for method in (self.doc.Save, self.doc.Close, self.doc.Undo,
                       self.doc.Application.Quit, self.doc.Application.Undo, self.doc.undo.Undo):
            method.assert_not_called()
        case.assertFalse(self.doc.TrackRevisions)


class ExcelFixture:
    name = "excel"
    unknown = x.WriteOutcomeUnknown
    unknown_message = "No retry or rollback"

    def __init__(self):
        self.journal = []
        self.sheet = excel_mocks.Sheet(self.journal)
        self.spec = dict(name="Data", code_name="Sheet1")
        self.target = dict(application="excel", pid=11, hwnd=22,
                           name="owned.xlsx", full_name="owned.xlsx")
        self.doc = N(Worksheets=N(Item=lambda name: self.sheet),
                     Application=N(Version="16.0", Quit=Mock()),
                     ReadOnly=False, ProtectStructure=False, HasVBProject=False,
                     MultiUserEditing=False, Saved=False, FileFormat=51,
                     Save=Mock(), SaveCopyAs=Mock(), Close=Mock(), Undo=Mock())
        self.bridge = excel_mocks.Bridge(self.doc, dict(self.target))
        self.req = dict(target=dict(self.target), operation=dict(
            kind="set_values", sheet=self.spec, address="B2:C2", values=[[12, 34]],
            expected_sha256="0" * 64))

    def snapshot(self, wb, sheet, address, **kwargs):
        sh = wb.Worksheets.Item(sheet)
        rng = sh.Range(address)
        cells = [[rng.item(r, c) for c in range(1, rng.Columns.Count + 1)]
                 for r in range(1, rng.Rows.Count + 1)]
        return dict(sheet=dict(name=sh.Name, code_name=sh.CodeName),
                    address=x.parse_address(address)["address"],
                    data=dict(values=[[c.Value2 for c in row] for row in cells],
                              formulas_or_constants=[[c.Formula for c in row] for row in cells]),
                    types=dict(values=[[dict(kind="empty" if c.Value2 is None else "number")
                                        for c in row] for row in cells]))

    def prepare(self):
        self.req["operation"]["expected_sha256"] = x.read_edit_range(
            self.doc, self.spec, "B2:C2")["sha256"]
        self.journal.clear()

    def apply(self, commit=False):
        return x.apply_request(self.bridge, self.req, commit=commit)

    def change_evidence(self):
        self.sheet.cells[2, 2].__dict__["Value2"] = 99

    def at_commit(self, callback):
        self.bridge.on_select = lambda n: callback() if n == 2 else None

    def at_rebind(self, callback):
        self.bridge.on_select = lambda n: callback() if n == 2 else None

    def native_identity(self):
        return self.bridge.identity

    def fail_write(self):
        self.sheet.cells[2, 2].__dict__["fail_after"] = True

    def writes(self):
        return self.journal

    def content(self):
        return [self.sheet.cells[2, c].Value2 for c in (2, 3)]

    def assert_no_lifecycle(self, case):
        for method in (self.doc.Save, self.doc.SaveCopyAs, self.doc.Close,
                       self.doc.Undo, self.doc.Application.Quit):
            method.assert_not_called()
        case.assertFalse(self.doc.Saved)


class PowerPointFixture:
    name = "powerpoint"
    unknown = p.WriteOutcomeUnknown
    unknown_message = "No retry or rollback"

    def __init__(self):
        self.shape = powerpoint_mocks.shape("Hello world")
        self.slide = N(SlideID=42, Shapes=powerpoint_mocks.Collection([self.shape]))
        self.doc = N(Slides=powerpoint_mocks.Collection([self.slide]),
                     HasVBProject=False, Saved=False, ReadOnly=0,
                     Save=Mock(), SaveCopyAs=Mock(), Close=Mock(), Undo=Mock(),
                     Application=N(Quit=Mock()))
        self.target = dict(application="powerpoint", pid=11, hwnd=22,
                           full_name="owned.pptx", name="owned.pptx",
                           document_hwnd=33, binding="mock")
        self.bridge = powerpoint_mocks.Bridge(self.doc, dict(self.target))
        self.req = dict(target=dict(self.target), operation=dict(
            kind="replace_text", slide_id=42, shape_id=7, expected="Hello world",
            text_sha256=p.text_hash("Hello world"), text="New", find="Hello"))

    def apply(self, commit=False):
        return p.apply_request(self.bridge, self.req, commit=commit)

    def change_evidence(self):
        self.shape.TextFrame.TextRange._text = "Other world"

    def at_commit(self, callback):
        self.bridge.on_select = lambda n: callback() if n == 2 else None

    def at_rebind(self, callback):
        self.bridge.on_select = lambda n: callback() if n == 2 else None

    def native_identity(self):
        return self.bridge.target

    def fail_write(self):
        self.shape.TextFrame.TextRange.fail_after_write = True

    def writes(self):
        return self.shape.TextFrame.TextRange.replace_calls

    def content(self):
        return self.shape.TextFrame.TextRange.Text

    def assert_no_lifecycle(self, case):
        for method in (self.doc.Save, self.doc.SaveCopyAs, self.doc.Close,
                       self.doc.Undo, self.doc.Application.Quit):
            method.assert_not_called()
        case.assertFalse(self.doc.Saved)


class OfficeSafetyContractTests(unittest.TestCase):
    def fixture(self, cls):
        fixture = cls()
        if isinstance(fixture, ExcelFixture):
            reader = patch.object(x, "read_range", side_effect=fixture.snapshot)
            reader.start()
            self.addCleanup(reader.stop)
            fixture.prepare()
        return fixture

    def test_default_preview_has_no_writes_or_lifecycle_calls(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                before = copy.deepcopy(f.req), copy.deepcopy(f.content())
                result = f.apply()
                self.assertFalse(result["committed"])
                self.assertFalse(result.get("saved_in_place", result.get("saved_by_script")))
                self.assertEqual(f.writes(), [])
                self.assertEqual((f.req, f.content()), before)
                if cls is WordFixture:
                    f.doc.undo.StartCustomRecord.assert_not_called()
                    f.doc.undo.EndCustomRecord.assert_not_called()
                self.assertEqual(f.bridge.calls, 1)
                f.assert_no_lifecycle(self)

    def test_explicit_commit_positive_control_mutates_without_saving(self):
        for cls, expected in ((WordFixture, "New world\r"), (ExcelFixture, [12, 34]),
                              (PowerPointFixture, "New world")):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                self.assertTrue(f.apply(commit=True)["committed"])
                self.assertEqual(f.content(), expected)
                self.assertEqual(len(f.writes()), 2 if cls is ExcelFixture else 1)
                f.assert_no_lifecycle(self)

    def test_stale_initial_evidence_rejected_without_writes(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.change_evidence()
                with self.assertRaises(ValueError):
                    f.apply(commit=True)
                self.assertEqual(f.writes(), [])
                f.assert_no_lifecycle(self)

    def test_evidence_change_at_last_guard_rejected_without_content_write(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.at_commit(f.change_evidence)
                # Word raises this inside its undo scope but before any mutation, so it
                # must stay a refusal rather than being reported as an unknown write.
                with self.assertRaises(ValueError) as error:
                    f.apply(commit=True)
                self.assertNotIsInstance(error.exception, w.WriteOutcomeUnknown)
                self.assertEqual(f.writes(), [])
                f.assert_no_lifecycle(self)
                if cls is WordFixture:
                    f.doc.undo.StartCustomRecord.assert_called_once()
                    f.doc.undo.EndCustomRecord.assert_called_once()

    def test_supplied_identity_mismatch_rejected_without_writes(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            for field in ("pid", "hwnd", "name", "full_name"):
                with self.subTest(app=cls.name, field=field):
                    f = self.fixture(cls)
                    value = f.req["target"][field]
                    f.req["target"][field] = value + 1 if type(value) is int else value + ".other"
                    with self.assertRaises(ValueError):
                        f.apply(commit=True)
                    self.assertEqual(f.writes(), [])
                    f.assert_no_lifecycle(self)

    def test_native_identity_rechecked_on_second_binding(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            for field in cls().target:
                with self.subTest(app=cls.name, field=field):
                    f = self.fixture(cls)
                    identity = f.native_identity()
                    value = identity[field]
                    f.at_rebind(lambda: identity.update({field: value + 1 if type(value) is int else value + ".other"}))
                    with self.assertRaises(ValueError):
                        f.apply(commit=True)
                    self.assertEqual(f.bridge.calls, 2)
                    self.assertEqual(f.writes(), [])
                    f.assert_no_lifecycle(self)

    def test_commit_uses_freshly_bound_object_not_preview_object(self):
        for cls, expected in ((WordFixture, "New world\r"), (ExcelFixture, [12, 34]),
                              (PowerPointFixture, "New world")):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                replacement = cls()
                # Same identity and identical local evidence, different proxy/object.
                # The commit must use the second select's object, not a cached plan.
                def replace_object():
                    if cls is ExcelFixture:
                        f.bridge.wb = replacement.doc
                    elif cls is WordFixture:
                        f.bridge.doc = replacement.doc
                    else:
                        f.bridge.presentation = replacement.doc
                f.at_rebind(replace_object)
                self.assertTrue(f.apply(commit=True)["committed"])
                self.assertEqual(f.writes(), [])
                self.assertEqual(replacement.content(), expected)
                self.assertEqual(len(replacement.writes()), 2 if cls is ExcelFixture else 1)
                f.assert_no_lifecycle(self)
                replacement.assert_no_lifecycle(self)

    def test_read_only_rejected_without_writes_in_preview_and_commit(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            for commit in (False, True):
                with self.subTest(app=cls.name, commit=commit):
                    f = self.fixture(cls)
                    f.doc.ReadOnly = -1 if cls is PowerPointFixture else True
                    with self.assertRaises(ValueError):
                        f.apply(commit=commit)
                    self.assertEqual(f.writes(), [])
                    f.assert_no_lifecycle(self)

    def test_partial_write_warns_unknown_and_is_not_retried_or_undone(self):
        for cls, expected in ((WordFixture, "New world\r"), (ExcelFixture, [12, None]),
                              (PowerPointFixture, "New world")):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.fail_write()
                with self.assertRaisesRegex(f.unknown, f.unknown_message) as error:
                    f.apply(commit=True)
                self.assertIsInstance(error.exception.__cause__, RuntimeError)
                self.assertEqual(len(f.writes()), 1)
                self.assertEqual(f.content(), expected)
                f.assert_no_lifecycle(self)
                if cls is WordFixture:
                    f.doc.undo.StartCustomRecord.assert_called_once()
                    f.doc.undo.EndCustomRecord.assert_called_once()

    def test_word_protection_rejected_before_undo_scope(self):
        for protection in (0, 1, 2, 3):
            with self.subTest(protection=protection):
                f = self.fixture(WordFixture)
                f.doc.ProtectionType = protection
                with self.assertRaises(ValueError):
                    f.apply(commit=True)
                self.assertEqual(f.writes(), [])
                f.doc.undo.StartCustomRecord.assert_not_called()
                f.assert_no_lifecycle(self)

    def test_excel_workbook_and_sheet_protection_rechecked_before_commit(self):
        for scope, field in (("doc", "ProtectStructure"), ("sheet", "ProtectContents"),
                             ("sheet", "ProtectDrawingObjects"), ("sheet", "ProtectScenarios")):
            with self.subTest(scope=scope, field=field):
                f = self.fixture(ExcelFixture)
                f.at_commit(lambda: setattr(getattr(f, scope), field, True))
                with self.assertRaises(ValueError):
                    f.apply(commit=True)
                self.assertEqual(f.writes(), [])
                self.assertEqual(f.bridge.calls, 2)
                f.assert_no_lifecycle(self)

    def test_excel_and_powerpoint_read_only_change_at_rebind_is_rejected(self):
        for cls in (ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.at_commit(lambda: setattr(f.doc, "ReadOnly", True))
                with self.assertRaises(ValueError):
                    f.apply(commit=True)
                self.assertEqual(f.writes(), [])
                f.assert_no_lifecycle(self)

    def test_postwrite_identity_loss_is_unknown(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                identity = f.native_identity()
                f.bridge.on_select = lambda n: identity.update(hwnd=999) if n == 3 else None
                with self.assertRaises(f.unknown):
                    f.apply(commit=True)
                self.assertEqual(f.bridge.calls, 3)
                self.assertEqual(len(f.writes()), 2 if cls is ExcelFixture else 1)
                f.assert_no_lifecycle(self)

    def test_postwrite_rebind_failure_never_claims_no_write(self):
        def lost(call):
            if call == 3:
                raise RuntimeError("no native window after the write")

        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.bridge.on_select = lost
                with self.assertRaises(f.unknown) as error:
                    f.apply(commit=True)
                text = str(error.exception)
                # The write was already issued: this must never read as a refusal.
                self.assertNotIn("未尝试写入", text)
                self.assertTrue("部分" in text or "partial" in text, text)
                self.assertIsInstance(error.exception.__cause__, RuntimeError)
                self.assertEqual(len(f.writes()), 2 if cls is ExcelFixture else 1)
                f.assert_no_lifecycle(self)

    def test_postwrite_identity_change_message_differs_from_refusal(self):
        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                identity = f.native_identity()
                f.bridge.on_select = lambda n: identity.update(hwnd=999) if n == 3 else None
                with self.assertRaises(f.unknown) as error:
                    f.apply(commit=True)
                self.assertNotIn("未尝试写入", str(error.exception))
                self.assertEqual(f.bridge.calls, 3)

    def test_word_rebind_failure_is_prewrite_refusal_not_unknown(self):
        f = self.fixture(WordFixture)
        f.bridge.fail_on = 2
        with self.assertRaises(ValueError) as error:
            f.apply(commit=True)
        self.assertNotIsInstance(error.exception, w.WriteOutcomeUnknown)
        self.assertIn("未尝试写入", str(error.exception))
        self.assertIsInstance(error.exception.__cause__, RuntimeError)
        self.assertEqual(f.writes(), [])
        f.doc.undo.StartCustomRecord.assert_not_called()
        with self.assertRaises(ValueError):
            w.apply_request(f.doc, "not-a-dict", f.req, True, bridge=f.bridge)
        self.assertEqual(f.writes(), [])
        f.assert_no_lifecycle(self)

    def test_one_unknown_outcome_type_is_shared_by_every_app(self):
        import office_com

        self.assertIs(w.WriteOutcomeUnknown, office_com.WriteOutcomeUnknown)
        self.assertIs(x.WriteOutcomeUnknown, office_com.WriteOutcomeUnknown)
        self.assertIs(p.WriteOutcomeUnknown, office_com.WriteOutcomeUnknown)
        self.assertTrue(issubclass(office_com.WriteOutcomeUnknown, RuntimeError))
        # A handler written for one application must catch another's outcome.
        f = self.fixture(ExcelFixture)
        f.fail_write()
        with self.assertRaises(w.WriteOutcomeUnknown) as error:
            f.apply(commit=True)
        self.assertIsInstance(error.exception, x.WriteOutcomeUnknown)
        self.assertEqual(len(f.writes()), 1)

    def test_every_bridge_and_helper_comes_from_the_shared_core(self):
        import excel_com
        import office_com
        import powerpoint_com

        for bridge in (w.WordBridge, excel_com.ExcelBridge, powerpoint_com.PowerPointBridge):
            self.assertIs(bridge.__mro__[1], office_com.OfficeBridge, bridge.__name__)
        self.assertIs(excel_com.normalized_hwnd, office_com.normalized_hwnd)
        self.assertIs(powerpoint_com.normalized_hwnd, office_com.normalized_hwnd)

    def test_text_hash_is_shared_but_powerpoint_refuses_damaged_text(self):
        import office_com

        for text in ("Hello", "中文😀\r", ""):
            self.assertEqual(w.text_hash(text), office_com.text_hash(text))
            self.assertEqual(p.text_hash(text), office_com.text_hash(text))
        # Word hashes a damaged surrogate so evidence still covers it; PowerPoint
        # refuses instead of turning damaged text into an editable evidence value.
        self.assertEqual(len(w.text_hash("\ud800")), 64)
        with self.assertRaises(UnicodeEncodeError):
            p.text_hash("\ud800")
        with self.assertRaises(ValueError):
            office_com.text_hash(3)

    def test_rebind_loss_is_a_marker_labelled_refusal_in_every_app(self):
        def lost(call):
            if call == 2:
                raise RuntimeError("no native window for this target")

        for cls in (WordFixture, ExcelFixture, PowerPointFixture):
            with self.subTest(app=cls.name):
                f = self.fixture(cls)
                f.bridge.on_select = lost
                with self.assertRaises(ValueError) as error:
                    f.apply(commit=True)
                # The same documented marker and the same non-unknown class everywhere.
                self.assertNotIsInstance(error.exception, w.WriteOutcomeUnknown)
                self.assertIn("未尝试写入", str(error.exception))
                self.assertIsInstance(error.exception.__cause__, RuntimeError)
                self.assertEqual(f.writes(), [])
                f.assert_no_lifecycle(self)

    def test_word_preview_rebinds_once_and_commit_resolves_again(self):
        f = self.fixture(WordFixture)
        f.apply()
        self.assertEqual(f.bridge.calls, 1)
        self.assertEqual(f.writes(), [])
        f.bridge.calls = 0
        self.assertTrue(f.apply(commit=True)["committed"])
        self.assertEqual(f.bridge.calls, 3)
        self.assertEqual(len(f.writes()), 1)

    def test_word_without_bridge_keeps_legacy_signature_working(self):
        f = self.fixture(WordFixture)
        f.bridge = None
        self.assertTrue(f.apply(commit=True)["committed"])
        self.assertEqual(f.content(), "New world\r")

    def test_word_legacy_partial_write_is_unknown_without_bridge(self):
        f = self.fixture(WordFixture)
        f.bridge = None
        f.fail_write()
        with self.assertRaises(w.WriteOutcomeUnknown):
            f.apply(commit=True)
        self.assertEqual(f.content(), "New world\r")

    def test_word_cleanup_failure_after_write_is_unknown(self):
        f = self.fixture(WordFixture)
        f.doc.undo.EndCustomRecord.side_effect = RuntimeError("EndCustomRecord failed")
        with self.assertRaises(w.WriteOutcomeUnknown) as error:
            f.apply(commit=True)
        self.assertEqual(len(f.writes()), 1)
        self.assertEqual(f.content(), "New world\r")
        self.assertIn("撤销记录", str(error.exception))
        self.assertIsInstance(error.exception.__cause__, RuntimeError)

    def test_word_cleanup_failure_does_not_inherit_outer_exception(self):
        f = self.fixture(WordFixture)
        f.doc.undo.EndCustomRecord.side_effect = RuntimeError("EndCustomRecord failed")
        outer = ValueError("caller failure")
        try:
            raise outer
        except ValueError:
            # Inside a caller's handler, the cleanup error must not be swallowed.
            with self.assertRaises(w.WriteOutcomeUnknown) as error:
                f.apply(commit=True)
        self.assertIsInstance(error.exception.__cause__, RuntimeError)
        self.assertNotIn("Word cleanup failed", getattr(outer, "__notes__", []))

    @unittest.skip("PowerPoint has a ReadOnly guard, no separate protection/Protected View guard")
    def test_powerpoint_protected_view_contract_unsupported(self):
        pass


if __name__ == "__main__":
    unittest.main()
