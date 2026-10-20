"""Shared attach-only COM substrate for the Word, Excel and PowerPoint adapters.

One unknown-outcome type, one text hash and one native identity rebinding
algorithm. Each application keeps its own window classes, attachment, evidence
and CLI; only this substrate is shared, so a caller can treat "write outcome
unknown" the same way for every application and catch it with a single class.

Nothing here opens, saves, closes or quits a document, and nothing here changes
application settings or trust settings.
"""
from __future__ import annotations
import ctypes
from ctypes import wintypes
import hashlib
import os

REFUSAL = '目标无法原生重绑定（{stage}），未尝试写入；重新 list/read 后再规划: {error}'
MISMATCH = '原生身份不匹配（{stage}），未尝试写入；重新读取文档'
MISSING_IDENTITY = '必须传入完整身份 dict 才能原生重绑定，未尝试写入'
POST_MISMATCH = '原生身份已不匹配'
WRITE_LOST = ('写入已尝试，但写后无法重新绑定原目标，可能有部分修改；'
              '不自动重试或撤销，请先在 {label} 中检查: {error}')


class WriteOutcomeUnknown(RuntimeError):
    """A write was attempted; inspect the live document before another request.

    Every adapter raises this one class, so a single except clause holds for
    Word, Excel and PowerPoint alike. Never retry or undo automatically: the
    write may have partially completed.
    """


def text_hash(text, strict=False):
    """SHA256 of the exact text, without newline or Unicode normalization.

    strict=True refuses unpaired surrogates instead of hashing their
    surrogatepass bytes; PowerPoint treats such text as damaged.
    """
    if not isinstance(text, str):
        raise ValueError('text_hash requires a str')
    payload = text.encode('utf-8', 'strict' if strict else 'surrogatepass')
    return hashlib.sha256(payload).hexdigest()


def normalized_hwnd(value):
    """Office exposes HWND as a signed Long; normalize its 32-bit representation."""
    return int(value) & 0xffffffff


class OfficeBridge:
    """Win32 signatures shared by every adapter; no application selection here."""

    def __init__(self):
        if os.name != 'nt':
            raise RuntimeError('仅支持 Windows 桌面交互会话')
        import comtypes
        from comtypes.automation import IDispatch
        from comtypes.client.dynamic import Dispatch
        self.comtypes, self.IDispatch, self.Dispatch = comtypes, IDispatch, Dispatch
        self.user32 = ctypes.WinDLL('user32', use_last_error=True)
        self.oleacc = ctypes.OleDLL('oleacc')
        self.callback = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        u = self.user32
        u.EnumWindows.argtypes = [self.callback, wintypes.LPARAM]
        u.EnumChildWindows.argtypes = [wintypes.HWND, self.callback, wintypes.LPARAM]
        u.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        u.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        u.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
        self.oleacc.AccessibleObjectFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.POINTER(comtypes.GUID), ctypes.POINTER(ctypes.POINTER(IDispatch))]
        self.oleacc.AccessibleObjectFromWindow.restype = ctypes.HRESULT

    def label(self, hwnd, class_name=False):
        buf = ctypes.create_unicode_buffer(2048)
        (self.user32.GetClassNameW if class_name else self.user32.GetWindowTextW)(hwnd, buf, len(buf))
        return buf.value

    def pid(self, hwnd):
        pid = wintypes.DWORD()
        self.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value


def _resolve(bridge, identity):
    """Raw native resolution; the caller decides what a failure means."""
    if not isinstance(identity, dict):
        raise ValueError('目标身份必须是 dict')
    if not hasattr(bridge, 'select'):
        raise ValueError('bridge 必须提供 select(pid, hwnd, document)')
    return bridge.select(identity.get('pid'), identity.get('hwnd'), identity.get('full_name'))


def rebind(bridge, identity, stage, refusal=REFUSAL, mismatch=MISMATCH):
    """Resolve the target again from native windows; never trust a cached object.

    Call this before a write. Both failure modes are refusals, because Windows
    handles are reused and a COM reference does not prove the window still hosts
    the requested document. rebind_after_write reuses the same resolution after a
    write, where the identical failure is an unknown outcome instead.
    """
    if not isinstance(identity, dict):
        raise ValueError(MISSING_IDENTITY)
    try:
        current_object, current = _resolve(bridge, identity)
    except Exception as exc:
        raise ValueError(refusal.format(stage=stage, error=exc)) from exc
    if current != identity:
        raise ValueError(mismatch.format(stage=stage))
    return current_object, current


def rebind_after_write(bridge, identity, label='Word', message=WRITE_LOST):
    """Re-resolve after a write; failure means the target may have changed.

    The write call was already issued, so nothing reported here may claim "no
    write attempted": the outcome is unknown, inspection is required, and the
    message templates of rebind() must not be reused for this path.
    """
    try:
        _, current = _resolve(bridge, identity)
        if current != identity:
            raise ValueError(POST_MISMATCH)
    except Exception as exc:
        raise WriteOutcomeUnknown(message.format(label=label, error=exc)) from exc
