"""
Test suite for mail2mail (mail_formats, mail2eml, mail2mbox, mail2pst, mail2mail).
Automatically provides lightweight COM stubs if pywin32 is not installed in the test environment.
"""

from __future__ import annotations

import os
import sys
import types

repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if sys.path[0] != repo_root:
    sys.path.insert(0, repo_root)


def ensure_pywin32_stubs() -> None:
    if "pythoncom" not in sys.modules:
        try:
            import pythoncom  # noqa: F401
        except Exception:
            pc = types.ModuleType("pythoncom")
            pc.CoInitialize = lambda: None
            pc.CoUninitialize = lambda: None
            pc.PumpWaitingMessages = lambda: None
            sys.modules["pythoncom"] = pc

    if "pywintypes" not in sys.modules:
        try:
            import pywintypes  # noqa: F401
        except Exception:
            pw = types.ModuleType("pywintypes")

            class FakeComError(Exception):
                def __init__(self, hresult: int = 0, msg: str = ""):
                    super().__init__(msg or f"COM error {hex(hresult)}")
                    self.hresult = hresult

            pw.com_error = FakeComError
            sys.modules["pywintypes"] = pw

    if "win32com" not in sys.modules:
        try:
            import win32com.client  # noqa: F401
        except Exception:
            w32 = types.ModuleType("win32com")
            w32c = types.ModuleType("win32com.client")
            w32.client = w32c
            sys.modules["win32com"] = w32
            sys.modules["win32com.client"] = w32c


ensure_pywin32_stubs()
