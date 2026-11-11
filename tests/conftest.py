"""Test setup: stub the Windows COM modules so the suite runs on any OS."""

import os
import sys
import tempfile
import types
from unittest import mock

# Keep the listing file of the tests away from the real one.
os.environ["LOCALAPPDATA"] = tempfile.mkdtemp(prefix="outlook-mcp-tests-")


class _ComError(Exception):
    """Stand-in for pythoncom.com_error."""

    def __init__(self, hresult=0, *args):
        super().__init__(hresult, *args)
        self.hresult = hresult


def _install(name):
    if name in sys.modules:
        return sys.modules[name]
    module = mock.MagicMock(name=name)
    sys.modules[name] = module
    return module


try:
    import pythoncom  # noqa: F401
except ImportError:
    for _name in ("pythoncom", "pywintypes", "win32com", "win32com.client", "win32api", "win32con"):
        _install(_name)
    sys.modules["pythoncom"].com_error = _ComError
    sys.modules["pywintypes"].com_error = _ComError
