"""Send files to the Recycle Bin on Windows; permanently delete working temps."""

from __future__ import annotations

import os
import stat
import sys
from pathlib import Path


def is_ephemeral_delete(path: Path) -> bool:
    """True for working temps — permanently removed, not recycled."""
    name = path.name.lower()
    temp_root = os.environ.get("TEMP", "")
    if temp_root:
        try:
            temp_dir = Path(temp_root).resolve()
            resolved = path.resolve()
            try:
                in_temp = resolved.is_relative_to(temp_dir)
            except AttributeError:
                in_temp = str(resolved).lower().startswith(str(temp_dir).lower())
            if in_temp:
                return True
        except (OSError, ValueError):
            pass

    if ".__opus_" not in name:
        return False
    if "_orig" in name:
        return False
    return True


def recycle_delete(path: Path) -> bool:
    """Send to Recycle Bin. Returns True on success.

    On Windows, never falls back to a permanent delete: if the shell operation
    fails the file is left in place and False is returned, so a file the caller
    wanted recycled is never silently destroyed.
    """
    if sys.platform != "win32":
        try:
            path.unlink()
            return True
        except OSError:
            return False

    import ctypes
    from ctypes import wintypes

    FO_DELETE = 0x0003
    FOF_ALLOWUNDO = 0x0040
    FOF_NOCONFIRMATION = 0x0010
    FOF_SILENT = 0x0004

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("wFunc", wintypes.UINT),
            ("pFrom", wintypes.LPCWSTR),
            ("pTo", wintypes.LPCWSTR),
            ("fFlags", wintypes.WORD),
            ("fAnyOperationsAborted", wintypes.BOOL),
            ("hNameMappings", wintypes.LPVOID),
            ("lpszProgressTitle", wintypes.LPCWSTR),
        ]

    try:
        if not path.exists():
            return True
        try:
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        except OSError:
            pass
        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = str(path) + "\0\0"
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
        if ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) != 0:
            return False
        if op.fAnyOperationsAborted:
            return False
        return True
    except OSError:
        # Do not fall back to a permanent unlink() — leave the file in place.
        return False


def safe_delete(path: Path) -> bool:
    """Recycle (or permanently remove working temps). Returns True on success.

    Errors are reported via the return value, not silently swallowed, and a
    file meant for the Recycle Bin is never permanently deleted on failure.
    """
    try:
        if not path.exists():
            return True
        if is_ephemeral_delete(path):
            path.unlink()
            return True
        return recycle_delete(path)
    except OSError:
        return False
