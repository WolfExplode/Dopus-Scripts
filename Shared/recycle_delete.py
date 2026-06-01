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


def recycle_delete(path: Path) -> None:
    if sys.platform != "win32":
        try:
            path.unlink()
        except OSError:
            pass
        return

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
            return
        try:
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
        except OSError:
            pass
        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = str(path) + "\0\0"
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
        if ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) != 0:
            raise OSError("SHFileOperationW failed")
        if op.fAnyOperationsAborted:
            raise OSError("SHFileOperationW aborted")
    except OSError:
        try:
            path.unlink()
        except OSError:
            pass


def safe_delete(path: Path) -> None:
    try:
        if not path.exists():
            return
        if is_ephemeral_delete(path):
            path.unlink()
        else:
            recycle_delete(path)
    except OSError:
        pass
