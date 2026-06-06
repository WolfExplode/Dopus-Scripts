"""
JPEG XL (cjxl) — entry point (GUI and CLI).

Bulk image → .jxl via libjxl cjxl. See cjxl_logic.py and cjxl_gui.py.
"""

from __future__ import annotations

import sys

from cjxl_gui import run_gui
from cjxl_logic import run_cli


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (OSError, ValueError):
                pass


if __name__ == "__main__":
    _configure_stdio_utf8()
    if len(sys.argv) > 1:
        raise SystemExit(run_cli(sys.argv[1:]))
    run_gui()
