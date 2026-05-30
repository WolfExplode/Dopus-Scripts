"""
FFmpeg Tool — entry point (GUI and CLI).

Video/audio conversion, cover split/combine, rotate, trim, merge, and more.
See ffmpeg_logic.py and ffmpeg_gui.py.
"""

from __future__ import annotations

import sys

from ffmpeg_gui import run_gui
from ffmpeg_logic import run_cli


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
