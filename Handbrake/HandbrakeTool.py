"""
HandBrake Tool -- entry point (GUI and headless repeat).

Batch-encodes with a HandBrake preset JSON from this folder.
See handbrake_logic.py and handbrake_gui.py.
"""

from __future__ import annotations

import sys

from handbrake_gui import run_gui
from handbrake_logic import run_repeat_last


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (OSError, ValueError):
                pass


def _parse_args(argv: list[str]) -> tuple[bool, str | None, list[str]]:
    repeat = False
    only_list: str | None = None
    only_files: list[str] = []
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "--repeat":
            repeat = True
        elif arg == "--only-list" and i + 1 < len(argv):
            i += 1
            only_list = argv[i]
        elif arg == "--only-file" and i + 1 < len(argv):
            i += 1
            only_files.append(argv[i])
        i += 1
    return repeat, only_list, only_files


if __name__ == "__main__":
    _configure_stdio_utf8()
    repeat, only_list, only_files = _parse_args(sys.argv[1:])
    if repeat:
        raise SystemExit(run_repeat_last(only_list, only_files))
    run_gui(only_list, only_files)
