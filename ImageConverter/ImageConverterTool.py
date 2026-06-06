"""Image Converter — entry point (GUI and CLI)."""

from __future__ import annotations

import sys


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
        from converter_logic import run_cli
        raise SystemExit(run_cli(sys.argv[1:]))
    from converter_gui import run_gui
    run_gui()
