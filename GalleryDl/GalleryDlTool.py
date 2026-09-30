"""gallery-dl Pinterest scraper: entry point. See gallerydl_logic.py and gallerydl_gui.py."""

from __future__ import annotations

import sys

if __name__ == "__main__":
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            if stream is not None and hasattr(stream, "reconfigure"):
                try:
                    stream.reconfigure(encoding="utf-8")
                except (OSError, ValueError):
                    pass
    from gallerydl_gui import run_gui

    run_gui()
