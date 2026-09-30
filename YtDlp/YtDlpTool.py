"""
yt-dlp Downloader: entry point (GUI and headless repeat).

  YtDlpTool.py [--gui] [--dest DIR]   open the window; URL comes from the clipboard
  YtDlpTool.py --repeat --dest DIR    download the clipboard URL with saved settings

See ytdlp_logic.py and ytdlp_gui.py.
"""

from __future__ import annotations

import argparse
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
    parser = argparse.ArgumentParser(description="yt-dlp downloader (GUI and headless).")
    parser.add_argument("--gui", action="store_true", help="Open the GUI (default).")
    parser.add_argument("--repeat", action="store_true", help="Download the clipboard URL with saved settings.")
    parser.add_argument("--dest", default="", metavar="DIR", help="Folder to save into.")
    args = parser.parse_args()
    if args.repeat:
        from ytdlp_logic import run_repeat

        raise SystemExit(run_repeat(args.dest))
    from ytdlp_gui import run_gui

    run_gui(args.dest)
