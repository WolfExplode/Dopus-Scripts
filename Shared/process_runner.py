"""Run a console program, stream its output line by line, and allow cancelling.

Used by the downloader tools (yt-dlp, gallery-dl). Output is decoded as UTF-8;
lines ending in a bare carriage return are progress redraws and replace the
previous line when progress_re says both lines are progress.
"""

from __future__ import annotations

import re
import subprocess
import sys
import threading
from typing import Callable, Optional, Pattern

OutputSink = Callable[[str, bool], None]


def win_no_window() -> int:
    return subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


class ProcessRunner:
    def __init__(self, emit: OutputSink, *, progress_re: Optional[Pattern[str]] = None, env: Optional[dict] = None):
        self.emit = emit
        self.progress_re = progress_re
        self.env = env
        self.cancelled = False
        self._proc: Optional[subprocess.Popen] = None
        self._lock = threading.Lock()

    def cancel(self) -> None:
        """Kill the running process tree; later run() calls return -1 at once."""
        self.cancelled = True
        with self._lock:
            proc = self._proc
        if proc is not None and proc.poll() is None:
            if sys.platform == "win32":
                subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True,
                               creationflags=win_no_window())
            else:
                proc.kill()

    def run(self, cmd: list[str], cwd: Optional[str] = None, *, on_line: Optional[Callable[[str], None]] = None) -> int:
        """Run cmd to completion; returns its exit code (-1 if it could not start or was cancelled)."""
        if self.cancelled:
            return -1
        try:
            proc = subprocess.Popen(cmd, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                    stdin=subprocess.DEVNULL, creationflags=win_no_window(), env=self.env)
        except OSError as ex:
            self.emit(f"Error: could not start {cmd[0]}: {ex}", False)
            return -1
        with self._lock:
            self._proc = proc
        last_progress = False
        carry = ""
        assert proc.stdout is not None
        while True:
            chunk = proc.stdout.read1(4096)
            if not chunk:
                break
            carry += chunk.decode("utf-8", errors="replace")
            parts = re.split(r"(\r\n|\n|\r)", carry)
            carry = parts.pop()  # unfinished tail
            for i in range(0, len(parts), 2):
                line = parts[i].rstrip()
                if not line:
                    continue
                if on_line is not None:
                    on_line(line)
                progress = bool(self.progress_re and self.progress_re.match(line))
                self.emit(line, progress and last_progress)
                last_progress = progress
        tail = carry.strip()
        if tail:
            if on_line is not None:
                on_line(tail)
            self.emit(tail, False)
        rc = proc.wait()
        with self._lock:
            self._proc = None
        return -1 if self.cancelled else rc
