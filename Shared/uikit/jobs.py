"""Run a blocking function on a worker thread and stream its output to the UI."""

from __future__ import annotations

import queue
import threading
import traceback
from dataclasses import dataclass
from typing import Any, Callable, Optional

from PySide6.QtCore import QObject, QTimer, Signal

OutputSink = Callable[[str, bool], None]


@dataclass
class JobError:
    """Returned as the job result when the worker raised."""

    message: str
    traceback: str


class JobRunner(QObject):
    """One job at a time. The worker calls emit(text, replace_last); the UI thread
    drains those lines on a timer, so a chatty process never floods the event loop."""

    output = Signal(str, bool)
    started = Signal()
    finished = Signal(object)

    def __init__(self, parent: Optional[QObject] = None, interval_ms: int = 40):
        super().__init__(parent)
        self._queue: queue.Queue[tuple[str, bool]] = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._result: Any = None
        self._timer = QTimer(self)
        self._timer.setInterval(interval_ms)
        self._timer.timeout.connect(self._pump)

    def is_running(self) -> bool:
        return self._thread is not None

    def start(self, fn: Callable[[OutputSink], Any]) -> bool:
        if self._thread is not None:
            return False
        self._queue = queue.Queue()
        self._result = None
        q = self._queue

        def emit(text: str, replace_last: bool = False) -> None:
            q.put((str(text), bool(replace_last)))

        def worker() -> None:
            try:
                self._result = fn(emit)
            except Exception as ex:  # noqa: BLE001 - surfaced to the user
                self._result = JobError(str(ex) or ex.__class__.__name__, traceback.format_exc())

        self._thread = threading.Thread(target=worker, daemon=True)
        self._thread.start()
        self._timer.start()
        self.started.emit()
        return True

    def wait(self, timeout: float) -> None:
        if self._thread is not None:
            self._thread.join(timeout)

    def _pump(self) -> None:
        drained = 0
        while drained < 400:
            try:
                text, replace_last = self._queue.get_nowait()
            except queue.Empty:
                break
            self.output.emit(text, replace_last)
            drained += 1
        thread = self._thread
        if thread is not None and not thread.is_alive() and self._queue.empty():
            self._thread = None
            self._timer.stop()
            self.finished.emit(self._result)


class JobHost(QObject):
    """Glue between a JobRunner, the activity console and the window.

    Starts a job with a console header, streams its output, disables the
    registered action widgets while it runs and reports the result
    (anything with .ok / .summary, or a JobError).
    """

    def __init__(self, window, console, parent: Optional[QObject] = None):
        super().__init__(parent or window)
        self.window = window
        self.console = console
        self.runner = JobRunner(self)
        self.runner.output.connect(console.append)
        self.runner.finished.connect(self._done)
        self._locked: list = []
        self._on_done: Optional[Callable[[Any], None]] = None

    def lock_while_running(self, *widgets) -> None:
        self._locked.extend(widgets)

    def is_running(self) -> bool:
        return self.runner.is_running()

    def start(
        self,
        title: str,
        fn: Callable[[OutputSink], Any],
        *,
        cancel: Optional[Callable[[], None]] = None,
        on_done: Optional[Callable[[Any], None]] = None,
    ) -> bool:
        if self.runner.is_running():
            self.window.toast("A job is already running")
            return False
        self._on_done = on_done
        self.console.begin(title, cancel)
        for w in self._locked:
            w.setEnabled(False)
        return self.runner.start(fn)

    def _done(self, result: Any) -> None:
        for w in self._locked:
            w.setEnabled(True)
        if isinstance(result, JobError):
            self.console.finish(False, f"{result.message}\n\n{result.traceback}")
            self.window.toast("Job failed")
        else:
            ok = getattr(result, "ok", True)
            summary = getattr(result, "summary", "")
            self.console.finish(bool(ok), summary)
            if summary:
                self.window.toast(summary.splitlines()[0][:120])
        cb, self._on_done = self._on_done, None
        if cb is not None:
            cb(result)
