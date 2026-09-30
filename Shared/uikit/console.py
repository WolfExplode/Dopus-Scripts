"""Activity panel: streaming log with a status pill, progress strip and actions."""

from __future__ import annotations

import re
from typing import Callable, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import QApplication, QPlainTextEdit, QProgressBar

from . import icons
from .app import palette
from .widgets import Card, Pill, button, icon_button

_ERROR_RE = re.compile(r"^\s*(error|failed|fail:|cannot|could not|✗)|\b(error|failed)\b[: ]", re.I)
_WARN_RE = re.compile(r"^\s*(warning|warn:|skip|skipped)|\bwarning\b", re.I)


class ConsolePanel(Card):
    """Right-hand activity log shared by every tool."""

    def __init__(self, title: str = "Activity", placeholder: str = "", parent=None):
        super().__init__(title, parent=parent, padding=16, spacing=10)
        p = palette()
        self.pill = Pill("Idle", "idle")
        self._stop_btn = button("Stop", "danger", icons.STOP, "Cancel the running job.", self._on_stop)
        self._stop_btn.hide()
        self.header_actions.addWidget(self.pill, 0, Qt.AlignVCenter)
        self.header_actions.addSpacing(4)
        self.header_actions.addWidget(self._stop_btn)
        self.header_actions.addWidget(icon_button(icons.COPY, "Copy log", self.copy))
        self.header_actions.addWidget(icon_button(icons.CLEAR, "Clear log", self.clear))

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setTextVisible(False)
        self.progress.hide()
        self.add(self.progress)

        self.view = QPlainTextEdit()
        self.view.setObjectName("Console")
        self.view.setReadOnly(True)
        self.view.setMaximumBlockCount(5000)
        self.view.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.view.setPlaceholderText(placeholder)
        self.view.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        self.add(self.view, 1)

        self._fmt = {
            "text": self._make_fmt(p.text),
            "muted": self._make_fmt(p.muted),
            "accent": self._make_fmt(p.accent, bold=True),
            "error": self._make_fmt(p.danger),
            "warn": self._make_fmt(p.warning),
            "ok": self._make_fmt(p.success, bold=True),
        }
        self._empty = True
        self._cancel: Optional[Callable[[], None]] = None
        self.was_stopped = False

    @staticmethod
    def _make_fmt(colour: str, bold: bool = False) -> QTextCharFormat:
        f = QTextCharFormat()
        f.setForeground(QColor(colour))
        if bold:
            f.setFontWeight(700)
        return f

    def _tone_for(self, text: str) -> str:
        if _ERROR_RE.search(text):
            return "error"
        if _WARN_RE.search(text):
            return "warn"
        return "text"

    # ---- writing --------------------------------------------------------
    def clear(self) -> None:
        self.view.clear()
        self._empty = True

    def copy(self) -> None:
        QApplication.clipboard().setText(self.view.toPlainText())

    def set_text(self, text: str, tone: Optional[str] = None) -> None:
        self.clear()
        for line in (text or "").splitlines() or [""]:
            self.append(line, tone=tone)
        self.view.moveCursor(QTextCursor.Start)

    def append(self, text: str, replace_last: bool = False, tone: Optional[str] = None) -> None:
        sb = self.view.verticalScrollBar()
        at_bottom = sb.value() >= sb.maximum() - 4
        fmt = self._fmt[tone or self._tone_for(text)]
        cur = QTextCursor(self.view.document())
        cur.movePosition(QTextCursor.End)
        if replace_last and not self._empty:
            cur.movePosition(QTextCursor.StartOfBlock, QTextCursor.KeepAnchor)
            cur.removeSelectedText()
        elif not self._empty:
            cur.insertBlock()
        cur.insertText(text, fmt)
        self._empty = False
        if at_bottom:
            sb.setValue(sb.maximum())

    # ---- job lifecycle --------------------------------------------------
    def begin(self, title: str, cancel: Optional[Callable[[], None]] = None) -> None:
        self.clear()
        self.append(title, tone="accent")
        self.pill.set_tone("busy", "Running")
        self.progress.show()
        self._cancel = cancel
        self.was_stopped = False
        self._stop_btn.setVisible(cancel is not None)
        self._stop_btn.setEnabled(True)

    def finish(self, ok: Optional[bool], summary: str = "") -> None:
        self.progress.hide()
        self._stop_btn.hide()
        self._cancel = None
        if summary:
            self.append("")
            for i, line in enumerate(summary.splitlines()):
                first = "warn" if self.was_stopped else "ok" if ok else "error" if ok is False else "text"
                self.append(line, tone=first if i == 0 else None)
        if self.was_stopped:
            self.pill.set_tone("warn", "Stopped")
        elif ok is None:
            self.pill.set_tone("idle", "Idle")
        else:
            self.pill.set_tone("ok" if ok else "error", "Done" if ok else "Failed")
        sb = self.view.verticalScrollBar()
        sb.setValue(sb.maximum())

    def set_status(self, tone: str, text: str) -> None:
        self.pill.set_tone(tone, text)

    def _on_stop(self) -> None:
        if self._cancel is not None:
            self.was_stopped = True
            self._stop_btn.setEnabled(False)
            self.pill.set_tone("warn", "Stopping")
            self._cancel()
