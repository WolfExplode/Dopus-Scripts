"""Path inputs: a drag-and-drop file list and a single path field."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Iterable

from PySide6.QtCore import QRect, QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont, QKeySequence, QPainter, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QVBoxLayout,
    QWidget,
)

from . import icons
from .app import palette
from .theme import ICON_FONT
from .widgets import button, icon_button, label, repolish

VIDEO_EXT = {
    ".mp4", ".m4v", ".mov", ".qt", ".mkv", ".webm", ".avi", ".wmv", ".asf", ".mpg", ".mpeg",
    ".vob", ".ts", ".mts", ".m2ts", ".3gp", ".flv", ".f4v", ".ogv", ".mxf",
}
AUDIO_EXT = {
    ".mp3", ".m4a", ".m4b", ".aac", ".flac", ".wav", ".aiff", ".aif", ".ogg", ".oga", ".opus",
    ".mka", ".wma", ".ac3", ".dts", ".ape", ".wv", ".weba",
}
IMAGE_EXT = {
    ".jpg", ".jpeg", ".jfif", ".png", ".apng", ".webp", ".bmp", ".gif", ".tif", ".tiff", ".heic",
    ".heif", ".avif", ".jxl", ".ico", ".psd", ".tga", ".dds", ".exr",
}
TEXT_EXT = {".txt", ".md", ".csv", ".json", ".log", ".srt", ".ass", ".vtt"}

ROLE_PATH = Qt.UserRole + 1


def glyph_for_path(path: str) -> str:
    p = Path(path)
    if p.is_dir():
        return icons.FOLDER
    ext = p.suffix.lower()
    if ext in VIDEO_EXT:
        return icons.VIDEO
    if ext in AUDIO_EXT:
        return icons.MUSIC
    if ext in IMAGE_EXT:
        return icons.PHOTO
    return icons.DOCUMENT


def dedupe(paths: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for raw in paths:
        s = str(raw).strip().strip('"')
        if not s:
            continue
        key = os.path.normcase(os.path.normpath(s))
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def browse_dir_hint(hint: str) -> str:
    hint = (hint or "").strip()
    if not hint:
        return ""
    p = Path(hint)
    if p.is_dir():
        return str(p)
    if p.parent.is_dir():
        return str(p.parent)
    return ""


def pick_files(parent: QWidget, title: str, hint: str = "", filter_: str = "") -> list[str]:
    paths, _ = QFileDialog.getOpenFileNames(parent, title, browse_dir_hint(hint), filter_)
    return [os.path.normpath(p) for p in paths]


def pick_folder(parent: QWidget, title: str, hint: str = "") -> str:
    path = QFileDialog.getExistingDirectory(parent, title, browse_dir_hint(hint))
    return os.path.normpath(path) if path else ""


def reveal_in_explorer(path: str) -> None:
    p = Path(path)
    try:
        if p.exists():
            subprocess.Popen(["explorer", "/select,", os.fspath(p)])
        elif p.parent.is_dir():
            os.startfile(os.fspath(p.parent))  # type: ignore[attr-defined]
    except OSError:
        pass


def paths_from_mime(mime) -> list[str]:
    if mime.hasUrls():
        out = [u.toLocalFile() for u in mime.urls() if u.isLocalFile()]
        return [os.path.normpath(p) for p in out if p]
    if mime.hasText():
        return [ln.strip().strip('"') for ln in mime.text().splitlines() if ln.strip()]
    return []


class _PathDelegate(QStyledItemDelegate):
    """One row: type glyph, file name, dimmed parent folder."""

    ROW_H = 30

    def sizeHint(self, option, index):  # noqa: N802
        return QSize(option.rect.width(), self.ROW_H)

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index) -> None:
        p = palette()
        path = index.data(ROLE_PATH) or ""
        text_items = bool(getattr(self.parent(), "text_items", False))
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawPrimitive(QStyle.PE_PanelItemViewItem, opt, painter, opt.widget)

        painter.save()
        r = option.rect.adjusted(8, 0, -8, 0)
        exists = Path(path).exists()
        plain = text_items and not exists
        gf = QFont(ICON_FONT)
        gf.setPixelSize(14)
        painter.setFont(gf)
        painter.setPen(QColor(p.accent if exists else p.muted if plain else p.danger))
        painter.drawText(QRect(r.left(), r.top(), 18, r.height()), Qt.AlignVCenter | Qt.AlignLeft,
                         glyph_for_path(path) if exists else icons.CHARACTERS if plain else icons.WARNING)

        name = path if plain else (Path(path).name or path)
        parent = "" if plain else (str(Path(path).parent) if Path(path).name else "")
        text_left = r.left() + 28
        avail = r.right() - text_left
        f = QFont(option.font)
        painter.setFont(f)
        fm = painter.fontMetrics()
        # +2: fractional glyph advances can exceed the integer width and trigger eliding.
        name_w = min(fm.horizontalAdvance(name) + 2, int(avail * 0.62) if parent else avail)
        painter.setPen(QColor(p.text if exists or plain else p.danger))
        painter.drawText(QRect(text_left, r.top(), name_w, r.height()), Qt.AlignVCenter | Qt.AlignLeft,
                         fm.elidedText(name, Qt.ElideMiddle, name_w))
        if parent:
            left = text_left + name_w + 12
            pw = r.right() - left
            if pw > 30:
                f2 = QFont(option.font)
                f2.setPointSizeF(max(7.5, option.font.pointSizeF() - 1))
                painter.setFont(f2)
                painter.setPen(QColor(p.faint))
                painter.drawText(QRect(left, r.top(), pw, r.height()), Qt.AlignVCenter | Qt.AlignLeft,
                                 painter.fontMetrics().elidedText(parent, Qt.ElideLeft, pw))
        painter.restore()


class _DropList(QListWidget):
    dropped = Signal(list)

    def __init__(self, empty_text: str, parent=None, text_items: bool = False):
        super().__init__(parent)
        self.text_items = text_items
        self.setObjectName("DropZone")
        self._empty_text = empty_text
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setItemDelegate(_PathDelegate(self))
        self.setUniformItemSizes(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setTextElideMode(Qt.ElideMiddle)

    def _set_dragging(self, on: bool) -> None:
        self.setProperty("dragging", "true" if on else "false")
        repolish(self)

    def dragEnterEvent(self, e):  # noqa: N802
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._set_dragging(True)
        else:
            e.ignore()

    def dragMoveEvent(self, e):  # noqa: N802
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dragLeaveEvent(self, e):  # noqa: N802
        self._set_dragging(False)

    def dropEvent(self, e):  # noqa: N802
        self._set_dragging(False)
        paths = paths_from_mime(e.mimeData())
        if paths:
            e.acceptProposedAction()
            self.dropped.emit(paths)

    def paintEvent(self, e):  # noqa: N802
        super().paintEvent(e)
        if self.count():
            return
        p = palette()
        painter = QPainter(self.viewport())
        r = self.viewport().rect()
        gf = QFont(ICON_FONT)
        gf.setPixelSize(26)
        painter.setFont(gf)
        painter.setPen(QColor(p.faint))
        painter.drawText(QRect(r.left(), r.center().y() - 38, r.width(), 32), Qt.AlignCenter, icons.DOWNLOAD)
        painter.setFont(self.font())
        painter.setPen(QColor(p.muted))
        painter.drawText(QRect(r.left() + 12, r.center().y() - 2, r.width() - 24, 40),
                         Qt.AlignHCenter | Qt.AlignTop | Qt.TextWordWrap, self._empty_text)


class PathList(QWidget):
    """Editable list of file/folder paths with drag-and-drop, pickers and a count."""

    changed = Signal()

    def __init__(
        self,
        paths: Iterable[str] = (),
        *,
        empty_text: str = "Drop files or folders here",
        files_title: str = "Select files",
        file_filter: str = "",
        allow_folders: bool = True,
        noun: str = "item",
        plural: str = "",
        text_items: bool = False,
        parent=None,
    ):
        """text_items: lines that are not paths are valid entries (shown as text, not errors)."""
        super().__init__(parent)
        self._files_title = files_title
        self._file_filter = file_filter
        self._noun = noun
        self._plural = plural or noun + "s"
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)

        self._text_items = text_items
        self.list = _DropList(empty_text, text_items=text_items)
        self.list.dropped.connect(self.add_paths)
        self.list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self._context_menu)
        self.list.itemDoubleClicked.connect(lambda it: reveal_in_explorer(it.data(ROLE_PATH)))
        lay.addWidget(self.list, 1)

        bar = QHBoxLayout()
        bar.setSpacing(6)
        self.count_label = label("", "Muted")
        bar.addWidget(self.count_label)
        bar.addStretch(1)
        bar.addWidget(button("Add files", "ghost", icons.ADD, "Append files with a file picker.", self.browse_files))
        if allow_folders:
            bar.addWidget(button("Add folder", "ghost", icons.FOLDER, "Append a folder.", self.browse_folder))
        bar.addWidget(icon_button(icons.CLEAR, "Remove all paths", self.clear))
        lay.addLayout(bar)

        QShortcut(QKeySequence.Delete, self.list, activated=self.remove_selected, context=Qt.WidgetShortcut)
        QShortcut(QKeySequence.Paste, self.list, activated=self._paste, context=Qt.WidgetShortcut)
        self.set_paths(paths)

    # ---- data -----------------------------------------------------------
    def paths(self) -> list[str]:
        return [self.list.item(i).data(ROLE_PATH) for i in range(self.list.count())]

    def text(self) -> str:
        return "\n".join(self.paths())

    def set_text(self, text: str) -> None:
        self.set_paths(ln for ln in (text or "").splitlines())

    def set_paths(self, paths: Iterable[str]) -> None:
        self.list.clear()
        for p in dedupe(paths):
            self._append_item(p)
        self._refresh()

    def add_paths(self, paths: Iterable[str]) -> None:
        self.set_paths([*self.paths(), *paths])

    def clear(self) -> None:
        self.set_paths([])

    def remove_selected(self) -> None:
        for it in self.list.selectedItems():
            self.list.takeItem(self.list.row(it))
        self._refresh()

    def first_path(self) -> str:
        return self.list.item(0).data(ROLE_PATH) if self.list.count() else ""

    def _append_item(self, path: str) -> None:
        it = QListWidgetItem()
        it.setData(ROLE_PATH, path)
        found = Path(path).exists()
        it.setToolTip(path if found or self._text_items else f"{path}\n(not found)")
        self.list.addItem(it)

    def _refresh(self) -> None:
        n = self.list.count()
        self.count_label.setText(f"{n} {self._noun if n == 1 else self._plural}" if n else "")
        self.list.viewport().update()
        self.changed.emit()

    # ---- actions --------------------------------------------------------
    def browse_files(self) -> None:
        picked = pick_files(self, self._files_title, self.first_path(), self._file_filter)
        if picked:
            self.add_paths(picked)

    def browse_folder(self) -> None:
        picked = pick_folder(self, "Select folder", self.first_path())
        if picked:
            self.add_paths([picked])

    def _paste(self) -> None:
        mime = QApplication.clipboard().mimeData()
        paths = paths_from_mime(mime)
        if paths:
            self.add_paths(paths)

    def _context_menu(self, pos) -> None:
        it = self.list.itemAt(pos)
        menu = QMenu(self)
        if it is not None:
            path = it.data(ROLE_PATH)
            menu.addAction("Show in Explorer", lambda: reveal_in_explorer(path))
            menu.addAction("Copy path", lambda: QApplication.clipboard().setText(path))
            menu.addAction("Remove", self.remove_selected)
            menu.addSeparator()
        menu.addAction("Paste paths", self._paste)
        menu.addAction("Clear list", self.clear)
        menu.exec(self.list.viewport().mapToGlobal(pos))


class PathField(QWidget):
    """Single path line edit with a browse button; accepts drops."""

    changed = Signal(str)

    def __init__(
        self,
        value: str = "",
        *,
        folder: bool = True,
        placeholder: str = "",
        title: str = "Select folder",
        tip: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self._folder = folder
        self._title = title
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        self.edit = QLineEdit(value)
        self.edit.setPlaceholderText(placeholder)
        if tip:
            self.edit.setToolTip(tip)
        self.edit.textChanged.connect(self.changed)
        self.edit.setAcceptDrops(False)
        lay.addWidget(self.edit, 1)
        lay.addWidget(icon_button(icons.BROWSE_FOLDER if folder else icons.OPEN_FILE, "Browse…", self.browse))
        self.setAcceptDrops(True)

    def value(self) -> str:
        return self.edit.text().strip()

    def set_value(self, value: str) -> None:
        self.edit.setText(value)

    def browse(self) -> None:
        if self._folder:
            picked = pick_folder(self, self._title, self.value())
        else:
            files = pick_files(self, self._title, self.value())
            picked = files[0] if files else ""
        if picked:
            self.set_value(picked)

    def _accept(self, paths: list[str]) -> None:
        for raw in paths:
            p = Path(raw)
            if self._folder:
                target = p if p.is_dir() else p.parent if p.is_file() else None
            else:
                target = p if p.is_file() else None
            if target is not None:
                self.set_value(os.fspath(target))
                return

    def dragEnterEvent(self, e):  # noqa: N802
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):  # noqa: N802
        self._accept(paths_from_mime(e.mimeData()))
