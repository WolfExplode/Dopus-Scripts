"""ToolWindow: the BetterGhub-style frame every tool lives in.

    ┌────────┬───────────────────────────────┬──────────────┐
    │ brand  │ OVERLINE                [act] │              │
    │        │ Page title                    │   Activity   │
    │ ▌page  │ ┌ inputs (optional) ────────┐ │   console    │
    │  page  │ └───────────────────────────┘ │              │
    │  page  │ ┌ page cards (scrolls) ─────┐ │              │
    │        │ └───────────────────────────┘ │              │
    │ footer │                               │              │
    └────────┴───────────────────────────────┴──────────────┘
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

from PySide6.QtCore import QByteArray, QRectF, QSize, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QKeySequence, QPainter, QShortcut
from PySide6.QtWidgets import (
    QAbstractButton,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from .app import apply_dark_title_bar, glyph_icon, palette
from .console import ConsolePanel
from .theme import ICON_FONT
from .widgets import glyph_pixmap, label


class NavButton(QAbstractButton):
    """Rail entry: glyph + label, accent bar and tint when selected."""

    def __init__(self, text: str, glyph: str, parent=None):
        super().__init__(parent)
        self.setText(text)
        self._glyph = glyph
        self.setCheckable(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(40)
        self._hover = False

    def sizeHint(self) -> QSize:  # noqa: N802
        return QSize(180, 40)

    def enterEvent(self, e):  # noqa: N802
        self._hover = True
        self.update()

    def leaveEvent(self, e):  # noqa: N802
        self._hover = False
        self.update()

    def paintEvent(self, e):  # noqa: N802
        p = palette()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        r = QRectF(self.rect())
        checked = self.isChecked()
        if checked or self._hover:
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(p.accent_dim if checked else p.surface))
            painter.drawRoundedRect(r, 10, 10)
        if checked:
            painter.setBrush(QColor(p.accent))
            painter.drawRoundedRect(QRectF(0, r.height() / 2 - 9, 3, 18), 1.5, 1.5)
        gf = QFont(ICON_FONT)
        gf.setPixelSize(15)
        painter.setFont(gf)
        painter.setPen(QColor(p.accent if checked else p.text if self._hover else p.muted))
        painter.drawText(QRectF(14, 0, 22, r.height()), Qt.AlignVCenter | Qt.AlignLeft, self._glyph)
        f = QFont(self.font())
        f.setWeight(QFont.DemiBold)
        painter.setFont(f)
        painter.setPen(QColor(p.text if checked or self._hover else p.muted))
        painter.drawText(QRectF(46, 0, r.width() - 52, r.height()), Qt.AlignVCenter | Qt.AlignLeft, self.text())


@dataclass
class _Page:
    key: str
    title: str
    overline: str
    nav: NavButton
    on_run: Optional[Callable[[], None]]
    show_inputs: bool


def page(*items: QWidget, spacing: int = 14) -> QWidget:
    """Vertical stack of cards for one rail page."""
    w = QWidget()
    lay = QVBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(spacing)
    for it in items:
        lay.addWidget(it)
    lay.addStretch(1)
    return w


class ToolWindow(QMainWindow):
    def __init__(
        self,
        *,
        title: str,
        tagline: str,
        glyph: str,
        config_dir: Path,
        inputs: Optional[QWidget] = None,
        console: Optional[ConsolePanel] = None,
        size: tuple[int, int] = (1260, 800),
    ):
        super().__init__()
        self.setWindowTitle(title)
        self.setWindowIcon(glyph_icon(glyph))
        self.setMinimumSize(980, 620)
        self.resize(*size)
        self._ui_path = Path(config_dir) / "ui.json"
        self._pages: dict[str, _Page] = {}
        self._order: list[str] = []
        self.on_close: Optional[Callable[[], None]] = None
        p = palette()

        root = QWidget()
        root.setObjectName("Root")
        # Clicking empty space clears focus; also keeps initial focus off the first button.
        root.setFocusPolicy(Qt.ClickFocus)
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── rail ──
        rail = QFrame()
        rail.setObjectName("Rail")
        rail.setFixedWidth(216)
        rl = QVBoxLayout(rail)
        rl.setContentsMargins(14, 18, 14, 16)
        rl.setSpacing(4)
        brand = QHBoxLayout()
        brand.setSpacing(11)
        tile = QLabel()
        tile.setFixedSize(34, 34)
        tile.setAlignment(Qt.AlignCenter)
        tile.setPixmap(glyph_pixmap(glyph, p.on_accent, 18))
        tile.setStyleSheet(f"background: {p.accent}; border-radius: 9px;")
        brand.addWidget(tile)
        names = QVBoxLayout()
        names.setSpacing(0)
        names.addWidget(label(title, "Brand"))
        sub = label(tagline, "BrandSub")
        names.addWidget(sub)
        brand.addLayout(names, 1)
        rl.addLayout(brand)
        rl.addSpacing(22)
        self._nav_layout = QVBoxLayout()
        self._nav_layout.setSpacing(3)
        rl.addLayout(self._nav_layout)
        rl.addStretch(1)
        self.rail_footer = QVBoxLayout()
        self.rail_footer.setSpacing(8)
        rl.addLayout(self.rail_footer)
        outer.addWidget(rail)
        self._nav_group = QButtonGroup(self)
        self._nav_group.setExclusive(True)

        # ── main column ──
        main = QWidget()
        ml = QVBoxLayout(main)
        ml.setContentsMargins(28, 20, 20, 20)
        ml.setSpacing(16)
        head = QHBoxLayout()
        head.setSpacing(10)
        titles = QVBoxLayout()
        titles.setSpacing(0)
        self._overline = label("", "Overline")
        self._title = label("", "H1")
        titles.addWidget(self._overline)
        titles.addWidget(self._title)
        head.addLayout(titles, 1)
        self.header_actions = QHBoxLayout()
        self.header_actions.setSpacing(8)
        head.addLayout(self.header_actions)
        ml.addLayout(head)

        self._hsplit = QSplitter(Qt.Horizontal)
        self._hsplit.setChildrenCollapsible(False)
        self._hsplit.setHandleWidth(14)

        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 0, 0, 0)
        ll.setSpacing(0)
        self._vsplit = QSplitter(Qt.Vertical)
        self._vsplit.setChildrenCollapsible(False)
        self._vsplit.setHandleWidth(14)
        self._inputs = inputs
        if inputs is not None:
            self._vsplit.addWidget(inputs)
        self._stack = QStackedWidget()
        self._vsplit.addWidget(self._stack)
        if inputs is not None:
            self._vsplit.setStretchFactor(0, 0)
            self._vsplit.setStretchFactor(1, 1)
        ll.addWidget(self._vsplit)
        self._hsplit.addWidget(left)

        self.console = console
        if console is not None:
            console.setMinimumWidth(300)
            self._hsplit.addWidget(console)
            self._hsplit.setStretchFactor(0, 3)
            self._hsplit.setStretchFactor(1, 2)
        ml.addWidget(self._hsplit, 1)
        outer.addWidget(main, 1)

        # ── toast ──
        self._toast = QLabel(root)
        self._toast.setObjectName("Toast")
        self._toast.hide()
        self._toast_timer = QTimer(self)
        self._toast_timer.setSingleShot(True)
        self._toast_timer.timeout.connect(self._toast.hide)

        QShortcut(QKeySequence("Ctrl+Return"), self, activated=self._run_current)
        QShortcut(QKeySequence("Ctrl+Enter"), self, activated=self._run_current)

    # ---- pages ----------------------------------------------------------
    def add_page(
        self,
        key: str,
        title: str,
        glyph: str,
        widget: QWidget,
        *,
        overline: str = "",
        on_run: Optional[Callable[[], None]] = None,
        show_inputs: bool = True,
        scroll: bool = True,
    ) -> None:
        if scroll:
            area = QScrollArea()
            area.setWidgetResizable(True)
            area.setFrameShape(QFrame.NoFrame)
            area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            area.setWidget(widget)
            host: QWidget = area
        else:
            host = widget
        self._stack.addWidget(host)
        nav = NavButton(title, glyph)
        n = len(self._order) + 1
        nav.setToolTip(f"{title}  (Ctrl+{n})" if n <= 9 else title)
        nav.clicked.connect(lambda _=False, k=key: self.select_page(k))
        self._nav_group.addButton(nav)
        self._nav_layout.addWidget(nav)
        self._pages[key] = _Page(key, title, overline, nav, on_run, show_inputs)
        self._order.append(key)
        if n <= 9:
            QShortcut(QKeySequence(f"Ctrl+{n}"), self, activated=lambda k=key: self.select_page(k))
        if len(self._order) == 1:
            self.select_page(key)

    def add_rail_note(self, widget: QWidget) -> None:
        self.rail_footer.addWidget(widget)

    def current_page(self) -> str:
        idx = self._stack.currentIndex()
        return self._order[idx] if 0 <= idx < len(self._order) else ""

    def select_page(self, key: str) -> None:
        pg = self._pages.get(key)
        if pg is None:
            return
        self._stack.setCurrentIndex(self._order.index(key))
        pg.nav.setChecked(True)
        self._overline.setText(pg.overline.upper())
        self._overline.setVisible(bool(pg.overline))
        self._title.setText(pg.title)
        if self._inputs is not None:
            self._inputs.setVisible(pg.show_inputs)

    def _run_current(self) -> None:
        pg = self._pages.get(self.current_page())
        if pg is not None and pg.on_run is not None:
            pg.on_run()

    # ---- feedback -------------------------------------------------------
    def toast(self, text: str, ms: int = 2600) -> None:
        self._toast.setText(text)
        self._toast.adjustSize()
        parent = self._toast.parentWidget()
        x = (parent.width() - self._toast.width()) // 2
        y = parent.height() - self._toast.height() - 28
        self._toast.move(max(12, x), y)
        self._toast.raise_()
        self._toast.show()
        self._toast_timer.start(ms)

    # ---- persistence ----------------------------------------------------
    def restore_ui(self, page: str = "", fallback_page: str = "") -> None:
        """Restore geometry/splitters. Opens `page` if given, else the page open
        last time, else `fallback_page`."""
        data: dict = {}
        if not _snapshot_mode():
            try:
                data = json.loads(self._ui_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                data = {}
        geo = data.get("geometry")
        if isinstance(geo, str):
            self.restoreGeometry(QByteArray.fromBase64(geo.encode("ascii")))
        for split, key in ((self._hsplit, "hsplit"), (self._vsplit, "vsplit")):
            sizes = data.get(key)
            if isinstance(sizes, list) and len(sizes) == split.count() and all(isinstance(s, int) for s in sizes):
                split.setSizes(sizes)
        if "vsplit" not in data and self._inputs is not None:
            self._vsplit.setSizes([230, 520])
        if "hsplit" not in data and self.console is not None:
            self._hsplit.setSizes([560, 420])
        for pg in (page, data.get("page"), fallback_page):
            if pg in self._pages:
                self.select_page(pg)
                break

    def save_ui(self) -> None:
        data = {
            "geometry": bytes(self.saveGeometry().toBase64()).decode("ascii"),
            "hsplit": self._hsplit.sizes(),
            "vsplit": self._vsplit.sizes(),
            "page": self.current_page(),
        }
        try:
            self._ui_path.parent.mkdir(parents=True, exist_ok=True)
            self._ui_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            pass

    def run_app(self) -> int:
        """Show the window and enter the event loop.

        Design review: with UIKIT_SNAPSHOT=<folder> set, render every page to
        <folder>/<page>.png and exit instead of staying open.
        """
        import os

        from PySide6.QtWidgets import QApplication

        self.show()
        self.centralWidget().setFocus()
        out = os.environ.get("UIKIT_SNAPSHOT", "").strip()
        if out:
            def shoot() -> None:
                folder = Path(out)
                folder.mkdir(parents=True, exist_ok=True)
                for key in self._order:
                    self.select_page(key)
                    QApplication.processEvents()
                    self.grab().save(str(folder / f"{key}.png"))
                self.on_close = None
                self.close()

            QTimer.singleShot(700, shoot)
        return QApplication.instance().exec()

    def showEvent(self, e):  # noqa: N802
        apply_dark_title_bar(self)
        super().showEvent(e)

    def closeEvent(self, e):  # noqa: N802
        if self.on_close is not None:
            try:
                self.on_close()
            except Exception:  # noqa: BLE001 - never block closing
                pass
        if not _snapshot_mode():
            self.save_ui()
        super().closeEvent(e)


def _snapshot_mode() -> bool:
    import os

    return bool(os.environ.get("UIKIT_SNAPSHOT", "").strip())
