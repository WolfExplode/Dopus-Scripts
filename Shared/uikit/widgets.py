"""Small building blocks: cards, fields, buttons, segmented control, pills."""

from __future__ import annotations

from typing import Callable, Iterable, Optional, Sequence

from PySide6.QtCore import QRectF, QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from .app import palette
from .theme import ICON_FONT


# ── glyphs ────────────────────────────────────────────────────────────────


def glyph_pixmap(glyph: str, color: str, px: int = 16) -> QPixmap:
    ratio = 2
    pm = QPixmap(px * ratio, px * ratio)
    pm.fill(Qt.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.TextAntialiasing)
    f = QFont(ICON_FONT)
    f.setPixelSize(int(px * ratio * 0.85))
    painter.setFont(f)
    painter.setPen(QColor(color))
    painter.drawText(QRectF(0, 0, px * ratio, px * ratio), Qt.AlignCenter, glyph)
    painter.end()
    pm.setDevicePixelRatio(ratio)
    return pm


def glyph_icon(glyph: str, color: str, px: int = 16) -> QIcon:
    return QIcon(glyph_pixmap(glyph, color, px))


class Glyph(QLabel):
    def __init__(self, glyph: str, size: int = 14, color: Optional[str] = None, parent=None):
        super().__init__(glyph, parent)
        # The app stylesheet sets font-family on every QWidget, which beats setFont(),
        # so the icon font has to come from a stylesheet too.
        colour = f"color: {color};" if color else ""
        self.setStyleSheet(
            f"font-family: '{ICON_FONT}'; font-size: {size}px; {colour} background: transparent;"
        )
        self.setAlignment(Qt.AlignCenter)


# ── text ──────────────────────────────────────────────────────────────────


def label(text: str, role: str = "", wrap: bool = False, tip: str = "") -> QLabel:
    lbl = QLabel(text)
    if role:
        lbl.setObjectName(role)
    lbl.setWordWrap(wrap)
    if tip:
        lbl.setToolTip(tip)
    lbl.setTextInteractionFlags(Qt.TextSelectableByMouse if wrap else Qt.NoTextInteraction)
    return lbl


def hint(text: str) -> QLabel:
    return label(text, "Hint", wrap=True)


class Divider(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Divider")
        self.setFrameShape(QFrame.NoFrame)


# ── buttons ───────────────────────────────────────────────────────────────


def button(
    text: str,
    kind: str = "secondary",
    glyph: Optional[str] = None,
    tip: str = "",
    on_click: Optional[Callable[[], None]] = None,
    min_width: int = 0,
) -> QPushButton:
    btn = QPushButton(text.replace("&", "&&"))
    btn.setCursor(Qt.PointingHandCursor)
    if kind != "secondary":
        btn.setProperty("kind", kind)
    if glyph:
        p = palette()
        colour = {"primary": p.on_accent, "ghost": p.muted, "danger": p.danger}.get(kind, p.text)
        btn.setIcon(glyph_icon(glyph, colour, 14))
        btn.setIconSize(QSize(14, 14))
    if tip:
        btn.setToolTip(tip)
    if on_click is not None:
        btn.clicked.connect(lambda _=False: on_click())
    if min_width:
        btn.setMinimumWidth(min_width)
    return btn


def icon_button(glyph: str, tip: str, on_click: Optional[Callable[[], None]] = None) -> QPushButton:
    btn = QPushButton()
    btn.setProperty("kind", "ghost")
    btn.setCursor(Qt.PointingHandCursor)
    btn.setIcon(glyph_icon(glyph, palette().muted, 14))
    btn.setIconSize(QSize(14, 14))
    btn.setFixedSize(32, 30)
    btn.setToolTip(tip)
    if on_click is not None:
        btn.clicked.connect(lambda _=False: on_click())
    return btn


class Segmented(QFrame):
    """Pill-style exclusive choice (e.g. Video | Audio)."""

    changed = Signal(str)

    def __init__(self, options: Sequence[tuple[str, str]], value: Optional[str] = None, parent=None):
        super().__init__(parent)
        self.setObjectName("Segmented")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(3, 3, 3, 3)
        lay.setSpacing(2)
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._buttons: dict[str, QPushButton] = {}
        for key, text in options:
            b = QPushButton(text.replace("&", "&&"))
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, k=key: self.changed.emit(k))
            self._group.addButton(b)
            self._buttons[key] = b
            lay.addWidget(b)
        self.set_value(value if value in self._buttons else options[0][0])
        self.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)

    def value(self) -> str:
        for key, b in self._buttons.items():
            if b.isChecked():
                return key
        return next(iter(self._buttons))

    def set_value(self, key: str) -> None:
        b = self._buttons.get(key)
        if b is not None:
            b.setChecked(True)


# ── inputs ────────────────────────────────────────────────────────────────


def line_edit(text: str = "", placeholder: str = "", tip: str = "", password: bool = False) -> QLineEdit:
    e = QLineEdit(text)
    if placeholder:
        e.setPlaceholderText(placeholder)
    if tip:
        e.setToolTip(tip)
    if password:
        e.setEchoMode(QLineEdit.Password)
    return e


def combo(items: Iterable[str], current: str = "", tip: str = "") -> QComboBox:
    c = QComboBox()
    c.addItems(list(items))
    if current:
        idx = c.findText(current)
        if idx >= 0:
            c.setCurrentIndex(idx)
    if tip:
        c.setToolTip(tip)
    c.setCursor(Qt.PointingHandCursor)
    return c


def spin(value: int, lo: int, hi: int, tip: str = "", suffix: str = "") -> QSpinBox:
    s = QSpinBox()
    s.setMaximumWidth(200)
    s.setRange(lo, hi)
    s.setValue(max(lo, min(hi, value)))
    if suffix:
        s.setSuffix(suffix)
    if tip:
        s.setToolTip(tip)
    return s


def dspin(value: float, lo: float, hi: float, step: float = 0.1, decimals: int = 1, tip: str = "") -> QDoubleSpinBox:
    s = QDoubleSpinBox()
    s.setMaximumWidth(200)
    s.setRange(lo, hi)
    s.setDecimals(decimals)
    s.setSingleStep(step)
    s.setValue(max(lo, min(hi, value)))
    if tip:
        s.setToolTip(tip)
    return s


def check(text: str, checked: bool = False, tip: str = "") -> QCheckBox:
    c = QCheckBox(text)
    c.setChecked(bool(checked))
    c.setCursor(Qt.PointingHandCursor)
    if tip:
        c.setToolTip(tip)
    return c


class OptionRow(QWidget):
    """Checkbox with a one-line explanation underneath."""

    def __init__(self, text: str, description: str = "", checked: bool = False, tip: str = "", parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(1)
        self.box = check(text, checked, tip)
        lay.addWidget(self.box)
        if description:
            d = hint(description)
            d.setContentsMargins(26, 0, 0, 0)
            lay.addWidget(d)

    def isChecked(self) -> bool:  # noqa: N802 - mirror QCheckBox
        return self.box.isChecked()

    def setChecked(self, value: bool) -> None:  # noqa: N802
        self.box.setChecked(value)


# ── layout helpers ────────────────────────────────────────────────────────


def field(
    title: str,
    widget: QWidget | QLayout,
    hint_text: str = "",
    tip: str = "",
) -> QWidget:
    """Label stacked above a control, with optional hint text below."""
    w = QWidget()
    lay = QVBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(5)
    lbl = label(title, "FieldLabel", tip=tip)
    lay.addWidget(lbl)
    if isinstance(widget, QLayout):
        lay.addLayout(widget)
    else:
        lay.addWidget(widget)
        if tip and not widget.toolTip():
            widget.setToolTip(tip)
    if hint_text:
        lay.addWidget(hint(hint_text))
    return w


def row(*items: QWidget | QLayout | int | None, spacing: int = 10) -> QHBoxLayout:
    """Horizontal layout. An int adds that much stretch; None adds stretch 1."""
    lay = QHBoxLayout()
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(spacing)
    for item in items:
        if item is None:
            lay.addStretch(1)
        elif isinstance(item, int):
            lay.addStretch(item)
        elif isinstance(item, QLayout):
            lay.addLayout(item)
        else:
            lay.addWidget(item)
    return lay


def grid_fields(*fields: QWidget, columns: int = 2, spacing: int = 12) -> QWidget:
    """Lay out fields in equal-width columns."""
    from PySide6.QtWidgets import QGridLayout

    w = QWidget()
    g = QGridLayout(w)
    g.setContentsMargins(0, 0, 0, 0)
    g.setHorizontalSpacing(spacing)
    g.setVerticalSpacing(spacing)
    for i, f in enumerate(fields):
        g.addWidget(f, i // columns, i % columns, Qt.AlignTop)
    for c in range(columns):
        g.setColumnStretch(c, 1)
    return w


class Card(QFrame):
    """Rounded panel with optional title/subtitle header and a body layout."""

    def __init__(
        self,
        title: str = "",
        subtitle: str = "",
        glyph: Optional[str] = None,
        parent=None,
        padding: int = 18,
        spacing: int = 14,
    ):
        super().__init__(parent)
        self.setObjectName("Card")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(padding, padding - 2, padding, padding)
        outer.setSpacing(spacing)
        self.header_actions = QHBoxLayout()
        self.header_actions.setSpacing(6)
        if title:
            head = QHBoxLayout()
            head.setSpacing(10)
            if glyph:
                tile = QLabel()
                tile.setFixedSize(30, 30)
                tile.setAlignment(Qt.AlignCenter)
                p = palette()
                tile.setPixmap(glyph_pixmap(glyph, p.accent, 16))
                tile.setStyleSheet(f"background: {p.accent_dim}; border-radius: 8px;")
                head.addWidget(tile, 0, Qt.AlignTop)
            titles = QVBoxLayout()
            titles.setSpacing(1)
            titles.addWidget(label(title, "CardTitle"))
            if subtitle:
                titles.addWidget(label(subtitle, "CardSubtitle", wrap=True))
            head.addLayout(titles, 1)
            head.addLayout(self.header_actions)
            outer.addLayout(head)
        self.body = QVBoxLayout()
        self.body.setContentsMargins(0, 0, 0, 0)
        self.body.setSpacing(spacing)
        outer.addLayout(self.body)

    def add(self, item: QWidget | QLayout, stretch: int = 0) -> None:
        if isinstance(item, QLayout):
            self.body.addLayout(item, stretch)
        else:
            self.body.addWidget(item, stretch)

    def add_actions(self, *buttons: QWidget, align_right: bool = True) -> QHBoxLayout:
        """Footer row of action buttons."""
        lay = QHBoxLayout()
        lay.setSpacing(8)
        if align_right:
            lay.addStretch(1)
        for b in buttons:
            lay.addWidget(b)
        self.body.addLayout(lay)
        return lay


class Pill(QLabel):
    def __init__(self, text: str = "", tone: str = "idle", parent=None):
        super().__init__(text, parent)
        self.setObjectName("Pill")
        self.set_tone(tone, text)

    def set_tone(self, tone: str, text: Optional[str] = None) -> None:
        if text is not None:
            self.setText(text)
        self.setProperty("tone", tone)
        self.style().unpolish(self)
        self.style().polish(self)


def repolish(w: QWidget) -> None:
    w.style().unpolish(w)
    w.style().polish(w)
    w.update()


class ActionRow(QWidget):
    """Title + description on the left, an action button on the right."""

    def __init__(self, title: str, description: str, action: QWidget, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 2, 0, 2)
        lay.setSpacing(14)
        texts = QVBoxLayout()
        texts.setSpacing(2)
        t = label(title)
        t.setStyleSheet("font-weight: 600;")
        texts.addWidget(t)
        if description:
            texts.addWidget(hint(description))
        lay.addLayout(texts, 1)
        lay.addWidget(action, 0, Qt.AlignVCenter)


class RailNote(QFrame):
    """Small info card for the bottom of the rail."""

    def __init__(self, title: str, text: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("RailFoot")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 11)
        lay.setSpacing(3)
        self.title = label(title.upper(), "Overline")
        lay.addWidget(self.title)
        self.body = label(text, "Muted", wrap=True)
        self.body.setTextInteractionFlags(Qt.NoTextInteraction)
        lay.addWidget(self.body)

    def set_text(self, text: str) -> None:
        self.body.setText(text)
