"""QApplication bootstrap: style, palette, fonts, taskbar identity, window icon."""

from __future__ import annotations

import sys
from typing import Optional

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPalette, QPixmap
from PySide6.QtWidgets import QApplication, QWidget

from .theme import BODY_FONT, ICON_FONT, Palette, build_qss

_palette: Palette = Palette()


def palette() -> Palette:
    """Palette of the running app (accent included)."""
    return _palette


def create_app(app_id: str, accent: str) -> QApplication:
    """Create (or reuse) the QApplication styled for one tool.

    app_id groups the tool's windows on the taskbar under its own icon instead of
    under python.exe.
    """
    global _palette
    _palette = Palette(accent=accent)

    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"DopusScripts.{app_id}")
        except (AttributeError, OSError):
            pass

    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    p = _palette
    qp = QPalette()
    for role, colour in (
        (QPalette.Window, p.bg),
        (QPalette.WindowText, p.text),
        (QPalette.Base, p.surface),
        (QPalette.AlternateBase, p.panel),
        (QPalette.Text, p.text),
        (QPalette.Button, p.surface2),
        (QPalette.ButtonText, p.text),
        (QPalette.Highlight, p.accent_dim),
        (QPalette.HighlightedText, p.text),
        (QPalette.ToolTipBase, p.surface2),
        (QPalette.ToolTipText, p.text),
        (QPalette.PlaceholderText, p.faint),
        (QPalette.Link, p.accent),
    ):
        qp.setColor(role, QColor(colour))
    qp.setColor(QPalette.Disabled, QPalette.Text, QColor(p.faint))
    qp.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(p.faint))
    qp.setColor(QPalette.Disabled, QPalette.WindowText, QColor(p.faint))
    app.setPalette(qp)

    font = QFont(BODY_FONT, 10)
    font.setHintingPreference(QFont.PreferNoHinting)
    app.setFont(font)
    app.setStyleSheet(build_qss(p))
    return app


def glyph_icon(glyph: str, accent: Optional[str] = None, size: int = 256) -> QIcon:
    """App icon: rounded accent tile with a white Segoe MDL2 glyph."""
    accent = accent or _palette.accent
    colour = QColor(accent)
    fg = QColor(Palette(accent=accent).on_accent)
    icon = QIcon()
    for px in (16, 24, 32, 48, 64, size):
        pm = QPixmap(px, px)
        pm.fill(Qt.transparent)
        painter = QPainter(pm)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(Qt.NoPen)
        painter.setBrush(colour)
        painter.drawRoundedRect(QRectF(0, 0, px, px), px * 0.24, px * 0.24)
        f = QFont(ICON_FONT)
        f.setPixelSize(max(8, int(px * 0.56)))
        painter.setFont(f)
        painter.setPen(fg)
        painter.drawText(QRectF(0, 0, px, px), Qt.AlignCenter, glyph)
        painter.end()
        icon.addPixmap(pm)
    return icon


def apply_dark_title_bar(window: QWidget) -> None:
    """Ask DWM for a dark caption so the native frame matches the app."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes

        hwnd = wintypes.HWND(int(window.winId()))
        dwm = ctypes.windll.dwmapi
        on = ctypes.c_int(1)
        # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Win10 20H1+), 19 on older builds.
        for attr in (20, 19):
            if dwm.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(on), ctypes.sizeof(on)) == 0:
                break
        # 35 = DWMWA_CAPTION_COLOR (Windows 11 only; ignored elsewhere). COLORREF is 0x00BBGGRR.
        bg = _palette.bg
        colorref = ctypes.c_int(int(bg[5:7], 16) << 16 | int(bg[3:5], 16) << 8 | int(bg[1:3], 16))
        dwm.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(colorref), ctypes.sizeof(colorref))
    except (AttributeError, OSError):
        pass
