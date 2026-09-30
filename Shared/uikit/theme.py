"""Design tokens and the Qt stylesheet shared by every tool.

The palette follows BetterGhub: near-black surfaces layered by lightness, one
accent colour per tool, Bahnschrift for display text and Segoe UI for body.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

ASSETS = Path(__file__).resolve().parent / "assets"

DISPLAY_FONT = "Bahnschrift"
BODY_FONT = "Segoe UI"
ICON_FONT = "Segoe MDL2 Assets"
MONO_FONT = "Cascadia Mono"


@dataclass(frozen=True)
class Palette:
    bg: str = "#0B0D10"
    rail: str = "#101318"
    panel: str = "#14181E"
    surface: str = "#1A1F26"
    surface2: str = "#222832"
    hover: str = "#2A313C"
    line: str = "#232933"
    line_strong: str = "#343C48"
    text: str = "#ECEFF3"
    muted: str = "#8D97A5"
    faint: str = "#5D6673"
    success: str = "#3DDC84"
    warning: str = "#F4B740"
    danger: str = "#F05A63"
    accent: str = "#2EC5EA"

    @property
    def accent_hover(self) -> str:
        return mix(self.accent, "#FFFFFF", 0.25)

    @property
    def accent_dim(self) -> str:
        return mix(self.bg, self.accent, 0.16)

    @property
    def on_accent(self) -> str:
        return mix(self.accent, "#000000", 0.88)

    @property
    def danger_dim(self) -> str:
        return mix(self.bg, self.danger, 0.18)

    @property
    def warning_dim(self) -> str:
        return mix(self.bg, self.warning, 0.16)

    @property
    def success_dim(self) -> str:
        return mix(self.bg, self.success, 0.14)


def mix(a: str, b: str, t: float) -> str:
    """Blend hex colour a toward b by t (0..1)."""
    ar, ag, ab = (int(a[i:i + 2], 16) for i in (1, 3, 5))
    br, bg, bb = (int(b[i:i + 2], 16) for i in (1, 3, 5))
    r = round(ar + (br - ar) * t)
    g = round(ag + (bg - ag) * t)
    b_ = round(ab + (bb - ab) * t)
    return f"#{r:02X}{g:02X}{b_:02X}"


def _asset(name: str) -> str:
    return (ASSETS / name).as_posix()


def build_qss(p: Palette) -> str:
    return f"""
* {{
    outline: none;
}}
QWidget {{
    color: {p.text};
    font-family: "{BODY_FONT}";
    font-size: 10pt;
}}
QMainWindow, QDialog, #Root {{
    background: {p.bg};
}}
QToolTip {{
    background: {p.surface2};
    color: {p.text};
    border: 1px solid {p.line_strong};
    border-radius: 6px;
    padding: 6px 8px;
}}

/* ── Rail ─────────────────────────────────────────────── */
#Rail {{
    background: {p.rail};
    border-right: 1px solid {p.line};
}}
#Brand {{
    font-family: "{DISPLAY_FONT}";
    font-size: 12pt;
    font-weight: 600;
}}
#BrandSub {{
    color: {p.faint};
    font-size: 8.5pt;
}}
#RailFoot {{
    background: {p.panel};
    border: 1px solid {p.line};
    border-radius: 12px;
}}

/* ── Typography ───────────────────────────────────────── */
#H1 {{
    font-family: "{DISPLAY_FONT}";
    font-size: 19pt;
    font-weight: 600;
}}
#Overline {{
    color: {p.muted};
    font-family: "{DISPLAY_FONT}";
    font-size: 8.5pt;
    font-weight: 600;
    letter-spacing: 1px;
}}
#CardTitle {{
    font-family: "{DISPLAY_FONT}";
    font-size: 11.5pt;
    font-weight: 600;
}}
#CardSubtitle, #Hint, #Muted {{
    color: {p.muted};
}}
#Hint {{
    font-size: 9pt;
}}
#FieldLabel {{
    color: {p.muted};
    font-size: 9pt;
    font-weight: 600;
}}
#Faint {{
    color: {p.faint};
}}

/* ── Cards ────────────────────────────────────────────── */
#Card {{
    background: {p.panel};
    border: 1px solid {p.line};
    border-radius: 14px;
}}
#Inset {{
    background: {p.surface};
    border: 1px solid {p.line};
    border-radius: 10px;
}}
#Divider {{
    background: {p.line};
    max-height: 1px;
    min-height: 1px;
}}
QScrollArea, QScrollArea > QWidget > QWidget {{
    background: transparent;
    border: none;
}}

/* ── Inputs ───────────────────────────────────────────── */
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit {{
    background: {p.surface};
    border: 1px solid {p.line_strong};
    border-radius: 8px;
    padding: 6px 10px;
    selection-background-color: {p.accent_dim};
    selection-color: {p.text};
}}
QPlainTextEdit, QTextEdit {{
    padding: 6px 8px;
}}
QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover {{
    border-color: {p.faint};
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QPlainTextEdit:focus, QTextEdit:focus {{
    border-color: {p.accent};
}}
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled {{
    color: {p.faint};
    background: {p.panel};
    border-color: {p.line};
}}
QLineEdit[placeholderText] {{
    color: {p.text};
}}
QComboBox {{
    padding-right: 28px;
}}
QComboBox::drop-down {{
    border: none;
    width: 28px;
}}
QComboBox::down-arrow {{
    image: url({_asset("chevron-down.svg")});
    width: 12px;
    height: 12px;
}}
QComboBox QAbstractItemView {{
    background: {p.surface};
    border: 1px solid {p.line_strong};
    border-radius: 8px;
    padding: 4px;
    selection-background-color: {p.accent_dim};
}}
QComboBox QAbstractItemView::item {{
    min-height: 26px;
    padding: 2px 8px;
    border-radius: 6px;
}}
QSpinBox::up-button, QSpinBox::down-button,
QDoubleSpinBox::up-button, QDoubleSpinBox::down-button {{
    border: none;
    width: 20px;
    background: transparent;
}}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{
    image: url({_asset("chevron-up.svg")});
    width: 10px;
    height: 10px;
}}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{
    image: url({_asset("chevron-down.svg")});
    width: 10px;
    height: 10px;
}}

QCheckBox, QRadioButton {{
    spacing: 9px;
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 16px;
    height: 16px;
    background: {p.surface};
    border: 1px solid {p.line_strong};
}}
QCheckBox::indicator {{
    border-radius: 5px;
}}
QRadioButton::indicator {{
    border-radius: 9px;
}}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
    border-color: {p.accent};
}}
QCheckBox::indicator:checked {{
    background: {p.accent};
    border-color: {p.accent};
    image: url({_asset("check.svg")});
}}
QRadioButton::indicator:checked {{
    background: {p.surface};
    border: 5px solid {p.accent};
    width: 8px;
    height: 8px;
}}

/* ── Buttons ──────────────────────────────────────────── */
QPushButton, QToolButton {{
    background: {p.surface2};
    color: {p.text};
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 7px 14px;
    font-weight: 600;
}}
QPushButton:hover, QToolButton:hover {{
    background: {p.hover};
}}
QPushButton:pressed, QToolButton:pressed {{
    background: {p.line_strong};
}}
QPushButton:focus {{
    border-color: {p.line_strong};
}}
QPushButton:disabled, QToolButton:disabled {{
    color: {p.faint};
    background: {p.panel};
}}
QPushButton[kind="primary"] {{
    background: {p.accent};
    color: {p.on_accent};
}}
QPushButton[kind="primary"]:hover {{
    background: {p.accent_hover};
}}
QPushButton[kind="primary"]:pressed {{
    background: {p.accent};
}}
QPushButton[kind="primary"]:disabled {{
    background: {p.accent_dim};
    color: {p.faint};
}}
QPushButton[kind="ghost"], QToolButton[kind="ghost"] {{
    background: transparent;
    color: {p.muted};
}}
QPushButton[kind="ghost"]:hover, QToolButton[kind="ghost"]:hover {{
    background: {p.surface2};
    color: {p.text};
}}
QPushButton[kind="danger"] {{
    background: transparent;
    color: {p.danger};
}}
QPushButton[kind="danger"]:hover {{
    background: {p.danger_dim};
}}
QPushButton[kind="tile"] {{
    background: {p.surface};
    border: 1px solid {p.line};
    border-radius: 10px;
    padding: 12px 10px;
}}
QPushButton[kind="tile"]:hover {{
    border-color: {p.accent};
    background: {p.surface2};
}}

/* Segmented control */
#Segmented {{
    background: {p.surface};
    border: 1px solid {p.line_strong};
    border-radius: 9px;
}}
#Segmented QPushButton {{
    background: transparent;
    color: {p.muted};
    border-radius: 7px;
    padding: 5px 12px;
}}
#Segmented QPushButton:hover {{
    color: {p.text};
}}
#Segmented QPushButton:checked {{
    background: {p.accent_dim};
    color: {p.accent};
}}

/* ── Lists ────────────────────────────────────────────── */
QListWidget, QTreeWidget {{
    background: {p.surface};
    border: 1px solid {p.line_strong};
    border-radius: 10px;
    padding: 4px;
}}
QListWidget::item, QTreeWidget::item {{
    border-radius: 6px;
    padding: 3px 6px;
}}
QListWidget::item:hover, QTreeWidget::item:hover {{
    background: {p.surface2};
}}
QListWidget::item:selected, QTreeWidget::item:selected {{
    background: {p.accent_dim};
    color: {p.text};
}}
#DropZone[dragging="true"] {{
    border: 1px solid {p.accent};
    background: {p.accent_dim};
}}

/* ── Console ──────────────────────────────────────────── */
#Console {{
    background: {p.bg};
    border: 1px solid {p.line};
    border-radius: 10px;
    font-family: "{MONO_FONT}", "Consolas";
    font-size: 9pt;
    padding: 8px 10px;
}}

/* ── Pills / badges ───────────────────────────────────── */
#Pill {{
    border-radius: 9px;
    padding: 2px 9px;
    font-size: 8.5pt;
    font-weight: 700;
}}
#Pill[tone="idle"] {{ background: {p.surface2}; color: {p.muted}; }}
#Pill[tone="busy"] {{ background: {p.accent_dim}; color: {p.accent}; }}
#Pill[tone="ok"] {{ background: {p.success_dim}; color: {p.success}; }}
#Pill[tone="warn"] {{ background: {p.warning_dim}; color: {p.warning}; }}
#Pill[tone="error"] {{ background: {p.danger_dim}; color: {p.danger}; }}

#Toast {{
    background: {p.surface2};
    border: 1px solid {p.line_strong};
    border-radius: 10px;
    padding: 10px 16px;
    font-weight: 600;
}}

/* ── Scrollbars ───────────────────────────────────────── */
QScrollBar:vertical {{
    background: transparent;
    width: 10px;
    margin: 2px;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 10px;
    margin: 2px;
}}
QScrollBar::handle {{
    background: {p.line_strong};
    border-radius: 3px;
    min-height: 32px;
    min-width: 32px;
}}
QScrollBar::handle:hover {{
    background: {p.faint};
}}
QScrollBar::add-line, QScrollBar::sub-line {{
    width: 0;
    height: 0;
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: transparent;
}}

QSplitter::handle {{
    background: transparent;
}}
QSplitter::handle:hover {{
    background: {p.accent_dim};
}}

QProgressBar {{
    background: {p.surface2};
    border: none;
    border-radius: 2px;
    max-height: 4px;
    min-height: 4px;
}}
QProgressBar::chunk {{
    background: {p.accent};
    border-radius: 2px;
}}

QMenu {{
    background: {p.surface};
    border: 1px solid {p.line_strong};
    border-radius: 8px;
    padding: 4px;
}}
QMenu::item {{
    padding: 6px 18px;
    border-radius: 6px;
}}
QMenu::item:selected {{
    background: {p.accent_dim};
}}
QMessageBox QLabel {{
    min-width: 320px;
}}
"""
