"""Shared PySide6 UI kit for the Directory Opus tools.

Every tool builds its window from these pieces so they share one look:
ToolWindow (rail + header + pages + activity console), Card, field helpers,
PathList (drag-and-drop file list), ConsolePanel and JobRunner.
"""

from . import dialogs, icons
from .app import create_app, glyph_icon, palette
from .console import ConsolePanel
from .jobs import JobError, JobHost, JobRunner
from .paths import PathField, PathList, pick_files, pick_folder, reveal_in_explorer
from .shell import ToolWindow, page
from .widgets import (
    ActionRow,
    Card,
    Divider,
    Glyph,
    OptionRow,
    Pill,
    RailNote,
    Segmented,
    button,
    check,
    combo,
    dspin,
    field,
    grid_fields,
    hint,
    icon_button,
    label,
    line_edit,
    row,
    spin,
)

__all__ = [
    "ActionRow", "RailNote", "Card", "ConsolePanel", "Divider", "Glyph", "JobError", "JobHost", "JobRunner", "OptionRow", "PathField",
    "PathList", "Pill", "Segmented", "ToolWindow", "button", "check", "combo", "create_app",
    "dialogs", "dspin", "field", "glyph_icon", "grid_fields", "hint", "icon_button", "icons",
    "label", "line_edit", "page", "palette", "pick_files", "pick_folder", "reveal_in_explorer",
    "row", "spin",
]
