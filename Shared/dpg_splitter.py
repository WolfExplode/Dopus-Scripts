"""Reusable draggable splitter for the two-panel Dear PyGui tools.

Dear PyGui has no native splitter widget. This adds a thin vertical handle
between a fixed-width left panel and a right panel (width=-1). While the handle
is held, the left panel width follows the mouse, which shrinks/grows the right
panel to match. The chosen width is persisted to a small JSON file so it is
restored on the next launch.

Usage:
    splitter = PanelSplitter("panel_actions", left_width=400, config_dir=CONFIG_DIR)
    splitter.build_theme()          # once, inside/after theme setup
    with dpg.group(horizontal=True):
        with dpg.child_window(width=splitter.left_width, ..., tag="panel_actions"):
            ...
        splitter.add_handle()       # between the two panels
        with dpg.child_window(width=-1, ..., tag="panel_output"):
            ...
    # in the render loop, once per frame:
    splitter.update()
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Optional


class PanelSplitter:
    def __init__(
        self,
        left_panel_tag: str,
        left_width: int = 400,
        min_left: int = 300,
        min_right: int = 260,
        handle_width: int = 8,
        config_dir: Optional[Path] = None,
    ) -> None:
        self.left_panel_tag = left_panel_tag
        self.min_left = int(min_left)
        self.min_right = int(min_right)
        self.handle_width = int(handle_width)
        self._handle_tag = f"{left_panel_tag}__splitter_handle"
        self._theme_tag = f"{left_panel_tag}__splitter_theme"
        self._config_path = (Path(config_dir) / "splitter.json") if config_dir else None
        self._dragging = False
        self._prev_left_down = False
        self._drag_start_mouse: Optional[float] = None
        self._drag_start_width = 0
        self._dirty = False

        self.left_width = max(self.min_left, int(left_width))
        saved = self._load_saved_width()
        if saved is not None:
            self.left_width = max(self.min_left, saved)

    # ---- persistence -----------------------------------------------------
    def _load_saved_width(self) -> Optional[int]:
        if not self._config_path or not self._config_path.is_file():
            return None
        try:
            data = json.loads(self._config_path.read_text(encoding="utf-8"))
            value = data.get(self.left_panel_tag)
            return int(value) if value is not None else None
        except (OSError, ValueError, TypeError):
            return None

    def save(self) -> None:
        if not self._config_path:
            return
        try:
            data: dict = {}
            if self._config_path.is_file():
                try:
                    data = json.loads(self._config_path.read_text(encoding="utf-8"))
                except (OSError, ValueError):
                    data = {}
            if not isinstance(data, dict):
                data = {}
            data[self.left_panel_tag] = int(self.left_width)
            self._config_path.parent.mkdir(parents=True, exist_ok=True)
            self._config_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            pass
        self._dirty = False

    # ---- ui --------------------------------------------------------------
    def build_theme(self) -> None:
        import dearpygui.dearpygui as dpg

        with dpg.theme(tag=self._theme_tag):
            with dpg.theme_component(dpg.mvButton):
                dpg.add_theme_color(dpg.mvThemeCol_Button, (44, 50, 66))
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (72, 168, 190))
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (92, 198, 220))
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 3)

    def add_handle(self) -> None:
        import dearpygui.dearpygui as dpg

        dpg.add_button(
            tag=self._handle_tag, label="", width=self.handle_width, height=-1,
        )
        dpg.bind_item_theme(self._handle_tag, self._theme_tag)
        with dpg.tooltip(self._handle_tag, delay=0.4):
            dpg.add_text("Drag to resize panels", color=(200, 208, 220))

    # ---- per-frame update ------------------------------------------------
    def update(self) -> None:
        """Poll each frame from the render loop. Resizes while the handle is held
        and persists the width when the drag ends."""
        import dearpygui.dearpygui as dpg

        try:
            left_down = dpg.is_mouse_button_down(dpg.mvMouseButton_Left)
        except Exception:
            return

        # Start a drag only on a fresh press that lands on the handle.
        if not self._dragging and left_down and not self._prev_left_down:
            try:
                if dpg.is_item_hovered(self._handle_tag):
                    self._dragging = True
                    self._drag_start_mouse = None
            except Exception:
                pass

        if self._dragging:
            if left_down:
                self._apply_drag()
            else:
                self._dragging = False
                self._drag_start_mouse = None
                if self._dirty:
                    self.save()

        self._prev_left_down = left_down

    def _apply_drag(self) -> None:
        import dearpygui.dearpygui as dpg

        try:
            mouse_x = dpg.get_mouse_pos(local=False)[0]
        except Exception:
            return

        # Latch the reference point on the first drag frame, then move by delta.
        if self._drag_start_mouse is None:
            self._drag_start_mouse = mouse_x
            self._drag_start_width = self.left_width
            return

        new_w = self._drag_start_width + (mouse_x - self._drag_start_mouse)
        client_w = dpg.get_viewport_client_width()
        max_w = max(self.min_left, client_w - self.min_right)
        new_w = int(max(self.min_left, min(new_w, max_w)))

        if new_w != self.left_width:
            self.left_width = new_w
            self._dirty = True
        dpg.set_item_width(self.left_panel_tag, new_w)
