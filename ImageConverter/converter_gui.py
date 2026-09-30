"""PySide6 front-end for the Image Converter."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from converter_logic import (
    CONFIG_DIR,
    DEFAULT_CJXL_BIN_DIR,
    DEFAULT_MAGICK_BIN_DIR,
    DEFAULT_TEXCONV_BIN_DIR,
    ICO_SIZE_KEYS,
    ICO_SIZE_LABELS,
    MAX_DIMENSION_KEYS,
    MAX_DIMENSION_LABELS,
    OUTPUT_FORMAT_KEYS,
    OUTPUT_FORMATS,
    SKYRIM_PRESET_LABELS,
    Settings,
    build_initial_files_text,
    cancel_running_jobs,
    config_load_settings,
    config_save_settings,
    ico_sizes_key_from_label,
    ico_sizes_label,
    max_dimension_key_from_label,
    max_dimension_label,
    run_convert,
    run_resize,
    skyrim_preset_key_from_label,
    skyrim_preset_label,
    validate_resize_settings,
    validate_settings,
)

import uikit as ui
from uikit import icons

ACCENT = "#A78BFA"

TIP_JXL_DISTANCE = (
    "0 = lossless (keeps ICC colour profiles)\n"
    "1 = visually lossless (recommended)\n"
    "2–5 = good quality, smaller files\n"
    "25 = maximum compression\n\n"
    "Lossy (distance > 0) may strip embedded ICC profiles."
)
TIP_JXL_EFFORT = "Higher effort = smaller files, slower encode.\n7 = high quality (slow for big batches)\n3–4 = faster, still good\n1 = fastest"
TIP_SKYRIM = (
    "Skyrim SE DDS via texconv, full mip chain.\n\n"
    "Auto: names ending _n or _msn are normal maps, everything else is diffuse. Both use BC7.\n\n"
    "Normal maps keep alpha as the specular mask. BC5 is not offered because Skyrim does not "
    "reconstruct Z.\n\nBC1: half the size of BC7, no alpha. BC3/DXT5: legacy format with alpha."
)


def _bounded_int(raw: str, default: int, lo: int, hi: int) -> int:
    try:
        n = int(float(raw.strip()))
    except (ValueError, AttributeError):
        return default
    return max(lo, min(hi, n))


def _bounded_float(raw: str, default: float, lo: float, hi: float) -> float:
    try:
        n = float(raw.strip())
    except (ValueError, AttributeError):
        return default
    return max(lo, min(hi, n))


class ConverterWindow(ui.ToolWindow):
    def __init__(self, settings: Settings, files_text: str, tab_folder: str):
        self.settings = settings
        self._tab_folder = tab_folder

        self.files = ui.PathList(
            empty_text="Drop images or folders here\nFolders include every image inside, recursively",
            files_title="Select image files",
            noun="path",
        )
        self.files.set_text(files_text)
        files_card = ui.Card("Images", "Files already in the target format are skipped.", icons.OPEN_FILE)
        files_card.add(self.files, 1)

        console = ui.ConsolePanel(placeholder="Choose a format and images, then Convert. Encoder output appears here.")
        super().__init__(
            title="Image Converter",
            tagline="ImageMagick · cjxl · texconv",
            glyph=icons.PHOTO,
            config_dir=CONFIG_DIR,
            inputs=files_card,
            console=console,
        )
        self.jobs = ui.JobHost(self, console)

        self.add_page("convert", "Convert", icons.PHOTO, self._build_convert(), overline="Change format",
                      on_run=self.convert)
        self.add_page("resize", "Resize", icons.RESIZE, self._build_resize(), overline="Scale images",
                      on_run=self.resize_images)
        self.add_page("tools", "Encoders", icons.SETTINGS, self._build_tools(), overline="Program folders",
                      show_inputs=False)

        if tab_folder:
            name = Path(tab_folder).name or tab_folder
            self.add_rail_note(ui.RailNote("Opus folder", f"With an empty list, every image in “{name}” is used."))
        self.on_close = self._save
        self._sync_format()

    # ── pages ──────────────────────────────────────────────────────────────
    def _build_convert(self):
        s = self.settings
        labels = [OUTPUT_FORMATS[k]["label"] for k in OUTPUT_FORMAT_KEYS]
        self.format = ui.combo(labels, OUTPUT_FORMATS[s.output_format]["label"], "Target format for every image.")
        self.format.currentTextChanged.connect(lambda _: self._sync_format())
        self.quality = ui.spin(_bounded_int(s.quality, 90, 1, 100), 1, 100, "1 = worst, 100 = best.")
        self.jxl_distance = ui.dspin(_bounded_float(s.jxl_distance, 1.0, 0.0, 25.0), 0.0, 25.0, tip=TIP_JXL_DISTANCE)
        self.jxl_effort = ui.spin(_bounded_int(s.jxl_effort, 7, 0, 9), 0, 9, TIP_JXL_EFFORT)
        self.ico_size = ui.combo([ICO_SIZE_LABELS[k] for k in ICO_SIZE_KEYS], ico_sizes_label(s.ico_sizes),
                                 "Square size of the .ico output.")
        self.skyrim = ui.combo(SKYRIM_PRESET_LABELS, skyrim_preset_label(s.skyrim_preset), TIP_SKYRIM)
        self.max_dim = ui.combo([MAX_DIMENSION_LABELS[k] for k in MAX_DIMENSION_KEYS],
                                max_dimension_label(s.max_dimension),
                                "Downscale anything bigger than this box (Lanczos). Smaller images are left alone.")
        self.replace_src = ui.OptionRow("Replace source file",
                                        "Send each original to the Recycle Bin after it converts.", s.replace_source)

        self.f_quality = ui.field("Quality", self.quality, "JPEG / WebP / AVIF quality.")
        self.f_jxl = ui.grid_fields(ui.field("Distance", self.jxl_distance, "0 lossless · 1 visually lossless"),
                                    ui.field("Effort", self.jxl_effort, "Higher is smaller and slower"))
        self.f_ico = ui.field("Icon size", self.ico_size)
        self.f_skyrim = ui.field("Texture type", self.skyrim, "Encoded by texconv with a full mip chain.")

        card = ui.Card("Output format", "Converted files are written next to the originals.", icons.PHOTO)
        card.add(ui.grid_fields(ui.field("Format", self.format), ui.field("Max dimensions", self.max_dim)))
        for w in (self.f_quality, self.f_jxl, self.f_ico, self.f_skyrim):
            card.add(w)
        card.add(self.replace_src)
        btn = ui.button("Convert", "primary", icons.PLAY, "Convert every listed image.", self.convert)
        self.jobs.lock_while_running(btn)
        card.add_actions(btn)
        return ui.page(card)

    def _build_resize(self):
        s = self.settings
        self.width = ui.spin(_bounded_int(s.resize_width, 0, 0, 16384), 0, 16384, suffix=" px")
        self.height = ui.spin(_bounded_int(s.resize_height, 0, 0, 16384), 0, 16384, suffix=" px")
        for sp in (self.width, self.height):
            sp.setSpecialValueText("auto")
        self.keep_aspect = ui.OptionRow(
            "Preserve aspect ratio",
            "With both sizes set, fit inside that box. Off forces the exact width and height.",
            s.resize_preserve_aspect,
        )
        card = ui.Card("Resize to a custom size",
                       "Set either side to auto to derive it from the aspect ratio. Writes <name>_resized next to "
                       "each image.", icons.RESIZE)
        card.add(ui.grid_fields(ui.field("Width", self.width), ui.field("Height", self.height)))
        card.add(self.keep_aspect)
        btn = ui.button("Resize", "primary", icons.RESIZE, "Resize every listed image.", self.resize_images)
        self.jobs.lock_while_running(btn)
        card.add_actions(btn)
        return ui.page(card)

    def _build_tools(self):
        s = self.settings
        self.magick_dir = ui.PathField(s.magick_bin_dir, placeholder=DEFAULT_MAGICK_BIN_DIR,
                                       title="Folder containing magick.exe")
        self.cjxl_dir = ui.PathField(s.cjxl_bin_dir, placeholder=DEFAULT_CJXL_BIN_DIR,
                                     title="Folder containing cjxl.exe")
        self.texconv_dir = ui.PathField(s.texconv_bin_dir, placeholder=DEFAULT_TEXCONV_BIN_DIR,
                                        title="Folder containing texconv")
        card = ui.Card("Encoder folders", "Saved when you convert or close the window.", icons.SETTINGS)
        card.add(ui.field("ImageMagick", self.magick_dir, "Folder with magick.exe (portable ImageMagick 7)."))
        card.add(ui.field("cjxl (libjxl)", self.cjxl_dir,
                          "When cjxl.exe is found, JXL output uses it directly: native --distance and faster."))
        card.add(ui.field("texconv (DirectXTex)", self.texconv_dir,
                          "Folder with Texconvx64.exe or texconv.exe, for Skyrim DDS."))
        return ui.page(card)

    # ── state ──────────────────────────────────────────────────────────────
    def _format_key(self) -> str:
        label = self.format.currentText()
        for k, v in OUTPUT_FORMATS.items():
            if v["label"] == label:
                return k
        return "jpeg"

    def _sync_format(self) -> None:
        encode = OUTPUT_FORMATS[self._format_key()]["encode"]
        self.f_quality.setVisible(encode == "quality")
        self.f_jxl.setVisible(encode == "jxl")
        self.f_ico.setVisible(encode == "ico")
        self.f_skyrim.setVisible(encode == "skyrim")

    def _collect(self) -> Settings:
        return Settings(
            output_format=self._format_key(),
            quality=str(self.quality.value()),
            jxl_distance=f"{self.jxl_distance.value():g}",
            jxl_effort=str(self.jxl_effort.value()),
            replace_source=self.replace_src.isChecked(),
            magick_bin_dir=self.magick_dir.value() or DEFAULT_MAGICK_BIN_DIR,
            cjxl_bin_dir=self.cjxl_dir.value() or DEFAULT_CJXL_BIN_DIR,
            texconv_bin_dir=self.texconv_dir.value() or DEFAULT_TEXCONV_BIN_DIR,
            skyrim_preset=skyrim_preset_key_from_label(self.skyrim.currentText()),
            max_dimension=max_dimension_key_from_label(self.max_dim.currentText()),
            ico_sizes=ico_sizes_key_from_label(self.ico_size.currentText()),
            resize_width=str(self.width.value()) if self.width.value() else "",
            resize_height=str(self.height.value()) if self.height.value() else "",
            resize_preserve_aspect=self.keep_aspect.isChecked(),
            files_text=self.files.text(),
        )

    # ── running ────────────────────────────────────────────────────────────
    def _start(self, verb: str, validate, run) -> None:
        if self.jobs.is_running():
            self.toast("A job is already running")
            return
        settings = self._collect()
        err = validate(settings)
        if err:
            ui.dialogs.error(self, "Image Converter", err)
            return
        paths_text = settings.files_text.strip()
        if not paths_text and not self._tab_folder:
            self.toast("Add images first")
            return
        self.settings = settings
        config_save_settings(settings)
        tab = self._tab_folder or None
        target = OUTPUT_FORMATS[settings.output_format]["label"] if verb == "Convert" else "custom size"
        self.jobs.start(
            f"{verb} → {target}",
            lambda emit: run(paths_text, settings, tab_folder=tab, on_output=emit),
            cancel=cancel_running_jobs,
            on_done=self._on_done,
        )

    def _on_done(self, result) -> None:
        if self.console.was_stopped:
            return
        if not getattr(result, "ok", True) and getattr(result, "summary", ""):
            ui.dialogs.error(self, "Image Converter", result.summary)

    def convert(self) -> None:
        self._start("Convert", validate_settings, run_convert)

    def resize_images(self) -> None:
        self._start("Resize", validate_resize_settings, run_resize)

    def _save(self) -> None:
        if self.jobs.is_running():
            cancel_running_jobs()
            self.jobs.runner.wait(3.0)
        try:
            config_save_settings(self._collect())
        except OSError:
            pass


def run_gui(
    initial_only_list: Optional[str] = None,
    initial_only_files: Optional[list[str]] = None,
    initial_tab_folder: Optional[str] = None,
) -> None:
    ui.create_app("ImageConverter", ACCENT)
    settings = config_load_settings()
    files_text = build_initial_files_text(
        settings.files_text, initial_only_list, initial_only_files, initial_tab_folder
    )
    win = ConverterWindow(settings, files_text, (initial_tab_folder or "").strip())
    win.restore_ui()
    win.run_app()
