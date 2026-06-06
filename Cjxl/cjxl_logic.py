"""Bulk image → JPEG XL via libjxl cjxl (ported from DOpus_cjxl.js)."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))
from recycle_delete import recycle_delete

DEFAULT_CJXL_BIN_DIR = r"C:\Users\WXP\Desktop\Tools\jxl-x64-windows-static\bin"
DEFAULT_MAGICK_BIN_DIR = (
    r"C:\Users\WXP\Desktop\Tools\ImageMagick-7.1.2-25-portable-Q16-HDRI-x64"
)
CJXL_EXIT_CONTROL_C = -1073741510

MAX_DIMENSION_PRESETS: dict[str, tuple[int, int] | None] = {
    "none": None,
    "2048": (2048, 2048),
    "4096": (4096, 4096),
    "1k": (1920, 1080),
    "2k": (2560, 1440),
    "4k": (3840, 2160),
    "8k": (7680, 4320),
}

MAX_DIMENSION_LABELS: dict[str, str] = {
    "none": "No limit",
    "2048": "2048",
    "4096": "4096",
    "1k": "1K (1920×1080)",
    "2k": "2K (2560×1440)",
    "4k": "4K (3840×2160)",
    "8k": "8K (7680×4320)",
}

MAX_DIMENSION_KEYS = tuple(MAX_DIMENSION_PRESETS.keys())

CJXL_NATIVE_EXTS = {
    ".png", ".jpg", ".jpeg", ".jpe", ".gif",
    ".ppm", ".pgm", ".pbm", ".pnm", ".pfm", ".pam", ".pgx", ".exr",
}

PRECONVERT_EXTS = {
    ".bmp", ".dib", ".tif", ".tiff", ".webp", ".ico", ".cur",
    ".heic", ".heif", ".avif", ".psd", ".jfif", ".jp2", ".j2k", ".tga", ".pcx", ".wbmp",
}

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "CjxlTool"
CONFIG_PATH = CONFIG_DIR / "settings.json"
LEGACY_INI_PATH = Path(os.environ.get("APPDATA", "")) / "DOpus_cjxl_settings.ini"

GUI_SECTION_DEFAULTS: dict[str, bool] = {
    "files": True,
    "encode": True,
}

OutputSink = Callable[[str, bool], None]


@dataclass
class Settings:
    encode_mode: int = 0
    quality: str = "90"
    distance: str = "1"
    effort: str = ""
    progressive: bool = False
    replace_source: bool = False
    cjxl_bin_dir: str = DEFAULT_CJXL_BIN_DIR
    magick_bin_dir: str = DEFAULT_MAGICK_BIN_DIR
    max_dimension: str = "none"
    files_text: str = ""
    gui_sections: dict[str, bool] = field(default_factory=lambda: dict(GUI_SECTION_DEFAULTS))


@dataclass
class ConvertResult:
    ok: bool
    summary: str
    log: list[str] = field(default_factory=list)
    converted: int = 0
    failed: int = 0


def file_ext_lower(path: Path | str) -> str:
    return Path(path).suffix.lower()


def image_input_kind(path: Path) -> str:
    ext = file_ext_lower(path)
    if ext == ".jxl":
        return "skip"
    if ext in CJXL_NATIVE_EXTS:
        return "native"
    if ext in PRECONVERT_EXTS:
        return "preconvert"
    return "unknown"


def is_image_input_path(path: Path, allow_unknown: bool = False) -> bool:
    kind = image_input_kind(path)
    if kind == "skip":
        return False
    if kind == "unknown":
        return allow_unknown
    return True


def dedupe_paths(paths: list[Path]) -> list[Path]:
    seen: set[str] = set()
    out: list[Path] = []
    for p in paths:
        key = os.fspath(p).casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def dedupe_path_lines(lines: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for ln in lines:
        key = ln.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def paths_from_text_lines(text: str) -> list[Path]:
    paths: list[Path] = []
    for ln in text.splitlines():
        s = ln.strip()
        if not s:
            continue
        p = Path(s)
        if p.is_file() or p.is_dir():
            paths.append(p.resolve())
    return dedupe_paths(paths)


def paths_from_only_list(
    list_path: Optional[str],
    only_files: Optional[list[str]],
) -> list[Path]:
    lines: list[str] = []
    if list_path:
        try:
            lines.extend(Path(list_path).read_text(encoding="utf-8").splitlines())
        except OSError:
            pass
    if only_files:
        lines.extend(only_files)
    return paths_from_text_lines("\n".join(lines))


def build_initial_files_text(
    saved: str,
    only_list: Optional[str],
    only_files: Optional[list[str]],
    tab_folder: Optional[str],
) -> str:
    from_dopus = paths_from_only_list(only_list, only_files)
    if from_dopus:
        return "\n".join(os.fspath(p) for p in from_dopus)
    tab = (tab_folder or "").strip()
    if tab and Path(tab).is_dir():
        return tab
    return saved.strip()


def list_image_files_in_folder(folder: Path, out: list[Path], seen: set[str]) -> None:
    if not folder.is_dir():
        return
    try:
        for child in folder.iterdir():
            if child.is_file():
                if not is_image_input_path(child, allow_unknown=False):
                    continue
                key = os.fspath(child).casefold()
                if key not in seen:
                    seen.add(key)
                    out.append(child.resolve())
            elif child.is_dir():
                list_image_files_in_folder(child, out, seen)
    except OSError:
        pass


def collect_image_inputs(paths_text: str, tab_folder: Optional[str] = None) -> list[Path]:
    out: list[Path] = []
    seen: set[str] = set()
    listed = paths_from_text_lines(paths_text)

    for p in listed:
        if p.is_file():
            if is_image_input_path(p, allow_unknown=False):
                key = os.fspath(p).casefold()
                if key not in seen:
                    seen.add(key)
                    out.append(p)
        elif p.is_dir():
            list_image_files_in_folder(p, out, seen)

    if not out and not listed:
        tab = (tab_folder or "").strip()
        if tab:
            list_image_files_in_folder(Path(tab).resolve(), out, seen)

    return out


def _migrate_legacy_ini() -> Optional[dict]:
    if not LEGACY_INI_PATH.is_file():
        return None
    out: dict = {}
    try:
        for line in LEGACY_INI_PATH.read_text(encoding="utf-8").splitlines():
            line = line.rstrip("\r")
            if "=" not in line:
                continue
            key, val = line.split("=", 1)
            out[key.strip()] = val.strip()
    except OSError:
        return None
    return out


def config_read() -> dict:
    if CONFIG_PATH.is_file():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            pass
    legacy = _migrate_legacy_ini()
    if legacy:
        encode_mode = int(legacy.get("encodeMode", 0) or 0)
        if encode_mode in (2, 3):
            encode_mode = 1
        elif encode_mode not in (0, 1):
            encode_mode = 0
        return {
            "encode_mode": encode_mode,
            "quality": legacy.get("quality", "90"),
            "distance": legacy.get("distance", "1") if encode_mode == 1 else legacy.get("distance", "1"),
            "effort": legacy.get("effort", ""),
            "progressive": legacy.get("progressive") == "1",
            "cjxl_bin_dir": DEFAULT_CJXL_BIN_DIR,
        }
    return {}


def config_write(data: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def config_load_settings() -> Settings:
    data = config_read()
    sections = data.get("gui_sections") or {}
    gui_sections = dict(GUI_SECTION_DEFAULTS)
    if isinstance(sections, dict):
        for k in GUI_SECTION_DEFAULTS:
            if k in sections:
                gui_sections[k] = bool(sections[k])
    encode_mode = int(data.get("encode_mode", 0) or 0)
    if encode_mode not in (0, 1):
        encode_mode = 0
    max_dim = str(data.get("max_dimension") or "none")
    if max_dim not in MAX_DIMENSION_PRESETS:
        max_dim = "none"
    return Settings(
        encode_mode=encode_mode,
        quality=str(data.get("quality") or "90") or "90",
        distance=str(data.get("distance") or "1") or "1",
        effort=str(data.get("effort") or ""),
        progressive=bool(data.get("progressive")),
        replace_source=bool(data.get("replace_source")),
        cjxl_bin_dir=str(data.get("cjxl_bin_dir") or DEFAULT_CJXL_BIN_DIR) or DEFAULT_CJXL_BIN_DIR,
        magick_bin_dir=str(data.get("magick_bin_dir") or DEFAULT_MAGICK_BIN_DIR) or DEFAULT_MAGICK_BIN_DIR,
        max_dimension=max_dim,
        files_text=str(data.get("files_text") or ""),
        gui_sections=gui_sections,
    )


def config_save_settings(settings: Settings) -> None:
    data = config_read()
    data["encode_mode"] = settings.encode_mode
    data["quality"] = settings.quality
    data["distance"] = settings.distance
    data["effort"] = settings.effort
    data["progressive"] = settings.progressive
    data["replace_source"] = settings.replace_source
    data["cjxl_bin_dir"] = settings.cjxl_bin_dir
    data["magick_bin_dir"] = settings.magick_bin_dir
    data["max_dimension"] = settings.max_dimension
    data["files_text"] = settings.files_text
    data["gui_sections"] = settings.gui_sections
    try:
        config_write(data)
    except OSError:
        pass


def parse_quality(raw: str) -> tuple[Optional[str], Optional[str]]:
    s = raw.strip()
    try:
        n = float(s)
    except ValueError:
        return None, "Quality must be a number from 1 to 100."
    if n < 1 or n > 100:
        return None, "Quality must be a number from 1 to 100."
    return str(n), None


def parse_distance(raw: str) -> tuple[Optional[str], Optional[str]]:
    s = raw.strip()
    try:
        n = float(s)
    except ValueError:
        return None, "Distance must be 0 or a positive number."
    if n < 0:
        return None, "Distance must be 0 or a positive number."
    return str(n), None


def parse_effort(raw: str) -> tuple[Optional[str], Optional[str]]:
    s = raw.strip()
    if not s:
        return "", None
    try:
        n = int(s)
    except ValueError:
        return None, "Effort must be blank or a whole number from 1 to 10."
    if n < 1 or n > 10:
        return None, "Effort must be blank or a whole number from 1 to 10."
    return str(n), None


def validate_settings(settings: Settings) -> Optional[str]:
    if settings.encode_mode == 1:
        _, err = parse_distance(settings.distance)
    else:
        _, err = parse_quality(settings.quality)
    if err:
        return err
    _, err = parse_effort(settings.effort)
    if err:
        return err
    return None


def build_cjxl_encode_args(settings: Settings) -> str:
    parts: list[str] = []
    if settings.encode_mode == 1:
        parts.append(f"--distance={settings.distance}")
    else:
        parts.append(f"--quality={settings.quality}")
    if settings.effort:
        parts.append(f"-e {settings.effort}")
    if settings.progressive:
        parts.append("--progressive")
    return " ".join(parts)


def output_path_for_input(input_path: Path) -> Path:
    return input_path.with_suffix(".jxl")


def temp_png_path(input_path: Path, seq: int) -> Path:
    base = input_path.stem
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in base) or "img"
    return Path(tempfile.gettempdir()) / f"DOpus_cjxl_{safe}_{seq}.png"


def max_dimension_box(key: str) -> tuple[int, int] | None:
    return MAX_DIMENSION_PRESETS.get(key, MAX_DIMENSION_PRESETS["none"])


def max_dimension_label(key: str) -> str:
    return MAX_DIMENSION_LABELS.get(key, MAX_DIMENSION_LABELS["none"])


def max_dimension_key_from_label(label: str) -> str:
    for key, text in MAX_DIMENSION_LABELS.items():
        if text == label:
            return key
    return "none"


def image_exceeds_box(width: int, height: int, box: tuple[int, int]) -> bool:
    return width > box[0] or height > box[1]


def resolve_magick_exe(settings: Settings) -> Optional[Path]:
    bin_dir = Path(os.path.expandvars(settings.magick_bin_dir.strip()))
    exe = bin_dir / "magick.exe"
    return exe if exe.is_file() else None


def magick_identify_size(magick: Path, input_path: Path) -> Optional[tuple[int, int]]:
    cmd = f'"{magick}" identify -ping -format "%w %h" "{input_path}"'
    try:
        proc = subprocess.run(
            cmd,
            shell=True,
            cwd=os.fspath(magick.parent),
            capture_output=True,
            text=True,
            creationflags=_win_subprocess_flags(),
        )
        if proc.returncode != 0:
            return None
        parts = proc.stdout.strip().split()
        if len(parts) != 2:
            return None
        return int(parts[0]), int(parts[1])
    except (OSError, ValueError):
        return None


def _run_magick(magick: Path, args: str, input_path: Path, output_png: Path) -> bool:
    cmd = f'"{magick}" "{input_path}" {args} "{output_png}"'
    try:
        proc = subprocess.run(
            cmd,
            shell=True,
            cwd=os.fspath(magick.parent),
            capture_output=True,
            creationflags=_win_subprocess_flags(),
        )
        return proc.returncode == 0 and output_png.is_file()
    except OSError:
        return False


def magick_prepare_png(
    magick: Path,
    input_path: Path,
    output_png: Path,
    box: tuple[int, int] | None,
) -> bool:
    resize = f'-filter Lanczos -resize "{box[0]}x{box[1]}>"' if box else ""
    return _run_magick(magick, resize, input_path, output_png)


def needs_magick_prepare(
    settings: Settings,
    input_path: Path,
    magick: Optional[Path],
) -> tuple[bool, tuple[int, int] | None]:
    kind = image_input_kind(input_path)
    if kind in ("skip", "unknown"):
        return False, None
    box = max_dimension_box(settings.max_dimension)
    needs_preconvert = kind == "preconvert"
    if not box:
        return needs_preconvert, None
    if not magick:
        return True, box
    size = magick_identify_size(magick, input_path)
    if size is None:
        return True, box
    if image_exceeds_box(size[0], size[1], box):
        return True, box
    return needs_preconvert, None


def resolve_cjxl_input_path(
    settings: Settings,
    input_path: Path,
    seq: int,
    magick: Optional[Path],
) -> tuple[Optional[Path], Optional[Path], str]:
    kind = image_input_kind(input_path)
    if kind == "skip":
        return None, None, ""
    use_magick, resize_box = needs_magick_prepare(settings, input_path, magick)
    if not use_magick and kind == "native":
        return input_path, None, ""
    if not magick:
        return None, None, ""
    temp_png = temp_png_path(input_path, seq)
    try:
        if temp_png.is_file():
            temp_png.unlink()
    except OSError:
        pass
    if not magick_prepare_png(magick, input_path, temp_png, resize_box):
        return None, None, ""
    if resize_box:
        label = max_dimension_label(settings.max_dimension)
        return temp_png, temp_png, f"resized to {label}"
    if kind == "preconvert":
        return temp_png, temp_png, "converted to PNG"
    return temp_png, temp_png, "converted to PNG"


def resolve_cjxl_exe(settings: Settings) -> Optional[Path]:
    bin_dir = Path(os.path.expandvars(settings.cjxl_bin_dir.strip()))
    exe = bin_dir / "cjxl.exe"
    return exe if exe.is_file() else None


def _win_subprocess_flags() -> int:
    if sys.platform != "win32":
        return 0
    return subprocess.CREATE_NO_WINDOW


def _stream_process_output(proc: subprocess.Popen, emit: OutputSink) -> None:
    stream = proc.stderr
    if stream is None:
        return
    carry = ""
    try:
        while True:
            chunk = stream.read(4096)
            if not chunk:
                break
            carry += chunk.decode("utf-8", errors="replace")
            while carry:
                r = carry.find("\r")
                n = carry.find("\n")
                if r == -1 and n == -1:
                    break
                if n != -1 and (r == -1 or n < r):
                    line, carry = carry[:n], carry[n + 1 :]
                    line = line.strip("\r")
                    if line:
                        emit(line, False)
                else:
                    line, carry = carry[:r], carry[r + 1 :]
                    if line:
                        emit(line, True)
        tail = carry.strip("\r\n")
        if tail:
            emit(tail, False)
    except OSError:
        pass


def _run_cjxl(
    exe: Path,
    encode_args: str,
    input_path: Path,
    output_path: Path,
    emit: OutputSink,
) -> int:
    cmd = f'"{exe}" {encode_args} "{input_path}" "{output_path}"'
    try:
        proc = subprocess.Popen(
            cmd,
            shell=True,
            cwd=os.fspath(exe.parent),
            stderr=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            creationflags=_win_subprocess_flags(),
        )
        reader = threading.Thread(target=_stream_process_output, args=(proc, emit), daemon=True)
        reader.start()
        try:
            rc = proc.wait()
        finally:
            if proc.stderr is not None:
                try:
                    proc.stderr.close()
                except OSError:
                    pass
            reader.join(timeout=2.0)
        return int(rc)
    except OSError as ex:
        emit(f"Error: {ex}", False)
        return -1


def run_convert(
    paths_text: str,
    settings: Settings,
    *,
    tab_folder: Optional[str] = None,
    on_output: Optional[OutputSink] = None,
) -> ConvertResult:
    log: list[str] = []

    def emit(text: str, replace_last: bool = False) -> None:
        if on_output:
            on_output(text, replace_last)
        else:
            if replace_last and log:
                log[-1] = text
            else:
                log.append(text)

    err = validate_settings(settings)
    if err:
        return ConvertResult(False, err, log)

    exe = resolve_cjxl_exe(settings)
    if not exe:
        msg = (
            f"cjxl.exe not found at:\n{settings.cjxl_bin_dir}\\cjxl.exe\n\n"
            "Extract jxl-x64-windows-static.zip from libjxl releases."
        )
        return ConvertResult(False, msg, log)

    inputs = collect_image_inputs(paths_text, tab_folder)
    if not inputs:
        return ConvertResult(
            False,
            "No image files to convert.\n\n"
            "Select image files or folders, or run from a folder that contains images.",
            log,
        )

    magick = resolve_magick_exe(settings)
    needs_magick = False
    for p in inputs:
        if image_input_kind(p) == "preconvert":
            needs_magick = True
            break
    if not needs_magick and max_dimension_box(settings.max_dimension):
        if not magick:
            needs_magick = True
        else:
            for p in inputs:
                use_magick, _ = needs_magick_prepare(settings, p, magick)
                if use_magick:
                    needs_magick = True
                    break
    if needs_magick and not magick:
        msg = (
            f"magick.exe not found at:\n{settings.magick_bin_dir}\\magick.exe\n\n"
            "ImageMagick is required for format conversion and max-dimension limits."
        )
        return ConvertResult(False, msg, log)

    encode_args = build_cjxl_encode_args(settings)
    mode_label = (
        f"distance {settings.distance}"
        if settings.encode_mode == 1
        else f"quality {settings.quality}"
    )
    emit(f"cjxl: {len(inputs)} file(s), {mode_label} ({encode_args})")

    ok = 0
    fail = 0
    temp_seq = 0

    for i, input_path in enumerate(inputs):
        output_path = output_path_for_input(input_path)
        temp_png: Optional[Path] = None

        temp_seq += 1
        cjxl_input, temp_png, prep_note = resolve_cjxl_input_path(
            settings, input_path, temp_seq, magick
        )
        if not cjxl_input:
            fail += 1
            msg = f"ImageMagick prepare failed.\n\nStopped after:\n{input_path}"
            emit(f"cjxl: ImageMagick failed: {input_path}")
            return ConvertResult(False, msg, log, ok, fail)

        if prep_note:
            emit(f"cjxl: {prep_note}: {input_path}")

        cmd = f'"{exe}" {encode_args} "{cjxl_input}" "{output_path}"'
        emit(f"cjxl [{i + 1}/{len(inputs)}]: {cmd}")
        rc = _run_cjxl(exe, encode_args, cjxl_input, output_path, emit)

        if temp_png:
            try:
                if temp_png.is_file():
                    temp_png.unlink()
            except OSError:
                pass

        if rc == CJXL_EXIT_CONTROL_C:
            msg = f"Conversion cancelled.\n\nStopped after:\n{input_path}"
            emit(f"cjxl: cancelled (exit {rc})")
            return ConvertResult(False, msg, log, ok, fail)

        if rc != 0:
            fail += 1
            msg = f"cjxl exited with code {rc}.\n\nStopped after:\n{input_path}"
            emit(f"cjxl: failed (exit {rc}): {input_path}")
            return ConvertResult(False, msg, log, ok, fail)

        ok += 1
        emit(f"cjxl: OK -> {output_path}")

        if settings.replace_source and output_path.is_file() and input_path.is_file():
            recycle_delete(input_path)
            if input_path.is_file():
                emit(f"cjxl: could not recycle source: {input_path}")
            else:
                emit(f"cjxl: recycled source: {input_path}")

    summary = f"Converted {ok} file(s) to JPEG XL."
    emit(f"cjxl: done. {summary}")
    return ConvertResult(True, summary, log, ok, fail)


def format_convert_log(result: ConvertResult) -> str:
    lines = [result.summary, ""]
    if result.log:
        lines.extend(result.log)
    return "\n".join(lines)


def run_cli(argv: list[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="JPEG XL (cjxl) tool.")
    parser.add_argument("--gui", action="store_true", help="Open Dear PyGui GUI.")
    parser.add_argument("--repeat", action="store_true", help="Convert with last saved settings.")
    parser.add_argument("--only-list", metavar="FILE", help="UTF-8 file, one path per line.")
    parser.add_argument("--only-file", action="append", default=[], metavar="PATH")
    parser.add_argument("--tab-folder", metavar="DIR", help="Tab folder when nothing is selected.")
    args = parser.parse_args(argv)

    settings = config_load_settings()

    if args.gui:
        from cjxl_gui import run_gui

        run_gui(
            initial_only_list=args.only_list,
            initial_only_files=args.only_file or None,
            initial_tab_folder=args.tab_folder,
        )
        return 0

    paths_text = build_initial_files_text(
        settings.files_text,
        args.only_list,
        args.only_file or None,
        args.tab_folder,
    )
    if not paths_text.strip():
        paths_text = settings.files_text
    result = run_convert(
        paths_text,
        settings,
        tab_folder=args.tab_folder,
    )
    print(format_convert_log(result))
    return 0 if result.ok else 1
