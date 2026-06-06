"""Bulk image conversion via ImageMagick."""

from __future__ import annotations

import concurrent.futures
import json
import math
import os
import subprocess
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

_REPO_SHARED = Path(__file__).resolve().parent.parent / "Shared"
if _REPO_SHARED.is_dir() and str(_REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(_REPO_SHARED))
from recycle_delete import recycle_delete

DEFAULT_MAGICK_BIN_DIR = (
    r"C:\Users\WXP\Desktop\Tools\ImageMagick-7.1.2-25-portable-Q16-HDRI-x64"
)

MAGICK_EXIT_CONTROL_C = -1073741510

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

ICO_SIZE_PRESETS: dict[str, list[int]] = {
    "16": [16],
    "32": [32],
    "48": [48],
    "64": [64],
    "128": [128],
    "256": [256],
    "512": [512],
}

ICO_SIZE_LABELS: dict[str, str] = {
    "16": "16",
    "32": "32",
    "48": "48",
    "64": "64",
    "128": "128",
    "256": "256",
    "512": "512",
}

ICO_SIZE_KEYS = tuple(ICO_SIZE_PRESETS.keys())

# All extensions ImageMagick can typically read
IMAGE_EXTS = {
    ".png", ".jpg", ".jpeg", ".jpe", ".jfif",
    ".gif", ".bmp", ".dib", ".tif", ".tiff",
    ".webp", ".ico", ".cur", ".heic", ".heif",
    ".avif", ".psd", ".jp2", ".j2k", ".tga",
    ".pcx", ".wbmp", ".ppm", ".pgm", ".pbm",
    ".pnm", ".pfm", ".pam", ".pgx", ".exr",
    ".jxl",
}

OUTPUT_FORMATS: dict[str, dict] = {
    "jpeg": {"label": "JPEG", "ext": ".jpg",  "encode": "quality"},
    "png":  {"label": "PNG",  "ext": ".png",  "encode": "none"},
    "webp": {"label": "WebP", "ext": ".webp", "encode": "quality"},
    "gif":  {"label": "GIF",  "ext": ".gif",  "encode": "none"},
    "bmp":  {"label": "BMP",  "ext": ".bmp",  "encode": "none"},
    "tiff": {"label": "TIFF", "ext": ".tiff", "encode": "none"},
    "avif": {"label": "AVIF", "ext": ".avif", "encode": "quality"},
    "jxl":  {"label": "JXL",  "ext": ".jxl",  "encode": "jxl"},
    "ico":  {"label": "ICO",  "ext": ".ico",  "encode": "ico"},
}

OUTPUT_FORMAT_KEYS = tuple(OUTPUT_FORMATS.keys())
OUTPUT_FORMAT_LABELS = [OUTPUT_FORMATS[k]["label"] for k in OUTPUT_FORMAT_KEYS]

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "ImageConverter"
CONFIG_PATH = CONFIG_DIR / "settings.json"

GUI_SECTION_DEFAULTS: dict[str, bool] = {
    "files": True,
    "encode": True,
    "resize": True,
}

OutputSink = Callable[[str, bool], None]


@dataclass
class Settings:
    output_format: str = "jpeg"
    quality: str = "90"
    jxl_distance: str = "1"
    jxl_effort: str = "7"
    replace_source: bool = False
    magick_bin_dir: str = DEFAULT_MAGICK_BIN_DIR
    max_dimension: str = "none"
    ico_sizes: str = "256"
    resize_width: str = ""
    files_text: str = ""
    workers: int = field(default_factory=lambda: os.cpu_count() or 4)
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


def is_image_input_path(path: Path, output_ext: str) -> bool:
    ext = file_ext_lower(path)
    if ext not in IMAGE_EXTS:
        return False
    return ext != output_ext.lower()


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


def list_image_files_in_folder(
    folder: Path, output_ext: str, out: list[Path], seen: set[str]
) -> None:
    if not folder.is_dir():
        return
    try:
        for child in folder.iterdir():
            if child.is_file():
                if not is_image_input_path(child, output_ext):
                    continue
                key = os.fspath(child).casefold()
                if key not in seen:
                    seen.add(key)
                    out.append(child.resolve())
            elif child.is_dir():
                list_image_files_in_folder(child, output_ext, out, seen)
    except OSError:
        pass


def collect_image_inputs(
    paths_text: str, output_ext: str, tab_folder: Optional[str] = None
) -> list[Path]:
    out: list[Path] = []
    seen: set[str] = set()
    listed = paths_from_text_lines(paths_text)

    for p in listed:
        if p.is_file():
            if is_image_input_path(p, output_ext):
                key = os.fspath(p).casefold()
                if key not in seen:
                    seen.add(key)
                    out.append(p)
        elif p.is_dir():
            list_image_files_in_folder(p, output_ext, out, seen)

    if not out and not listed:
        tab = (tab_folder or "").strip()
        if tab:
            list_image_files_in_folder(Path(tab).resolve(), output_ext, out, seen)

    return out


def config_read() -> dict:
    if CONFIG_PATH.is_file():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            pass
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
    output_format = str(data.get("output_format") or "jpeg")
    if output_format not in OUTPUT_FORMATS:
        output_format = "jpeg"
    max_dim = str(data.get("max_dimension") or "none")
    if max_dim not in MAX_DIMENSION_PRESETS:
        max_dim = "none"
    ico_sizes = str(data.get("ico_sizes") or "256")
    if ico_sizes not in ICO_SIZE_PRESETS:
        ico_sizes = "256"
    resize_width = str(data.get("resize_width") or "")
    cpu_default = os.cpu_count() or 4
    saved_workers = int(data.get("workers") or 0)
    workers = saved_workers if saved_workers >= 1 else cpu_default
    return Settings(
        output_format=output_format,
        quality=str(data.get("quality") or "90") or "90",
        jxl_distance=str(data.get("jxl_distance") or "1") or "1",
        jxl_effort=str(data.get("jxl_effort") or "7") or "7",
        replace_source=bool(data.get("replace_source")),
        magick_bin_dir=str(data.get("magick_bin_dir") or DEFAULT_MAGICK_BIN_DIR) or DEFAULT_MAGICK_BIN_DIR,
        max_dimension=max_dim,
        ico_sizes=ico_sizes,
        resize_width=resize_width,
        files_text=str(data.get("files_text") or ""),
        workers=workers,
        gui_sections=gui_sections,
    )


def config_save_settings(settings: Settings) -> None:
    data = config_read()
    data["output_format"] = settings.output_format
    data["quality"] = settings.quality
    data["jxl_distance"] = settings.jxl_distance
    data["jxl_effort"] = settings.jxl_effort
    data["replace_source"] = settings.replace_source
    data["magick_bin_dir"] = settings.magick_bin_dir
    data["max_dimension"] = settings.max_dimension
    data["ico_sizes"] = settings.ico_sizes
    data["resize_width"] = settings.resize_width
    data["workers"] = settings.workers
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
    return str(int(n)), None


def parse_jxl_distance(raw: str) -> tuple[Optional[str], Optional[str]]:
    s = raw.strip()
    try:
        n = float(s)
    except ValueError:
        return None, "Distance must be a number from 0 to 25."
    if n < 0 or n > 25:
        return None, "Distance must be between 0 and 25."
    return str(n), None


def parse_jxl_effort(raw: str) -> tuple[Optional[str], Optional[str]]:
    s = raw.strip()
    if not s:
        return "7", None
    try:
        n = int(s)
    except ValueError:
        return None, "Effort must be a whole number from 0 to 9."
    if n < 0 or n > 9:
        return None, "Effort must be between 0 and 9."
    return str(n), None


def validate_settings(settings: Settings) -> Optional[str]:
    fmt = OUTPUT_FORMATS.get(settings.output_format)
    if not fmt:
        return f"Unknown output format: {settings.output_format}"
    if fmt["encode"] == "quality":
        _, err = parse_quality(settings.quality)
        if err:
            return err
    if fmt["encode"] == "jxl":
        _, err = parse_jxl_distance(settings.jxl_distance)
        if err:
            return err
        _, err = parse_jxl_effort(settings.jxl_effort)
        if err:
            return err
    if fmt["encode"] == "ico":
        if settings.ico_sizes not in ICO_SIZE_PRESETS:
            return f"Unknown ICO size preset: {settings.ico_sizes}"
    return None


def validate_resize_settings(settings: Settings) -> Optional[str]:
    rw = settings.resize_width.strip()
    if not rw:
        return "Resize width is required for resizing."
    try:
        w = int(rw)
        if w < 1:
            return "Resize width must be at least 1."
    except ValueError:
        return "Resize width must be a positive integer."
    return None


def effective_workers(settings: Settings) -> int:
    return max(1, settings.workers)


def _jxl_distance_to_quality(distance: float) -> int:
    # ImageMagick's JXL coder ignores -define jxl:distance and only reads -quality.
    # Invert its internal quality→distance formula so the user-facing distance value
    # maps to the correct -quality argument.
    if distance <= 0.1:
        return 100
    if distance <= 6.4:
        q = 100.0 - (distance - 0.1) / 0.09
        return max(30, min(100, round(q)))
    q = 30.0 - 5.0 * math.log((distance - 6.4) * 6.25) / math.log(2.5)
    return max(0, min(29, round(q)))


def build_magick_encode_args(settings: Settings) -> str:
    fmt = OUTPUT_FORMATS[settings.output_format]
    parts: list[str] = []
    if fmt["encode"] == "quality":
        parts.append(f"-quality {settings.quality}")
    elif fmt["encode"] == "jxl":
        dist, _ = parse_jxl_distance(settings.jxl_distance)
        quality = _jxl_distance_to_quality(float(dist or "1"))
        parts.append(f"-quality {quality}")
        effort, _ = parse_jxl_effort(settings.jxl_effort)
        parts.append(f"-define jxl:effort={effort or '7'}")
    elif fmt["encode"] == "ico":
        sizes = ico_sizes_list(settings.ico_sizes)
        parts.append(f"-define icon:auto-resize={','.join(str(s) for s in sizes)}")
    return " ".join(parts)


def output_path_for_input(input_path: Path, settings: Settings) -> Path:
    ext = OUTPUT_FORMATS[settings.output_format]["ext"]
    return input_path.with_suffix(ext)


def max_dimension_box(key: str) -> tuple[int, int] | None:
    return MAX_DIMENSION_PRESETS.get(key, MAX_DIMENSION_PRESETS["none"])


def max_dimension_label(key: str) -> str:
    return MAX_DIMENSION_LABELS.get(key, MAX_DIMENSION_LABELS["none"])


def max_dimension_key_from_label(label: str) -> str:
    for key, text in MAX_DIMENSION_LABELS.items():
        if text == label:
            return key
    return "none"


def ico_sizes_list(key: str) -> list[int]:
    return ICO_SIZE_PRESETS.get(key, ICO_SIZE_PRESETS["256"])


def ico_sizes_label(key: str) -> str:
    return ICO_SIZE_LABELS.get(key, ICO_SIZE_LABELS["256"])


def ico_sizes_key_from_label(label: str) -> str:
    for key, text in ICO_SIZE_LABELS.items():
        if text == label:
            return key
    return "256"


def resolve_magick_exe(settings: Settings) -> Optional[Path]:
    bin_dir = Path(os.path.expandvars(settings.magick_bin_dir.strip()))
    exe = bin_dir / "magick.exe"
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
                    line, carry = carry[:n], carry[n + 1:]
                    line = line.strip("\r")
                    if line:
                        emit(line, False)
                else:
                    line, carry = carry[:r], carry[r + 1:]
                    if line:
                        emit(line, True)
        tail = carry.strip("\r\n")
        if tail:
            emit(tail, False)
    except OSError:
        pass


def _run_magick(
    exe: Path,
    input_path: Path,
    resize_arg: str,
    encode_args: str,
    output_path: Path,
    emit: OutputSink,
    thread_limit: int = 0,
) -> int:
    parts = [f'"{exe}"']
    if thread_limit > 0:
        parts.append(f"-limit thread {thread_limit}")
    parts.append(f'"{input_path}"')
    if resize_arg:
        parts.append(resize_arg)
    if encode_args:
        parts.append(encode_args)
    parts.append(f'"{output_path}"')
    cmd = " ".join(parts)
    try:
        proc = subprocess.Popen(
            cmd,
            shell=True,
            cwd=os.fspath(exe.parent),
            stderr=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            creationflags=_win_subprocess_flags(),
        )
        reader = threading.Thread(
            target=_stream_process_output, args=(proc, emit), daemon=True
        )
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
    log_lock = threading.Lock()

    def emit(text: str, replace_last: bool = False) -> None:
        if on_output:
            on_output(text, replace_last)
        else:
            with log_lock:
                if replace_last and log:
                    log[-1] = text
                else:
                    log.append(text)

    err = validate_settings(settings)
    if err:
        return ConvertResult(False, err, log)

    magick = resolve_magick_exe(settings)
    if not magick:
        msg = (
            f"magick.exe not found at:\n{settings.magick_bin_dir}\\magick.exe\n\n"
            "Download portable ImageMagick 7 and update the folder in settings."
        )
        return ConvertResult(False, msg, log)

    fmt = OUTPUT_FORMATS[settings.output_format]
    output_ext = fmt["ext"]
    inputs = collect_image_inputs(paths_text, output_ext, tab_folder)
    if not inputs:
        return ConvertResult(
            False,
            "No image files to convert.\n\n"
            "Select image files or folders, or run from a folder that contains images.\n"
            f"Files already in {fmt['label']} format are skipped.",
            log,
        )

    encode_args = build_magick_encode_args(settings)
    box = max_dimension_box(settings.max_dimension)
    resize_arg = f'-filter Lanczos -resize "{box[0]}x{box[1]}>"' if box else ""

    workers_count = effective_workers(settings)
    cpu_count = os.cpu_count() or 4
    # When running multiple workers, cap threads per process so they share cores evenly.
    thread_limit = max(1, cpu_count // workers_count) if workers_count > 1 else 0
    total = len(inputs)

    emit(f"convert: {total} file(s) → {fmt['label']} ({encode_args or 'default'}) — {workers_count} worker(s)")

    ok = 0
    fail = 0
    cancel_event = threading.Event()

    def convert_one(i: int, input_path: Path) -> tuple[int, Path, Path, int]:
        if cancel_event.is_set():
            return i, input_path, output_path_for_input(input_path, settings), MAGICK_EXIT_CONTROL_C
        output_path = output_path_for_input(input_path, settings)
        emit(f"convert [{i + 1}/{total}]: {input_path.name} → {output_path.name}")
        # Suppress replace_last in parallel mode to avoid interleaved display glitches.
        worker_emit: OutputSink = (lambda t, r: emit(t, False)) if workers_count > 1 else emit
        rc = _run_magick(magick, input_path, resize_arg, encode_args, output_path, worker_emit, thread_limit)
        return i, input_path, output_path, rc

    error_result: Optional[ConvertResult] = None
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers_count) as pool:
        fs = [pool.submit(convert_one, i, path) for i, path in enumerate(inputs)]
        for future in concurrent.futures.as_completed(fs):
            _, input_path, output_path, rc = future.result()

            if cancel_event.is_set() and rc == MAGICK_EXIT_CONTROL_C:
                continue

            if rc == MAGICK_EXIT_CONTROL_C:
                cancel_event.set()
                for f in fs:
                    f.cancel()
                msg = f"Conversion cancelled.\n\nStopped after:\n{input_path}"
                emit(f"convert: cancelled (exit {rc})")
                error_result = ConvertResult(False, msg, log, ok, fail)
                break

            if rc != 0:
                cancel_event.set()
                fail += 1
                for f in fs:
                    f.cancel()
                msg = f"magick exited with code {rc}.\n\nStopped after:\n{input_path}"
                emit(f"convert: failed (exit {rc}): {input_path}")
                error_result = ConvertResult(False, msg, log, ok, fail)
                break

            ok += 1
            emit(f"convert: OK → {output_path}")

            if settings.replace_source and output_path.is_file() and input_path.is_file():
                recycle_delete(input_path)
                if input_path.is_file():
                    emit(f"convert: could not recycle source: {input_path}")
                else:
                    emit(f"convert: recycled source: {input_path}")

    if error_result is not None:
        return error_result

    summary = f"Converted {ok} file(s) to {fmt['label']}."
    emit(f"convert: done. {summary}")
    return ConvertResult(True, summary, log, ok, fail)


def run_resize(
    paths_text: str,
    settings: Settings,
    *,
    tab_folder: Optional[str] = None,
    on_output: Optional[OutputSink] = None,
) -> ConvertResult:
    log: list[str] = []
    log_lock = threading.Lock()

    def emit(text: str, replace_last: bool = False) -> None:
        if on_output:
            on_output(text, replace_last)
        else:
            with log_lock:
                if replace_last and log:
                    log[-1] = text
                else:
                    log.append(text)

    err = validate_resize_settings(settings)
    if err:
        return ConvertResult(False, err, log)

    magick = resolve_magick_exe(settings)
    if not magick:
        msg = (
            f"magick.exe not found at:\n{settings.magick_bin_dir}\\magick.exe\n\n"
            "Download portable ImageMagick 7 and update the folder in settings."
        )
        return ConvertResult(False, msg, log)

    inputs = collect_image_inputs(paths_text, "", tab_folder)
    if not inputs:
        return ConvertResult(
            False,
            "No image files to resize.\n\n"
            "Select image files or folders, or run from a folder that contains images.",
            log,
        )

    rw = settings.resize_width.strip()
    width = int(rw)
    resize_arg = f"-resize {width}x"

    workers_count = effective_workers(settings)
    cpu_count = os.cpu_count() or 4
    thread_limit = max(1, cpu_count // workers_count) if workers_count > 1 else 0
    total = len(inputs)

    emit(f"resize: {total} file(s) → width {width}px — {workers_count} worker(s)")

    ok = 0
    fail = 0
    cancel_event = threading.Event()

    def resize_one(i: int, input_path: Path) -> tuple[int, Path, Path, int]:
        if cancel_event.is_set():
            output_path = input_path.parent / (input_path.stem + "_resized" + input_path.suffix)
            return i, input_path, output_path, MAGICK_EXIT_CONTROL_C
        output_path = input_path.parent / (input_path.stem + "_resized" + input_path.suffix)
        emit(f"resize [{i + 1}/{total}]: {input_path.name} → {output_path.name}")
        worker_emit: OutputSink = (lambda t, r: emit(t, False)) if workers_count > 1 else emit
        rc = _run_magick(magick, input_path, resize_arg, "", output_path, worker_emit, thread_limit)
        return i, input_path, output_path, rc

    error_result: Optional[ConvertResult] = None
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers_count) as pool:
        fs = [pool.submit(resize_one, i, path) for i, path in enumerate(inputs)]
        for future in concurrent.futures.as_completed(fs):
            _, input_path, output_path, rc = future.result()

            if cancel_event.is_set() and rc == MAGICK_EXIT_CONTROL_C:
                continue

            if rc == MAGICK_EXIT_CONTROL_C:
                cancel_event.set()
                for f in fs:
                    f.cancel()
                msg = f"Resize cancelled.\n\nStopped after:\n{input_path}"
                emit(f"resize: cancelled (exit {rc})")
                error_result = ConvertResult(False, msg, log, ok, fail)
                break

            if rc != 0:
                cancel_event.set()
                fail += 1
                for f in fs:
                    f.cancel()
                msg = f"magick exited with code {rc}.\n\nStopped after:\n{input_path}"
                emit(f"resize: failed (exit {rc}): {input_path}")
                error_result = ConvertResult(False, msg, log, ok, fail)
                break

            ok += 1
            emit(f"resize: OK → {output_path}")

            if settings.replace_source and output_path.is_file() and input_path.is_file():
                recycle_delete(input_path)
                if input_path.is_file():
                    emit(f"resize: could not recycle source: {input_path}")
                else:
                    emit(f"resize: recycled source: {input_path}")

    if error_result is not None:
        return error_result

    summary = f"Resized {ok} file(s) to width {width}px."
    emit(f"resize: done. {summary}")
    return ConvertResult(True, summary, log, ok, fail)


def format_convert_log(result: ConvertResult) -> str:
    lines = [result.summary, ""]
    if result.log:
        lines.extend(result.log)
    return "\n".join(lines)


def run_cli(argv: list[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Image converter (ImageMagick).")
    parser.add_argument("--gui", action="store_true", help="Open Dear PyGui GUI.")
    parser.add_argument("--repeat", action="store_true", help="Convert with last saved settings.")
    parser.add_argument("--only-list", metavar="FILE", help="UTF-8 file, one path per line.")
    parser.add_argument("--only-file", action="append", default=[], metavar="PATH")
    parser.add_argument("--tab-folder", metavar="DIR", help="Tab folder when nothing is selected.")
    args = parser.parse_args(argv)

    settings = config_load_settings()

    if args.gui:
        from converter_gui import run_gui
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
