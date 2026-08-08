"""Translate file names to English via the DeepSeek or Kimi API, with batch rename support."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

DEEPSEEK_API_URL = "https://api.deepseek.com/chat/completions"
KIMI_API_URL = "https://api.moonshot.ai/v1/chat/completions"

PROVIDERS = {
    "deepseek": {
        "label": "DeepSeek",
        "url": DEEPSEEK_API_URL,
        "default_model": "deepseek-chat",  # deepseek-reasoner is the smarter/slower alternative
        "env_key": "DEEPSEEK_API_KEY",
    },
    "kimi": {
        "label": "Kimi",
        "url": KIMI_API_URL,
        "default_model": "kimi-k3",
        "env_key": "KIMI_API_KEY",
    },
}
DEFAULT_PROVIDER = "kimi" #deepseek #kimi
DEFAULT_MODEL = PROVIDERS[DEFAULT_PROVIDER]["default_model"]


def default_model_for(provider: str) -> str:
    return PROVIDERS.get(provider, PROVIDERS[DEFAULT_PROVIDER])["default_model"]

MAX_NAME_LEN = 255  # Windows NTFS filename-component limit.

CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "TranslateFilename"
CONFIG_PATH = CONFIG_DIR / "settings.json"

SYSTEM_PROMPT = (
    "You translate file names into English and also extract a short natural-language "
    "phrase from that translation. You will be given a full file name, including its "
    "extension, in whatever language it is written in — the extension is context, not "
    "something to translate. Respond with ONLY a JSON object with exactly two keys:\n"
    '- "translation": the file name\'s title (i.e. everything except the extension) '
    "translated into natural, concise English, suitable for use as a file name. Do NOT "
    "include the extension in this value. Preserve any numbers, dates, episode/chapter "
    "markers, version tags, bracketed codes, and technical suffixes exactly as given, in "
    "their original position. If the title is already in English, return it unchanged.\n"
    '- "text_only": just the "natural language" portion of your translation'
    ' basically strip any non language elements leaving only the translated portion'
    ' you may include brackets or parentheses if they are part of the sentence or add context.'
    ' Most of the time, the entire translation may read as acceptable "natural language"'
    ' in that case, the translation itself is simply used for "text_only"\n'
    "Respond with raw JSON only — no markdown, no code fences, no explanations."
)

FEWSHOT_USER = "Female 샐루 심박수 170회 이상 (운동) [159,155-164bpm]_bpm_plot.mp4"
FEWSHOT_ASSISTANT = json.dumps(
    {
        "translation": "Female Fast Heart Rate 170+ (exercise) [159,155-164bpm]_bpm_plot",
        "text_only": "Female Fast Heart Rate 170+ (exercise)",
    },
    ensure_ascii=False,
)

INVALID_FILENAME_CHARS = '<>:"/\\|?*'


@dataclass
class Settings:
    provider: str = DEFAULT_PROVIDER
    api_key: str = ""  # DeepSeek key
    kimi_api_key: str = ""
    model: str = DEFAULT_MODEL
    auto_rename: bool = False
    append_mode: bool = False
    inputs_text: str = ""

    def active_api_key(self) -> str:
        return self.kimi_api_key if self.provider == "kimi" else self.api_key


@dataclass
class TranslateResult:
    ok: bool
    original: str = ""
    translation: str = ""
    text_only: str = ""
    error: str = ""


@dataclass
class BatchItemResult:
    path: str
    ok: bool
    original_name: str = ""
    new_name: str = ""
    warning: str = ""
    error: str = ""


# ---------------------------------------------------------------------------
# Config / persistent settings + untranslate history
# ---------------------------------------------------------------------------

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
    provider = str(data.get("provider") or DEFAULT_PROVIDER)
    if provider not in PROVIDERS:
        provider = DEFAULT_PROVIDER
    api_key = str(data.get("api_key") or os.environ.get("DEEPSEEK_API_KEY") or "")
    kimi_api_key = str(data.get("kimi_api_key") or os.environ.get("KIMI_API_KEY") or "")
    model = str(data.get("model") or default_model_for(provider)) or default_model_for(provider)
    return Settings(
        provider=provider,
        api_key=api_key,
        kimi_api_key=kimi_api_key,
        model=model,
        auto_rename=bool(data.get("auto_rename")),
        append_mode=bool(data.get("append_mode")),
        inputs_text=str(data.get("inputs_text") or ""),
    )


def config_save_settings(settings: Settings) -> None:
    data = config_read()
    data["provider"] = settings.provider
    data["api_key"] = settings.api_key
    data["kimi_api_key"] = settings.kimi_api_key
    data["model"] = settings.model
    data["auto_rename"] = settings.auto_rename
    data["append_mode"] = settings.append_mode
    data["inputs_text"] = settings.inputs_text
    try:
        config_write(data)
    except OSError:
        pass


def _history_key(path: Path) -> str:
    return os.fspath(path).casefold()


def get_history_entry(path: Path) -> Optional[dict]:
    data = config_read()
    return (data.get("history") or {}).get(_history_key(path))


def record_rename(old_path: Path, new_path: Path) -> None:
    data = config_read()
    history = data.setdefault("history", {})
    if not isinstance(history, dict):
        history = {}
        data["history"] = history
    existing = history.pop(_history_key(old_path), None)
    original_name = (existing or {}).get("original") or old_path.name
    history[_history_key(new_path)] = {
        "path": os.fspath(new_path),
        "original": original_name,
        "last_translated": new_path.name,
    }
    try:
        config_write(data)
    except OSError:
        pass


def record_untranslate(path: Path) -> None:
    data = config_read()
    history = data.get("history")
    if isinstance(history, dict):
        history.pop(_history_key(path), None)
        try:
            config_write(data)
        except OSError:
            pass


def forget_all_history() -> None:
    data = config_read()
    data["history"] = {}
    try:
        config_write(data)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Filename building / clamping
# ---------------------------------------------------------------------------

def sanitize_filename_component(s: str) -> str:
    out = "".join("_" if c in INVALID_FILENAME_CHARS else c for c in s)
    return out.strip().rstrip(" .")


def clamp_filename(stem: str, ext: str) -> tuple[str, bool]:
    """Return (name, was_truncated). Hard-truncates the stem to fit MAX_NAME_LEN."""
    name = stem + ext
    if len(name) <= MAX_NAME_LEN:
        return name, False
    budget = MAX_NAME_LEN - len(ext)
    if budget <= 0:
        return name[:MAX_NAME_LEN], True
    return stem[:budget] + ext, True


def build_output_filename(
    original_name: str, translation: str, text_only: str, append_mode: bool, is_dir: bool = False
) -> tuple[str, Optional[str]]:
    if is_dir:
        ext = ""
        base_stem = original_name
    else:
        p = Path(original_name)
        ext = p.suffix
        base_stem = p.stem
    if append_mode:
        candidate_stem = f"{base_stem} {text_only}".strip() if text_only else base_stem
    else:
        candidate_stem = translation or base_stem
    candidate_stem = sanitize_filename_component(candidate_stem) or base_stem
    name, truncated = clamp_filename(candidate_stem, ext)
    warning = f"Filename too long — truncated to fit {MAX_NAME_LEN} characters." if truncated else None
    return name, warning


# ---------------------------------------------------------------------------
# LLM API (DeepSeek / Kimi — both OpenAI-compatible chat/completions)
# ---------------------------------------------------------------------------

def _clean_json_content(text: str) -> str:
    s = (text or "").strip()
    if s.startswith("```"):
        s = s.strip("`")
        if s.lower().startswith("json"):
            s = s[4:]
    return s.strip()


def call_llm_translate(
    name: str, provider: str, api_key: str, model: str
) -> tuple[Optional[dict], Optional[str]]:
    info = PROVIDERS.get(provider, PROVIDERS[DEFAULT_PROVIDER])
    label = info["label"]
    if not api_key.strip():
        return None, f"{label} API key is not set."
    if not name.strip():
        return None, "Nothing to translate."

    import requests

    resolved_model = model or info["default_model"]
    # kimi-k3 is a reasoning model that only accepts temperature=1.
    temperature = 1 if resolved_model.startswith("kimi-k3") else 0.3

    payload = {
        "model": resolved_model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": FEWSHOT_USER},
            {"role": "assistant", "content": FEWSHOT_ASSISTANT},
            {"role": "user", "content": name},
        ],
        "temperature": temperature,
        "stream": False,
        "response_format": {"type": "json_object"},
    }
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(info["url"], json=payload, headers=headers, timeout=30)
    except requests.RequestException as exc:
        return None, f"Network error contacting {label}: {exc}"

    if resp.status_code != 200:
        detail = ""
        try:
            detail = resp.json().get("error", {}).get("message", "")
        except (ValueError, AttributeError):
            detail = resp.text[:200]
        return None, f"{label} API error {resp.status_code}: {detail}".strip()

    try:
        data = resp.json()
        content = data["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError):
        return None, f"Unexpected response from {label} API."

    try:
        parsed = json.loads(_clean_json_content(content))
    except (ValueError, TypeError):
        return None, f"{label} returned malformed JSON."

    translation = str(parsed.get("translation") or "").strip()
    text_only = str(parsed.get("text_only") or "").strip()
    if not translation:
        return None, f"{label} returned an empty translation."
    return {"translation": translation, "text_only": text_only}, None


def translate_name(name: str, settings: Settings) -> TranslateResult:
    data, err = call_llm_translate(name, settings.provider, settings.active_api_key(), settings.model)
    if err:
        return TranslateResult(ok=False, original=name, error=err)
    return TranslateResult(
        ok=True,
        original=name,
        translation=data["translation"],
        text_only=data["text_only"],
    )


# ---------------------------------------------------------------------------
# Rename / untranslate operations
# ---------------------------------------------------------------------------

def _long_path(path: Path) -> str:
    """Prefix with \\\\?\\ so Windows accepts paths beyond MAX_PATH (260 chars)."""
    s = os.fspath(path.resolve())
    if os.name == "nt" and not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def rename_file_apply(path: Path, new_name: str) -> tuple[Optional[Path], Optional[str]]:
    if not path.exists():
        return None, "File or folder not found."
    new_path = path.with_name(new_name)
    if new_path == path:
        return path, None
    if new_path.exists():
        return None, f"Target already exists: {new_name}"
    try:
        os.rename(_long_path(path), _long_path(new_path))
    except OSError as exc:
        return None, f"Rename failed: {exc}"
    record_rename(path, new_path)
    return new_path, None


def untranslate_file(path: Path) -> tuple[Optional[Path], Optional[str]]:
    if not path.exists():
        return None, "File or folder not found."
    hist = get_history_entry(path)
    if not hist:
        return None, "No translation history for this file."
    original_name = str(hist.get("original") or "").strip()
    if not original_name:
        return None, "No original name recorded."
    new_path = path.with_name(original_name)
    if new_path == path:
        record_untranslate(path)
        return path, None
    if new_path.exists():
        return None, f"Cannot revert — target already exists: {original_name}"
    try:
        os.rename(_long_path(path), _long_path(new_path))
    except OSError as exc:
        return None, f"Revert failed: {exc}"
    record_untranslate(path)
    return new_path, None


def translate_and_maybe_rename(
    path: Path, settings: Settings, rename: bool, append_mode: bool
) -> BatchItemResult:
    if not path.exists():
        return BatchItemResult(path=str(path), ok=False, error="File or folder not found.")
    is_dir = path.is_dir()
    result = translate_name(path.name, settings)
    if not result.ok:
        return BatchItemResult(path=str(path), ok=False, original_name=path.name, error=result.error)
    new_name, warning = build_output_filename(
        path.name, result.translation, result.text_only, append_mode, is_dir=is_dir
    )
    if not rename:
        return BatchItemResult(
            path=str(path), ok=True, original_name=path.name, new_name=new_name, warning=warning or ""
        )
    new_path, err = rename_file_apply(path, new_name)
    if err:
        return BatchItemResult(
            path=str(path), ok=False, original_name=path.name, new_name=new_name, error=err
        )
    return BatchItemResult(
        path=str(new_path), ok=True, original_name=path.name, new_name=new_name, warning=warning or ""
    )


# ---------------------------------------------------------------------------
# Batch input helpers (shared by GUI + CLI)
# ---------------------------------------------------------------------------

def dedupe_lines(lines: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for ln in lines:
        key = ln.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(ln)
    return out


def paths_from_only_list(list_path: Optional[str], only_files: Optional[list[str]]) -> list[Path]:
    lines: list[str] = []
    if list_path:
        try:
            lines.extend(Path(list_path).read_text(encoding="utf-8-sig").splitlines())
        except OSError:
            pass
    if only_files:
        lines.extend(only_files)
    out: list[Path] = []
    seen: set[str] = set()
    for ln in dedupe_lines([s.strip() for s in lines if s.strip()]):
        p = Path(ln)
        if p.is_file() or p.is_dir():
            key = os.fspath(p).casefold()
            if key not in seen:
                seen.add(key)
                out.append(p)
    return out


def build_initial_inputs_text(
    saved: str, only_list: Optional[str], only_files: Optional[list[str]]
) -> str:
    from_dopus = paths_from_only_list(only_list, only_files)
    if from_dopus:
        return "\n".join(os.fspath(p) for p in from_dopus)
    return saved.strip()


def run_cli(argv: list[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Translate file names to English via DeepSeek or Kimi.")
    parser.add_argument("--gui", action="store_true", help="Open Dear PyGui GUI.")
    parser.add_argument("--only-list", metavar="FILE", help="UTF-8 file, one path per line.")
    parser.add_argument("--only-file", action="append", default=[], metavar="PATH")
    args = parser.parse_args(argv)

    settings = config_load_settings()

    if args.gui:
        from translate_gui import run_gui
        run_gui(initial_only_list=args.only_list, initial_only_files=args.only_file or None)
        return 0

    paths = paths_from_only_list(args.only_list, args.only_file or None)
    if not paths:
        print("No files given. Pass --only-list/--only-file, or use --gui.")
        return 1

    failures = 0
    for path in paths:
        result = translate_and_maybe_rename(path, settings, rename=True, append_mode=settings.append_mode)
        if result.ok:
            note = f" ({result.warning})" if result.warning else ""
            print(f"{result.original_name} -> {result.new_name}{note}")
        else:
            failures += 1
            print(f"{result.original_name or path.name}: ERROR - {result.error}")
    return 0 if not failures else 1
