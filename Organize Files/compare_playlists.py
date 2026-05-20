"""Compare two text files line-by-line: exact match, then word-based fuzzy.

Library API: run_compare(), format_compare_report()
CLI: python compare_playlists.py file_a.txt file_b.txt [-o output_dir]
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from rapidfuzz import fuzz

FUZZY_THRESHOLD = 80
_WIN_BAD_FILENAME = re.compile(r'[<>:"/\\|?*]')
_WORD_SPLIT = re.compile(r"[\s,]+")
AUDIO_EXTS = {".mp3", ".opus", ".m4a", ".flac", ".wav", ".ogg", ".aac", ".wma", ".webm"}


@dataclass
class Line:
    raw: str


@dataclass(frozen=True)
class CompareOptions:
    strip_extensions: bool = True
    romaji_compare: bool = True


_kakasi_instance = None


def missing_from_report_path(output_dir: Path, other_file: Path) -> Path:
    stem = _WIN_BAD_FILENAME.sub("", other_file.stem).strip() or "file"
    return output_dir / f"missing-from-{stem}.txt"


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("utf-8", errors="replace")


def read_lines(path: Path) -> list[Line]:
    lines = []
    for raw in read_text(path).splitlines():
        line = raw.strip()
        if line:
            lines.append(Line(line))
    return lines


def strip_audio_ext(text: str) -> str:
    path = Path(text)
    if path.suffix.lower() in AUDIO_EXTS:
        return path.stem
    return text


def compare_text(line: str, *, strip_extensions: bool) -> str:
    text = line.strip()
    if strip_extensions:
        text = strip_audio_ext(text)
    return text


def exact_key(line: str, *, strip_extensions: bool) -> str:
    return compare_text(line, strip_extensions=strip_extensions).casefold()


def split_words(line: str) -> list[str]:
    return [w for w in _WORD_SPLIT.split(line) if w]


def word_pair_score(a: str, b: str) -> float:
    if a.casefold() == b.casefold():
        return 100.0
    return float(fuzz.ratio(a.casefold(), b.casefold()))


def word_coverage(
    source: list[str], target: list[str], word_threshold: float
) -> float:
    """Share of ``source`` words that match a word in ``target``."""
    if not source:
        return 100.0 if not target else 0.0

    used: set[int] = set()
    matched = 0
    for word in source:
        best = -1.0
        best_j = -1
        for j, other in enumerate(target):
            if j in used:
                continue
            score = word_pair_score(word, other)
            if score > best:
                best = score
                best_j = j
        if best >= word_threshold and best_j >= 0:
            matched += 1
            used.add(best_j)

    return 100.0 * matched / len(source)


def _kakasi():
    global _kakasi_instance
    if _kakasi_instance is None:
        import pykakasi

        _kakasi_instance = pykakasi.kakasi()
    return _kakasi_instance


def to_romaji(text: str) -> str:
    """Japanese (and kana) to Hepburn romaji; ASCII passes through."""
    try:
        items = _kakasi().convert(text)
    except ImportError:
        return text
    parts = [item.get("hepburn", "") for item in items if item.get("hepburn")]
    return " ".join(parts)


def _word_match_score_text(text_a: str, text_b: str, threshold: float) -> float:
    wa = split_words(text_a)
    wb = split_words(text_b)
    if not wa and not wb:
        return 100.0
    if not wa or not wb:
        return 0.0
    return max(word_coverage(wa, wb, threshold), word_coverage(wb, wa, threshold))


def word_match_score(
    line_a: str,
    line_b: str,
    threshold: float,
    options: CompareOptions,
) -> float:
    """Max word coverage either way; optional romaji pass takes the higher score."""
    text_a = compare_text(line_a, strip_extensions=options.strip_extensions)
    text_b = compare_text(line_b, strip_extensions=options.strip_extensions)
    score = _word_match_score_text(text_a, text_b, threshold)
    if not options.romaji_compare:
        return score
    romaji_a = to_romaji(text_a)
    romaji_b = to_romaji(text_b)
    return max(score, _word_match_score_text(romaji_a, romaji_b, threshold))


def build_exact_index(lines: list[Line], options: CompareOptions) -> dict[str, list[Line]]:
    index: dict[str, list[Line]] = {}
    for item in lines:
        index.setdefault(
            exact_key(item.raw, strip_extensions=options.strip_extensions), []
        ).append(item)
    return index


def find_exact_match(
    line: Line, index: dict[str, list[Line]], options: CompareOptions
) -> Line | None:
    hits = index.get(exact_key(line.raw, strip_extensions=options.strip_extensions))
    return hits[0] if hits else None


def find_best_fuzzy(
    line: Line,
    candidates: list[Line],
    used: set[int],
    word_threshold: float,
    options: CompareOptions,
) -> tuple[Line, float] | None:
    best: Line | None = None
    best_score = -1.0
    for cand in candidates:
        if id(cand) in used:
            continue
        score = word_match_score(line.raw, cand.raw, word_threshold, options)
        if score > best_score:
            best_score = score
            best = cand
    if best is None or best_score < 0:
        return None
    return best, best_score


def find_fuzzy_match(
    line: Line,
    candidates: list[Line],
    used: set[int],
    threshold: float,
    options: CompareOptions,
) -> tuple[Line, float] | None:
    best = find_best_fuzzy(line, candidates, used, threshold, options)
    if best is None:
        return None
    other, score = best
    if score >= threshold:
        return other, score
    return None


def collect_fuzzy_mismatches(
    lines: list[Line],
    candidates: list[Line],
    threshold: float,
    options: CompareOptions,
) -> list[tuple[Line, Line, float]]:
    """Best candidate pairs with line score below ``threshold`` (debug)."""
    pairs: list[tuple[Line, Line, float]] = []
    for line in lines:
        best = find_best_fuzzy(line, candidates, set(), threshold, options)
        if best is None:
            continue
        other, score = best
        if score < threshold:
            pairs.append((line, other, score))
    return pairs


def match_a_to_b(
    lines_a: list[Line],
    lines_b: list[Line],
    index_b: dict[str, list[Line]],
    threshold: float,
    options: CompareOptions,
) -> tuple[list[Line], list[tuple[Line, Line, float]], set[int]]:
    """Lines in A missing from B, fuzzy pairs, used ids from B."""
    used_b: set[int] = set()
    missing: list[Line] = []
    fuzzy_pairs: list[tuple[Line, Line, float]] = []

    for line in lines_a:
        hit = find_exact_match(line, index_b, options)
        if hit is not None and id(hit) not in used_b:
            used_b.add(id(hit))
            continue

        fuzzy = find_fuzzy_match(line, lines_b, used_b, threshold, options)
        if fuzzy is not None:
            other, score = fuzzy
            used_b.add(id(other))
            fuzzy_pairs.append((line, other, score))
            continue

        missing.append(line)

    return missing, fuzzy_pairs, used_b


def write_lines(path: Path, lines: list[Line]) -> None:
    path.write_text(
        "\n".join(item.raw for item in lines) + ("\n" if lines else ""),
        encoding="utf-8",
    )


def write_fuzzy_matches(path: Path, pairs: list[tuple[Line, Line, float]]) -> None:
    out = []
    for left, right, score in pairs:
        out.append(f"{score:.0f}%  {left.raw}")
        out.append(f"       -> {right.raw}")
    path.write_text("\n".join(out) + ("\n" if out else ""), encoding="utf-8")


@dataclass
class CompareResult:
    file_a: Path
    file_b: Path
    output_dir: Path
    threshold: int
    count_a: int
    count_b: int
    exact_matched: int
    missing_from_b: list[Line]
    missing_from_a: list[Line]
    fuzzy_pairs: list[tuple[Line, Line, float]]
    fuzzy_mismatches: list[tuple[Line, Line, float]]
    missing_from_b_path: Path
    missing_from_a_path: Path
    fuzzy_path: Path | None
    fuzzy_mismatches_path: Path | None
    error: str | None = None

    @property
    def total_matched(self) -> int:
        return self.count_a - len(self.missing_from_b)

    @property
    def fuzzy_matched(self) -> int:
        return len(self.fuzzy_pairs)


def run_compare(
    file_a: Path,
    file_b: Path,
    *,
    output_dir: Path | None = None,
    threshold: int = FUZZY_THRESHOLD,
    write_reports: bool = True,
    debug: bool = False,
    strip_extensions: bool = True,
    romaji_compare: bool = True,
) -> CompareResult:
    options = CompareOptions(
        strip_extensions=strip_extensions,
        romaji_compare=romaji_compare,
    )
    out_default = output_dir or file_a.parent
    missing_from_b_path = missing_from_report_path(out_default, file_b)
    missing_from_a_path = missing_from_report_path(out_default, file_a)
    fuzzy_path = out_default / "fuzzy-matches.txt" if debug else None
    fuzzy_mismatches_path = out_default / "fuzzy-mismatches.txt" if debug else None

    base = CompareResult(
        file_a=file_a,
        file_b=file_b,
        output_dir=out_default,
        threshold=threshold,
        count_a=0,
        count_b=0,
        exact_matched=0,
        missing_from_b=[],
        missing_from_a=[],
        fuzzy_pairs=[],
        fuzzy_mismatches=[],
        missing_from_b_path=missing_from_b_path,
        missing_from_a_path=missing_from_a_path,
        fuzzy_path=fuzzy_path,
        fuzzy_mismatches_path=fuzzy_mismatches_path,
    )

    if not file_a.is_file():
        base.error = f"File not found:\n{file_a}"
        return base
    if not file_b.is_file():
        base.error = f"File not found:\n{file_b}"
        return base

    lines_a = read_lines(file_a)
    lines_b = read_lines(file_b)
    index_b = build_exact_index(lines_b, options)

    missing_from_b, fuzzy_pairs, used_b = match_a_to_b(
        lines_a, lines_b, index_b, float(threshold), options
    )
    missing_from_a = [line for line in lines_b if id(line) not in used_b]
    exact_matched = len(lines_a) - len(missing_from_b) - len(fuzzy_pairs)

    thr = float(threshold)
    fuzzy_mismatches: list[tuple[Line, Line, float]] = []
    if debug:
        fuzzy_mismatches = collect_fuzzy_mismatches(
            missing_from_b, lines_b, thr, options
        ) + collect_fuzzy_mismatches(missing_from_a, lines_a, thr, options)

    out_dir = output_dir or file_a.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    missing_from_b_path = missing_from_report_path(out_dir, file_b)
    missing_from_a_path = missing_from_report_path(out_dir, file_a)
    fuzzy_path = out_dir / "fuzzy-matches.txt" if debug else None
    fuzzy_mismatches_path = out_dir / "fuzzy-mismatches.txt" if debug else None

    if write_reports:
        write_lines(missing_from_b_path, missing_from_b)
        write_lines(missing_from_a_path, missing_from_a)
        if debug:
            if fuzzy_path is not None:
                write_fuzzy_matches(fuzzy_path, fuzzy_pairs)
            if fuzzy_mismatches_path is not None:
                write_fuzzy_matches(fuzzy_mismatches_path, fuzzy_mismatches)

    return CompareResult(
        file_a=file_a,
        file_b=file_b,
        output_dir=out_dir,
        threshold=threshold,
        count_a=len(lines_a),
        count_b=len(lines_b),
        exact_matched=exact_matched,
        missing_from_b=missing_from_b,
        missing_from_a=missing_from_a,
        fuzzy_pairs=fuzzy_pairs,
        fuzzy_mismatches=fuzzy_mismatches,
        missing_from_b_path=missing_from_b_path,
        missing_from_a_path=missing_from_a_path,
        fuzzy_path=fuzzy_path,
        fuzzy_mismatches_path=fuzzy_mismatches_path,
    )


def format_compare_report(result: CompareResult, *, preview_limit: int = 15) -> str:
    if result.error:
        return result.error

    lines = [
        "Text compare\n",
        f"File A:    {result.file_a}",
        f"File B:    {result.file_b}",
        f"Reports:   {result.output_dir}",
        f"Threshold: {result.threshold}%",
        "",
        f"Lines in A:      {result.count_a}",
        f"Lines in B:      {result.count_b}",
        f"Exact matched:   {result.exact_matched}",
        f"Fuzzy matched:   {result.fuzzy_matched}",
        f"Total matched:   {result.total_matched}",
        "",
        f"Missing from {result.file_b.name} ({len(result.missing_from_b)}):",
        f"  -> {result.missing_from_b_path}",
    ]
    for item in result.missing_from_b[:preview_limit]:
        lines.append(f"     {item.raw}")
    if len(result.missing_from_b) > preview_limit:
        lines.append(
            f"     ... and {len(result.missing_from_b) - preview_limit} more"
        )

    lines.append("")
    lines.append(f"Missing from {result.file_a.name} ({len(result.missing_from_a)}):")
    lines.append(f"  -> {result.missing_from_a_path}")
    for item in result.missing_from_a[:preview_limit]:
        lines.append(f"     {item.raw}")
    if len(result.missing_from_a) > preview_limit:
        lines.append(
            f"     ... and {len(result.missing_from_a) - preview_limit} more"
        )

    if result.fuzzy_pairs and result.fuzzy_path is not None:
        lines.append("")
        lines.append(f"Fuzzy matches ({len(result.fuzzy_pairs)}):")
        lines.append(f"  -> {result.fuzzy_path}")
        for left, right, score in result.fuzzy_pairs[:preview_limit]:
            lines.append(f"  {score:.0f}%  {left.raw}")
            lines.append(f"       -> {right.raw}")
        if len(result.fuzzy_pairs) > preview_limit:
            lines.append(
                f"     ... and {len(result.fuzzy_pairs) - preview_limit} more"
            )

    if result.fuzzy_mismatches and result.fuzzy_mismatches_path is not None:
        lines.append("")
        lines.append(
            f"Fuzzy mismatches below {result.threshold}% ({len(result.fuzzy_mismatches)}):"
        )
        lines.append(f"  -> {result.fuzzy_mismatches_path}")
        for left, right, score in result.fuzzy_mismatches[:preview_limit]:
            lines.append(f"  {score:.0f}%  {left.raw}")
            lines.append(f"       -> {right.raw}")
        if len(result.fuzzy_mismatches) > preview_limit:
            lines.append(
                f"     ... and {len(result.fuzzy_mismatches) - preview_limit} more"
            )

    return "\n".join(lines) + "\n"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Compare two text files (exact line match, then word-based fuzzy)."
    )
    p.add_argument("file_a", type=Path, help="First text file (one line per entry)")
    p.add_argument("file_b", type=Path, help="Second text file")
    p.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Write report files here (default: same folder as file A)",
    )
    p.add_argument(
        "--threshold",
        type=int,
        default=FUZZY_THRESHOLD,
        help=f"Word fuzzy minimum score (default: {FUZZY_THRESHOLD})",
    )
    p.add_argument(
        "--debug",
        action="store_true",
        help="Write fuzzy-matches.txt and fuzzy-mismatches.txt",
    )
    p.add_argument(
        "--strip-extensions",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Strip known audio extensions before comparing (default: on)",
    )
    p.add_argument(
        "--romaji",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Also compare Japanese converted to romaji (default: on)",
    )
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    args = parse_args(argv)
    result = run_compare(
        args.file_a,
        args.file_b,
        output_dir=args.output_dir,
        threshold=args.threshold,
        debug=args.debug,
        strip_extensions=args.strip_extensions,
        romaji_compare=args.romaji,
    )
    text = format_compare_report(result)
    try:
        print(text.rstrip("\n"))
    except UnicodeEncodeError:
        enc = sys.stdout.encoding or "utf-8"
        print(
            text.encode(enc, errors="replace")
            .decode(enc, errors="replace")
            .rstrip("\n")
        )
    return 1 if result.error else 0


if __name__ == "__main__":
    raise SystemExit(main())
