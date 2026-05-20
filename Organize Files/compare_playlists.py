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


@dataclass
class Line:
    raw: str


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


def exact_key(line: str) -> str:
    return line.casefold()


def split_words(line: str) -> list[str]:
    return line.split()


def word_pair_score(a: str, b: str) -> float:
    if a.casefold() == b.casefold():
        return 100.0
    return float(fuzz.ratio(a.casefold(), b.casefold()))


def word_match_score(line_a: str, line_b: str, threshold: float) -> float:
    """Share of words in the longer line that match a word in the other line."""
    wa = split_words(line_a)
    wb = split_words(line_b)
    if not wa and not wb:
        return 100.0
    if not wa or not wb:
        return 0.0

    used_b: set[int] = set()
    matched = 0
    for word in wa:
        best = -1.0
        best_j = -1
        for j, other in enumerate(wb):
            if j in used_b:
                continue
            score = word_pair_score(word, other)
            if score > best:
                best = score
                best_j = j
        if best >= threshold and best_j >= 0:
            matched += 1
            used_b.add(best_j)

    return 100.0 * matched / max(len(wa), len(wb))


def build_exact_index(lines: list[Line]) -> dict[str, list[Line]]:
    index: dict[str, list[Line]] = {}
    for item in lines:
        index.setdefault(exact_key(item.raw), []).append(item)
    return index


def find_exact_match(line: Line, index: dict[str, list[Line]]) -> Line | None:
    hits = index.get(exact_key(line.raw))
    return hits[0] if hits else None


def find_fuzzy_match(
    line: Line,
    candidates: list[Line],
    used: set[int],
    threshold: float,
) -> tuple[Line, float] | None:
    best: Line | None = None
    best_score = -1.0
    for cand in candidates:
        if id(cand) in used:
            continue
        score = word_match_score(line.raw, cand.raw, threshold)
        if score >= threshold and score > best_score:
            best_score = score
            best = cand
    if best is None:
        return None
    return best, best_score


def match_a_to_b(
    lines_a: list[Line],
    lines_b: list[Line],
    index_b: dict[str, list[Line]],
    threshold: float,
) -> tuple[list[Line], list[tuple[Line, Line, float]], set[int]]:
    """Lines in A missing from B, fuzzy pairs, used ids from B."""
    used_b: set[int] = set()
    missing: list[Line] = []
    fuzzy_pairs: list[tuple[Line, Line, float]] = []

    for line in lines_a:
        hit = find_exact_match(line, index_b)
        if hit is not None and id(hit) not in used_b:
            used_b.add(id(hit))
            continue

        fuzzy = find_fuzzy_match(line, lines_b, used_b, threshold)
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
    missing_from_b_path: Path
    missing_from_a_path: Path
    fuzzy_path: Path
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
) -> CompareResult:
    out_default = output_dir or file_a.parent
    missing_from_b_path = missing_from_report_path(out_default, file_b)
    missing_from_a_path = missing_from_report_path(out_default, file_a)
    fuzzy_path = out_default / "fuzzy-matches.txt"

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
        missing_from_b_path=missing_from_b_path,
        missing_from_a_path=missing_from_a_path,
        fuzzy_path=fuzzy_path,
    )

    if not file_a.is_file():
        base.error = f"File not found:\n{file_a}"
        return base
    if not file_b.is_file():
        base.error = f"File not found:\n{file_b}"
        return base

    lines_a = read_lines(file_a)
    lines_b = read_lines(file_b)
    index_b = build_exact_index(lines_b)

    missing_from_b, fuzzy_pairs, used_b = match_a_to_b(
        lines_a, lines_b, index_b, float(threshold)
    )
    missing_from_a = [line for line in lines_b if id(line) not in used_b]
    exact_matched = len(lines_a) - len(missing_from_b) - len(fuzzy_pairs)

    out_dir = output_dir or file_a.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    missing_from_b_path = missing_from_report_path(out_dir, file_b)
    missing_from_a_path = missing_from_report_path(out_dir, file_a)
    fuzzy_path = out_dir / "fuzzy-matches.txt"

    if write_reports:
        write_lines(missing_from_b_path, missing_from_b)
        write_lines(missing_from_a_path, missing_from_a)
        write_fuzzy_matches(fuzzy_path, fuzzy_pairs)

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
        missing_from_b_path=missing_from_b_path,
        missing_from_a_path=missing_from_a_path,
        fuzzy_path=fuzzy_path,
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

    if result.fuzzy_pairs:
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
