"""Stage 2 quality assurance checks for Stage 1 output.

This script compares text coverage between a Stage 1 DOCX and source documents:
- original HTML (recommended)
- original PDF (optional, requires pypdf)

The goal is to highlight potentially missing words in the DOCX output.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup

WORD_RE = re.compile(r"[A-Za-z0-9]+(?:['-][A-Za-z0-9]+)?")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DEFAULT_QA_DIR = Path("outputs") / "qa"


@dataclass
class ComparisonResult:
    source_type: str
    source_path: str
    total_source_tokens: int
    total_docx_tokens: int
    matched_tokens: int
    missing_tokens_total: int
    missing_unique_tokens: int
    extra_tokens_total: int
    extra_unique_tokens: int
    coverage_ratio: float
    top_missing: list[tuple[str, int]]
    top_extra: list[tuple[str, int]]


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def tokenize(text: str) -> list[str]:
    return [m.group(0).lower() for m in WORD_RE.finditer(text)]


def extract_docx_text(docx_path: str | Path) -> str:
    text_parts: list[str] = []
    with zipfile.ZipFile(docx_path, "r") as zf:
        xml_members = [
            name
            for name in zf.namelist()
            if name.startswith("word/")
            and name.endswith(".xml")
            and "/_rels/" not in name
            and name not in {"word/styles.xml", "word/settings.xml", "word/fontTable.xml", "word/numbering.xml"}
        ]
        for member in xml_members:
            try:
                root = ET.fromstring(zf.read(member))
            except ET.ParseError:
                continue
            for t in root.findall(f".//{{{W_NS}}}t"):
                if t.text:
                    text_parts.append(t.text)
    return normalize_text(" ".join(text_parts))


def extract_html_text(html_path: str | Path) -> str:
    html = Path(html_path).read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(html, "lxml")
    for noisy in soup(["script", "style", "noscript"]):
        noisy.decompose()
    return normalize_text(soup.get_text(" ", strip=True))


def extract_pdf_text(pdf_path: str | Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError(
            "PDF comparison requires pypdf. Install it with: pip install pypdf"
        ) from exc

    reader = PdfReader(str(pdf_path))
    page_text: list[str] = []
    for page in reader.pages:
        extracted = page.extract_text() or ""
        page_text.append(extracted)
    return normalize_text(" ".join(page_text))


def compare_token_coverage(docx_tokens: list[str], source_tokens: list[str], source_type: str, source_path: str, top_n: int) -> ComparisonResult:
    docx_counter = Counter(docx_tokens)
    source_counter = Counter(source_tokens)

    missing_counter = source_counter - docx_counter
    extra_counter = docx_counter - source_counter

    matched_tokens = sum(min(source_counter[token], docx_counter[token]) for token in source_counter)
    total_source = len(source_tokens)
    coverage_ratio = (matched_tokens / total_source) if total_source else 1.0

    return ComparisonResult(
        source_type=source_type,
        source_path=str(source_path),
        total_source_tokens=total_source,
        total_docx_tokens=len(docx_tokens),
        matched_tokens=matched_tokens,
        missing_tokens_total=sum(missing_counter.values()),
        missing_unique_tokens=len(missing_counter),
        extra_tokens_total=sum(extra_counter.values()),
        extra_unique_tokens=len(extra_counter),
        coverage_ratio=coverage_ratio,
        top_missing=missing_counter.most_common(top_n),
        top_extra=extra_counter.most_common(top_n),
    )


def default_report_path(docx_path: str | Path, source_type: str) -> Path:
    return DEFAULT_QA_DIR / f"{Path(docx_path).stem}_stage2_qa_report_{source_type}.json"


def resolve_report_path(
    docx_path: str | Path,
    source_type: str,
    report_override: str | None,
    multiple_sources: bool,
) -> Path:
    if not report_override:
        return default_report_path(docx_path, source_type)

    base = Path(report_override).expanduser().resolve()
    suffix = base.suffix or ".json"
    stem = base.stem if base.suffix else base.name
    if multiple_sources:
        return base.with_name(f"{stem}_{source_type}{suffix}")
    return base if base.suffix else base.with_suffix(".json")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Stage 2 QA: compare Stage 1 DOCX against original HTML/PDF for missing-word checks."
    )
    parser.add_argument("docx", help="Path to Stage 1 output .docx file")
    parser.add_argument("--html", help="Path to original HTML source")
    parser.add_argument("--pdf", help="Path to original PDF source")
    parser.add_argument(
        "--report",
        help=(
            "Base path for JSON QA report(s). Defaults to outputs/qa/<docx_stem>_stage2_qa_report_<html|pdf>.json. "
            "When both HTML and PDF are provided, _html/_pdf is appended to the base name."
        ),
    )
    parser.add_argument("--top", type=int, default=50, help="Number of top missing/extra tokens to include in report")
    parser.add_argument(
        "--min-coverage",
        type=float,
        default=None,
        help="Optional minimum acceptable token coverage ratio per source",
    )
    parser.add_argument("--fail-on-missing", action="store_true", help="Exit non-zero if any source has missing tokens")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.html and not args.pdf:
        parser.error("Provide at least one source for comparison: --html and/or --pdf")

    docx_path = Path(args.docx).expanduser().resolve()
    if not docx_path.exists():
        print(f"DOCX file not found: {docx_path}")
        return 2

    docx_text = extract_docx_text(docx_path)
    docx_tokens = tokenize(docx_text)

    comparisons: list[ComparisonResult] = []

    if args.html:
        html_path = Path(args.html).expanduser().resolve()
        if not html_path.exists():
            print(f"HTML file not found: {html_path}")
            return 2
        html_tokens = tokenize(extract_html_text(html_path))
        comparisons.append(compare_token_coverage(docx_tokens, html_tokens, "html", str(html_path), args.top))

    if args.pdf:
        pdf_path = Path(args.pdf).expanduser().resolve()
        if not pdf_path.exists():
            print(f"PDF file not found: {pdf_path}")
            return 2
        try:
            pdf_tokens = tokenize(extract_pdf_text(pdf_path))
        except RuntimeError as exc:
            print(f"PDF extraction unavailable: {exc}")
            return 2
        comparisons.append(compare_token_coverage(docx_tokens, pdf_tokens, "pdf", str(pdf_path), args.top))

    multiple_sources = len(comparisons) > 1
    for item in comparisons:
        report_path = resolve_report_path(docx_path, item.source_type, args.report, multiple_sources)
        report_path.parent.mkdir(parents=True, exist_ok=True)

        report = {
            "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "docx_path": str(docx_path),
            "docx_token_count": len(docx_tokens),
            "source_type": item.source_type,
            "source_path": item.source_path,
            "min_coverage_threshold": args.min_coverage,
            "comparison": asdict(item),
        }
        report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
        print(f"Stage 2 QA report written ({item.source_type}): {report_path}")

    for item in comparisons:
        print(
            f"[{item.source_type}] coverage={item.coverage_ratio:.4f}, "
            f"missing={item.missing_tokens_total}, extra={item.extra_tokens_total}"
        )

    below_threshold = []
    if args.min_coverage is not None:
        below_threshold = [item for item in comparisons if item.coverage_ratio < args.min_coverage]
    any_missing = any(item.missing_tokens_total > 0 for item in comparisons)

    if below_threshold:
        print("Coverage check failed for one or more sources.")
        return 1

    if args.fail_on_missing and any_missing:
        print("Missing-token check failed.")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
