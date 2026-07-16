"""Diagnostic utility for analysing EUR-Lex XHTML structure.

This module can be used in two ways:
1. Imported from stage1.py via run_diagnostics(...)
2. Run standalone from CLI
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

from bs4 import BeautifulSoup, Tag


OUTPUT_FILENAMES = {
    "pre_article": "diagnostic_pre_article.csv",
    "recital_candidates": "diagnostic_recital_candidates.csv",
    "annex_candidates": "diagnostic_annex_candidates.csv",
    "footnotes": "diagnostic_footnotes.csv",
    "structure_classes": "diagnostic_structure_classes.csv",
}



def classes_of(tag: Tag) -> str:
    cls = tag.get("class", [])
    return " ".join(cls)



def text_of(tag: Tag) -> str:
    return " ".join(tag.get_text(" ", strip=True).split())



def load_document(path: str | Path) -> BeautifulSoup:
    html = Path(path).read_text(encoding="utf-8", errors="replace")
    return BeautifulSoup(html, "lxml")



def find_first_article(soup: BeautifulSoup) -> Tag | None:
    return soup.select_one(".oj-ti-art")



def _write_csv(rows: list[dict[str, str]], output_path: Path, fieldnames: list[str]) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    return output_path



def dump_pre_article_content(soup: BeautifulSoup, output_dir: Path) -> Path:
    article = find_first_article(soup)
    rows: list[dict[str, str]] = []

    if article:
        for tag in article.find_all_previous():
            if not isinstance(tag, Tag):
                continue
            text = text_of(tag)
            if not text:
                continue
            rows.append(
                {
                    "tag": tag.name,
                    "classes": classes_of(tag),
                    "id": tag.get("id", ""),
                    "text": text[:250],
                }
            )
    rows.reverse()
    out = _write_csv(rows, output_dir / OUTPUT_FILENAMES["pre_article"], ["tag", "classes", "id", "text"])
    print(f"Pre-Article elements written: {len(rows)} -> {out}")
    return out



def dump_recital_candidates(soup: BeautifulSoup, output_dir: Path) -> Path:
    rows: list[dict[str, str]] = []
    for tag in soup.find_all(True):
        text = text_of(tag)
        if text and text.startswith("("):
            rows.append(
                {
                    "tag": tag.name,
                    "classes": classes_of(tag),
                    "id": tag.get("id", ""),
                    "text": text[:250],
                }
            )

    out = _write_csv(
        rows,
        output_dir / OUTPUT_FILENAMES["recital_candidates"],
        ["tag", "classes", "id", "text"],
    )
    print(f"Recital candidates written: {len(rows)} -> {out}")
    return out



def dump_annex_candidates(soup: BeautifulSoup, output_dir: Path) -> Path:
    rows: list[dict[str, str]] = []
    for tag in soup.find_all(True):
        text = text_of(tag)
        if "ANNEX" in text.upper():
            rows.append(
                {
                    "tag": tag.name,
                    "classes": classes_of(tag),
                    "parent_classes": classes_of(tag.parent) if isinstance(tag.parent, Tag) else "",
                    "id": tag.get("id", ""),
                    "text": text[:250],
                }
            )

    out = _write_csv(
        rows,
        output_dir / OUTPUT_FILENAMES["annex_candidates"],
        ["tag", "classes", "parent_classes", "id", "text"],
    )
    print(f"Annex candidates written: {len(rows)} -> {out}")
    return out



def dump_footnotes(soup: BeautifulSoup, output_dir: Path) -> Path:
    rows: list[dict[str, str]] = []
    for note in soup.select(".oj-note"):
        rows.append(
            {
                "tag": note.name,
                "classes": classes_of(note),
                "id": note.get("id", ""),
                "text": text_of(note)[:500],
            }
        )

    out = _write_csv(rows, output_dir / OUTPUT_FILENAMES["footnotes"], ["tag", "classes", "id", "text"])
    print(f"Footnotes found: {len(rows)} -> {out}")
    return out



def dump_structure_classes(soup: BeautifulSoup, output_dir: Path) -> Path:
    interesting = [
        "eli-main-title",
        "eli-title",
        "eli-subdivision",
        "oj-doc-ti",
        "oj-ti-art",
        "oj-sti-art",
        "oj-ti-section-1",
        "oj-ti-section-2",
        "oj-note",
        "oj-table",
    ]

    rows: list[dict[str, str]] = []
    for cls in interesting:
        for tag in soup.select(f".{cls}"):
            rows.append(
                {
                    "class": cls,
                    "tag": tag.name,
                    "id": tag.get("id", ""),
                    "text": text_of(tag)[:250],
                }
            )

    out = _write_csv(rows, output_dir / OUTPUT_FILENAMES["structure_classes"], ["class", "tag", "id", "text"])
    print(f"Structure rows written: {len(rows)} -> {out}")
    return out



def run_diagnostics(source: str | Path, output_dir: str | Path) -> list[Path]:
    soup = load_document(source)
    out_dir = Path(output_dir)
    outputs = [
        dump_pre_article_content(soup, out_dir),
        dump_recital_candidates(soup, out_dir),
        dump_annex_candidates(soup, out_dir),
        dump_footnotes(soup, out_dir),
        dump_structure_classes(soup, out_dir),
    ]
    print("Diagnostics complete.")
    return outputs



def main() -> int:
    parser = argparse.ArgumentParser(description="Run EUR-Lex structure diagnostics.")
    parser.add_argument("source", help="Path to EUR-Lex XHTML file")
    parser.add_argument(
        "--output-dir",
        default="outputs/diagnostics",
        help="Directory for diagnostic CSV files (default: outputs/diagnostics)",
    )
    args = parser.parse_args()

    run_diagnostics(args.source, args.output_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
