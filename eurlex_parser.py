"""
eurlex_parser.py

EUR-Lex XHTML parser for Official Journal-style HTML files such as:
    L_2017168EN.01001201.xml.html

This parser is deliberately class-driven.  It uses EUR-Lex/OJ CSS classes found in
Simon McKnight's sample file, including:

    eli-main-title
    eli-title
    eli-subdivision
    oj-doc-ti
    oj-hd-ti
    oj-normal
    oj-ti-art
    oj-sti-art
    oj-ti-section-1
    oj-ti-section-2
    oj-table
    oj-tbl-hdr
    oj-tbl-txt
    oj-note
    oj-note-tag
    oj-super

Purpose
-------
Stage 1 of the legislative consolidation project:

    EUR-Lex XHTML -> structured legislative model

The output model is intended to be consumed by a separate Word writer and a later
Stage 2 QA process.  This file does not generate Word itself.

Design principles
-----------------
1. Parse by EUR-Lex classes rather than by generic paragraph extraction.
2. Keep the document order stable.
3. Preserve enough structure to support later amendment targeting.
4. Treat tables and footnotes as first-class objects.
5. Record parser warnings rather than silently discarding uncertainty.

Important limitation
--------------------
This is a strong first EUR-Lex-aware draft, but the exact boundaries between
recitals, articles and annexes may still need tuning against the real source HTML.
The next quality leap will come from inspecting representative snippets around:
    - the first recital;
    - Article 1;
    - an Annex heading;
    - a complex annex table;
    - footnotes.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Iterable, List, Optional, Sequence
import json
import re

from bs4 import BeautifulSoup, NavigableString, Tag


# =============================================================================
# Data model
# =============================================================================

@dataclass
class Footnote:
    """A footnote or endnote detected in the EUR-Lex source."""
    note_id: Optional[str]
    marker: str
    text: str


@dataclass
class TableCell:
    """A simplified table cell."""
    text: str
    is_header: bool = False
    colspan: int = 1
    rowspan: int = 1
    classes: List[str] = field(default_factory=list)


@dataclass
class TableBlock:
    """A table block preserving row/cell/text structure."""
    provision_id: str
    rows: List[List[TableCell]] = field(default_factory=list)
    caption: Optional[str] = None
    classes: List[str] = field(default_factory=list)


@dataclass
class TextBlock:
    """A text block within a provision, article, recital or annex."""
    provision_id: str
    text: str
    role: str = "paragraph"
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Recital:
    """A recital."""
    provision_id: str
    number: Optional[str]
    text: str
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Article:
    """An Article with heading, optional subtitle and child blocks."""
    provision_id: str
    number: str
    heading: str
    subtitle: Optional[str] = None
    blocks: List[TextBlock | TableBlock] = field(default_factory=list)
    section_context: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Section:
    """A heading/subdivision in the main body of the act."""
    provision_id: str
    label: str
    level: int
    text: str
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Annex:
    """An annex with child blocks."""
    provision_id: str
    heading: str
    blocks: List[TextBlock | TableBlock] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Legislation:
    """Structured representation of the EUR-Lex document."""
    title: str
    metadata: dict[str, Any] = field(default_factory=dict)
    sections: List[Section] = field(default_factory=list)
    recitals: List[Recital] = field(default_factory=list)
    articles: List[Article] = field(default_factory=list)
    annexes: List[Annex] = field(default_factory=list)
    footnotes: List[Footnote] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# =============================================================================
# Class map and constants
# =============================================================================

# Classes identified in the sample class inventory.
LEGAL_STRUCTURE_CLASSES = {
    "eli-main-title",
    "eli-title",
    "eli-subdivision",
    "oj-doc-ti",
    "oj-hd-ti",
    "oj-hd-date",
    "oj-hd-lg",
    "oj-hd-oj",
    "oj-ti-art",
    "oj-sti-art",
    "oj-ti-section-1",
    "oj-ti-section-2",
    "oj-normal",
    "oj-enumeration-spacing",
    "oj-doc-end",
    "oj-doc-sep",
    "oj-final",
    "oj-signatory",
}

TABLE_CLASSES = {
    "oj-table",
    "oj-tbl-hdr",
    "oj-tbl-txt",
    "oj-ti-tbl",
}

FOOTNOTE_CLASSES = {
    "oj-note",
    "oj-note-tag",
    "oj-super",
}

# Classes that are layout/navigation rather than legal content.
IGNORE_CLASSES = {
    "Wrapper",
    "affix-top",
    "clearfix",
    "col-md-3",
    "col-md-9",
    "collapse",
    "container-fluid",
    "fa",
    "fa-angle-right",
    "fa-times",
    "nav",
    "row",
    "row-offcanvas",
    "toc-eli-label",
    "toc-eli-subdivisions",
    "toc-sidebar",
    "toc-sidenav",
    "tocWrapper",
}

ARTICLE_HEADING_CLASS = "oj-ti-art"
ARTICLE_SUBTITLE_CLASS = "oj-sti-art"
MAIN_TITLE_CLASS = "eli-main-title"
NORMAL_TEXT_CLASS = "oj-normal"
SUBDIVISION_CLASS = "eli-subdivision"
SECTION_1_CLASS = "oj-ti-section-1"
SECTION_2_CLASS = "oj-ti-section-2"

ARTICLE_RE = re.compile(r"\bArticle\s+(\d+[A-Za-z]?)\b", re.I)
ANNEX_RE = re.compile(r"^\s*ANNEX\s+([IVXLCDM]+|\d+)?\b", re.I)
RECITAL_RE = re.compile(r"^\s*\(?\s*(\d+)\s*\)?\s+(.*)", re.S)
POINT_RE = re.compile(r"^\s*\(?([a-z])\)\s+", re.I)
SUBPOINT_RE = re.compile(r"^\s*\(?([ivxlcdm]+)\)\s+", re.I)
PARA_NUM_RE = re.compile(r"^\s*(\d+)\.\s+")


# =============================================================================
# Basic helpers
# =============================================================================

def read_html(path: str | Path) -> str:
    """Read a EUR-Lex XHTML/HTML file with safe encoding fallbacks."""
    path = Path(path)
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_bytes().decode("utf-8", errors="replace")


def load_soup(path: str | Path) -> BeautifulSoup:
    """Load the file into BeautifulSoup using lxml's HTML parser."""
    return BeautifulSoup(read_html(path), "lxml")


def classes_of(tag: Tag | None) -> List[str]:
    """Return the tag's classes as a clean list."""
    if not isinstance(tag, Tag):
        return []
    classes = tag.get("class", [])
    if isinstance(classes, str):
        return classes.split()
    return list(classes)


def has_class(tag: Tag | None, class_name: str) -> bool:
    """True if tag has a given CSS class."""
    return class_name in classes_of(tag)


def has_any_class(tag: Tag | None, class_names: Iterable[str]) -> bool:
    """True if tag has any of the supplied CSS classes."""
    cls = set(classes_of(tag))
    return bool(cls.intersection(class_names))


def is_layout_or_nav(tag: Tag | None) -> bool:
    """True if a tag looks like navigation/layout rather than legal content."""
    return has_any_class(tag, IGNORE_CLASSES)


def clean_text(text: str | None) -> str:
    """Normalise whitespace without rewriting legal text."""
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = text.replace("\u2011", "-")
    text = re.sub(r"[\t\r\f\v ]+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def element_text(tag: Tag) -> str:
    """Extract visible text from a tag."""
    return clean_text(tag.get_text(" ", strip=True))


def source_id(tag: Tag | None) -> Optional[str]:
    """Return the best available source identifier for traceability."""
    if not isinstance(tag, Tag):
        return None
    for attr in ("id", "name", "eId", "xml:id"):
        if tag.get(attr):
            return str(tag.get(attr))
    return None


def slug(value: str) -> str:
    """Create a stable identifier fragment."""
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "unknown"


def dedupe_preserve_order(items: Sequence[str]) -> List[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            out.append(item)
            seen.add(item)
    return out


# =============================================================================
# Source area selection
# =============================================================================

def find_legal_root(soup: BeautifulSoup) -> Tag:
    """Find the main legal text container.

    The sample has class 'eli-container' several times.  We choose the candidate
    with most legal-structure descendants.  If none exists, use body/soup.
    """
    candidates = soup.select(".eli-container")
    if not candidates:
        return soup.body or soup

    def score(tag: Tag) -> int:
        return (
            len(tag.select(".oj-ti-art")) * 10
            + len(tag.select(".eli-main-title")) * 5
            + len(tag.select(".eli-title")) * 3
            + len(tag.select(".oj-normal"))
            + len(tag.select(".oj-table"))
        )

    return max(candidates, key=score)


# =============================================================================
# Metadata/title extraction
# =============================================================================

def extract_title(root: Tag, soup: BeautifulSoup) -> str:
    """Extract main title, preferring EUR-Lex title classes."""
    for selector in (".eli-main-title", ".oj-hd-ti", ".eli-title", ".oj-doc-ti"):
        found = root.select_one(selector) or soup.select_one(selector)
        if found:
            text = element_text(found)
            if text:
                return text
    if soup.title and soup.title.text:
        return clean_text(soup.title.text)
    return "Untitled EUR-Lex document"


def extract_metadata(root: Tag, soup: BeautifulSoup, input_path: str | Path) -> dict[str, Any]:
    """Extract lightweight metadata useful for QA and traceability."""
    metadata: dict[str, Any] = {
        "source_path": str(input_path),
        "parser": "eurlex_parser.py",
        "format": "EUR-Lex XHTML / OJ HTML",
    }
    for cls, key in (
        ("oj-hd-oj", "official_journal"),
        ("oj-hd-date", "oj_date"),
        ("oj-hd-lg", "language"),
    ):
        found = root.select_one(f".{cls}") or soup.select_one(f".{cls}")
        if found:
            metadata[key] = element_text(found)
    return metadata


# =============================================================================
# Linear legal element stream
# =============================================================================

def is_relevant_legal_element(tag: Tag) -> bool:
    """True if tag is one of the legal-content elements we want to stream."""
    if not isinstance(tag, Tag):
        return False
    if is_layout_or_nav(tag):
        return False
    return has_any_class(tag, LEGAL_STRUCTURE_CLASSES | TABLE_CLASSES | FOOTNOTE_CLASSES)


def should_skip_nested_duplicate(tag: Tag) -> bool:
    """Avoid emitting high-level containers whose descendants carry the real content.

    This prevents eli-subdivision containers from duplicating all child text where the
    same content is already present as oj-ti-art/oj-normal/table elements.
    """
    if has_class(tag, SUBDIVISION_CLASS):
        # Keep subdivision only if it has little/no legal children underneath.  Most
        # subdivisions are containers; their title children are usually eli-title or
        # oj-ti-section-* elements.
        legal_children = tag.find_all(
            lambda t: isinstance(t, Tag)
            and t is not tag
            and has_any_class(t, LEGAL_STRUCTURE_CLASSES | TABLE_CLASSES)
        )
        return len(legal_children) > 0
    return False


def legal_stream(root: Tag) -> List[Tag]:
    """Return relevant legal elements in source order."""
    stream: List[Tag] = []
    for tag in root.find_all(True):
        if not is_relevant_legal_element(tag):
            continue
        if should_skip_nested_duplicate(tag):
            continue
        text = element_text(tag)
        # Keep tables even where text extraction is sparse.
        if not text and not has_any_class(tag, TABLE_CLASSES):
            continue
        stream.append(tag)
    return stream


# =============================================================================
# Footnotes
# =============================================================================

def extract_footnotes(root: Tag) -> List[Footnote]:
    """Extract footnotes/endnotes from oj-note elements."""
    footnotes: List[Footnote] = []
    for idx, note in enumerate(root.select(".oj-note"), start=1):
        marker = ""
        marker_tag = note.select_one(".oj-note-tag")
        if marker_tag:
            marker = element_text(marker_tag)
            marker_tag.extract()  # remove marker from note body for cleaner text
        text = element_text(note)
        footnotes.append(
            Footnote(
                note_id=source_id(note) or f"note_{idx}",
                marker=marker or str(idx),
                text=text,
            )
        )
    return footnotes


# =============================================================================
# Tables
# =============================================================================

def parse_table(table_tag: Tag, provision_id: str) -> TableBlock:
    """Parse a table bearing oj-table/table-related classes."""
    classes = classes_of(table_tag)
    caption = None

    # Some EUR-Lex tables have a separate title immediately before the table.
    previous = table_tag.find_previous_sibling()
    if isinstance(previous, Tag) and has_class(previous, "oj-ti-tbl"):
        caption = element_text(previous)

    rows: List[List[TableCell]] = []
    for tr in table_tag.find_all("tr"):
        row: List[TableCell] = []
        for cell in tr.find_all(["th", "td"], recursive=False):
            cell_classes = classes_of(cell)
            is_header = cell.name == "th" or has_class(cell, "oj-tbl-hdr")
            text = element_text(cell)
            try:
                colspan = int(cell.get("colspan", 1))
            except ValueError:
                colspan = 1
            try:
                rowspan = int(cell.get("rowspan", 1))
            except ValueError:
                rowspan = 1
            row.append(
                TableCell(
                    text=text,
                    is_header=is_header,
                    colspan=colspan,
                    rowspan=rowspan,
                    classes=cell_classes,
                )
            )
        if row:
            rows.append(row)

    # Fallback for unusual markup: preserve table text as one row.
    if not rows:
        text = element_text(table_tag)
        if text:
            rows = [[TableCell(text=text, classes=classes)]]

    return TableBlock(
        provision_id=provision_id,
        rows=rows,
        caption=caption,
        classes=classes,
    )


# =============================================================================
# Recitals
# =============================================================================

def find_first_article_index(stream: Sequence[Tag]) -> Optional[int]:
    for i, tag in enumerate(stream):
        if has_class(tag, ARTICLE_HEADING_CLASS):
            return i
    return None


def is_probable_recital(tag: Tag, before_first_article: bool) -> bool:
    """Heuristic recital detection within the pre-Article part of the act."""
    if not before_first_article:
        return False
    if not has_class(tag, NORMAL_TEXT_CLASS):
        return False
    text = element_text(tag)
    if not text:
        return False
    # Most recitals in OJ HTML begin with a bracketed number or number-like marker.
    if re.match(r"^\(?\d+\)?\s+", text):
        return True
    # Fallback: long preamble paragraphs before the first Article may be recitals,
    # but exclude formal formulae and short headers.
    if len(text) > 80 and not text.upper().startswith(("HAVE ADOPTED", "HAS ADOPTED")):
        return True
    return False


def extract_recitals(stream: Sequence[Tag]) -> List[Recital]:
    """Extract recitals from oj-normal blocks before the first Article."""
    first_article_idx = find_first_article_index(stream)
    if first_article_idx is None:
        return []

    recitals: List[Recital] = []
    for tag in stream[:first_article_idx]:
        if not is_probable_recital(tag, before_first_article=True):
            continue
        text = element_text(tag)
        match = RECITAL_RE.match(text)
        number = match.group(1) if match else None
        provision_id = f"recital_{number}" if number else f"recital_{len(recitals)+1}"
        recitals.append(
            Recital(
                provision_id=provision_id,
                number=number,
                text=text,
                classes=classes_of(tag),
                source_id=source_id(tag),
            )
        )
    return recitals


# =============================================================================
# Sections and structural headings
# =============================================================================

def classify_section_level(tag: Tag) -> int:
    """Map section/title classes to structural levels."""
    if has_class(tag, "eli-title"):
        return 1
    if has_class(tag, SECTION_1_CLASS):
        return 2
    if has_class(tag, SECTION_2_CLASS):
        return 3
    if has_class(tag, "oj-doc-ti"):
        return 1
    return 9


def is_section_heading(tag: Tag) -> bool:
    """True if the tag is a structural heading other than article/annex heading."""
    if has_class(tag, ARTICLE_HEADING_CLASS) or has_class(tag, ARTICLE_SUBTITLE_CLASS):
        return False
    if has_any_class(tag, {"eli-title", SECTION_1_CLASS, SECTION_2_CLASS, "oj-doc-ti"}):
        return True
    return False


def extract_sections(stream: Sequence[Tag]) -> List[Section]:
    """Extract major headings/sections in source order."""
    sections: List[Section] = []
    for idx, tag in enumerate(stream, start=1):
        if not is_section_heading(tag):
            continue
        text = element_text(tag)
        if not text:
            continue
        sections.append(
            Section(
                provision_id=f"section_{idx}_{slug(text)[:40]}",
                label=text.split(" ", 1)[0],
                level=classify_section_level(tag),
                text=text,
                classes=classes_of(tag),
                source_id=source_id(tag),
            )
        )
    return sections


# =============================================================================
# Articles
# =============================================================================

def article_number_from_heading(heading: str, fallback_index: int) -> str:
    """Extract Article number from heading text."""
    match = ARTICLE_RE.search(heading)
    if match:
        return match.group(1)
    return str(fallback_index)


def text_block_role(text: str, classes: Sequence[str]) -> str:
    """Classify a normal/article text block more precisely."""
    if PARA_NUM_RE.match(text):
        return "article_paragraph"
    if POINT_RE.match(text):
        return "point"
    if SUBPOINT_RE.match(text):
        return "subpoint"
    if "oj-enumeration-spacing" in classes:
        return "enumeration"
    if "oj-final" in classes:
        return "final"
    if "oj-signatory" in classes:
        return "signatory"
    return "paragraph"


def iter_between(stream: Sequence[Tag], start_idx: int, stop_predicate) -> Iterable[Tag]:
    """Yield tags after start_idx until stop_predicate(tag) is true."""
    for tag in stream[start_idx + 1:]:
        if stop_predicate(tag):
            break
        yield tag


def is_article_boundary(tag: Tag) -> bool:
    """True if a tag starts a new Article."""
    return has_class(tag, ARTICLE_HEADING_CLASS)


def is_annex_start(tag: Tag) -> bool:
    text = element_text(tag)
    return bool(ANNEX_RE.match(text))


def extract_articles(stream: Sequence[Tag]) -> List[Article]:
    """Extract Articles by oj-ti-art boundaries."""
    articles: List[Article] = []
    article_indices = [i for i, tag in enumerate(stream) if has_class(tag, ARTICLE_HEADING_CLASS)]

    for art_count, idx in enumerate(article_indices, start=1):
        art_tag = stream[idx]
        heading = element_text(art_tag)
        number = article_number_from_heading(heading, art_count)
        article = Article(
            provision_id=f"article_{slug(number)}",
            number=number,
            heading=heading,
            source_id=source_id(art_tag),
        )

        block_counter = 1
        table_counter = 1

        # Stop at next article or first annex heading.  Annexes are parsed separately.
        def stop(t: Tag) -> bool:
            return is_article_boundary(t) or is_annex_start(t)

        for child in iter_between(stream, idx, stop):
            text = element_text(child)
            cls = classes_of(child)

            if has_class(child, ARTICLE_SUBTITLE_CLASS):
                if text:
                    article.subtitle = text
                continue

            if has_any_class(child, TABLE_CLASSES):
                if child.name == "table" or has_class(child, "oj-table"):
                    article.blocks.append(
                        parse_table(child, f"article_{slug(number)}_table_{table_counter}")
                    )
                    table_counter += 1
                continue

            # Skip footnote blocks in the main article content stream; they are
            # extracted separately into legislation.footnotes.
            if has_any_class(child, FOOTNOTE_CLASSES):
                continue

            if has_class(child, NORMAL_TEXT_CLASS) or has_class(child, "oj-enumeration-spacing"):
                if text:
                    article.blocks.append(
                        TextBlock(
                            provision_id=f"article_{slug(number)}_block_{block_counter}",
                            text=text,
                            role=text_block_role(text, cls),
                            classes=cls,
                            source_id=source_id(child),
                        )
                    )
                    block_counter += 1
                continue

        articles.append(article)

    return articles


# =============================================================================
# Annexes
# =============================================================================

def annex_label_from_heading(heading: str, fallback_index: int) -> str:
    match = ANNEX_RE.match(heading)
    if match and match.group(1):
        return match.group(1)
    return str(fallback_index)


def find_annex_indices(stream: Sequence[Tag]) -> List[int]:
    """Find likely Annex heading positions."""
    indices = []
    for i, tag in enumerate(stream):
        text = element_text(tag)
        if not text:
            continue
        if ANNEX_RE.match(text):
            # Prefer formal title/subdivision classes, but allow oj-doc-ti/eli-title.
            if has_any_class(tag, {"eli-title", "oj-doc-ti", SECTION_1_CLASS, SECTION_2_CLASS, NORMAL_TEXT_CLASS}):
                indices.append(i)
    return indices


def extract_annexes(stream: Sequence[Tag]) -> List[Annex]:
    """Extract annexes from the first ANNEX heading onwards."""
    annexes: List[Annex] = []
    annex_indices = find_annex_indices(stream)
    if not annex_indices:
        return annexes

    annex_set = set(annex_indices)

    for annex_count, idx in enumerate(annex_indices, start=1):
        heading_tag = stream[idx]
        heading = element_text(heading_tag)
        label = annex_label_from_heading(heading, annex_count)
        annex = Annex(
            provision_id=f"annex_{slug(label)}",
            heading=heading,
            source_id=source_id(heading_tag),
        )
        block_counter = 1
        table_counter = 1

        stop_idx = annex_indices[annex_count] if annex_count < len(annex_indices) else len(stream)
        for child in stream[idx + 1:stop_idx]:
            if has_any_class(child, FOOTNOTE_CLASSES):
                continue
            text = element_text(child)
            cls = classes_of(child)
            if has_any_class(child, TABLE_CLASSES):
                if child.name == "table" or has_class(child, "oj-table"):
                    annex.blocks.append(
                        parse_table(child, f"annex_{slug(label)}_table_{table_counter}")
                    )
                    table_counter += 1
                continue
            if has_class(child, ARTICLE_HEADING_CLASS):
                # Defensive: articles should not normally be within annexes here.
                break
            if text and has_any_class(child, LEGAL_STRUCTURE_CLASSES):
                annex.blocks.append(
                    TextBlock(
                        provision_id=f"annex_{slug(label)}_block_{block_counter}",
                        text=text,
                        role=text_block_role(text, cls),
                        classes=cls,
                        source_id=source_id(child),
                    )
                )
                block_counter += 1

        annexes.append(annex)

    return annexes


# =============================================================================
# Diagnostics / reports
# =============================================================================

def build_structure_summary(legislation: Legislation) -> dict[str, Any]:
    """Generate a compact structure summary for QA."""
    return {
        "title": legislation.title,
        "recitals": len(legislation.recitals),
        "sections": len(legislation.sections),
        "articles": len(legislation.articles),
        "annexes": len(legislation.annexes),
        "footnotes": len(legislation.footnotes),
        "article_numbers": [a.number for a in legislation.articles],
        "annex_headings": [a.heading for a in legislation.annexes],
        "warnings": legislation.warnings,
    }


def legislation_to_dict(legislation: Legislation) -> dict[str, Any]:
    """Convert dataclass model to JSON-serialisable dict."""
    return asdict(legislation)


def write_structure_report(legislation: Legislation, output_path: str | Path) -> Path:
    """Write a JSON structure report for Stage 2 QA."""
    output_path = Path(output_path)
    report = {
        "summary": build_structure_summary(legislation),
        "model": legislation_to_dict(legislation),
    }
    output_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return output_path


# =============================================================================
# Main parse entry point
# =============================================================================

def parse_eurlex_document(input_path: str | Path) -> Legislation:
    """Parse a EUR-Lex XHTML/OJ HTML file into a Legislation model."""
    soup = load_soup(input_path)
    root = find_legal_root(soup)
    stream = legal_stream(root)

    title = extract_title(root, soup)
    metadata = extract_metadata(root, soup, input_path)
    metadata["legal_stream_elements"] = len(stream)

    legislation = Legislation(title=title, metadata=metadata)
    legislation.footnotes = extract_footnotes(root)
    legislation.recitals = extract_recitals(stream)
    legislation.sections = extract_sections(stream)
    legislation.articles = extract_articles(stream)
    legislation.annexes = extract_annexes(stream)

    # Diagnostics: these are not fatal, but they are important for Stage 2 QA.
    if not legislation.articles:
        legislation.warnings.append("No Articles detected using class oj-ti-art.")
    if not legislation.recitals:
        legislation.warnings.append("No Recitals detected before first Article.")
    if not legislation.annexes:
        legislation.warnings.append("No Annexes detected using ANNEX heading pattern.")
    if not legislation.footnotes:
        legislation.warnings.append("No footnotes detected using class oj-note.")

    return legislation


# =============================================================================
# Command-line utility for parser testing
# =============================================================================

def _default_report_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}_structure_report.json")


def main(argv: Optional[list[str]] = None) -> int:
    """CLI for testing the parser independently of Word generation.

    Usage:
        python eurlex_parser.py L_2017168EN.01001201.xml.html
        python eurlex_parser.py L_2017168EN.01001201.xml.html --report report.json
    """
    import argparse

    parser = argparse.ArgumentParser(description="Parse EUR-Lex XHTML/OJ HTML and generate a structure report.")
    parser.add_argument("source", help="Path to EUR-Lex XHTML/HTML file")
    parser.add_argument("--report", "-r", help="Optional JSON structure report path")
    args = parser.parse_args(argv)

    source = Path(args.source).expanduser().resolve()
    report_path = Path(args.report).expanduser().resolve() if args.report else _default_report_path(source)

    legislation = parse_eurlex_document(source)
    write_structure_report(legislation, report_path)

    summary = build_structure_summary(legislation)
    print("EUR-Lex parse completed.")
    print(f"Source: {source}")
    print(f"Report: {report_path}")
    print(f"Title: {summary['title']}")
    print(f"Recitals: {summary['recitals']}")
    print(f"Sections/headings: {summary['sections']}")
    print(f"Articles: {summary['articles']}")
    print(f"Annexes: {summary['annexes']}")
    print(f"Footnotes: {summary['footnotes']}")
    if legislation.warnings:
        print("Warnings:")
        for warning in legislation.warnings:
            print(f" - {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
