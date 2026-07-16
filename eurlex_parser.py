"""
eurlex_parser.py

Revised EUR-Lex XHTML parser for Official Journal-style HTML files such as:
    L_2017168EN.01001201.xml.html

Revision in this version
------------------------
This version reflects the diagnostic results from the first class-driven run:

    Recitals: 103, but expected 89
    Articles: 49, expected 49
    Annexes: 0, but expected 6
    Footnotes: 28, but expected 27

The changes are therefore deliberately targeted:

1. Recitals
   The previous parser treated too many pre-Article oj-normal paragraphs as recitals.
   This version captures only the contiguous numbered recital sequence before the
   first Article.  It starts at recital (1) / 1 and continues while the numbers are
   sequential.  It therefore excludes legal bases, formulae and other preamble text.

2. Annexes
   The previous parser skipped many eli-subdivision containers to avoid duplicate
   text.  That was too aggressive because Annex headings may be encoded on or inside
   those containers.  This version has a dedicated annex discovery pass over the
   original BeautifulSoup tree and does not rely only on the flattened legal stream.

3. Footnotes
   The previous parser could over-count because EUR-Lex note structures may contain
   nested note-like elements.  This version extracts only top-level oj-note elements,
   removes the marker only from a copied text computation, and de-duplicates by
   source id / marker / text.

4. Diagnostics
   The JSON structure report now includes more useful diagnostics, including likely
   recital numbers, article numbers, annex headings and footnote markers.

This file remains parser-only. It does not generate Word. The intended usage is:

    EUR-Lex XHTML -> Legislation model -> Stage 1 formatter and QA outputs
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Iterable, List, Optional, Sequence, Callable
import copy
import json
import re

from bs4 import BeautifulSoup, Tag


# =============================================================================
# Data model
# =============================================================================

@dataclass
class Footnote:
    note_id: Optional[str]
    marker: str
    text: str


@dataclass
class TableCell:
    text: str
    is_header: bool = False
    colspan: int = 1
    rowspan: int = 1
    classes: List[str] = field(default_factory=list)


@dataclass
class TableBlock:
    provision_id: str
    rows: List[List[TableCell]] = field(default_factory=list)
    caption: Optional[str] = None
    classes: List[str] = field(default_factory=list)


@dataclass
class TextBlock:
    provision_id: str
    text: str
    role: str = "paragraph"
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Recital:
    provision_id: str
    number: Optional[str]
    text: str
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Article:
    provision_id: str
    number: str
    heading: str
    subtitle: Optional[str] = None
    blocks: List[TextBlock | TableBlock] = field(default_factory=list)
    section_context: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Section:
    provision_id: str
    label: str
    level: int
    text: str
    classes: List[str] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Annex:
    provision_id: str
    heading: str
    blocks: List[TextBlock | TableBlock] = field(default_factory=list)
    source_id: Optional[str] = None


@dataclass
class Legislation:
    title: str
    metadata: dict[str, Any] = field(default_factory=dict)
    sections: List[Section] = field(default_factory=list)
    recitals: List[Recital] = field(default_factory=list)
    articles: List[Article] = field(default_factory=list)
    annexes: List[Annex] = field(default_factory=list)
    footnotes: List[Footnote] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# =============================================================================
# Classes / constants
# =============================================================================

LEGAL_STRUCTURE_CLASSES = {
    "eli-container",
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
RECITAL_NUMBER_RE = re.compile(r"^\s*\(?\s*(\d{1,3})\s*\)?\s+(.*)", re.S)
POINT_RE = re.compile(r"^\s*\(?([a-z])\)\s+", re.I)
SUBPOINT_RE = re.compile(r"^\s*\(?([ivxlcdm]+)\)\s+", re.I)
PARA_NUM_RE = re.compile(r"^\s*(\d+)\.\s+")


# =============================================================================
# Basic helpers
# =============================================================================

def read_html(path: str | Path) -> str:
    path = Path(path)
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return path.read_bytes().decode("utf-8", errors="replace")


def load_soup(path: str | Path) -> BeautifulSoup:
    return BeautifulSoup(read_html(path), "lxml")


def classes_of(tag: Tag | None) -> List[str]:
    if not isinstance(tag, Tag):
        return []
    classes = tag.get("class", [])
    if isinstance(classes, str):
        return classes.split()
    return list(classes)


def has_class(tag: Tag | None, class_name: str) -> bool:
    return class_name in classes_of(tag)


def has_any_class(tag: Tag | None, class_names: Iterable[str]) -> bool:
    return bool(set(classes_of(tag)).intersection(class_names))


def is_layout_or_nav(tag: Tag | None) -> bool:
    return has_any_class(tag, IGNORE_CLASSES)


def clean_text(text: str | None) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ").replace("\u2011", "-")
    text = re.sub(r"[\t\r\f\v ]+", " ", text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def element_text(tag: Tag) -> str:
    return clean_text(tag.get_text(" ", strip=True))


def element_text_without_child_classes(tag: Tag, child_classes_to_remove: Iterable[str]) -> str:
    """Return tag text after removing descendants with specified classes.

    This is used for footnotes so the note marker is not duplicated in the body text.
    It uses a copy, so the BeautifulSoup tree is not mutated.
    """
    tag_copy = copy.copy(tag)
    soup_copy = BeautifulSoup(str(tag_copy), "lxml")
    for removable in soup_copy.find_all(lambda t: isinstance(t, Tag) and has_any_class(t, child_classes_to_remove)):
        removable.decompose()
    return clean_text(soup_copy.get_text(" ", strip=True))


def source_id(tag: Tag | None) -> Optional[str]:
    if not isinstance(tag, Tag):
        return None
    for attr in ("id", "name", "eId", "xml:id"):
        if tag.get(attr):
            return str(tag.get(attr))
    return None


def slug(value: str) -> str:
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "unknown"


def truncated(text: str, length: int = 120) -> str:
    text = clean_text(text)
    return text if len(text) <= length else text[: length - 1] + "…"


# =============================================================================
# Root selection
# =============================================================================

def find_legal_root(soup: BeautifulSoup) -> Tag:
    candidates = soup.select(".eli-container")
    if not candidates:
        return soup.body or soup

    def score(tag: Tag) -> int:
        return (
            len(tag.select(".oj-ti-art")) * 100
            + len(tag.select(".eli-main-title")) * 25
            + len(tag.select(".eli-title")) * 10
            + len(tag.select(".oj-normal"))
            + len(tag.select(".oj-table")) * 5
            + len(tag.select(".oj-note")) * 2
        )

    return max(candidates, key=score)


# =============================================================================
# Metadata/title
# =============================================================================

def extract_title(root: Tag, soup: BeautifulSoup) -> str:
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
# Legal stream
# =============================================================================

def is_relevant_legal_element(tag: Tag) -> bool:
    if not isinstance(tag, Tag):
        return False
    if is_layout_or_nav(tag):
        return False
    return has_any_class(tag, LEGAL_STRUCTURE_CLASSES | TABLE_CLASSES | FOOTNOTE_CLASSES)


def should_emit_element(tag: Tag) -> bool:
    """Avoid duplicate container text, but do not suppress potential Annex containers.

    The earlier version skipped eli-subdivision containers entirely where they had
    legal descendants.  That accidentally hid annex-related structure.  In this
    version, the stream omits eli-subdivision containers to avoid duplicate body text,
    but annex discovery is performed separately on the original tree.  This keeps the
    article stream clean while allowing robust annex detection.
    """
    if has_class(tag, SUBDIVISION_CLASS):
        return False
    return True


def legal_stream(root: Tag) -> List[Tag]:
    stream: List[Tag] = []
    for tag in root.find_all(True):
        if not is_relevant_legal_element(tag):
            continue
        if not should_emit_element(tag):
            continue
        text = element_text(tag)
        if not text and not has_any_class(tag, TABLE_CLASSES):
            continue
        stream.append(tag)
    return stream


# =============================================================================
# Footnotes
# =============================================================================

def is_top_level_note(note: Tag) -> bool:
    """Return True only for oj-note elements not nested inside another oj-note."""
    parent = note.parent
    while isinstance(parent, Tag):
        if has_class(parent, "oj-note"):
            return False
        parent = parent.parent
    return True


def extract_footnotes(root: Tag) -> List[Footnote]:
    """Extract unique top-level oj-note footnotes.

    Over-counting is avoided by:
    - selecting only .oj-note, not .oj-note-tag or .oj-super;
    - ignoring oj-note elements nested inside other oj-note elements;
    - de-duplicating on source id / marker / text.
    """
    footnotes: List[Footnote] = []
    seen: set[tuple[str, str, str]] = set()

    for idx, note in enumerate(root.select(".oj-note"), start=1):
        if not is_top_level_note(note):
            continue

        marker_tag = note.select_one(".oj-note-tag")
        marker = element_text(marker_tag) if marker_tag else str(idx)
        text = element_text_without_child_classes(note, {"oj-note-tag"})
        note_id = source_id(note) or f"note_{marker or idx}"
        key = (note_id or "", marker or "", text)
        if key in seen:
            continue
        seen.add(key)
        footnotes.append(Footnote(note_id=note_id, marker=marker, text=text))

    return footnotes


# =============================================================================
# Tables
# =============================================================================

def parse_table(table_tag: Tag, provision_id: str) -> TableBlock:
    classes = classes_of(table_tag)
    caption = None
    previous = table_tag.find_previous_sibling()
    if isinstance(previous, Tag) and has_class(previous, "oj-ti-tbl"):
        caption = element_text(previous)

    rows: List[List[TableCell]] = []
    for tr in table_tag.find_all("tr"):
        row: List[TableCell] = []
        for cell in tr.find_all(["th", "td"], recursive=False):
            cell_classes = classes_of(cell)
            is_header = cell.name == "th" or has_class(cell, "oj-tbl-hdr")
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
                    text=element_text(cell),
                    is_header=is_header,
                    colspan=colspan,
                    rowspan=rowspan,
                    classes=cell_classes,
                )
            )
        if row:
            rows.append(row)

    if not rows:
        text = element_text(table_tag)
        if text:
            rows = [[TableCell(text=text, classes=classes)]]

    return TableBlock(provision_id=provision_id, rows=rows, caption=caption, classes=classes)


# =============================================================================
# Recitals
# =============================================================================

def find_first_article_index(stream: Sequence[Tag]) -> Optional[int]:
    for i, tag in enumerate(stream):
        if has_class(tag, ARTICLE_HEADING_CLASS):
            return i
    return None


def extract_numbered_recital_candidate(tag: Tag) -> Optional[tuple[int, str]]:
    """Return (number, text) if the tag is a numbered recital candidate."""
    if not has_class(tag, NORMAL_TEXT_CLASS):
        return None
    text = element_text(tag)
    if not text:
        return None
    match = RECITAL_NUMBER_RE.match(text)
    if not match:
        return None
    try:
        number = int(match.group(1))
    except ValueError:
        return None
    return number, text


def extract_recitals(stream: Sequence[Tag]) -> List[Recital]:
    """Extract the contiguous numbered recital sequence before Article 1.

    This deliberately avoids the previous over-inclusive approach which captured
    legal bases and other preamble paragraphs.  It starts at the first numbered
    oj-normal block whose number is 1 and then continues only while the numbering is
    sequential.  For the Prospectus Regulation sample this should move the count
    towards the expected 89 recitals.
    """
    first_article_idx = find_first_article_index(stream)
    if first_article_idx is None:
        return []

    pre_article = stream[:first_article_idx]
    recitals: List[Recital] = []
    expected = 1
    started = False

    for tag in pre_article:
        candidate = extract_numbered_recital_candidate(tag)
        if candidate is None:
            if started:
                # Once the numbered sequence has started, non-numbered material ends it.
                # This protects against picking up unrelated numbering later.
                continue
            continue

        number, text = candidate
        if not started:
            if number != 1:
                continue
            started = True
            expected = 1

        if number == expected:
            recitals.append(
                Recital(
                    provision_id=f"recital_{number}",
                    number=str(number),
                    text=text,
                    classes=classes_of(tag),
                    source_id=source_id(tag),
                )
            )
            expected += 1
        elif started and number > expected:
            # Numbering gap: stop, because the official recital sequence should be contiguous.
            break
        else:
            # Duplicate/lower number after start: ignore.
            continue

    return recitals


# =============================================================================
# Sections
# =============================================================================

def classify_section_level(tag: Tag) -> int:
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
    if has_class(tag, ARTICLE_HEADING_CLASS) or has_class(tag, ARTICLE_SUBTITLE_CLASS):
        return False
    # Annex headings are extracted separately; exclude them from ordinary sections.
    if ANNEX_RE.match(element_text(tag)):
        return False
    return has_any_class(tag, {"eli-title", SECTION_1_CLASS, SECTION_2_CLASS, "oj-doc-ti"})


def extract_sections(stream: Sequence[Tag]) -> List[Section]:
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
    match = ARTICLE_RE.search(heading)
    if match:
        return match.group(1)
    return str(fallback_index)


def text_block_role(text: str, classes: Sequence[str]) -> str:
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


def iter_between(stream: Sequence[Tag], start_idx: int, stop_predicate: Callable[[Tag], bool]) -> Iterable[Tag]:
    for tag in stream[start_idx + 1:]:
        if stop_predicate(tag):
            break
        yield tag


def is_article_boundary(tag: Tag) -> bool:
    return has_class(tag, ARTICLE_HEADING_CLASS)


def is_annex_start_text(tag: Tag) -> bool:
    return bool(ANNEX_RE.match(element_text(tag)))


def extract_articles(stream: Sequence[Tag]) -> List[Article]:
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

        def stop(t: Tag) -> bool:
            return is_article_boundary(t) or is_annex_start_text(t)

        for child in iter_between(stream, idx, stop):
            text = element_text(child)
            cls = classes_of(child)

            if has_class(child, ARTICLE_SUBTITLE_CLASS):
                if text:
                    article.subtitle = text
                continue

            if has_any_class(child, TABLE_CLASSES):
                if child.name == "table" or has_class(child, "oj-table"):
                    article.blocks.append(parse_table(child, f"article_{slug(number)}_table_{table_counter}"))
                    table_counter += 1
                continue

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

        articles.append(article)

    return articles


# =============================================================================
# Annexes — revised dedicated tree-based discovery
# =============================================================================

def is_annex_heading_element(tag: Tag) -> bool:
    """Return True where a tag itself appears to be an Annex heading."""
    if not isinstance(tag, Tag) or is_layout_or_nav(tag):
        return False
    text = element_text(tag)
    if not ANNEX_RE.match(text):
        return False
    # Annex headings are normally structural titles/subdivisions.  Be deliberately
    # permissive because the prior result was zero annexes.
    return has_any_class(
        tag,
        {
            "eli-subdivision",
            "eli-title",
            "oj-doc-ti",
            "oj-ti-section-1",
            "oj-ti-section-2",
            "oj-normal",
            "oj-hd-ti",
        },
    )


def find_annex_heading_tags(root: Tag) -> List[Tag]:
    """Find distinct Annex heading tags in source order.

    This searches the original tree, not the legal stream, because Annex headings may
    be encoded on or inside eli-subdivision containers that the stream suppresses to
    avoid duplicate text.
    """
    candidates: List[Tag] = []
    seen_ids: set[int] = set()

    for tag in root.find_all(True):
        if not is_annex_heading_element(tag):
            continue
        text = element_text(tag)
        # Exclude table/body paragraphs that merely refer to an annex, if any.  The
        # heading is usually short: 'ANNEX I', 'ANNEX II', etc.  Allow some expanded
        # headings but avoid long body text.
        if len(text) > 120:
            continue
        ident = id(tag)
        if ident not in seen_ids:
            candidates.append(tag)
            seen_ids.add(ident)

    # De-duplicate by heading text while preserving order.  This protects against a
    # container and its child both reading as 'ANNEX I'.
    out: List[Tag] = []
    seen_headings: set[str] = set()
    for tag in candidates:
        key = element_text(tag).upper()
        if key in seen_headings:
            continue
        seen_headings.add(key)
        out.append(tag)

    return out


def annex_label_from_heading(heading: str, fallback_index: int) -> str:
    match = ANNEX_RE.match(heading)
    if match and match.group(1):
        return match.group(1)
    return str(fallback_index)


def nearest_annex_container(heading_tag: Tag) -> Tag:
    """Return the smallest useful container for an annex heading."""
    current = heading_tag
    # Prefer an ancestor eli-subdivision if present because it will normally contain
    # the annex heading and body.
    parent = heading_tag.parent
    while isinstance(parent, Tag):
        if has_class(parent, "eli-subdivision"):
            return parent
        parent = parent.parent
    return current


def following_elements_until_next_annex(root: Tag, heading_tag: Tag, next_heading: Optional[Tag]) -> List[Tag]:
    """Collect source-order elements after heading_tag until next annex heading.

    Uses next_elements so that it works whether the annex body is inside the same
    subdivision container or follows as siblings.
    """
    collected: List[Tag] = []
    for el in heading_tag.next_elements:
        if not isinstance(el, Tag):
            continue
        if next_heading is not None and el is next_heading:
            break
        if is_annex_heading_element(el) and el is not heading_tag:
            break
        if is_relevant_legal_element(el) and should_emit_element(el):
            collected.append(el)
    return collected


def extract_annexes(root: Tag) -> List[Annex]:
    """Extract Annexes I-VI etc. using dedicated tree-based discovery."""
    annexes: List[Annex] = []
    heading_tags = find_annex_heading_tags(root)
    if not heading_tags:
        return annexes

    for idx, heading_tag in enumerate(heading_tags):
        heading = element_text(heading_tag)
        label = annex_label_from_heading(heading, idx + 1)
        next_heading = heading_tags[idx + 1] if idx + 1 < len(heading_tags) else None

        annex = Annex(
            provision_id=f"annex_{slug(label)}",
            heading=heading,
            source_id=source_id(heading_tag),
        )
        block_counter = 1
        table_counter = 1
        seen_tag_ids: set[int] = set()

        for child in following_elements_until_next_annex(root, heading_tag, next_heading):
            if id(child) in seen_tag_ids:
                continue
            seen_tag_ids.add(id(child))

            if child is heading_tag:
                continue
            if has_any_class(child, FOOTNOTE_CLASSES):
                continue

            text = element_text(child)
            cls = classes_of(child)

            if has_any_class(child, TABLE_CLASSES):
                if child.name == "table" or has_class(child, "oj-table"):
                    annex.blocks.append(parse_table(child, f"annex_{slug(label)}_table_{table_counter}"))
                    table_counter += 1
                continue

            # Avoid recursively adding very large containers as text blocks.
            if has_class(child, "eli-container") or has_class(child, "eli-subdivision"):
                continue

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
# Reports
# =============================================================================

def build_structure_summary(legislation: Legislation) -> dict[str, Any]:
    return {
        "title": legislation.title,
        "recitals": len(legislation.recitals),
        "sections": len(legislation.sections),
        "articles": len(legislation.articles),
        "annexes": len(legislation.annexes),
        "footnotes": len(legislation.footnotes),
        "recital_numbers": [r.number for r in legislation.recitals],
        "article_numbers": [a.number for a in legislation.articles],
        "annex_headings": [a.heading for a in legislation.annexes],
        "footnote_markers": [f.marker for f in legislation.footnotes],
        "warnings": legislation.warnings,
    }


def legislation_to_dict(legislation: Legislation) -> dict[str, Any]:
    return asdict(legislation)


def write_structure_report(legislation: Legislation, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    report = {"summary": build_structure_summary(legislation), "model": legislation_to_dict(legislation)}
    output_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return output_path


# =============================================================================
# Main parse entry point
# =============================================================================

def parse_eurlex_document(input_path: str | Path) -> Legislation:
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
    legislation.annexes = extract_annexes(root)

    if not legislation.articles:
        legislation.warnings.append("No Articles detected using class oj-ti-art.")
    if not legislation.recitals:
        legislation.warnings.append("No Recitals detected using contiguous numbered oj-normal sequence before first Article.")
    if not legislation.annexes:
        legislation.warnings.append("No Annexes detected using tree-based ANNEX heading discovery.")
    if not legislation.footnotes:
        legislation.warnings.append("No footnotes detected using top-level class oj-note.")

    # Soft diagnostics useful for the Prospectus Regulation sample, but not fatal for
    # future acts where the counts will differ.
    if legislation.recitals:
        nums = [int(r.number) for r in legislation.recitals if r.number and r.number.isdigit()]
        if nums and nums != list(range(1, len(nums) + 1)):
            legislation.warnings.append("Detected recital numbers are not a simple contiguous sequence.")

    return legislation


# =============================================================================
# CLI
# =============================================================================

def _default_report_path(source: Path) -> Path:
    return source.with_name(f"{source.stem}_structure_report.json")


def main(argv: Optional[list[str]] = None) -> int:
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
