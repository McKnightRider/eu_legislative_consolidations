"""
Stage 1 helper functions: XML/HTML legislative source -> structured intermediate model.

Purpose
-------
This module deliberately separates parsing from Word generation.  The output of the
parser is a simple, inspectable data model made of dataclasses.  That makes the later
Stage 2 quality-check easier, because the same structured model can be compared
against the source document and the generated Word document.

The parser is intentionally conservative:
- it preserves source order;
- it avoids making legal assumptions where the source is ambiguous;
- it records warnings rather than silently discarding content;
- it is designed to work with EUR-Lex HTML/XML but also with reasonably clean generic
  legislative HTML/XML.

Known limitations of this Stage 1 version
-----------------------------------------
1. Native Word footnotes are not automatically reconstructed from every possible
   EUR-Lex footnote format.  Footnote-like links are currently preserved inline.
2. Highly complex tables are preserved as Word tables, but nested tables and unusual
   colspan/rowspan structures are simplified.
3. The parser identifies likely recitals/articles/annexes using common text patterns
   and HTML/XML tags.  It is therefore a robust starting point, not a substitute for
   Stage 2 legal QA.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re
from typing import Iterable, List, Optional, Sequence, Tuple

from bs4 import BeautifulSoup, NavigableString, Tag
from lxml import etree


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class TableCell:
    """A simplified representation of a source table cell."""
    text: str


@dataclass
class TableBlock:
    """A simplified representation of a source table."""
    rows: List[List[TableCell]]


@dataclass
class ParagraphBlock:
    """A paragraph-like block of legislative text."""
    text: str
    role: str = "paragraph"  # title, recital, article_heading, article_paragraph, annex_heading, paragraph
    source_id: Optional[str] = None


@dataclass
class LegislativeDocument:
    """Structured intermediate representation used by the Word writer."""
    title: str = "Converted Legislative Document"
    metadata: dict = field(default_factory=dict)
    blocks: List[ParagraphBlock | TableBlock] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Source loading and parsing helpers
# ---------------------------------------------------------------------------

def read_source_file(input_path: str | Path) -> str:
    """Read the source file as text, trying UTF-8 first and falling back safely.

    Parameters
    ----------
    input_path:
        Path to the XML or HTML source file.
    """
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    # Final fallback should be very rare, but prevents hard failure on odd files.
    return path.read_bytes().decode("utf-8", errors="replace")


def detect_source_type(input_path: str | Path, raw_text: str) -> str:
    """Return 'xml' or 'html' using file extension and lightweight content checks."""
    suffix = Path(input_path).suffix.lower()
    if suffix in {".xml", ".akn"}:
        return "xml"
    if suffix in {".html", ".htm", ".xhtml"}:
        return "html"
    # Content-based fallback.
    stripped = raw_text.lstrip()
    if stripped.startswith("<?xml") or re.search(r"<\w+:?law|<\w+:?act|<\w+:?article", stripped[:2000], re.I):
        return "xml"
    return "html"


def make_soup(raw_text: str, source_type: str) -> BeautifulSoup:
    """Create a BeautifulSoup object using a suitable parser.

    For XML, BeautifulSoup's XML parser keeps tag names simple enough for our first
    pass.  If the XML is malformed, we fall back to html5lib so that the script can
    still produce a diagnostic Word document rather than failing completely.
    """
    if source_type == "xml":
        try:
            # lxml XML parse first: useful as a validity check.
            etree.fromstring(raw_text.encode("utf-8"))
            return BeautifulSoup(raw_text, "xml")
        except Exception:
            return BeautifulSoup(raw_text, "html5lib")
    return BeautifulSoup(raw_text, "html5lib")


# ---------------------------------------------------------------------------
# Text normalisation
# ---------------------------------------------------------------------------

_WHITESPACE_RE = re.compile(r"[\t\r\f\v ]+")
_MULTI_NL_RE = re.compile(r"\n{3,}")


def normalise_text(text: str) -> str:
    """Normalise text while preserving legally meaningful punctuation.

    This is intentionally light-touch.  It removes layout noise but does not rewrite
    legislative language.
    """
    if text is None:
        return ""
    text = text.replace("\xa0", " ")
    text = text.replace("\u2011", "-")
    text = _WHITESPACE_RE.sub(" ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = _MULTI_NL_RE.sub("\n\n", text)
    return text.strip()


def tag_text(tag: Tag) -> str:
    """Extract visible text from a tag with conservative spacing."""
    return normalise_text(tag.get_text(" ", strip=True))


def is_noise(text: str) -> bool:
    """Return True for empty or non-substantive navigation text."""
    if not text:
        return True
    lowered = text.lower()
    noise_phrases = {
        "languages, formats and link to oj",
        "multilingual display",
        "display all documents",
        "print",
        "save to my items",
    }
    return lowered in noise_phrases


# ---------------------------------------------------------------------------
# Legislative role detection
# ---------------------------------------------------------------------------

_RECITAL_RE = re.compile(r"^\(?\d+\)?\s+")
_ARTICLE_RE = re.compile(r"^Article\s+\d+[a-zA-Z]?\b", re.I)
_ANNEX_RE = re.compile(r"^ANNEX\b|^Annex\b")
_TITLE_HINT_RE = re.compile(r"REGULATION\s+\(EU\)|DIRECTIVE\s+\(EU\)|DECISION\s+\(EU\)", re.I)


def classify_text_block(text: str, previous_role: Optional[str] = None) -> str:
    """Classify a text block into a legal/structural role.

    The classification is deliberately heuristic at Stage 1.  It is designed to
    provide useful Word styles and a basis for Stage 2 QA rather than to constitute
    a legal conclusion about the provision.
    """
    if _TITLE_HINT_RE.search(text[:200]):
        return "title"
    if _ARTICLE_RE.match(text):
        return "article_heading"
    if _ANNEX_RE.match(text):
        return "annex_heading"
    if _RECITAL_RE.match(text) and previous_role in {"title", "recital", "paragraph", None}:
        # Recitals are commonly represented as numbered paragraphs before Article 1.
        return "recital"
    if previous_role == "article_heading":
        return "article_paragraph"
    return "paragraph"


def likely_title(soup: BeautifulSoup, fallback: str) -> str:
    """Find a plausible document title from title/h1/early text."""
    for selector in ["title", "h1", "p.title", "div.title"]:
        found = soup.select_one(selector)
        if found:
            text = tag_text(found)
            if text and len(text) > 10:
                return text[:250]

    # EUR-Lex sources often have the legal title in early paragraphs.
    for tag in soup.find_all(["p", "div", "heading", "title"], limit=40):
        text = tag_text(tag)
        if _TITLE_HINT_RE.search(text):
            return text[:250]
    return Path(fallback).stem


# ---------------------------------------------------------------------------
# Block extraction
# ---------------------------------------------------------------------------

_BLOCK_TAGS = {
    "p", "div", "section", "article", "recital", "paragraph", "num", "heading",
    "h1", "h2", "h3", "h4", "h5", "h6", "li", "table"
}


def _is_leafish_block(tag: Tag) -> bool:
    """Return True if this tag should be emitted as one block.

    We treat paragraphs/list items/headings/tables as atomic enough for Stage 1.
    Container divs/sections are emitted only if they do not contain lower-level block
    elements, to avoid duplicating text.
    """
    if tag.name == "table":
        return True
    if tag.name in {"p", "li", "heading", "h1", "h2", "h3", "h4", "h5", "h6", "recital"}:
        return True
    if tag.name in {"div", "section", "article", "paragraph", "num"}:
        return not any(child.name in _BLOCK_TAGS for child in tag.find_all(recursive=False) if isinstance(child, Tag))
    return False


def extract_table(tag: Tag) -> Optional[TableBlock]:
    """Convert a source HTML/XML table into a simplified TableBlock."""
    rows: List[List[TableCell]] = []
    for tr in tag.find_all("tr"):
        cells: List[TableCell] = []
        for cell in tr.find_all(["th", "td"], recursive=False):
            cells.append(TableCell(tag_text(cell)))
        if cells:
            rows.append(cells)
    if not rows:
        return None
    return TableBlock(rows=rows)


def extract_blocks(soup: BeautifulSoup) -> Tuple[List[ParagraphBlock | TableBlock], List[str]]:
    """Extract source-order paragraph and table blocks from the parsed document."""
    blocks: List[ParagraphBlock | TableBlock] = []
    warnings: List[str] = []
    previous_role: Optional[str] = None

    body = soup.body or soup
    seen_text_hashes = set()

    for tag in body.find_all(list(_BLOCK_TAGS)):
        if not isinstance(tag, Tag) or not _is_leafish_block(tag):
            continue

        if tag.name == "table":
            table = extract_table(tag)
            if table:
                blocks.append(table)
            else:
                warnings.append("Skipped an empty or unsupported table.")
            previous_role = "table"
            continue

        text = tag_text(tag)
        if is_noise(text):
            continue

        # Avoid obvious duplicate container text, but do not dedupe short legal text.
        text_hash = hash(text)
        if len(text) > 80 and text_hash in seen_text_hashes:
            continue
        seen_text_hashes.add(text_hash)

        role = classify_text_block(text, previous_role)
        source_id = tag.get("id") or tag.get("name") or tag.get("eId") or tag.get("xml:id")
        blocks.append(ParagraphBlock(text=text, role=role, source_id=source_id))
        previous_role = role

    if not blocks:
        # Last-resort fallback: split body text into paragraphs.
        text = normalise_text(body.get_text("\n", strip=True))
        for para in [p.strip() for p in text.split("\n") if p.strip()]:
            blocks.append(ParagraphBlock(text=para, role=classify_text_block(para)))
        warnings.append("No structured blocks detected; used plain-text fallback extraction.")

    return blocks, warnings


def parse_legislative_source(input_path: str | Path) -> LegislativeDocument:
    """Parse XML/HTML source into a LegislativeDocument model."""
    raw_text = read_source_file(input_path)
    source_type = detect_source_type(input_path, raw_text)
    soup = make_soup(raw_text, source_type)
    title = likely_title(soup, str(input_path))
    blocks, warnings = extract_blocks(soup)

    metadata = {
        "source_path": str(input_path),
        "source_type": source_type,
        "block_count": len(blocks),
    }
    return LegislativeDocument(title=title, metadata=metadata, blocks=blocks, warnings=warnings)
