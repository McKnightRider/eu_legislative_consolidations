
"""
formatting_stage2_v5.py

Overlay update for formatting_stage2_v3.py.

Run this file instead of v3/v4, in the same folder as formatting_stage2_v3.py:

    python formatting_stage2_v5.py L_2017168EN.01001201.xml.html Prospectus_Regulation_Stage2.docx

Changes from v4:
    1. Fix Annex I-V corruption caused by merging consecutive Annex level-1 markers
       such as "I." and "II." into "I. II.".
    2. Preserve Annex level-1 markers stored in .tit_ nodes, but only merge a
       standalone marker with the next paragraph where the next paragraph is not
       itself another Annex marker/heading.
    3. Format Annex VI correlation table without visible borders: no surrounding
       border, no row borders and no vertical borders between entries.
    4. Keep only the Annex VI header row bold; ordinary correlation rows are not bold.
    5. Continue to justify native Word footnote text via OOXML w:jc="both".

This file intentionally reuses formatting_stage2_v3.py for the rest of the
conversion logic and patches only the Annex and footnote/table behaviours.
"""

from __future__ import annotations

from pathlib import Path
import argparse
import re
from bs4 import Tag
from lxml import etree
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import formatting_stage2_v3 as base


# =============================================================================
# Annex I-V helpers
# =============================================================================

def _clean(text: str) -> str:
    return base.clean_text(text)


def _is_annex_level_marker(text: str) -> bool:
    """True for standalone Annex level-1 markers such as A., B., I., II."""
    return bool(re.match(r"^([A-Z]\.|[IVXLCDM]+\.)$", _clean(text), flags=re.I))


def _is_annex_marker_heading(text: str) -> bool:
    """True for Annex marker headings, standalone or combined.

    Examples:
        A.
        A. Offer statistics
        I.
        I. Persons responsible
    """
    return bool(re.match(r"^([A-Z]\.|[IVXLCDM]+\.)(\s+.+)?$", _clean(text), flags=re.I))


def _normalise_annex_marker_spacing(text: str) -> str:
    """Ensure combined Annex marker headings use a tab after the marker."""
    text = _clean(text)
    text = re.sub(r"^([A-Z]\.|[IVXLCDM]+\.)\s+", r"\1\t", text, flags=re.I)
    return text


def emit_annex_paragraph_sequence(doc, paragraphs: list[Tag]) -> None:
    """Emit Annex I-V paragraphs without corrupting consecutive Roman headings.

    The key rule is:
      - merge "A." + "Offer statistics" into "A.\tOffer statistics";
      - do not merge "I." + "II." or "I." + "II. Something".

    This prevents the v4 error where consecutive level-1 Annex headings were
    combined as "I. II.".
    """
    i = 0
    current_body_indent = 0

    while i < len(paragraphs):
        txt = base.text_with_footnote_tokens(paragraphs[i])
        txt = _clean(txt)
        if not txt:
            i += 1
            continue

        if _is_annex_level_marker(txt) and i + 1 < len(paragraphs):
            nxt = _clean(base.text_with_footnote_tokens(paragraphs[i + 1]))

            # Only merge where next paragraph is actual heading/body text, not the
            # next Annex level marker/heading.  This is the critical v5 fix.
            if nxt and not _is_annex_marker_heading(nxt):
                combined = f"{txt}\t{nxt}"
                base.add_para(doc, combined, left_cm=1, first_line_cm=-1)
                current_body_indent = 1
                i += 2
                continue

            # Standalone marker followed by another marker/marker-heading: keep it.
            base.add_para(doc, txt, left_cm=1, first_line_cm=-1)
            current_body_indent = 1
            i += 1
            continue

        if _is_annex_marker_heading(txt):
            combined = _normalise_annex_marker_spacing(txt)
            base.add_para(doc, combined, left_cm=1, first_line_cm=-1)
            current_body_indent = 1
            i += 1
            continue

        base.add_para(doc, txt, left_cm=current_body_indent)
        i += 1


def add_annexes(doc, soup) -> None:
    """Replacement Annex handler for v5."""
    for annex in soup.find_all(id=re.compile(r"^anx_[IVXLCDM]+$")):
        heading = annex.select_one("p.oj-doc-ti")
        heading_text = _clean(heading.get_text(" ", strip=True)) if heading else str(annex.get("id", "Annex"))

        hp = base.add_center_heading(doc, heading_text)
        hp.paragraph_format.page_break_before = True

        if annex.get("id") == "anx_VI":
            build_annex_vi_table(doc, annex)
            continue

        paras: list[Tag] = []
        for p in annex.find_all("p"):
            if base.has_class(p, "oj-note"):
                continue

            txt = _clean(base.text_with_footnote_tokens(p))
            if not txt or txt == heading_text:
                continue

            pid = str(p.get("id", ""))
            classes = base.tag_classes(p)

            # Preserve Annex .tit_ nodes only where they are visible marker headings.
            # Skip other metadata title nodes.
            if ".tit_" in pid and not _is_annex_marker_heading(txt):
                continue

            if (
                _is_annex_marker_heading(txt)
                or any(c in classes for c in [
                    "oj-normal",
                    "oj-enumeration-spacing",
                    "oj-doc-ti",
                    "oj-ti-section-1",
                    "oj-ti-section-2",
                ])
            ):
                paras.append(p)

        emit_annex_paragraph_sequence(doc, paras)


# =============================================================================
# Annex VI borderless table
# =============================================================================

def remove_table_borders(table) -> None:
    """Remove all visible borders from a python-docx table."""
    tbl = table._tbl
    tblPr = tbl.tblPr

    # Remove existing borders if present, then add nil borders.
    for existing in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(existing)

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    tblPr.append(borders)


def remove_cell_borders(cell) -> None:
    """Remove cell-level borders defensively, in case a table style adds them."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    tcPr.append(borders)


def set_cell_font(cell, *, bold=False, size=11) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            base.set_run_font(run, size=size, bold=bold)


def cell_text_from_html(cell: Tag) -> str:
    children = cell.find_all(["p", "div"], recursive=False)
    parts = [base.text_with_footnote_tokens(child) for child in children] if children else [base.text_with_footnote_tokens(cell)]
    return "\n".join([p for p in parts if p])


def find_largest_html_table(annex: Tag):
    tables = annex.select("table.oj-table") or annex.find_all("table")
    if not tables:
        return None
    return max(tables, key=lambda t: len(t.find_all("tr")) * 10 + len(t.find_all(["td", "th"])))


def build_annex_vi_table(doc, annex: Tag) -> None:
    """Build Annex VI as a two-column table with no visible borders."""
    html_table = find_largest_html_table(annex)
    if html_table is None:
        base.add_para(doc, base.text_with_footnote_tokens(annex), left_cm=0)
        return

    table = doc.add_table(rows=0, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    remove_table_borders(table)

    row_idx = 0
    for tr in html_table.find_all("tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue

        extracted = [cell_text_from_html(c) for c in cells]
        extracted = [e for e in extracted if e]
        if not extracted:
            continue

        row = table.add_row()
        row.cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        row.cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        row.cells[0].text = extracted[0]
        row.cells[1].text = "\n".join(extracted[1:]) if len(extracted) > 1 else ""

        # No surrounding border, no row borders, no borders between entries.
        remove_cell_borders(row.cells[0])
        remove_cell_borders(row.cells[1])

        # Keep only header row bold.
        is_header = row_idx == 0
        set_cell_font(row.cells[0], bold=is_header)
        set_cell_font(row.cells[1], bold=is_header)
        row_idx += 1


# =============================================================================
# Footnotes justified
# =============================================================================

def create_footnotes_xml(id_to_text: dict[int, str]) -> bytes:
    """Replacement native footnotes XML builder with justified footnotes."""
    root = etree.Element(base.w_tag("footnotes"), nsmap={"w": base.W_NS})

    for fid, ftype, marker in [(-1, "separator", "separator"), (0, "continuationSeparator", "continuationSeparator")]:
        fn = etree.SubElement(root, base.w_tag("footnote"))
        fn.set(base.w_tag("id"), str(fid))
        fn.set(base.w_tag("type"), ftype)
        p = etree.SubElement(fn, base.w_tag("p"))
        r = etree.SubElement(p, base.w_tag("r"))
        etree.SubElement(r, base.w_tag(marker))

    for fid, text in sorted(id_to_text.items()):
        fn = etree.SubElement(root, base.w_tag("footnote"))
        fn.set(base.w_tag("id"), str(fid))

        p = etree.SubElement(fn, base.w_tag("p"))
        ppr = etree.SubElement(p, base.w_tag("pPr"))
        etree.SubElement(ppr, base.w_tag("pStyle")).set(base.w_tag("val"), "FootnoteText")

        # Justify footnote paragraph text.
        jc = etree.SubElement(ppr, base.w_tag("jc"))
        jc.set(base.w_tag("val"), "both")

        ind = etree.SubElement(ppr, base.w_tag("ind"))
        ind.set(base.w_tag("left"), "567")
        ind.set(base.w_tag("hanging"), "567")

        spacing = etree.SubElement(ppr, base.w_tag("spacing"))
        spacing.set(base.w_tag("before"), "0")
        spacing.set(base.w_tag("after"), "0")
        spacing.set(base.w_tag("line"), "240")
        spacing.set(base.w_tag("lineRule"), "auto")

        # Footnote number: Arial 9pt superscript.
        r_ref = etree.SubElement(p, base.w_tag("r"))
        rpr = etree.SubElement(r_ref, base.w_tag("rPr"))
        etree.SubElement(rpr, base.w_tag("rStyle")).set(base.w_tag("val"), "FootnoteReference")
        rfonts_ref = etree.SubElement(rpr, base.w_tag("rFonts"))
        rfonts_ref.set(base.w_tag("ascii"), "Arial")
        rfonts_ref.set(base.w_tag("hAnsi"), "Arial")
        sz_ref = etree.SubElement(rpr, base.w_tag("sz"))
        sz_ref.set(base.w_tag("val"), "18")
        szcs_ref = etree.SubElement(rpr, base.w_tag("szCs"))
        szcs_ref.set(base.w_tag("val"), "18")
        vert_ref = etree.SubElement(rpr, base.w_tag("vertAlign"))
        vert_ref.set(base.w_tag("val"), "superscript")
        etree.SubElement(r_ref, base.w_tag("footnoteRef"))

        r_tab = etree.SubElement(p, base.w_tag("r"))
        etree.SubElement(r_tab, base.w_tag("tab"))

        # Footnote body: Arial 9pt.
        r_text = etree.SubElement(p, base.w_tag("r"))
        rpr_text = etree.SubElement(r_text, base.w_tag("rPr"))
        rfonts = etree.SubElement(rpr_text, base.w_tag("rFonts"))
        rfonts.set(base.w_tag("ascii"), "Arial")
        rfonts.set(base.w_tag("hAnsi"), "Arial")
        sz = etree.SubElement(rpr_text, base.w_tag("sz"))
        sz.set(base.w_tag("val"), "18")
        szcs = etree.SubElement(rpr_text, base.w_tag("szCs"))
        szcs.set(base.w_tag("val"), "18")
        t = etree.SubElement(r_text, base.w_tag("t"))
        t.text = text

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


# =============================================================================
# Entrypoint
# =============================================================================

def build_document(html_path: str | Path, output_path: str | Path) -> Path:
    """Run v3 conversion with v5 overrides applied."""
    base.add_annexes = add_annexes
    base.build_annex_vi_table = build_annex_vi_table
    base.create_footnotes_xml = create_footnotes_xml
    return base.build_document(html_path, output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="EUR-Lex XHTML to formatted Word with v5 Annex and borderless Annex VI fixes.")
    parser.add_argument("source", help="Path to EUR-Lex XHTML/HTML input file")
    parser.add_argument("output", help="Path to output .docx file")
    args = parser.parse_args()
    output = build_document(args.source, args.output)
    print(f"Stage 2 document saved: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
