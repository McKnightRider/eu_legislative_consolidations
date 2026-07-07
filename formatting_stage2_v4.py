
"""
formatting_stage2_v4.py

Small overlay update for formatting_stage2_v3.py.

Place this file in the same folder as formatting_stage2_v3.py and run this file
instead of v3:

    python formatting_stage2_v4.py L_2017168EN.01001201.xml.html Prospectus_Regulation_Stage2.docx

Changes from v3:
    1. Annexes I-V: preserve level-1 annex markers such as A., B., C. even where
       EUR-Lex stores them in .tit_ nodes.  In v3, .tit_ nodes were excluded too
       aggressively, which could remove those level-1 markers.
    2. Annexes I-V: merge standalone A./B./C. marker paragraphs with the following
       paragraph using a tab.
    3. Footnotes: explicitly justify footnote paragraphs by adding w:jc="both" to
       footnote paragraph properties.

This overlay deliberately reuses v3 for the rest of the conversion logic so the
latest working behaviour is preserved.
"""

from __future__ import annotations

from pathlib import Path
import argparse
import re
from bs4 import Tag
from lxml import etree

import formatting_stage2_v3 as base


def _is_annex_level_marker(text: str) -> bool:
    """True for Annex level-1 markers such as A., B., I., II., etc."""
    t = base.clean_text(text)
    return bool(re.match(r"^([A-Z]\.|[IVXLCDM]+\.)$", t, flags=re.I))


def _is_annex_heading_or_marker(text: str) -> bool:
    """True for standalone or combined Annex level-1 headings.

    Examples:
        A.
        A. Offer statistics
        I.
        I. Risk factors
    """
    t = base.clean_text(text)
    return bool(re.match(r"^([A-Z]\.|[IVXLCDM]+\.)(\s+.+)?$", t, flags=re.I))


def add_annexes(doc, soup) -> None:
    """Replacement v4 Annex handler.

    v3 skipped all .tit_ nodes.  That is correct for operative Article metadata
    such as art_31.tit_1, but too aggressive for Annexes I-V because the visible
    level-1 Annex markers may be stored in title nodes.  This version preserves
    title nodes where their visible text is an Annex marker or Annex marker heading.
    """
    for annex in soup.find_all(id=re.compile(r"^anx_[IVXLCDM]+$")):
        heading = annex.select_one("p.oj-doc-ti")
        heading_text = base.clean_text(heading.get_text(" ", strip=True)) if heading else str(annex.get("id", "Annex"))

        hp = base.add_center_heading(doc, heading_text)
        hp.paragraph_format.page_break_before = True

        if annex.get("id") == "anx_VI":
            base.build_annex_vi_table(doc, annex)
            continue

        paras: list[Tag] = []
        for p in annex.find_all("p"):
            if base.has_class(p, "oj-note"):
                continue

            txt = base.text_with_footnote_tokens(p)
            if not txt or txt == heading_text:
                continue

            pid = str(p.get("id", ""))
            classes = base.tag_classes(p)

            # Preserve Annex .tit_ nodes only where they are visible Annex headings
            # or marker headings, e.g. A. / A. Offer statistics.  Continue to skip
            # other title metadata.
            if ".tit_" in pid and not _is_annex_heading_or_marker(txt):
                continue

            if (
                _is_annex_heading_or_marker(txt)
                or any(c in classes for c in ["oj-normal", "oj-enumeration-spacing", "oj-doc-ti", "oj-ti-section-1", "oj-ti-section-2"])
            ):
                paras.append(p)

        # Pass the full Annex paragraph sequence to preserve and merge A./B./etc.
        # with the following paragraph using a tab.
        base.emit_numbered_paragraph_sequence(doc, paras, annex_mode=True)


def create_footnotes_xml(id_to_text: dict[int, str]) -> bytes:
    """Replacement v4 native footnotes XML builder with justified footnotes."""
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

        # NEW in v4: justify footnote text.
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

        # Footnote number: Arial 9 pt superscript.
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

        # Footnote body: Arial 9 pt.
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


def build_document(html_path: str | Path, output_path: str | Path) -> Path:
    """Run v3 conversion with v4 overrides applied."""
    base.add_annexes = add_annexes
    base.create_footnotes_xml = create_footnotes_xml
    return base.build_document(html_path, output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="EUR-Lex XHTML to formatted Word with v4 Annex and justified-footnote fixes.")
    parser.add_argument("source", help="Path to EUR-Lex XHTML/HTML input file")
    parser.add_argument("output", help="Path to output .docx file")
    args = parser.parse_args()
    output = build_document(args.source, args.output)
    print(f"Stage 2 document saved: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
