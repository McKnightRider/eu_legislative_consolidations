"""
formatting_stage2.py

Stage 2 formatting engine:
- Rebuilds clean Word document
- Converts footnotes to native Word footnotes
- Parses Article hierarchy
- Restores inline formatting (basic)
- Builds Annex VI table
"""

from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


# -----------------------------
# LOAD XHTML
# -----------------------------

def load_html(path):
    with open(path, encoding="utf-8") as f:
        return BeautifulSoup(f, "lxml")


# -----------------------------
# WORD SETUP
# -----------------------------

def new_document():
    doc = Document()

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Arial"
    font.size = Pt(11)

    return doc


def add_para(doc, text, bold=False, italic=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    return p


# -----------------------------
# FOOTNOTE ENGINE
# -----------------------------

def extract_footnotes(soup):
    notes = {}
    for i, p in enumerate(soup.select("p.oj-note"), start=1):
        notes[str(i)] = p.get_text(" ", strip=True)
    return notes


def add_footnote(paragraph, text):
    """
    Adds a Word footnote (low-level XML).
    """
    run = paragraph.add_run()
    footnote = run._element.add_footnote_reference()
    footnote.text = text


def process_text_with_footnotes(doc, text, footnotes):
    import re

    p = doc.add_paragraph()

    parts = re.split(r"\((\d+)\)", text)

    for i, part in enumerate(parts):
        if i % 2 == 0:
            p.add_run(part)
        else:
            add_footnote(p, footnotes.get(part, ""))

    return p


# -----------------------------
# RECITALS
# -----------------------------

def add_recital(doc, text):
    import re

    match = re.match(r"\((\d+)\)\s*(.*)", text)

    if not match:
        doc.add_paragraph(text)
        return

    number, body = match.groups()

    p = doc.add_paragraph()

    p.add_run(f"({number})\t")

    run = p.add_run(body)

    fmt = p.paragraph_format
    fmt.left_indent = Cm(1)
    fmt.first_line_indent = Cm(-1)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


# -----------------------------
# ARTICLE PARSING
# -----------------------------

def detect_level(text):
    import re

    if re.match(r"^\(\d+\)", text):
        return 1
    if re.match(r"^\([a-z]\)", text):
        return 2
    if re.match(r"^\([ivx]+\)", text):
        return 3
    return 0


def add_article_paragraph(doc, text):

    level = detect_level(text)

    p = doc.add_paragraph(text)

    fmt = p.paragraph_format
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if level == 1:
        fmt.left_indent = Cm(1)
        fmt.first_line_indent = Cm(-1)
    elif level == 2:
        fmt.left_indent = Cm(2)
        fmt.first_line_indent = Cm(-1)
    elif level == 3:
        fmt.left_indent = Cm(3)
        fmt.first_line_indent = Cm(-1)
    else:
        fmt.left_indent = Cm(1)


# -----------------------------
# INLINE FORMATTING
# -----------------------------

def add_formatted_text(doc, element):
    p = doc.add_paragraph()

    for node in element.descendants:

        if node.name == "b":
            p.add_run(node.text).bold = True

        elif node.name == "i":
            p.add_run(node.text).italic = True

        elif node.string:
            p.add_run(node.string)

    return p


# -----------------------------
# ANNEX VI TABLE
# -----------------------------

def build_annex_vi_table(doc, annex):

    rows = annex.get_text("\n").split("\n")

    table = doc.add_table(rows=0, cols=2)

    for row in rows:

        if not row.strip():
            continue

        cells = table.add_row().cells

        if "Directive" in row:
            left, right = row.split(" ", 1)
        else:
            parts = row.split(" ", 1)
            left = parts[0]
            right = parts[1] if len(parts) > 1 else ""

        cells[0].text = left.strip()
        cells[1].text = right.strip()


# -----------------------------
# MAIN PIPELINE
# -----------------------------

def build_document(html_path, output_path):

    soup = load_html(html_path)
    doc = new_document()

    footnotes = extract_footnotes(soup)

    # TITLE
    title = soup.select_one(".eli-main-title")
    if title:
        p = doc.add_heading(title.get_text(" "), 0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # RECITALS
    doc.add_heading("Recitals", 1)

    for r in soup.select('div[id^="rct_"]'):
        text = r.get_text(" ", strip=True)
        add_recital(doc, text)

    # ARTICLES
    doc.add_heading("Articles", 1)

    for art in soup.select('div[id^="art_"]'):

        title_tag = soup.find(id=f"{art['id']}.tit_1")

        if title_tag:
            h = doc.add_heading(title_tag.text, 2)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER

        paragraphs = art.get_text("\n").split("\n")

        for ptext in paragraphs:
            if ptext.strip():
                add_article_paragraph(doc, ptext.strip())

    # ANNEXES
    doc.add_heading("Annexes", 1)

    for annex in soup.select('div[id^="anx_"]'):

        heading = annex.select_one("p.oj-doc-ti")

        if heading:
            doc.add_heading(heading.text, 2)

        if annex["id"] == "anx_VI":
            build_annex_vi_table(doc, annex)
        else:
            doc.add_paragraph(annex.get_text(" "))

    # FOOTNOTES handled inline via references

    doc.save(output_path)

    print(f"Stage 2 document saved: {output_path}")


# -----------------------------
# RUN
# -----------------------------

if __name__ == "__main__":

    build_document(
        "L_2017168EN.01001201.xml.html",
        "Prospectus_Regulation_Stage2.docx"
    )
