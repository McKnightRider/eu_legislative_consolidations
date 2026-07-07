"""
diagnostic_eurlex_structure.py

Diagnostic utility for analysing EUR-Lex XHTML files.

Purpose:
    Understand how recitals, annexes and footnotes are actually
    represented in the XHTML before changing the parser again.
"""

from pathlib import Path
from bs4 import BeautifulSoup, Tag
import csv


def classes_of(tag):
    cls = tag.get("class", [])
    return " ".join(cls)


def text_of(tag):
    return " ".join(tag.get_text(" ", strip=True).split())


def load_document(path):
    html = Path(path).read_text(
        encoding="utf-8",
        errors="replace"
    )
    return BeautifulSoup(html, "lxml")


def find_first_article(soup):
    return soup.select_one(".oj-ti-art")


def dump_pre_article_content(soup):
    """
    Show all elements before the first Article.

    This should reveal how recitals are encoded.
    """

    article = find_first_article(soup)

    rows = []

    for tag in article.find_all_previous():

        if not isinstance(tag, Tag):
            continue

        text = text_of(tag)

        if not text:
            continue

        rows.append({
            "tag": tag.name,
            "classes": classes_of(tag),
            "id": tag.get("id", ""),
            "text": text[:250]
        })

    rows.reverse()

    with open(
        "diagnostic_pre_article.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=rows[0].keys()
        )

        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Pre-Article elements written: "
        f"{len(rows)}"
    )


def dump_recital_candidates(soup):
    """
    Find anything that looks remotely like a recital.
    """

    rows = []

    for tag in soup.find_all(True):

        text = text_of(tag)

        if not text:
            continue

        if text.startswith("("):

            rows.append({
                "tag": tag.name,
                "classes": classes_of(tag),
                "id": tag.get("id", ""),
                "text": text[:250]
            })

    with open(
        "diagnostic_recital_candidates.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=rows[0].keys()
        )

        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Recital candidates written: "
        f"{len(rows)}"
    )


def dump_annex_candidates(soup):
    """
    Show every element containing ANNEX.
    """

    rows = []

    for tag in soup.find_all(True):

        text = text_of(tag)

        if "ANNEX" in text.upper():

            rows.append({
                "tag": tag.name,
                "classes": classes_of(tag),
                "parent_classes":
                    classes_of(tag.parent)
                    if isinstance(tag.parent, Tag)
                    else "",
                "id": tag.get("id", ""),
                "text": text[:250]
            })

    with open(
        "diagnostic_annex_candidates.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=rows[0].keys()
        )

        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Annex candidates written: "
        f"{len(rows)}"
    )


def dump_footnotes(soup):
    """
    Examine actual footnote structure.
    """

    rows = []

    notes = soup.select(".oj-note")

    for note in notes:

        rows.append({
            "tag": note.name,
            "classes": classes_of(note),
            "id": note.get("id", ""),
            "text": text_of(note)[:500]
        })

    with open(
        "diagnostic_footnotes.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=rows[0].keys()
        )

        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Footnotes found: "
        f"{len(rows)}"
    )


def dump_structure_classes(soup):
    """
    Show all EUR-Lex structural classes.
    """

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
        "oj-table"
    ]

    rows = []

    for cls in interesting:

        for tag in soup.select(f".{cls}"):

            rows.append({
                "class": cls,
                "tag": tag.name,
                "id": tag.get("id", ""),
                "text": text_of(tag)[:250]
            })

    with open(
        "diagnostic_structure_classes.csv",
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=rows[0].keys()
        )

        writer.writeheader()
        writer.writerows(rows)

    print(
        f"Structure rows written: "
        f"{len(rows)}"
    )


def main():

    import argparse

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "source",
        help="EUR-Lex XHTML file"
    )

    args = parser.parse_args()

    soup = load_document(args.source)

    dump_pre_article_content(soup)

    dump_recital_candidates(soup)

    dump_annex_candidates(soup)

    dump_footnotes(soup)

    dump_structure_classes(soup)

    print("Diagnostics complete.")


if __name__ == "__main__":
    main()
