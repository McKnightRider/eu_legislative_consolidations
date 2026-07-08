# EU Legislative Consolidations
This repository contains code that (a) converts a piece of EU legislation into Word and (b) shows changes to that EU legislation in different colours.

## Steps
There are several steps to the process:
1. Python-based conversion - structured XML/HTML (not pdf) -> identify provisions (e.g. recitals, articles, footnotes, annexes) -> Word document.
2. Quality check - checks Word document against original legislation (perhaps the pdf version as that can be used for Litera comparisons).
3. Python-based consolidation engine - identify amending regulation -> add recitals -> parse amending provisions -> apply amendments -> new Word document.
4. Quality check - checks Word document changes (perhaps against consolidated pdf version on Europa website).
5. Legislative consolidation assistant - agent that locates documents -> identifies amendment requirements -> invokes consolidation scripts -> return finished Word document.

There are potentially future developments, such as the agent noting where particular amendments came from and official EU/ESMA commentary.

## Setting up Conda Environment
INSERT

---

## HTML structure and Word formatting reference (`formatting_stage2_v6.py`)

### Document sections (identified by `id` attributes)

| HTML `id` pattern | Content | Word treatment |
|---|---|---|
| `cit_N` | Citations ("Having regard to…") | Flush left, justified |
| `rct_N` | Recitals ("(1) …") | Hanging indent: left 1 cm, first-line −1 cm |
| `cpt_ROMAN` | Chapter container | Chapter heading: number not bold, title bold, centred |
| `art_N[A-Z]?` | Article container | Article heading: number italic, title bold, centred; then body paragraphs |
| `anx_ROMAN` | Annex container | Annex heading on new page (label uppercase italic, title bold, centred) |

### Paragraph classes inside articles

All body paragraphs inside articles carry class `oj-normal`.  **Nesting level is
not encoded in the CSS class** — it is encoded by the number of `<table>` ancestors
between the `<p>` and its article container:

| `<table>` ancestors | List level | Left indent | Examples |
|---|---|---|---|
| 0 | Top-level | 0 cm | Digit markers `1.`, `2.`; chapeau and continuation text |
| 1 | Letter level | 1 cm | Items `(a)`, `(b)`, …, `(i)`, `(j)` |
| 2 | Roman level | 2 cm | Sub-items `(i)`, `(ii)`, `(iii)` etc. |
| 3+ | Further nesting | 3 cm, 4 cm, … | (Uncommon) |

All marker paragraphs use a **hanging indent of −1 cm**, so the text body aligns
1 cm to the right of the marker.  Continuation paragraphs (non-marker text at the
same table-nesting depth) sit at the **same** `left_cm` as the marker level with no
extra indent.

Standalone markers (a `<p>` containing only e.g. `(a)`) are automatically combined
with the immediately following body `<p>` into a single hanging-indent paragraph
separated by a tab.

### Heading classes

| CSS class | Content | Word treatment |
|---|---|---|
| `eli-main-title` | Regulation main title | Bold, centred; automatic line breaks before "of 14 June 2017" and "(Text with EEA relevance)" |
| `oj-ti-section-1` / `oj-ti-section-2` | Chapter/section titles | Bold, centred |
| `oj-sti-art` | Article sub-title | Bold, centred |
| `oj-doc-ti` (inside annex) | Annex label e.g. "ANNEX I" | Uppercase italic, centred; triggers page break before |
| `oj-ti-tbl` (Annex VI only) | Annex VI table headings | Non-bold, non-italic, centred |

### Footnotes

Footnote references are `<sup>` or `.oj-super` elements containing the number.
Footnote bodies are `p.oj-note` paragraphs with a `.oj-note-tag` child.

The converter tokenises references as `[[FN:N]]`, then replaces them with native
OOXML footnote references in a post-processing zip-patch step.  Footnote text is
justified (`w:jc val="both"`) with a 1 cm hanging indent.

### Annex structure (Annexes I–V)

Inside an `anx_ROMAN` container:

| Element | Treatment |
|---|---|
| `p.oj-doc-ti` | Annex label/title heading |
| `div.oj-enumeration-spacing` | Section heading (`I.\tSummary`) + body paragraphs |
| `table` (direct child) | Rows emitted as tab-separated text at 2 cm indent |
| `p.oj-normal` (direct child) | Explanatory text, flush left |

**Annex VI** is rendered as a two-column Word table with a top border and vertical
divider only (no outer side/bottom borders); header row gets a bottom border.

### Source anchors

Each emitted paragraph contains a hidden run `[[SRC:<id>.<index>]]` (1 pt font,
`w:vanish`) invisible in normal view but present in the document XML for automated
indent validation.  The validation script writes results to
`/home/swm35/article_misindent_report.txt`.
