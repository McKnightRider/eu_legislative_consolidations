# EU Legislative Consolidations

This repository supports a multi-stage legislative workflow from EUR-Lex source files to quality-checked consolidation outputs.

## Workflow Stages

### Stage 1 - Conversion and Formatting (active)

Purpose:
- Convert EUR-Lex XML/HTML into a formatted Word document.
- Generate a parser structure report.
- Generate structural diagnostics CSVs.

Primary files:
- `stage1.py` (main entrypoint)
- `eurlex_parser.py` (parser and report generation; standalone or imported)
- `diagnostic_eurlex_structure.py` (diagnostics; standalone or imported)

Run Stage 1:

```bash
python stage1.py /path/to/source.html EUPR_Stage1.docx
```

Output behavior:
- File-name outputs are written to `outputs/` by default.
- Parser JSON report: `outputs/<source_stem>_structure_report.json`
- Diagnostics CSVs: `outputs/diagnostics/`

Optional flags:

```bash
python stage1.py /path/to/source.html EUPR_Stage1.docx --skip-diagnostics
python stage1.py /path/to/source.html EUPR_Stage1.docx --parser-report /custom/path/report.json
```

Standalone parser/diagnostics:

```bash
python eurlex_parser.py /path/to/source.html
python diagnostic_eurlex_structure.py /path/to/source.html --output-dir outputs/diagnostics
```

### Stage 1 cheat sheet: which one to run when

Use `eurlex_parser.py` when you want parsed legal structure data for the pipeline.

- Run this when you need the canonical parse output (recitals, sections, articles, annexes, footnotes).
- Run this when you want the JSON structure report used for downstream validation and debugging.

```bash
python eurlex_parser.py /path/to/source.html
```

Use `diagnostic_eurlex_structure.py` when you want investigative CSVs to diagnose extraction issues.

- Run this when counts look wrong (for example: too many recitals, missing annexes, odd footnote totals).
- Run this when you need to inspect candidate elements and class distribution quickly in spreadsheets.

```bash
python diagnostic_eurlex_structure.py /path/to/source.html --output-dir outputs/diagnostics
```

Use `stage1.py` when you want the full Stage 1 deliverable in one run.

- Runs conversion/formatting to DOCX.
- Runs parser report generation.
- Runs diagnostics CSV generation (unless `--skip-diagnostics` is provided).

```bash
python stage1.py /path/to/source.html EUPR_Stage1.docx
```

Quick decision rule:

- Need final Stage 1 output: run `stage1.py`.
- Need structured parse only: run `eurlex_parser.py`.
- Need troubleshooting evidence: run `diagnostic_eurlex_structure.py`.

### Stage 2 - Quality Assurance (active)

Purpose:
- Validate Stage 1 output against source structure and expected formatting behavior.
- Detect missing-word risks before downstream use.

Primary file:
- `stage2.py` (DOCX coverage QA against HTML and/or PDF)

Run Stage 2:

```bash
python stage2.py /path/to/stage1_output.docx --html /path/to/source.html
python stage2.py /path/to/stage1_output.docx --html /path/to/source.html --pdf /path/to/source.pdf
```

Output behavior:
- JSON report(s): `outputs/qa/<docx_stem>_stage2_qa_report_html.json` and/or `outputs/qa/<docx_stem>_stage2_qa_report_pdf.json`
- Console summary per source with coverage and missing token counts.

Useful flags:

```bash
python stage2.py /path/to/stage1_output.docx --html /path/to/source.html --fail-on-missing
python stage2.py /path/to/stage1_output.docx --html /path/to/source.html --min-coverage 0.999
```

By default, Stage 2 writes the report and returns success. Use `--fail-on-missing` and/or
`--min-coverage` when you want CI-style failure conditions.

Notes:
- HTML comparison is usually the most reliable baseline for text coverage.
- PDF text extraction can differ due to OCR, ligatures, hyphenation, and layout artifacts.

### Stage 3 - Consolidation Engine (draft)

Purpose:
- Identify amending instruments.
- Parse amendment provisions.
- Apply amendment layers to produce a consolidated Word version.

Primary file:
- `stage3.py` (draft amendment layering and word-level change analysis)

Run Stage 3:

```bash
python stage3.py /path/to/current_version.docx /path/to/amending_regulation.html red
```

Inputs:
- Current consolidated DOCX (Stage 1 output for first run, then previous Stage 3 output)
- Amending regulation HTML
- Color for this amendment layer

Outputs:
- Consolidated DOCX: auto-numbered as `outputs/stage3_<n>.docx`
- Analysis report: `outputs/qa/<output_stem>_amendment_analysis.json`

Numbering rule:
- If base file is `stage1.docx`, output is `stage3_1.docx`.
- If base file is `stage3_1.docx`, output is `stage3_1.docx` (overwrite).
- If base file is `stage3_2.docx`, output is `stage3_2.docx` (overwrite), and so on.

If the target `stage3_<n>.docx` already exists, it is overwritten.

Current draft behavior:
- Inserts amending title below existing title.
- Inserts amending recitals below existing recitals.
- Adds careful word-level analysis for amending article blocks against closest prior text.
- Enforces a one-go full-application check: Stage 3 fails if detected substantive amendment items remain `analysis_only`.

Deterministic safety layer (current):
- Stage 3 now emits an operation-proof record in the analysis JSON (`operation_proof`).
- Each detected amendment row is represented as a typed operation (`replace`, `delete`, `insert`, `unknown`) with:
	- target path (for example `article/5/paragraph/1` or `annex/ix/point/ii`)
	- required preconditions (target scope present, paragraph scope for subparagraph instructions, non-empty amendment payload)
	- proof status (`proven` vs `unresolved`) and applied mode.
- Hard precondition gate: with full-application mode, Stage 3 fails if any operation misses required preconditions.
- Fail-closed behavior: unresolved substantive operations also fail the one-go application check.

Where to inspect this in output:
- `outputs/qa/<output_stem>_amendment_analysis.json`
	- `operation_proof.summary`
	- `operation_proof.precondition_failure_examples`
	- `operation_proof.operations`
	- `one_go_application_check`

If you intentionally need partial/staggered application in a specific workflow, you can opt out:

```bash
python stage3.py /path/to/current_version.docx /path/to/amending_regulation.html red --allow-partial-application
```

### Stage 4 - Consolidation QA (active)

Purpose:
- Check whether Stage 3 insertions/replacements were applied and visibly revision-formatted.
- Check whether the effective final text (after removing strike-through deletions) matches the official EU consolidated HTML text.

Primary file:
- `stage4.py`

Required/expected inputs:
- Stage 3 output DOCX (e.g. `outputs/stage3_1.docx`)
- Stage 3 analysis JSON (auto-detected by default at `outputs/qa/<docx_stem>_amendment_analysis.json`)
- Official EU consolidated HTML (recommended under new `consolidated/` folder)

Run Stage 4:

```bash
python stage4.py outputs/stage3_1.docx --html consolidated/stage3_1.html --pdf consolidated/stage3_1.pdf
```

Convenience behavior:
- If `--html` is omitted, Stage 4 tries `consolidated/<docx_stem>.html` automatically.
- If `--pdf` is omitted, Stage 4 tries `consolidated/<docx_stem>.pdf` automatically.

Output behavior:
- JSON report: `outputs/qa/<docx_stem>_stage4_qa_report.json`

Useful flags:

```bash
python stage4.py outputs/stage3_1.docx --html consolidated/stage3_1.html --pdf consolidated/stage3_1.pdf --min-coverage 0.999
python stage4.py outputs/stage3_1.docx --html consolidated/stage3_1.html --pdf consolidated/stage3_1.pdf --fail-on-issues
python stage4.py outputs/stage3_1.docx --analysis-json outputs/qa/stage3_1_amendment_analysis.json
```

Recurring run pattern examples:

```bash
# After Stage 3 run for stage3_2:
python stage4.py outputs/stage3_2.docx --html consolidated/stage3_2.html --pdf consolidated/stage3_2.pdf --fail-on-issues

# After Stage 3 run for stage3_3:
python stage4.py outputs/stage3_3.docx --html consolidated/stage3_3.html --pdf consolidated/stage3_3.pdf --fail-on-issues

# If naming is consistent, --html/--pdf can be omitted (auto-lookup in consolidated/):
python stage4.py outputs/stage3_3.docx --fail-on-issues
```

Watch out (common but not universal):
- A single amending act can produce multiple official consolidated snapshots with different application dates (staggered implementation).
- In those cases, run Stage 4 as separate checkpoint comparisons (one Stage 3 output matched to one official consolidated HTML snapshot).
- This pattern appears in some EUR-Lex regulations, but not all. Do not assume one amendment always maps to one final consolidated state.

Folder convention:
- `inputs/` keeps source/amending inputs for processing stages.
- `consolidated/` stores official EU consolidated HTML files used by Stage 4 equivalence checks.

### Stage 5 - Assistant Orchestration (planned)

Purpose:
- Automate retrieval, amendment detection, pipeline invocation, and delivery of final outputs.

## Environment Setup

```bash
conda env create -f environment.yml
conda activate legislative-consolidator
```

---

## Stage 1 HTML Structure and Formatting Reference (`stage1.py`)

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
indent validation. The optional diagnostics are written under `outputs/diagnostics/`.
