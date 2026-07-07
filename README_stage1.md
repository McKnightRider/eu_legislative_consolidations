# Stage 1 — XML/HTML to Word conversion

This is the first stage of the legislative consolidation workflow.  It converts a source XML/HTML legislative document into a canonical `.docx` baseline suitable for Stage 2 QA and later consolidation.

## Files

- `environment.yml` — Conda environment definition.
- `stage1_conversion_helpers.py` — source loading, parsing, block extraction and classification.
- `stage1_word_writer.py` — Word styles and `.docx` generation.
- `run_stage1_conversion.py` — short command-line entry point.

## Set up the environment

```bash
conda env create -f environment.yml
conda activate legislative-consolidator
```

## Run the conversion

```bash
python run_stage1_conversion.py /path/to/source.html
```

or specify the output:

```bash
python run_stage1_conversion.py /path/to/source.xml --output /path/to/EUPR_v0_Original.docx
```

## Important QA point

The generated Word document should not be treated as an authoritative master text until it has passed Stage 2 QA.  This script is intentionally conservative and records warnings where extraction is uncertain.
