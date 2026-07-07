"""
Command-line entry point for Stage 1 conversion.

Usage
-----
    python run_stage1_conversion.py path/to/source.xml
    python run_stage1_conversion.py path/to/source.html --output path/to/output.docx

The script takes one required input: the path to the XML/HTML file.
If --output is not supplied, the output is written next to the source file with a
"_stage1_converted.docx" suffix.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from stage1_conversion_helpers import parse_legislative_source
from stage1_word_writer import write_legislative_docx


def default_output_path(input_path: Path) -> Path:
    """Create a default output path next to the input source file."""
    return input_path.with_name(f"{input_path.stem}_stage1_converted.docx")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Stage 1: convert legislative XML/HTML source into a canonical Word document."
    )
    parser.add_argument(
        "source",
        help="Path to the XML or HTML legislative source file."
    )
    parser.add_argument(
        "--output",
        "-o",
        help="Optional path for the generated .docx file. Defaults to SOURCE_stem_stage1_converted.docx."
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    source_path = Path(args.source).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve() if args.output else default_output_path(source_path)

    try:
        model = parse_legislative_source(source_path)
        written = write_legislative_docx(model, output_path)
    except Exception as exc:
        print(f"Stage 1 conversion failed: {exc}", file=sys.stderr)
        return 1

    print("Stage 1 conversion completed.")
    print(f"Source: {source_path}")
    print(f"Output: {written}")
    print(f"Detected source type: {model.metadata.get('source_type')}")
    print(f"Blocks written: {model.metadata.get('block_count')}")
    if model.warnings:
        print("Warnings:")
        for warning in model.warnings:
            print(f" - {warning}")
    else:
        print("Warnings: none")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
