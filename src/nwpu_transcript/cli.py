from __future__ import annotations

import argparse
from pathlib import Path
from typing import Callable, Dict, List, Tuple

from .excel import write_to_template
from .parser import parse_chinese, parse_english


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert NWPU transcript PDFs into the Excel upload template."
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("课程分学期模版.xlsx"),
        help="Path to the Excel template workbook.",
    )
    parser.add_argument(
        "--chinese",
        type=Path,
        help="Path to the Chinese transcript PDF.",
    )
    parser.add_argument(
        "--english",
        type=Path,
        help="Path to the English transcript PDF.",
    )
    parser.add_argument(
        "--output-chinese",
        type=Path,
        default=Path("transcript_chinese.xlsx"),
        help="Output path for the Chinese Excel file.",
    )
    parser.add_argument(
        "--output-english",
        type=Path,
        default=Path("transcript_english.xlsx"),
        help="Output path for the English Excel file.",
    )
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()

    template_path = args.template.resolve()
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    tasks: List[Tuple[Callable[[Path], List[Dict[str, str]]], Path, Path]] = []

    if args.chinese:
        tasks.append((parse_chinese, args.chinese.resolve(), args.output_chinese))

    if args.english:
        tasks.append((parse_english, args.english.resolve(), args.output_english))

    if not tasks:
        raise SystemExit("No transcripts provided. Use --chinese and/or --english.")

    for parser_func, pdf_path, output_path in tasks:
        if not pdf_path.exists():
            raise FileNotFoundError(f"Transcript not found: {pdf_path}")
        records = parser_func(pdf_path)
        write_to_template(template_path, records, output_path)
        print(f"Wrote {len(records)} rows to {output_path}")


if __name__ == "__main__":  # pragma: no cover
    main()

