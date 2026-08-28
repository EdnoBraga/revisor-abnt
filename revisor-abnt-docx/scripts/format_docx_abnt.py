#!/usr/bin/env python3
"""Aplica regras ABNT verificáveis a uma cópia de um DOCX e registra cada ação."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from abnt_engine import DEFAULT_FONT, ReviewConfig, apply_formatting
from pdf_report import build_pdf_report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--document-type", choices=("tcc", "article"), default="tcc")
    parser.add_argument("--citation-system", choices=("auto", "author-date", "numeric"), default="auto")
    parser.add_argument("--font", choices=("Times New Roman", "Arial"), default=DEFAULT_FONT)
    parser.add_argument("--toc-mode", choices=("audit", "insert-if-empty"), default="insert-if-empty")
    parser.add_argument("--pagination-mode", choices=("audit", "request"), default="request")
    parser.add_argument("--do-not-order-references", action="store_true")
    parser.add_argument("--institution", choices=("generic", "cgaem"), default="generic")
    args = parser.parse_args()

    config = ReviewConfig(
        document_type=args.document_type,
        citation_system=args.citation_system,
        font=args.font,
        toc_mode=args.toc_mode,
        pagination_mode=args.pagination_mode,
        order_references=not args.do_not_order_references,
        institution=args.institution,
    )
    report = apply_formatting(args.input.resolve(), args.out.resolve(), config)
    report_path = args.out.resolve().with_name(f"{args.out.stem}_abnt_report.json")
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    pdf_report_path = args.out.resolve().with_name(f"{args.out.stem}_abnt_report.pdf")
    build_pdf_report(report, pdf_report_path, original_filename=args.input.name)
    print(args.out.resolve())
    print(report_path)
    print(pdf_report_path)


if __name__ == "__main__":
    main()
