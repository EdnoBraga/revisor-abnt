#!/usr/bin/env python3
"""Gera auditoria ABNT estruturada de um DOCX sem modificar o original."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document

from abnt_engine import ReviewConfig, scan_document


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--document-type", choices=("tcc", "article"), default="tcc")
    parser.add_argument("--citation-system", choices=("auto", "author-date", "numeric"), default="auto")
    args = parser.parse_args()

    source = args.input.resolve()
    output = args.out or source.with_name(f"{source.stem}_abnt_report.json")
    report = scan_document(Document(source), ReviewConfig(document_type=args.document_type, citation_system=args.citation_system))
    report["input"] = str(source)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
