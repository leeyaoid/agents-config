#!/usr/bin/env python3
"""Build an isolated public-document template from Word's Normal styles."""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document

from generate_official_doc import DEFAULT_TEMPLATE_PATH, _configure_page, create_official_styles


def build_template(output_path: str | Path = DEFAULT_TEMPLATE_PATH) -> str:
    document = Document()
    _configure_page(document)
    create_official_styles(document)
    document.core_properties.title = "党政机关公文Normal基础模板"
    document.core_properties.subject = "GB/T 9704-2012公文格式"
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    return str(output)


def main():
    parser = argparse.ArgumentParser(description="构建公文Normal基础模板")
    parser.add_argument("--output", default=str(DEFAULT_TEMPLATE_PATH))
    args = parser.parse_args()
    print(build_template(args.output))


if __name__ == "__main__":
    main()
