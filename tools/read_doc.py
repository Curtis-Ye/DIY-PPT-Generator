"""Read .docx / .pdf files and extract text with structure.

Usage:
    python tools/read_doc.py <file_path>
    python tools/read_doc.py <file_path> --pretty  # formatted output

Output (JSON to stdout):
    {
      "title": "文档标题",
      "headings": [{"level": 1, "text": "...", "order": 0}, ...],
      "paragraphs": [{"text": "...", "order": 0}, ...],
      "tables": [[["cell", ...], ...], ...],
      "stats": {"paragraphs": 42, "headings": 7, "tables": 3, "chars": 12345}
    }
"""

import json
import sys
import os
from pathlib import Path


def read_docx(path: str) -> dict:
    from docx import Document

    doc = Document(path)

    # Extract title from document properties or first heading
    title = doc.core_properties.title or ""
    if not title:
        for para in doc.paragraphs:
            if para.style.name.startswith("Heading") or para.style.name.startswith("Title"):
                title = para.text.strip()
                break
        if not title:
            title = Path(path).stem

    headings = []
    paragraphs = []
    tables = []
    para_order = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style = para.style.name if para.style else ""

        if style.startswith("Heading"):
            level = 1
            try:
                level = int(style.split()[-1])
            except ValueError:
                level = 1
            headings.append({
                "level": min(level, 6),
                "text": text,
                "order": para_order,
            })

        paragraphs.append({
            "text": text,
            "order": para_order,
            "is_heading": style.startswith("Heading"),
            "style": style,
        })
        para_order += 1

    for table in doc.tables:
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])
        tables.append(rows)

    return {
        "title": title,
        "headings": headings,
        "paragraphs": paragraphs,
        "tables": tables,
        "stats": {
            "paragraphs": len(paragraphs),
            "headings": len(headings),
            "tables": len(tables),
            "chars": sum(len(p["text"]) for p in paragraphs),
        },
    }


def read_pdf(path: str) -> dict:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        return {
            "title": Path(path).stem,
            "headings": [],
            "paragraphs": [],
            "tables": [],
            "error": "PyPDF2 not installed. Run: pip install PyPDF2",
        }

    reader = PdfReader(path)
    title = Path(path).stem
    paragraphs = []
    para_order = 0

    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            # Heuristic: short lines at page top are likely headings
            is_likely_heading = len(line) < 60 and line[0].isdigit()
            paragraphs.append({
                "text": line,
                "order": para_order,
                "is_heading": is_likely_heading,
                "page": i + 1,
            })
            para_order += 1

    # Extract headings (short lines or numbered lines)
    headings = []
    for p in paragraphs:
        text = p["text"]
        if len(text) < 60 and (
            text[0].isdigit()
            or text.startswith("第")
            or text.endswith("章")
            or text.endswith("节")
            or len(text) < 30
        ):
            headings.append({
                "level": 1,
                "text": text,
                "order": p["order"],
                "page": p.get("page", 1),
            })

    return {
        "title": title,
        "headings": headings,
        "paragraphs": paragraphs,
        "tables": [],
        "stats": {
            "paragraphs": len(paragraphs),
            "headings": len(headings),
            "tables": 0,
            "chars": sum(len(p["text"]) for p in paragraphs),
        },
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: python read_doc.py <file_path> [--pretty]", file=sys.stderr)
        sys.exit(1)

    file_path = sys.argv[1]
    pretty = "--pretty" in sys.argv

    if not os.path.exists(file_path):
        print(f"Error: file not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    ext = Path(file_path).suffix.lower()

    try:
        if ext == ".docx":
            result = read_docx(file_path)
        elif ext == ".pdf":
            result = read_pdf(file_path)
        elif ext == ".doc":
            # .doc (old format) — try python-docx first, fall back to message
            print(
                "Warning: .doc is a legacy format. For best results, convert to .docx first.",
                file=sys.stderr,
            )
            print(
                "You can use LibreOffice headless: libreoffice --headless --convert-to docx file.doc",
                file=sys.stderr,
            )
            sys.exit(1)
        else:
            print(f"Error: unsupported format '{ext}'. Supported: .docx, .pdf", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"Error reading file: {e}", file=sys.stderr)
        sys.exit(1)

    indent = 2 if pretty else None
    print(json.dumps(result, ensure_ascii=False, indent=indent))


if __name__ == "__main__":
    main()
