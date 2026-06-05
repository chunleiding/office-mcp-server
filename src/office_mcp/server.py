#!/usr/bin/env python3
"""Office MCP Server - Main entry point.

Supports Word/Excel/PPT document automation via MCP tools.
"""

# Add src/ to sys.path BEFORE any other imports so that
# "import office_mcp.*" and "from office_mcp.* import ..." work
# when the file is executed as `python -m office_mcp.server`.
import sys
from pathlib import Path

_SRC_DIR = str(Path(__file__).resolve().parent.parent)
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

import argparse
from fastmcp import FastMCP

# Build the combined server
mcp = FastMCP("office-mcp-server")


# =============================================================================
# Word Tools
# =============================================================================

@mcp.tool()
def word_clone_template(
    template_path: str,
    output_path: str,
    title: str = None,
    date: str = None,
    main_topic: str = None,
    process_record: str = None,
) -> dict:
    """
    Clone a Word template document and replace specified fields.

    Perfect for reusing document templates (e.g. meeting minutes,
    lesson plans) with new content while preserving all formatting.

    Args:
        template_path: Path to the .docx template file.
        output_path: Where to save the new .docx file.
        title: (Optional) New title text (replaces centered bold title).
        date: (Optional) New date string (replaces date in table).
        main_topic: (Optional) New main topic (replaces topic in table).
        process_record: (Optional) Full process record text (replaces merged cell content).

    Returns:
        Dict with success status and output path.
    """
    try:
        sys.path.insert(0, "src")
        from office_mcp.word.template import clone_word_template

        replacements = {}
        if title:
            replacements["title"] = title
        if date:
            replacements["date"] = date
        if main_topic:
            replacements["main_topic"] = main_topic
        if process_record:
            replacements["process_record"] = process_record

        result = clone_word_template(
            template_path=template_path,
            output_path=output_path,
            replacements=replacements,
        )
        return {"success": True, "output_path": result}
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def word_analyze_template(
    template_path: str,
) -> dict:
    """
    Analyze a Word template and return its structure.

    Use this to understand what fields can be replaced in a template
    before calling word_clone_template.

    Args:
        template_path: Path to the .docx template file.

    Returns:
        Dict describing the document structure and replaceable fields.
    """
    try:
        sys.path.insert(0, "src")
        from office_mcp.word.template import analyze_template

        result = analyze_template(template_path)
        return {"success": True, "structure": result}
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def word_create_document(
    output_path: str,
    title: str = None,
    content: str = None,
) -> dict:
    """
    Create a new Word document from scratch.

    Args:
        output_path: Where to save the .docx file.
        title: (Optional) Document title.
        content: (Optional) Document content (use \\n to separate paragraphs).

    Returns:
        Dict with success status and output path.
    """
    try:
        from docx import Document
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()

        if title:
            heading = doc.add_heading(title, level=0)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if content:
            for line in content.split("\\n"):
                if line.strip():
                    p = doc.add_paragraph(line)
                    for run in p.runs:
                        run.font.name = "宋体"
                        run.font.size = Pt(12)

        import os
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
        return {"success": True, "output_path": output_path}
    except Exception as e:
        return {"success": False, "error": str(e)}


@mcp.tool()
def word_read_document(
    document_path: str,
) -> dict:
    """
    Read and extract text content from a Word document.

    Args:
        document_path: Path to the .docx file.

    Returns:
        Dict with document text content and structure info.
    """
    try:
        from docx import Document

        doc = Document(document_path)

        paragraphs = []
        for i, para in enumerate(doc.paragraphs):
            if para.text.strip():
                paragraphs.append({
                    "index": i,
                    "style": para.style.name if para.style else "Normal",
                    "text": para.text,
                })

        tables = []
        for ti, table in enumerate(doc.tables):
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            tables.append({
                "table_index": ti,
                "rows": len(table.rows),
                "cols": len(table.columns),
                "data": table_data,
            })

        return {
            "success": True,
            "paragraph_count": len(paragraphs),
            "table_count": len(tables),
            "paragraphs": paragraphs,
            "tables": tables,
        }
    except Exception as e:
        return {"success": False, "error": str(e)}


# =============================================================================
# Main
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Office MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="Transport protocol (default: stdio)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for SSE transport (default: 8000)",
    )
    args = parser.parse_args()

    print(
        f"🚀 Starting Office MCP Server (transport={args.transport})",
        file=sys.stderr,
    )

    if args.transport == "sse":
        mcp.run(transport="sse", host="127.0.0.1", port=args.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
