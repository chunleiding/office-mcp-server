"""Office MCP Server — single entry point, all tools defined here."""
import sys
from pathlib import Path
from fastmcp import FastMCP

mcp = FastMCP("office-mcp-server")

# Lazy imports — only load heavy deps when a tool is actually called

def _norm(text: str) -> str:
    """Normalize chars that break JSON-RPC (Chinese quotes → ASCII)."""
    return text.translate(str.maketrans('\u201c\u201d\u2018\u2019', "\"\"\'\'"))


@mcp.tool()
def word_clone_template(
    template_path: str,
    output_path: str,
    title: str = None,
    date: str = None,
    main_topic: str = None,
    process_record: str = None,
    process_record_file: str = None,
) -> dict:
    """Clone a Word template and replace fields. Preserves all formatting."""
    try:
        from office_mcp.word.template import clone_word_template
        rp = {}
        if title: rp["title"] = title
        if date: rp["date"] = date
        if main_topic: rp["main_topic"] = main_topic
        pr = process_record
        if process_record_file:
            with open(process_record_file, "r", encoding="utf-8") as f:
                pr = f.read()
        if pr:
            rp["process_record"] = _norm(pr)
        r = clone_word_template(template_path, output_path, rp)
        return {"ok": True, "path": r}
    except Exception as e:
        return {"ok": False, "err": str(e)}


@mcp.tool()
def word_analyze_template(template_path: str) -> dict:
    """Analyze a Word template structure and table formats."""
    try:
        from docx import Document
        from office_mcp.word.template import analyze_template
        from office_mcp.word.table_parser import format_table_analysis
        struct = analyze_template(template_path)
        doc = Document(template_path)
        tf = [format_table_analysis(t) for t in doc.tables]
        return {"ok": True, "struct": struct, "tf": tf}
    except Exception as e:
        return {"ok": False, "err": str(e)}


@mcp.tool()
def word_create_document(output_path: str, title: str = None, content: str = None) -> dict:
    """Create a new Word document from scratch."""
    try:
        from docx import Document
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        doc = Document()
        if title:
            h = doc.add_heading(title, level=0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if content:
            for line in content.split("\n"):
                if line.strip():
                    p = doc.add_paragraph(line)
                    for r in p.runs:
                        r.font.name = "宋体"
                        r.font.size = Pt(12)
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(output_path)
        return {"ok": True, "path": output_path}
    except Exception as e:
        return {"ok": False, "err": str(e)}


@mcp.tool()
def word_read_document(document_path: str) -> dict:
    """Read a Word document: paragraphs + tables with format info."""
    try:
        from docx import Document
        from office_mcp.word.table_parser import format_table_analysis
        doc = Document(document_path)
        paras = [{"i": i, "s": p.style.name, "t": p.text}
                 for i, p in enumerate(doc.paragraphs) if p.text.strip()]
        tables = []
        for ti, tbl in enumerate(doc.tables):
            data = [[c.text.strip() for c in row.cells] for row in tbl.rows]
            tables.append({"i": ti, "rows": len(tbl.rows), "cols": len(tbl.columns),
                           "data": data, "fmt": format_table_analysis(tbl)})
        return {"ok": True, "pn": len(paras), "tn": len(tables), "paras": paras, "tables": tables}
    except Exception as e:
        return {"ok": False, "err": str(e)}


if __name__ == "__main__":
    mcp.run(transport="stdio")
