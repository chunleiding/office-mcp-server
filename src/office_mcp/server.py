"""Office MCP Server — universal Word document engine.

2 core tools:
  word_analyze  → analyze any Word doc, return text index
  word_replace  → universal text replacement with addressing scheme

2 legacy tools (backward compat, internally call core tools):
  word_clone_template  → table-template clone (title/date/topic/record)
  word_read_document   → read document content
"""
from fastmcp import FastMCP

mcp = FastMCP("office-mcp-server")


def _norm(text: str) -> str:
    """Normalize chars that break JSON-RPC (Chinese quotes → ASCII)."""
    return text.translate(str.maketrans('\u201c\u201d\u2018\u2019', "\"\"\'\'"))


# ─── Core Tool 1: Analyze ───────────────────────────────────────────────────

@mcp.tool()
def word_analyze(path: str) -> dict:
    """Analyze a Word document structure. Returns text index for word_replace.

    Addressing scheme for replacements:
      p:{i}              → body paragraph
      c:{t}:{r}:{c}      → entire cell (auto line-map, use \\n for multi-line)
      c:{t}:{r}:{c}:{p}  → specific paragraph inside a cell

    Optional format hints (in fmt_hints param):
      "_fmt:p:{i}" → {"runs": [{"text": "中一", "underline": true}]}
    """
    try:
        from office_mcp.word.engine import analyze
        result = analyze(path)
        return {"ok": True, **result}
    except Exception as e:
        return {"ok": False, "err": str(e)}


# ─── Core Tool 2: Replace ───────────────────────────────────────────────────

@mcp.tool()
def word_replace(
    path: str,
    output_path: str,
    replacements: str,
    fmt_hints: str = None,
) -> dict:
    """Replace text in a Word document, preserving all formatting.

    Args:
        path: source .docx path
        output_path: where to save the new .docx
        replacements: JSON string with replacement map, e.g.:
            {"p:0": "new title", "c:0:3:0": "line1\\nline2", "c:0:2:1:0": "topic"}
        fmt_hints: optional JSON string with run-format hints, e.g.:
            {"_fmt:p:1": {"runs": [{"text": "中一", "underline": true}]}}
    """
    try:
        import json
        from office_mcp.word.engine import replace_text

        rmap = json.loads(_norm(replacements))
        fh = json.loads(_norm(fmt_hints)) if fmt_hints else None

        # Normalize all replacement values
        for k, v in rmap.items():
            if isinstance(v, str):
                rmap[k] = _norm(v)

        result = replace_text(path, output_path, rmap, fh)
        return {"ok": True, "path": result}
    except Exception as e:
        return {"ok": False, "err": str(e)}


# ─── Legacy Tool: Clone Template ────────────────────────────────────────────

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
    """Clone a table-template Word doc and replace fields. Preserves all formatting.

    (Legacy tool — for new templates, prefer word_analyze + word_replace)
    """
    try:
        from office_mcp.word.engine import clone_word_template

        rp = {}
        if title: rp["title"] = _norm(title)
        if date: rp["date"] = _norm(date)
        if main_topic: rp["main_topic"] = _norm(main_topic)
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


# ─── Legacy Tool: Read Document ─────────────────────────────────────────────

@mcp.tool()
def word_read_document(document_path: str) -> dict:
    """Read a Word document: paragraphs + tables."""
    try:
        from docx import Document
        doc = Document(document_path)
        paras = [{"i": i, "s": p.style.name, "t": p.text}
                 for i, p in enumerate(doc.paragraphs) if p.text.strip()]
        tables = []
        for ti, tbl in enumerate(doc.tables):
            data = [[c.text.strip() for c in row.cells] for row in tbl.rows]
            tables.append({"i": ti, "rows": len(tbl.rows), "cols": len(tbl.columns),
                           "data": data})
        return {"ok": True, "pn": len(paras), "tn": len(tables),
                "paras": paras, "tables": tables}
    except Exception as e:
        return {"ok": False, "err": str(e)}


if __name__ == "__main__":
    mcp.run(transport="stdio")
