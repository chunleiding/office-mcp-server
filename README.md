# Office MCP Server

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Python 3.10+](https://img.shields.io/badge/python-3.10+-green.svg)

MCP (Model Context Protocol) Server for Office document automation — **create, read, and clone Word/Excel/PPT documents** with AI assistants.

> 🎯 **Core Feature: Template Cloning** — Reuse any `.docx`/`.xlsx`/`.pptx` template, replace content while **preserving all formatting**. Perfect for meeting minutes, lesson plans, reports, and recurring documents.

---

## ✨ Features

### 📝 Word Documents (Phase 1 — Current)
- ✅ **Clone template** — Replace fields in a Word template while keeping all formatting (tables, merged cells, fonts, alignment)
- ✅ **Read document** — Extract text content and structure
- ✅ **Create document** — Build a new Word doc from scratch
- ✅ **Analyze template** — Understand a template's structure before cloning

### 📊 Excel Documents (Phase 3 — Planned)
- 🔜 Create/workbooks with formatted data
- 🔜 Read cell values and formulas
- 🔜 Clone Excel templates

### 📽️ PowerPoint (Phase 4 — Planned)
- 🔜 Create presentations with custom layouts
- 🔜 Clone PPT templates
- 🔜 Add slides with formatted content

---

## 🚀 Quick Start

### Install

```bash
# Clone the repo
git clone https://github.com/chunleiding/office-mcp-server.git
cd office-mcp-server

# Install with uv (recommended)
uv sync

# Or with pip
pip install -e .
```

### Use with WorkBuddy / Claude Desktop

Add to your MCP config (`~/.workbuddy/mcp.json` or Claude Desktop config):

```json
{
  "mcpServers": {
    "office-mcp-server": {
      "command": "uv",
      "args": [
        "--directory", "/path/to/office-mcp-server",
        "run", "python", "-m", "office_mcp.server"
      ]
    }
  }
}
```

Or with `uvx` (no local install):

```json
{
  "mcpServers": {
    "office-mcp-server": {
      "command": "uvx",
      "args": ["--from", "office-mcp-server", "office-mcp-server"]
    }
  }
}
```

---

## 🛠️ Tools Reference

### `word_clone_template`

Clone a Word template and replace fields:

```json
{
  "template_path": "/path/to/template.docx",
  "output_path": "/path/to/output.docx",
  "title": "My New Title",
  "date": "2026年6月5日 星期五 下午",
  "main_topic": "Recent Issues and Countermeasures for Children's Drinking Water",
  "process_record": "一、本周回顾\n1. ..."
}
```

### `word_analyze_template`

Analyze a template to understand its replaceable fields:

```json
{
  "template_path": "/path/to/template.docx"
}
```

Returns structure info and suggested replaceable fields.

### `word_read_document`

Read text content from a Word document:

```json
{
  "document_path": "/path/to/document.docx"
}
```

### `word_create_document`

Create a new Word document from scratch:

```json
{
  "output_path": "/path/to/new.docx",
  "title": "Document Title",
  "content": "Line 1\nLine 2\nLine 3"
}
```

---

## 🏗️ Architecture

```
office-mcp-server/
├── src/office_mcp/
│   ├── server.py          # FastMCP entry point
│   ├── word/             # Word tools (✅ Done)
│   │   ├── tools.py      # MCP tool definitions
│   │   └── template.py  # Template cloning engine
│   ├── excel/            # Excel tools (🔜 Phase 3)
│   ├── ppt/             # PPT tools (🔜 Phase 4)
│   └── shared/           # Shared utilities
└── tests/
```

---

## 🗺️ Roadmap

| Phase | Feature | Status |
|-------|---------|--------|
| 1 | Word template cloning + basic tools | ✅ In Progress |
| 2 | Architecture refactor (modular) | 🔜 Next |
| 3 | Excel support | 🔜 Planned |
| 4 | PPT support | 🔜 Planned |
| 5 | Cross-format batch operations | 💡 Future |

---

## 📋 Requirements

- Python 3.10+
- [FastMCP](https://github.com/modelcontextprotocol/python-sdk)
- `python-docx` (Word)
- `openpyxl` (Excel — Phase 3)
- `python-pptx` (PPT — Phase 4)

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.

---

## 💡 Why This Instead of Forking?

The original [`Office-Word-MCP-Server`](https://github.com/GongRzhe/Office-Word-MCP-Server) is **archived (read-only) since March 2026**, with 53 open issues and no maintenance. 

This project is a **clean rewrite** with:
- ✅ Modular architecture from day one (Word/Excel/PPT separated)
- ✅ Template cloning as a first-class feature
- ✅ Python 3.13 compatibility (fixed lxml issues)
- ✅ Active maintenance
