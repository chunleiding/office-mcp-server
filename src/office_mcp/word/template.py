"""Word document template cloning engine."""

import copy
from pathlib import Path
from docx import Document
from docx.shared import Pt
from lxml import etree


def clone_word_template(
    template_path: str,
    output_path: str,
    replacements: dict,
    date_str: str = None,
) -> str:
    """
    Clone a Word template document and replace specified content.
    
    Args:
        template_path: Path to the template .docx file.
        output_path: Path to save the new .docx file.
        replacements: Dict mapping field names to new content.
                     Supported keys:
                     - "title": Document title (paragraph 1, bold centered)
                     - "main_topic": Main content/topic (table cell)
                     - "process_record": Full process record text (table cell)
                     - "date": Date string (overrides date_str)
        date_str: Date string to replace in the date cell (e.g. "2026年6月5日 星期五 下午").
                   Deprecated: use replacements["date"] instead.
    
    Returns:
        Path to the saved document.
    """
    doc = Document(template_path)
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # --- Replace title (paragraph 1, centered, bold) ---
    if "title" in replacements:
        _replace_title(doc, replacements["title"])

    # --- Replace date ---
    effective_date = replacements.get("date", date_str)
    if effective_date:
        _replace_date(doc, effective_date, ns_w)

    # --- Replace main topic in table ---
    if "main_topic" in replacements and doc.tables:
        _replace_main_topic(doc, replacements["main_topic"], ns_w)

    # --- Replace process record (full merged cell content) ---
    if "process_record" in replacements and doc.tables:
        _replace_process_record(doc, replacements["process_record"])

    # --- Save ---
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out)


def _replace_title(doc: Document, title: str):
    """Replace the centered bold title paragraph."""
    # Find the title paragraph (typically paragraph 1, centered)
    target_para = None
    for i, para in enumerate(doc.paragraphs):
        if i == 1 and para.alignment and para.alignment == 1:  # CENTER
            target_para = para
            break

    if target_para is None and len(doc.paragraphs) > 1:
        target_para = doc.paragraphs[1]

    if target_para is None:
        return

    # Clear ALL runs, then set the first run to the new title
    for run in target_para.runs:
        run.text = ''
    if target_para.runs:
        target_para.runs[0].text = title
        target_para.runs[0].bold = True
    else:
        run = target_para.add_run(title)
        run.bold = True


def _replace_date(doc: Document, date_str: str, ns_w: str):
    """Replace date in the first table's date cell."""
    if not doc.tables:
        return
    table = doc.tables[0]
    # Date is typically in row 0, col 1 (merged)
    if len(table.rows) > 0 and len(table.rows[0].cells) > 1:
        cell = table.rows[0].cells[1]
        para = cell.paragraphs[0]
        # 修改第一个 run 的文本，保留格式
        if para.runs:
            para.runs[0].text = date_str
        else:
            run = para.add_run(date_str)
            run.font.name = '宋体'
            run.font.size = Pt(12)


def _replace_main_topic(doc: Document, topic: str, ns_w: str):
    """Replace main topic in table row 2, col 1 (merged)."""
    if not doc.tables:
        return
    table = doc.tables[0]
    if len(table.rows) > 2 and len(table.rows[2].cells) > 1:
        cell = table.rows[2].cells[1]
        # 修改第一个 run 的文本，保留格式
        if cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = topic
        else:
            _clear_cell_text(cell, ns_w)
            para = cell.paragraphs[0]
            run = para.add_run(topic)
            run.font.name = '宋体'
            run.font.size = Pt(12)


def _replace_process_record(doc: Document, content: str):
    """Replace the full process record in the merged cell (row 3, col 0)."""
    if not doc.tables:
        return
    table = doc.tables[0]
    if len(table.rows) <= 3:
        return
    
    cell = table.rows[3].cells[0]
    ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    # 保存模板段落的 pPr (段落格式 XML)
    template_pPr_list = []
    for para in cell.paragraphs:
        p_elem = para._element
        pPr = p_elem.find(f'{{{ns_w}}}pPr')
        if pPr is not None:
            template_pPr_list.append(copy.deepcopy(pPr))
        else:
            template_pPr_list.append(None)
    
    # 清除所有段落的 XML 元素
    tc = cell._tc
    for para in list(cell.paragraphs):
        p_elem = para._element
        parent = p_elem.getparent()
        if parent is not None:
            parent.remove(p_elem)
    
    # 添加新段落，应用原格式
    lines = content.strip().split('\n')
    for i, line in enumerate(lines):
        new_para = cell.add_paragraph()
        
        # 应用模板段落的 pPr (段落格式)
        if i < len(template_pPr_list):
            pPr = template_pPr_list[i]
        elif template_pPr_list:
            pPr = template_pPr_list[-1]
        else:
            pPr = None
        
        if pPr is not None:
            # 将保存的 pPr 应用到新段落
            new_p_elem = new_para._element
            # 移除现有的 pPr
            existing_pPr = new_p_elem.find(f'{{{ns_w}}}pPr')
            if existing_pPr is not None:
                new_p_elem.remove(existing_pPr)
            # 插入保存的 pPr (必须是第一个子元素)
            new_p_elem.insert(0, copy.deepcopy(pPr))
        
        # 添加 run 并设置字体
        run = new_para.add_run(line)
        run.font.name = '宋体'
        run.font.size = Pt(12)


def _clear_cell_text(cell, ns_w: str):
    """Clear all text content in a table cell via XML."""
    for p in cell._tc.findall(f'.//{{{ns_w}}}p'):
        for r in p.findall(f'.//{{{ns_w}}}r'):
            for t in r.findall(f'.//{{{ns_w}}}t'):
                t.text = ''


def analyze_template(template_path: str) -> dict:
    """
    Analyze a Word template and return its structure as a dict.
    Useful for understanding what fields can be replaced.
    
    Returns:
        {
            "paragraphs": [{"index": 0, "style": "...", "text": "..."}],
            "tables": [{"rows": N, "cols": M, "cells": [...]}],
            "replaceable_fields": ["title", "date", "main_topic", "process_record"]
        }
    """
    doc = Document(template_path)
    result = {
        "paragraphs": [],
        "tables": [],
        "replaceable_fields": []
    }

    for i, para in enumerate(doc.paragraphs):
        result["paragraphs"].append({
            "index": i,
            "style": para.style.name if para.style else "Normal",
            "alignment": str(para.alignment),
            "text": para.text[:100] if para.text else "",
        })

    for ti, table in enumerate(doc.tables):
        table_info = {
            "table_index": ti,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "sample_cells": []
        }
        for ri, row in enumerate(table.rows[:3]):  # first 3 rows
            for ci, cell in enumerate(row.cells[:3]):  # first 3 cols
                table_info["sample_cells"].append({
                    "row": ri, "col": ci,
                    "text": cell.text.strip()[:50]
                })
        result["tables"].append(table_info)

    # Infer replaceable fields based on content
    text = '\n'.join(p.text for p in doc.paragraphs)
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                text += '\n' + cell.text
    
    if '教研活动记录' in text or '活动' in text:
        result["replaceable_fields"] = [
            "title", "date", "location", "host", 
            "recorder", "main_topic", "process_record", "summary"
        ]

    return result
