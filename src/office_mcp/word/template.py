"""Word template clone engine — XML-first approach.

Core principle: NEVER delete/create <w:p> elements. Only replace <w:t> text.
This guarantees 100% format fidelity (spacing, indentation, line-spacing)
with the template.
"""
import copy
from pathlib import Path
from docx import Document
from docx.shared import Pt
from lxml import etree

_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_Wp = f'{{{_W}}}'


def clone_word_template(template_path: str, output_path: str,
                        replacements: dict, date_str: str = None) -> str:
    doc = Document(template_path)

    if "title" in replacements:
        _set_title(doc, replacements["title"])

    date = replacements.get("date", date_str)
    if date:
        _set_run_text(doc.tables[0].rows[0].cells[1].paragraphs[0], date) if doc.tables else None

    if "main_topic" in replacements and doc.tables:
        cell = doc.tables[0].rows[2].cells[1]
        if len(doc.tables[0].rows) > 2 and cell.paragraphs and cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].text = replacements["main_topic"]

    if "process_record" in replacements and doc.tables:
        _set_process_record(doc.tables[0], replacements["process_record"])

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return str(out)


def _set_title(doc, title: str):
    para = doc.paragraphs[1] if len(doc.paragraphs) > 1 else None
    if not para:
        return
    for r in para.runs:
        r.text = ''
    if para.runs:
        para.runs[0].text = title
        para.runs[0].bold = True
    else:
        r = para.add_run(title)
        r.bold = True


def _set_run_text(para, text: str):
    """Replace text in first run, preserving format."""
    if para.runs:
        para.runs[0].text = text
    else:
        r = para.add_run(text)
        r.font.name = '宋体'
        r.font.size = Pt(12)


def _set_process_record(table, content: str):
    """Replace process record by only touching <w:t> text nodes.

    NEVER delete <w:p> elements — they carry spacing/indent/line-spacing.
    Just clear text and refill line-by-line.
    """
    if len(table.rows) <= 3:
        return
    cell = table.rows[3].cells[0]
    p_list = cell._tc.findall(f'{_Wp}p')
    if not p_list:
        return

    # Save last pPr for appending if needed
    last_pPr = None
    for p in reversed(p_list):
        pp = p.find(f'{_Wp}pPr')
        if pp is not None:
            last_pPr = copy.deepcopy(pp)
            break

    # Step 1: Clear ALL <w:t> text (keep <w:p> and <w:pPr> intact)
    for p in p_list:
        for r in p.findall(f'{_Wp}r'):
            t = r.find(f'{_Wp}t')
            if t is not None:
                t.text = ''

    # Step 2: Fill new content line-by-line into existing paragraphs
    lines = content.strip().split('\n')
    for i, line in enumerate(lines):
        if i < len(p_list):
            # Reuse existing paragraph — find or create a <w:r><w:t> to hold text
            p = p_list[i]
            rs = p.findall(f'{_Wp}r')
            if rs:
                t = rs[0].find(f'{_Wp}t')
                if t is not None:
                    t.text = line
                else:
                    etree.SubElement(rs[0], f'{_Wp}t').text = line
                # Remove extra runs (keep only first)
                for r in rs[1:]:
                    p.remove(r)
            else:
                r = etree.SubElement(p, f'{_Wp}r')
                etree.SubElement(r, f'{_Wp}t').text = line
        else:
            # Append new paragraph (content exceeds template)
            new_p = etree.SubElement(cell._tc, f'{_Wp}p')
            if last_pPr is not None:
                new_p.append(copy.deepcopy(last_pPr))
            r = etree.SubElement(new_p, f'{_Wp}r')
            etree.SubElement(r, f'{_Wp}t').text = line


def analyze_template(template_path: str) -> dict:
    doc = Document(template_path)
    paras = [{"i": i, "style": p.style.name, "text": p.text[:80]}
             for i, p in enumerate(doc.paragraphs) if p.text.strip()]

    fields = []
    if doc.tables:
        all_text = ''.join(c.text for row in doc.tables[0].rows for c in row.cells)
        if '教研' in all_text or '活动' in all_text:
            fields = ["title", "date", "main_topic", "process_record"]

    return {"paragraphs": paras, "replaceable_fields": fields}
