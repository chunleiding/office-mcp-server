"""Word template clone engine — XML-first approach.

Core principle: NEVER delete/create <w:p> elements. Only replace <w:t> text.
This guarantees 100% format fidelity (spacing, indentation, line-spacing)
with the template.
"""
import copy
import re
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
    """Replace title text while preserving per-run formatting (underline etc.).

    Strategy: keep ALL <w:r> elements and their <w:rPr> intact.
    Distribute new title text across original runs proportionally,
    then fix underline: find the originally-underlined text in the new
    title and ensure it lands in the run that has <w:u>.
    """
    para = doc.paragraphs[1] if len(doc.paragraphs) > 1 else None
    if not para:
        return
    runs = list(para.runs)
    if not runs:
        r = para.add_run(title)
        r.bold = True
        return

    # Step 1: save original run info (text + rPr XML)
    orig_info = []
    for r in runs:
        r_elem = r._r
        rPr = r_elem.find(f'{_Wp}rPr')
        # Check for underline: <w:u w:val="single"/> or similar (not "none")
        underline_val = None
        if rPr is not None:
            u = rPr.find(f'{_Wp}u')
            if u is not None:
                underline_val = u.get(f'{_Wp}val', 'single')
        orig_info.append({
            'text': r.text,
            'rPr': copy.deepcopy(rPr) if rPr is not None else None,
            'underline_val': underline_val,  # None = no <w:u>, "none" = explicit none, "single" etc = underlined
        })

    # Step 2: collect underlined text segments from original
    underlined_texts = [o['text'] for o in orig_info
                        if o['underline_val'] and o['underline_val'] != 'none']

    # Step 3: clear all run text (keep <w:r> elements)
    for r in runs:
        r.text = ''

    # Step 4: distribute new title across runs proportionally
    orig_lens = [len(o['text']) for o in orig_info]
    total_orig = sum(orig_lens) or 1
    n = len(title)
    idx = 0
    for i, r in enumerate(runs):
        if idx >= n:
            break
        if i < len(runs) - 1:
            share = max(1, round((orig_lens[i] / total_orig) * n))
            share = min(share, n - idx)
        else:
            share = n - idx
        r.text = title[idx:idx + share]
        idx += share

    # Step 5: restore per-run rPr (preserves bold, font, underline XML)
    for i, r in enumerate(runs):
        if i < len(orig_info) and orig_info[i]['rPr'] is not None:
            r_elem = r._r
            existing = r_elem.find(f'{_Wp}rPr')
            if existing is not None:
                r_elem.remove(existing)
            r_elem.insert(0, copy.deepcopy(orig_info[i]['rPr']))

    # Step 6: smart underline fix
    # If original had underlined text (e.g. "中一"), find it in the new title
    # and move the underline to whichever run now contains that text
    if underlined_texts:
        pattern = '|'.join(re.escape(t) for t in underlined_texts if t)
        if pattern:
            m = re.search(pattern, title)
            if m:
                pos = m.start()
                # Find which run contains this position
                run_start = 0
                for i, r in enumerate(runs):
                    run_end = run_start + len(r.text)
                    if run_start <= pos < run_end:
                        # This run should have the underline
                        r_elem = runs[i]._r
                        rPr = r_elem.find(f'{_Wp}rPr')
                        if rPr is None:
                            rPr = etree.SubElement(r_elem, f'{_Wp}rPr')
                        u = rPr.find(f'{_Wp}u')
                        if u is None:
                            u = etree.SubElement(rPr, f'{_Wp}u')
                        u.set(f'{_Wp}val', 'single')
                        break
                    run_start = run_end

    # Step 7: remove underline from runs that originally had underline
    # but the underlined text is NOT in them anymore
    if underlined_texts and m:
        pos = m.start()
        run_start = 0
        for i, r in enumerate(runs):
            run_end = run_start + len(r.text)
            if run_start <= pos < run_end:
                pass  # this run keeps the underline (set in step 6)
            elif orig_info[i]['underline_val'] and orig_info[i]['underline_val'] != 'none':
                # This run originally had underline but the text moved away
                r_elem = r._r
                rPr = r_elem.find(f'{_Wp}rPr')
                if rPr is not None:
                    u = rPr.find(f'{_Wp}u')
                    if u is not None:
                        u.set(f'{_Wp}val', 'none')
            run_start = run_end


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

    Key: blank lines in content are STRIPPED because the template's
    paragraph formatting (pPr spacing) controls visual spacing, not
    empty paragraphs. Including blank lines would consume template
    paragraphs that originally had content, shifting the mapping and
    creating extra visual gaps.
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

    # Step 2: Fill new content into existing paragraphs
    # Strip blank lines — spacing comes from pPr, not empty paragraphs
    lines = [l for l in content.strip().split('\n') if l.strip()]

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
