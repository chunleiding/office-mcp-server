"""Table format parser — optimized for minimal MCP response size."""
from docx.table import Table

_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_Wp = f'{{{_W}}}'

_ALIGN_MAP = {'left': 'L', 'center': 'C', 'right': 'R', 'both': 'J', 'justify': 'J', 'distribute': 'D'}
_VALIGN_MAP = {'top': 'T', 'center': 'C', 'bottom': 'B'}
_PARA_ALIGN = {0: 'L', 1: 'C', 2: 'R', 3: 'J'}


def parse_table_format(table: Table) -> list:
    """Parse cell formats. Returns compact list of rows."""
    if not table:
        return []
    result = []
    for ri, row in enumerate(table.rows):
        seen = set()
        cells = []
        for cell in row.cells:
            cid = id(cell)
            if cid in seen:
                continue
            seen.add(cid)
            cells.append(_cell_fmt(cell))
        result.append({"r": ri, "c": cells})
    return result


def _cell_fmt(cell) -> dict:
    tc = cell._tc
    d = {"t": cell.text.strip()[:60]}

    # Merge info
    gs = tc.find(f'.//{_Wp}gridSpan')
    if gs is not None:
        d["cs"] = int(gs.get(f'{_Wp}val', '1'))

    vm = tc.find(f'.//{_Wp}vMerge')
    if vm is not None:
        val = vm.get(f'{_Wp}val', 'continue')
        d["ms"] = val == 'restart'  # merge-start

    # Background
    shd = tc.find(f'.//{_Wp}shd')
    if shd is not None:
        fill = shd.get(f'{_Wp}fill')
        if fill and fill != 'auto':
            d["bg"] = fill

    # Alignment
    jc = tc.find(f'.//{_Wp}jc')
    if jc is not None:
        d["a"] = _ALIGN_MAP.get(jc.get(f'{_Wp}val', 'left'), 'L')

    va = tc.find(f'.//{_Wp}vAlign')
    if va is not None:
        d["va"] = _VALIGN_MAP.get(va.get(f'{_Wp}val', 'top'), 'T')

    # Paragraphs (only non-empty)
    paras = []
    for p in cell.paragraphs:
        if not p.text.strip():
            continue
        pi = {"t": p.text.strip()[:40]}
        if p.alignment is not None:
            pi["a"] = _PARA_ALIGN.get(p.alignment, 'L')
        # Only include runs with non-default formatting
        for r in p.runs:
            if r.bold or r.italic or r.font.size:
                ri = {"t": r.text[:20]}
                if r.bold: ri["b"] = True
                if r.italic: ri["i"] = True
                if r.font.size: ri["sz"] = r.font.size
                if r.font.name and r.font.name != '宋体': ri["fn"] = r.font.name
                pi.setdefault("runs", []).append(ri)
        paras.append(pi)
    if paras:
        d["p"] = paras

    return d


def analyze_table_merge(table: Table) -> dict:
    """Compact merge map — only stores non-default values."""
    rows_n, cols_n = len(table.rows), len(table.columns)
    merges = []
    for ri, row in enumerate(table.rows):
        ci = 0
        for cell in row.cells:
            while ci < cols_n:
                occupied = False
                for m in merges:
                    if m["r"] <= ri < m["r"] + m["rs"] and m["c"] <= ci < m["c"] + m["cs"]:
                        occupied = True
                        break
                if not occupied:
                    break
                ci += 1
            if ci >= cols_n:
                break
            tc = cell._tc
            cs = 1
            gs = tc.find(f'.//{_Wp}gridSpan')
            if gs is not None:
                try: cs = int(gs.get(f'{_Wp}val', '1'))
                except: pass
            rs = 1
            vm = tc.find(f'.//{_Wp}vMerge')
            if vm is not None and vm.get(f'{_Wp}val', '') == 'restart':
                for ro in range(1, rows_n - ri):
                    nr = table.rows[ri + ro]
                    nvm = nr.cells[0]._tc.find(f'.//{_Wp}vMerge')
                    if nvm is not None and nvm.get(f'{_Wp}val', '') == 'continue':
                        rs += 1
                    else:
                        break
            if cs > 1 or rs > 1:
                merges.append({"r": ri, "c": ci, "rs": rs, "cs": cs})
            ci += cs
    return {"rows": rows_n, "cols": cols_n, "merges": merges}


def format_table_analysis(table: Table) -> dict:
    """Combined analysis — no XML preview (saves tokens)."""
    return {
        "cells": parse_table_format(table),
        "merge": analyze_table_merge(table),
    }
