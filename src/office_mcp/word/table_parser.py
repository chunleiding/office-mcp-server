"""
Word table format parser - based on docx XML.

Parses table cells to extract format information:
- Cell merge (gridSpan, vMerge)
- Border styles (tcBorders)
- Background color (shd)
- Alignment (jc, vAlign)
- Paragraph formats inside cells
"""

from docx.table import Table
from lxml import etree


# XML namespaces for docx
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'


def parse_table_format(table: Table) -> list:
    """
    Parse a docx table and return format information for each cell.
    
    Args:
        table: docx Table object
        
    Returns:
        List of table format info, each table has rows -> cells structure.
        Format:
        [
            {
                "row_index": 0,
                "cells": [
                    {
                        "text": "cell text",
                        "colspan": 1,  # gridSpan
                        "rowspan": 1,  # vMerge
                        "border": {...},
                        "bg_color": None,
                        "align": "left" | "center" | "right" | "justify",
                        "valign": "top" | "center" | "bottom",
                        "paragraphs": [...],  # paragraph formats
                        "is_merged_start": True,  # first cell of merged region
                        "is_merged_continue": False,  # continued cell of merged region
                    }
                ]
            }
        ]
    """
    if not table:
        return []
    
    result = []
    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    for row_idx, row in enumerate(table.rows):
        row_data = {"row_index": row_idx, "cells": []}
        
        # Get unique cells (docx repeats cells for merged regions)
        seen_cells = set()
        
        for cell in row.cells:
            # Skip duplicate cells in same row
            cell_id = id(cell)
            if cell_id in seen_cells:
                continue
            seen_cells.add(cell_id)
            
            cell_info = _parse_cell_format(cell, ns_w)
            row_data["cells"].append(cell_info)
        
        result.append(row_data)
    
    return result


def _parse_cell_format(cell, ns_w: str) -> dict:
    """
    Parse format information from a single table cell.
    
    Args:
        cell: docx Table cell object
        ns_w: XML namespace string
        
    Returns:
        dict with format information
    """
    # Get the XML element for this cell
    tc = cell._tc
    
    # Basic info
    cell_info = {
        "text": cell.text.strip(),
        "colspan": 1,
        "rowspan": 1,
        "border": {},
        "bg_color": None,
        "align": None,
        "valign": None,
        "paragraphs": [],
        "is_merged_start": False,
        "is_merged_continue": False,
    }
    
    # --- 1. Parse cell merge (gridSpan, vMerge) ---
    grid_span = tc.find(f'.//{ns_w}gridSpan')
    if grid_span is not None:
        try:
            cell_info["colspan"] = int(grid_span.get(f'{ns_w}val', '1'))
        except (ValueError, TypeError):
            cell_info["colspan"] = 1
    
    v_merge = tc.find(f'.//{ns_w}vMerge')
    if v_merge is not None:
        merge_val = v_merge.get(f'{ns_w}val', 'continue')
        if merge_val == 'restart':
            cell_info["is_merged_start"] = True
            cell_info["rowspan"] = 1  # Will need table-level analysis for true rowspan
        elif merge_val == 'continue':
            cell_info["is_merged_continue"] = True
    
    # --- 2. Parse border styles (tcBorders) ---
    tc_borders = tc.find(f'.//{ns_w}tcBorders')
    if tc_borders is not None:
        border_info = {}
        for side in ['top', 'right', 'bottom', 'left']:
            border_elem = tc_borders.find(f'{ns_w}{side}')
            if border_elem is not None:
                border_info[side] = {
                    "val": border_elem.get(f'{ns_w}val', ''),
                    "sz": border_elem.get(f'{ns_w}sz', ''),
                    "space": border_elem.get(f'{ns_w}space', ''),
                    "color": border_elem.get(f'{ns_w}color', ''),
                }
        cell_info["border"] = border_info
    
    # --- 3. Parse background color (shd) ---
    shd = tc.find(f'.//{ns_w}shd')
    if shd is not None:
        cell_info["bg_color"] = shd.get(f'{ns_w}fill', None)
    
    # --- 4. Parse alignment (jc, vAlign) ---
    # Horizontal alignment (jc) - can be in tc, tcPr, or pPr
    jc = tc.find(f'.//{ns_w}jc')
    if jc is not None:
        # Map XML alignment values to our standard values
        align_map_xml = {
            'left': 'left',
            'center': 'center',
            'right': 'right',
            'both': 'justify',
            'justify': 'justify',
            'distribute': 'distribute',
        }
        xml_val = jc.get(f'{ns_w}val', 'left')
        cell_info["align"] = align_map_xml.get(xml_val, 'left')
    
    # Vertical alignment (vAlign)
    v_align = tc.find(f'.//{ns_w}vAlign')
    if v_align is not None:
        cell_info["valign"] = v_align.get(f'{ns_w}val', 'top')
    
    # --- 5. Parse paragraph formats inside cell ---
    paragraphs = []
    for para in cell.paragraphs:
        para_info = {
            "text": para.text.strip(),
            "alignment": None,
            "runs": []
        }
        
        # Paragraph alignment
        if para.alignment is not None:
            align_map = {
                0: 'left',
                1: 'center',
                2: 'right',
                3: 'justify',
            }
            para_info["alignment"] = align_map.get(para.alignment, 'left')
        
        # Run formats
        for run in para.runs:
            run_info = {
                "text": run.text,
                "bold": run.bold,
                "italic": run.italic,
                "underline": run.underline,
                "font_name": run.font.name,
                "font_size": run.font.size,
                "font_color": None,
            }
            
            # Font color
            if run.font.color and run.font.color.rgb:
                run_info["font_color"] = str(run.font.color.rgb)
            
            para_info["runs"].append(run_info)
        
        paragraphs.append(para_info)
    
    cell_info["paragraphs"] = paragraphs
    
    return cell_info


def analyze_table_merge(table: Table) -> dict:
    """
    Analyze table merge information at table level.
    
    This is more accurate than cell-level parsing because it can
    calculate true rowspan by scanning the table.
    
    Args:
        table: docx Table object
        
    Returns:
        dict with merge analysis result
    """
    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    # Build a 2D array to track merge information
    rows = len(table.rows)
    cols = len(table.columns)
    
    merge_map = [[None for _ in range(cols)] for _ in range(rows)]
    
    for row_idx, row in enumerate(table.rows):
        col_idx = 0
        for cell in row.cells:
            # Find the actual column index (skip merged cells)
            while col_idx < cols and merge_map[row_idx][col_idx] is not None:
                col_idx += 1
            
            if col_idx >= cols:
                break
            
            tc = cell._tc
            
            # Check colspan (gridSpan)
            grid_span = tc.find(f'.//{ns_w}gridSpan')
            colspan = 1
            if grid_span is not None:
                try:
                    colspan = int(grid_span.get(f'{ns_w}val', '1'))
                except (ValueError, TypeError):
                    colspan = 1
            
            # Check rowspan (vMerge)
            v_merge = tc.find(f'.//{ns_w}vMerge')
            rowspan = 1
            is_merge_start = False
            is_merge_continue = False
            
            if v_merge is not None:
                merge_val = v_merge.get(f'{ns_w}val', 'continue')
                if merge_val == 'restart':
                    is_merge_start = True
                    # Scan down to find true rowspan
                    for r_offset in range(1, rows - row_idx):
                        next_row = table.rows[row_idx + r_offset]
                        if r_offset < len(next_row.cells):
                            next_tc = next_row.cells[0]._tc
                            next_v_merge = next_tc.find(f'.//{ns_w}vMerge')
                            if next_v_merge is not None and next_v_merge.get(f'{ns_w}val', '') == 'continue':
                                rowspan += 1
                            else:
                                break
                elif merge_val == 'continue':
                    is_merge_continue = True
            
            # Fill merge_map
            for r in range(rowspan):
                for c in range(colspan):
                    if row_idx + r < rows and col_idx + c < cols:
                        merge_map[row_idx + r][col_idx + c] = {
                            "row": row_idx,
                            "col": col_idx,
                            "rowspan": rowspan,
                            "colspan": colspan,
                            "is_merge_start": is_merge_start and r == 0 and c == 0,
                            "is_merge_continue": is_merge_continue,
                        }
            
            col_idx += colspan
    
    return {
        "rows": rows,
        "cols": cols,
        "merge_map": merge_map,
    }


def format_table_analysis(table: Table) -> dict:
    """
    Complete table format analysis.
    
    Combines cell-level format parsing and table-level merge analysis.
    
    Args:
        table: docx Table object
        
    Returns:
        Complete table format information
    """
    # Cell-level format parsing
    cell_formats = parse_table_format(table)
    
    # Table-level merge analysis
    merge_analysis = analyze_table_merge(table)
    
    return {
        "cell_formats": cell_formats,
        "merge_analysis": merge_analysis,
        "table_xml_preview": etree.tostring(table._tbl, pretty_print=True).decode('utf-8')[:500],
    }
