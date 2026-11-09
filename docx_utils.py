# -*- coding: utf-8 -*-
"""Utilidades de formato Word (estilo institucional DITEL — réplica del formato editado de Melgar)."""
import os
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_DARK_GRAY = RGBColor(0x33, 0x33, 0x33)

HEX_HEADER_BG = "E2E8F0"
HEX_ALT_ROW = "F8FAFC"
HEX_WHITE = "FFFFFF"
HEX_RED_BG = "FEE2E2"
HEX_YELLOW_BG = "FEF3C7"
HEX_GREEN_BG = "DCFCE7"
HEX_BLUE_BG = "DBEAFE"
HEX_LIME_BG = "ECFCCB"
SEM_BG = {"VERDE": HEX_GREEN_BG, "AMARILLO": HEX_YELLOW_BG, "ROJO": HEX_RED_BG}
CAT_BG = {"CUMPLE": HEX_GREEN_BG, "CERCA": HEX_LIME_BG, "MODERADO": HEX_YELLOW_BG, "BRECHA ALTA": HEX_RED_BG}

BANNER_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Reporte", "banner_cabecera.png")


def n0(x):
    """Entero con separador de miles (.)"""
    return f"{int(round(x)):,}".replace(",", ".")


def n1(x):
    return f"{x:.1f}"


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))


def set_cell_margins(cell, top=100, bottom=100, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'))


def set_table_borders(table, color="94A3B8", sz="4"):
    table._tbl.tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:left w:val="none"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>'))


def fijar_anchos(tbl, widths):
    """Diseño fijo + gridCol (necesario para que Word y LibreOffice respeten los anchos)."""
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    for j, w in enumerate(widths):
        tbl.columns[j].width = Inches(w)
        for cell in tbl.columns[j].cells:
            cell.width = Inches(w)


def sin_espacio(p):
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0


def repetir_encabezado(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))


def no_partir_fila(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))


def _run(p, text, size=10, bold=False, italic=False, color=COLOR_BLACK):
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color
    return r


def add_callout_box(doc, title, text, bg_hex="F8FAFC", border_color="475569"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    fijar_anchos(tbl, [7.0])
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=180)
    cell._tc.get_or_add_tcPr().append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="20" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'))
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    _run(p, f"{title}\n", 10, bold=True)
    _run(p, text, 9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    _run(p, text, 13, bold=True)
    return p


def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    _run(p, text, 11, bold=True)
    return p


def add_body_p(doc, text, bold_prefix="", space_after=5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        _run(p, bold_prefix, 10, bold=True)
    _run(p, text, 10)
    return p


def add_figura(doc, path, caption, width=6.2):
    if not path or not os.path.exists(path):
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Inches(width))
    pc = doc.add_paragraph()
    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pc.paragraph_format.space_after = Pt(8)
    _run(pc, caption, 8, italic=True, color=COLOR_DARK_GRAY)


def add_tabla(doc, headers, rows, widths, aligns=None, font=8, fills=None, bold_cols=(), header_font=None,
              pad=60):
    """rows: lista de listas de str. fills: lista (por fila) de dict {col: hex} o None."""
    tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_borders(tbl)
    aligns = aligns or ["L"] * len(headers)
    hf = header_font or font
    for j, h in enumerate(headers):
        c = tbl.cell(0, j)
        c.width = Inches(widths[j])
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, top=80, bottom=80, left=70, right=70)
        p = c.paragraphs[0]
        sin_espacio(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, h, hf, bold=True)
    repetir_encabezado(tbl.rows[0])
    fijar_anchos(tbl, widths)
    for i, row in enumerate(rows):
        no_partir_fila(tbl.rows[i + 1])
        for j, val in enumerate(row):
            c = tbl.cell(i + 1, j)
            c.width = Inches(widths[j])
            bg = (fills[i] or {}).get(j) if fills else None
            set_cell_background(c, bg or (HEX_ALT_ROW if i % 2 else HEX_WHITE))
            set_cell_margins(c, top=pad, bottom=pad, left=70, right=70)
            p = c.paragraphs[0]
            sin_espacio(p)
            p.alignment = {"C": WD_ALIGN_PARAGRAPH.CENTER, "R": WD_ALIGN_PARAGRAPH.RIGHT}.get(aligns[j],
                                                                                          WD_ALIGN_PARAGRAPH.LEFT)
            _run(p, str(val), font, bold=(j in bold_cols))
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    return tbl

