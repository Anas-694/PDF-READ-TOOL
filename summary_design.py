from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------- Design tokens (match these to your brand) ----------------
NAVY = RGBColor(0x12, 0x23, 0x3F)
NAVY_LIGHT = RGBColor(0x1F, 0x3A, 0x5F)
GOLD = RGBColor(0xB9, 0x86, 0x2E)
GREEN = RGBColor(0x1C, 0x7A, 0x4D)
RED = RGBColor(0xB2, 0x3A, 0x2E)
GREY_TEXT = RGBColor(0x3A, 0x3F, 0x47)
MUTED = RGBColor(0x9A, 0xA1, 0xAC)

LIGHT_BG_HEX = "F2F4F7"
NAVY_HEX = "12233F"
WHITE_HEX = "FFFFFF"

FONT_NAME = "Calibri"


# ---------------- Low-level OOXML helpers ----------------
def set_cell_shading(cell, hex_color: str):
    """Fill a table cell with a solid background color."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def set_cell_border(cell, color="D8DCE3", size=4):
    """Add a thin border on all sides of a cell (for KPI cards)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:color"), color)
        borders.append(el)
    tc_pr.append(borders)


def add_page_number_field(paragraph):
    """Insert a live 'PAGE' field (updates automatically in Word)."""
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr)
    run._r.append(fld_char2)


def add_left_accent_border(paragraph, color=NAVY_HEX, size=24):
    """Gold/navy accent bar on the left of a heading paragraph (like a Notion callout)."""
    p_pr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), color)
    borders.append(left)
    p_pr.append(borders)


# ---------------- High-level building blocks ----------------
def add_section_heading(doc: Document, text: str):
    """Section heading with a colored left accent bar — visually anchors each section."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    add_left_accent_border(p, color="B9862E")  # gold bar
    run = p.add_run("  " + text)
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = NAVY
    run.font.name = FONT_NAME
    return p


def add_kpi_row(doc: Document, cards):
    """
    cards = [{"title": "Revenue", "value": "Rs. 49,716M", "change": "35.3%", "positive": True}, ...]
    Renders a 1-row table where each cell looks like a metric card.
    """
    n = len(cards)
    table = doc.add_table(rows=1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for i, card in enumerate(cards):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, LIGHT_BG_HEX)
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # clear default empty paragraph, then build 3 lines: label / value / change
        cell.paragraphs[0].text = ""
        p_label = cell.paragraphs[0]
        r = p_label.add_run(card["title"].upper())
        r.bold = True
        r.font.size = Pt(7.5)
        r.font.color.rgb = MUTED
        r.font.name = FONT_NAME

        p_value = cell.add_paragraph()
        r = p_value.add_run(card["value"])
        r.bold = True
        r.font.size = Pt(15)
        r.font.color.rgb = NAVY
        r.font.name = FONT_NAME

        p_change = cell.add_paragraph()
        arrow = "\u25B2" if card["positive"] else "\u25BC"  # ▲ / ▼
        r = p_change.add_run(f"{arrow} {card['change']} YoY")
        r.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = GREEN if card["positive"] else RED
        r.font.name = FONT_NAME
    return table


def add_data_table(doc: Document, headers, rows, first_col_left=True):
    """
    headers = ["Indicator", "Jun-26 (Rs.)", "Jun-25 (Rs.)", "Change"]
    rows    = [["Revenue", "49,715,861,103", "36,739,108,828", "+35.3%"], ...]
    Auto color-codes any cell containing a % sign: green if it starts with '+',
    red if it starts with '-' or is wrapped in parentheses.
    """
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"

    # header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, NAVY_HEX)
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(h)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.name = FONT_NAME

    # body rows
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        bg = WHITE_HEX if ri % 2 == 0 else LIGHT_BG_HEX
        for ci, val in enumerate(row):
            if ci >= len(cells):
                break
            cell = cells[ci]
            set_cell_shading(cell, bg)
            cell.paragraphs[0].text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            r = p.add_run(str(val))
            r.font.size = Pt(9.5)
            r.font.name = FONT_NAME
            if ci == 0:
                r.bold = True
                r.font.color.rgb = NAVY
            elif "%" in str(val):
                is_neg = str(val).strip().startswith("-") or "(" in str(val)
                r.bold = True
                r.font.color.rgb = RED if is_neg else GREEN
            else:
                r.font.color.rgb = GREY_TEXT
    return table


def add_footer_page_number(section, brand_text="AI-generated summary — verify against source"):
    """Adds 'Page X of Y' style footer with a muted disclaimer, centered."""
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.text = ""
    r = p.add_run("Page ")
    r.font.size = Pt(7.5)
    r.font.color.rgb = MUTED
    add_page_number_field(p)
    r2 = p.add_run(f"   |   {brand_text}")
    r2.font.size = Pt(7.5)
    r2.font.color.rgb = MUTED
