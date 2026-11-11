"""Shared DOCX styling helpers for customer report renderers."""
from __future__ import annotations

from typing import Any, Mapping, Sequence

from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from app.customer_report.report_summary_helpers import normalize_text


CONTENT_WIDTH_IN = 6.5
CONTENT_WIDTH_DXA = 9360
TABLE_INDENT_DXA = 120
FONT_NAME = "宋体"
FONT_SIZE_BODY = 10.0
FONT_SIZE_LABEL = 9.5
FONT_SIZE_SMALL = 9.0
FONT_SIZE_TABLE = 9.5
SECTION_ORDER = (
    "AI综合智能分析建议",
    "客户全貌",
    "股权关联",
    "经营能力",
    "财务分析",
    "风险提示",
    "行内合作情况",
    "访谈脚本",
    "营销方案",
)
SECTION_TONES = {
    "AI综合智能分析建议": "7B2E5A",
    "客户全貌": "1F5F8B",
    "股权关联": "28745A",
    "经营能力": "9A651A",
    "财务分析": "4A6FA5",
    "风险提示": "A43C3C",
    "行内合作情况": "315C7A",
    "访谈脚本": "6851A0",
    "营销方案": "1E6F75",
}
SECTION_SOURCE_COUNTS = {
    "AI综合智能分析建议": "综合判断",
    "客户全貌": "11 类接口",
    "股权关联": "8 类接口",
    "经营能力": "17 类接口",
    "财务分析": "财务数据",
    "风险提示": "27 类接口",
    "行内合作情况": "CRM 数据",
    "访谈脚本": "客群题库",
    "营销方案": "产品配置",
}


def _style_table(table, *, widths: Sequence[float], header: bool) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    _set_table_geometry(table, widths)
    for row_index, row in enumerate(table.rows):
        for cell_index, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_margins(cell, top=80, bottom=80, start=120, end=120)
            _set_cell_border(cell, "D9E0E8")
            if header and row_index == 0:
                _set_cell_shading(cell, "F2F4F7")


def _set_table_geometry(table, widths: Sequence[float]) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(CONTENT_WIDTH_DXA))
    tbl_layout = OxmlElement("w:tblLayout")
    tbl_layout.set(qn("w:type"), "fixed")
    tbl_pr.append(tbl_layout)
    tbl_ind = OxmlElement("w:tblInd")
    tbl_ind.set(qn("w:w"), str(TABLE_INDENT_DXA))
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_pr.append(tbl_ind)
    grid = tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(int(width * 1440)))
        grid.append(grid_col)
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = Inches(width)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.first_child_found_in("w:tcW")
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:type"), "dxa")
            tc_w.set(qn("w:w"), str(int(width * 1440)))


def _cell_text(cell, value: Any, *, bold: bool = False, size: float = FONT_SIZE_TABLE, color: str = "26313C", align=None, blank_as_dash: bool = True) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align or WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.10
    text = normalize_text(value)
    if blank_as_dash and not text:
        text = "-"
    for index, part in enumerate(str(text).split("\n")):
        if index:
            paragraph.add_run().add_break()
        run = paragraph.add_run(part)
        run.bold = bold
        run.font.name = FONT_NAME
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
        _set_run_east_asia_font(run, FONT_NAME)


def _section_heading_cell_text(cell, title: Any, source_count: Any) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.10
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(6.05), WD_TAB_ALIGNMENT.RIGHT)

    title_run = paragraph.add_run(normalize_text(title) or "-")
    title_run.bold = True
    title_run.font.name = FONT_NAME
    title_run.font.size = Pt(13)
    title_run.font.color.rgb = RGBColor.from_string("16202A")
    _set_run_east_asia_font(title_run, FONT_NAME)

    paragraph.add_run("\t")
    count_run = paragraph.add_run(normalize_text(source_count) or "-")
    count_run.font.name = FONT_NAME
    count_run.font.size = Pt(8.5)
    count_run.font.color.rgb = RGBColor.from_string("657180")
    _set_run_east_asia_font(count_run, FONT_NAME)


def _merge_row(cells) -> None:
    if len(cells) <= 1:
        return
    cells[0].merge(cells[-1])


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_border(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def _set_cell_margins(cell, *, top: int, bottom: int, start: int, end: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for name, value in (("top", top), ("bottom", bottom), ("start", start), ("end", end)):
        node = margins.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_east_asia_font(style, font_name: str) -> None:
    r_pr = style.element.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.append(fonts)
    fonts.set(qn("w:eastAsia"), font_name)


def _set_run_east_asia_font(run, font_name: str) -> None:
    r_pr = run._element.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.append(fonts)
    fonts.set(qn("w:eastAsia"), font_name)
