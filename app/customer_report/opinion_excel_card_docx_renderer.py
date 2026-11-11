"""Render the opinion Excel report with card-like visual style matching the sample report format."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from app.customer_report.customer_report_docx_renderer import (
    CONTENT_WIDTH_IN,
    FONT_NAME,
    FONT_SIZE_BODY,
    FONT_SIZE_LABEL,
    FONT_SIZE_SMALL,
    FONT_SIZE_TABLE,
    _cell_text,
    _set_cell_shading,
    _set_east_asia_font,
    _set_run_east_asia_font,
    _style_table,
)
from app.customer_report.report_summary_helpers import normalize_text
from app.customer_report.opinion_excel_docx_renderer import (
    DEFAULT_PRODUCT_RECOMMENDATION_PATH,
    OpinionExcelDOCXRenderer,
    OpinionReportField,
    RiskDetailTable,
    SECTION_TONES,
    BID_DETAIL_HEADERS,
    _bid_record_value,
    _bid_records,
    _change_records,
    _equity_change_detail_records,
    _group_fields,
    _investor_detail_records,
    _is_displayable_value,
    _is_long_text_field,
    _filter_displayable_field_rows,
    _layout_row_count,
    _outbound_investment_detail_records,
    _overall_risk_warning_level,
    _risk_record_value,
    _should_hide_submodule_heading,
    _title_with_warning,
    generate_ai_analysis_items,
    generate_marketing_recommendation,
)


TEXT_DARK = "202937"
TEXT_MUTED = "6B778C"
ACCENT_RED = "D9233F"
ACCENT_BLUE = "2F6FEB"
ACCENT_GREEN = "159B68"
ACCENT_ORANGE = "E58A17"
CARD_RED_BG = "FFF1F3"
CARD_BLUE_BG = "EDF5FF"
CARD_GREEN_BG = "EAF8F2"
CARD_ORANGE_BG = "FFF6E7"
CARD_BLUE_LIGHT = "EAF1F7"
CARD_TEXT = "203040"

SECTION_SUBTITLES: dict[str, str] = {
    "AI综合智能分析建议": "综合判断与行动建议",
    "客户全貌": "基础信息与工商变更",
    "股权关联": "股东、实控人与对外投资",
    "经营能力": "招投标、融资与经营活跃度",
    "财务分析": "年报财务摘要与核心指标",
    "风险提示": "涉诉、处罚及经营风险",
    "行内合作情况": "CRM 行内合作数据",
    "访谈脚本": "客户经理现场访谈问题",
    "营销方案": "对公产品推荐与营销建议",
}
SECTION_NUMBERS: dict[str, str] = {
    "AI综合智能分析建议": "01",
    "客户全貌": "02",
    "股权关联": "03",
    "经营能力": "04",
    "财务分析": "05",
    "风险提示": "06",
    "行内合作情况": "07",
    "访谈脚本": "08",
    "营销方案": "09",
}


def _add_fallback_paragraph_for_module(doc: Document, module_name: str) -> None:
    """为无数据的新 9 模块输出兜底文案."""
    from app.module_registry import MODULE_REGISTRY

    for mod_id, mod_def in MODULE_REGISTRY.items():
        if mod_def["name"] == module_name:
            fallback = mod_def.get("fallback")
            if fallback:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                r = p.add_run(fallback)
                r.font.name = FONT_NAME
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
                _set_run_east_asia_font(r, FONT_NAME)
            break


class OpinionExcelCardDOCXRenderer(OpinionExcelDOCXRenderer):
    """Visual-variant renderer matching the sample report card style."""

    def _setup_document(self, doc: Document) -> None:
        section = doc.sections[0]
        section.top_margin = Inches(0.58)
        section.right_margin = Inches(0.72)
        section.bottom_margin = Inches(0.58)
        section.left_margin = Inches(0.72)

    def _setup_styles(self, doc: Document) -> None:
        normal = doc.styles["Normal"]
        normal.font.name = FONT_NAME
        normal.font.size = Pt(FONT_SIZE_BODY)
        normal.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_east_asia_font(normal, FONT_NAME)
        normal.paragraph_format.space_after = Pt(5)
        normal.paragraph_format.line_spacing = 1.25

    def _add_header(self, doc: Document, company_name: str, generated_at) -> None:
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(10)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run("客户画像报告")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(22)
        run.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(run, FONT_NAME)

        subtitle = doc.add_paragraph()
        subtitle.paragraph_format.space_after = Pt(10)
        run1 = subtitle.add_run(company_name)
        run1.bold = True
        run1.font.name = FONT_NAME
        run1.font.size = Pt(16)
        run1.font.color.rgb = RGBColor.from_string(ACCENT_RED)
        _set_run_east_asia_font(run1, FONT_NAME)
        run2 = subtitle.add_run("  /  企业全景洞察与营销建议")
        run2.font.name = FONT_NAME
        run2.font.size = Pt(FONT_SIZE_SMALL)
        run2.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
        _set_run_east_asia_font(run2, FONT_NAME)

        meta = doc.add_table(rows=1, cols=2)
        _style_table(meta, widths=[0.8, CONTENT_WIDTH_IN - 0.8], header=False)
        _set_cell_shading(meta.rows[0].cells[0], CARD_BLUE_LIGHT)
        _cell_text(meta.rows[0].cells[0], "报告日期", bold=True, size=FONT_SIZE_SMALL, color=TEXT_DARK)
        _cell_text(meta.rows[0].cells[1], f"{generated_at:%Y-%m-%d}", size=FONT_SIZE_SMALL)
        doc.add_paragraph()

    def _add_ai_insight_card(self, doc: Document, field_values: Mapping[str, Any]) -> None:
        industry = normalize_text(field_values.get("国标行业")) or ""
        status = normalize_text(field_values.get("企业经营状态")) or ""
        rating = normalize_text(field_values.get("信用评级")) or ""
        parts = []
        if status and status != "-":
            parts.append(f"经营状态为{status}")
        if industry:
            parts.append(f"所属{industry}")
        if rating:
            parts.append(f"信用评级{rating}")
        if not parts:
            return

        table = doc.add_table(rows=1, cols=1)
        _style_table(table, widths=[CONTENT_WIDTH_IN], header=False)
        _set_cell_shading(table.rows[0].cells[0], CARD_RED_BG)
        cell = table.rows[0].cells[0]
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        prefix = p.add_run("AI 客户洞察  ")
        prefix.font.name = FONT_NAME
        prefix.font.size = Pt(FONT_SIZE_BODY)
        prefix.bold = True
        prefix.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(prefix, FONT_NAME)
        note = p.add_run("基于现有资料自动归纳")
        note.font.name = FONT_NAME
        note.font.size = Pt(FONT_SIZE_SMALL)
        note.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
        _set_run_east_asia_font(note, FONT_NAME)

        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(0)
        summary = p2.add_run("企业" + "、".join(parts) + "。")
        summary.font.name = FONT_NAME
        summary.font.size = Pt(FONT_SIZE_LABEL)
        summary.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(summary, FONT_NAME)
        doc.add_paragraph()

    def _add_identity_cards(self, doc: Document, field_values: Mapping[str, Any], generated_at) -> None:
        credit_rating = normalize_text(field_values.get("信用评级"))
        reg_capital = normalize_text(field_values.get("注册 / 实缴资本（万元）"))
        establish_date = normalize_text(field_values.get("工商核准日期"))
        company_status = normalize_text(field_values.get("企业经营状态"))
        risk_level = normalize_text(field_values.get("风险提示")) or normalize_text(field_values.get("风险等级"))

        years = ""
        if establish_date and len(establish_date) >= 4:
            try:
                y = int(establish_date[:4])
                years = f"{generated_at.year - y} 年"
            except ValueError:
                pass

        cards: list[tuple[str, str, str, str, str]] = []
        if _is_displayable_value(credit_rating):
            cards.append((_format_card_value(credit_rating), "信用评级", establish_date, CARD_BLUE_BG, ACCENT_BLUE))
        if _is_displayable_value(reg_capital):
            cards.append((_format_card_value(reg_capital), "注册资本", "", CARD_RED_BG, ACCENT_RED))
        if _is_displayable_value(years):
            cards.append((years, "成立年限", f"{establish_date} 成立" if establish_date else "", CARD_GREEN_BG, ACCENT_GREEN))
        if _is_displayable_value(company_status):
            cards.append((company_status, "经营状态", "", CARD_ORANGE_BG, ACCENT_ORANGE))
        elif _is_displayable_value(risk_level):
            cards.append((risk_level, "风险等级", "", CARD_ORANGE_BG, ACCENT_ORANGE))

        if not cards:
            return

        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(5)
        title.paragraph_format.space_after = Pt(4)
        run_icon = title.add_run("▸ ")
        run_icon.font.name = FONT_NAME
        run_icon.font.size = Pt(FONT_SIZE_BODY)
        run_icon.bold = True
        run_icon.font.color.rgb = RGBColor.from_string(ACCENT_RED)
        _set_run_east_asia_font(run_icon, FONT_NAME)
        run_label = title.add_run("企业身份概览")
        run_label.font.name = FONT_NAME
        run_label.font.size = Pt(11)
        run_label.bold = True
        run_label.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(run_label, FONT_NAME)

        col_count = min(len(cards), 4)
        table = doc.add_table(rows=1, cols=col_count)
        _style_table(table, widths=[CONTENT_WIDTH_IN / col_count] * col_count, header=False)

        for i, (value, label, detail, bg, accent) in enumerate(cards[:col_count]):
            _set_cell_shading(table.rows[0].cells[i], bg)
            cell = table.rows[0].cells[i]
            cell.text = ""
            p1 = cell.paragraphs[0]
            p1.paragraph_format.space_after = Pt(1)
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(value)
            r1.bold = True
            r1.font.name = FONT_NAME
            r1.font.size = Pt(12)
            r1.font.color.rgb = RGBColor.from_string(TEXT_DARK)
            _set_run_east_asia_font(r1, FONT_NAME)

            p2 = cell.add_paragraph()
            p2.paragraph_format.space_after = Pt(1)
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(label)
            r2.font.name = FONT_NAME
            r2.font.size = Pt(FONT_SIZE_SMALL)
            r2.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
            _set_run_east_asia_font(r2, FONT_NAME)

            if detail:
                p3 = cell.add_paragraph()
                p3.paragraph_format.space_after = Pt(0)
                p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r3 = p3.add_run(detail)
                r3.bold = True
                r3.font.name = FONT_NAME
                r3.font.size = Pt(FONT_SIZE_SMALL)
                r3.font.color.rgb = RGBColor.from_string(accent)
                _set_run_east_asia_font(r3, FONT_NAME)

        doc.add_paragraph()

    def _add_tag_bar(self, doc: Document, field_values: Mapping[str, Any]) -> None:
        tags = []
        company_type = normalize_text(field_values.get("企业类型"))
        if company_type:
            tags.append(company_type)
        scale = normalize_text(field_values.get("企业规模"))
        if scale:
            tags.append(scale)
        rating = normalize_text(field_values.get("信用评级"))
        if rating and rating != "-":
            tags.append(f"{rating} 评级")
        tax = normalize_text(field_values.get("税务评级"))
        if tax and tax != "-":
            tags.append(f"纳税{tax}级")
        status = normalize_text(field_values.get("企业经营状态"))
        if status:
            tags.append(status)
        if not tags:
            return

        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(5)
        title.paragraph_format.space_after = Pt(4)
        run_icon = title.add_run("▸ ")
        run_icon.font.name = FONT_NAME
        run_icon.font.size = Pt(10)
        run_icon.bold = True
        run_icon.font.color.rgb = RGBColor.from_string(ACCENT_RED)
        _set_run_east_asia_font(run_icon, FONT_NAME)
        run_label = title.add_run("客户标签")
        run_label.font.name = FONT_NAME
        run_label.font.size = Pt(11)
        run_label.bold = True
        run_label.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(run_label, FONT_NAME)

        tag_style = [
            (CARD_RED_BG, ACCENT_RED),
            (CARD_BLUE_BG, ACCENT_BLUE),
            (CARD_GREEN_BG, ACCENT_GREEN),
            (CARD_ORANGE_BG, ACCENT_ORANGE),
        ]
        table = doc.add_table(rows=1, cols=len(tags))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(tags)] * len(tags), header=False)
        for i, tag in enumerate(tags):
            bg, accent = tag_style[i % len(tag_style)]
            _set_cell_shading(table.rows[0].cells[i], bg)
            _cell_text(table.rows[0].cells[i], tag, bold=True, size=FONT_SIZE_SMALL, color=accent, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()

    def _add_financial_cards(self, doc: Document, field_values: Mapping[str, Any]) -> None:
        revenue = normalize_text(field_values.get("近 1-3 年营收数据及变动情况"))
        industry = normalize_text(field_values.get("企业所属行业") or field_values.get("国标行业"))
        bidding_text = normalize_text(field_values.get("企业招投标信息"))
        debt_ratio = normalize_text(field_values.get("资产负债率"))

        if _is_displayable_value(revenue):
            try:
                n = float(revenue.replace(",", "").replace("，", ""))
                if abs(n) >= 100000000:
                    revenue = f"{n / 100000000:.2f} 亿"
                elif abs(n) >= 10000:
                    revenue = f"{n / 10000:.2f} 万"
                else:
                    revenue = f"{n:,.2f}"
            except (ValueError, AttributeError):
                pass

        cards: list[tuple[str, str, str, str]] = []
        if _is_displayable_value(revenue):
            cards.append((revenue, "营业总收入", CARD_BLUE_BG, ACCENT_BLUE))
        if _is_displayable_value(industry):
            display_industry = industry if len(industry) <= 12 else f"{industry[:12]}…"
            cards.append((display_industry, "所属行业", CARD_ORANGE_BG, ACCENT_ORANGE))
        if _is_displayable_value(bidding_text):
            cards.append((bidding_text, "招投标", CARD_GREEN_BG, ACCENT_GREEN))
        if _is_displayable_value(debt_ratio):
            cards.append((debt_ratio, "资产负债率", CARD_RED_BG, ACCENT_RED))

        if not cards:
            return

        col_count = min(len(cards), 4)
        table = doc.add_table(rows=1, cols=col_count)
        _style_table(table, widths=[CONTENT_WIDTH_IN / col_count] * col_count, header=False)
        for i, (value, label, bg, accent) in enumerate(cards[:col_count]):
            _set_cell_shading(table.rows[0].cells[i], bg)
            cell = table.rows[0].cells[i]
            cell.text = ""
            p1 = cell.paragraphs[0]
            p1.paragraph_format.space_after = Pt(1)
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p1.add_run(value)
            r1.bold = True
            r1.font.name = FONT_NAME
            r1.font.size = Pt(12)
            r1.font.color.rgb = RGBColor.from_string(TEXT_DARK)
            _set_run_east_asia_font(r1, FONT_NAME)
            p2 = cell.add_paragraph()
            p2.paragraph_format.space_after = Pt(0)
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = p2.add_run(label)
            r2.font.name = FONT_NAME
            r2.font.size = Pt(FONT_SIZE_SMALL)
            r2.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
            _set_run_east_asia_font(r2, FONT_NAME)
        doc.add_paragraph()

    def _add_section_heading(self, doc: Document, title: str, source_count: str, tone: str) -> None:
        clean_title = normalize_text(title).replace("（无数据）", "").strip()
        num = SECTION_NUMBERS.get(clean_title, "")
        subtitle = SECTION_SUBTITLES.get(clean_title, "")

        table = doc.add_table(rows=1, cols=1)
        _style_table(table, widths=[CONTENT_WIDTH_IN], header=False)
        _set_cell_shading(table.rows[0].cells[0], "FFFFFF")
        cell = table.rows[0].cells[0]
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_before = Pt(10)
        paragraph.paragraph_format.space_after = Pt(7)

        prefix = f"{num}  " if num else ""
        r1 = paragraph.add_run(f"{prefix}{title}")
        r1.bold = True
        r1.font.name = FONT_NAME
        r1.font.size = Pt(15)
        r1.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(r1, FONT_NAME)

        if subtitle:
            r2 = paragraph.add_run(f"   {subtitle}")
            r2.font.name = FONT_NAME
            r2.font.size = Pt(FONT_SIZE_SMALL)
            r2.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
            _set_run_east_asia_font(r2, FONT_NAME)

    def _add_submodule_heading(self, doc: Document, text: str) -> None:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(5)
        paragraph.paragraph_format.space_after = Pt(4)
        run_icon = paragraph.add_run("▸ ")
        run_icon.font.name = FONT_NAME
        run_icon.font.size = Pt(10)
        run_icon.bold = True
        run_icon.font.color.rgb = RGBColor.from_string(ACCENT_RED)
        _set_run_east_asia_font(run_icon, FONT_NAME)
        run_label = paragraph.add_run(normalize_text(text))
        run_label.font.name = FONT_NAME
        run_label.font.size = Pt(11)
        run_label.bold = True
        run_label.font.color.rgb = RGBColor.from_string(TEXT_DARK)
        _set_run_east_asia_font(run_label, FONT_NAME)

    def _add_submodule_table(
        self, doc, module_name, submodule_name, rows, *,
        interface_responses, field_values, annual_financial_table,
    ) -> None:
        if module_name == "经营能力" and submodule_name == "营收与财报":
            if annual_financial_table:
                if not _should_hide_submodule_heading(module_name, submodule_name):
                    self._add_submodule_heading(doc, submodule_name)
                self._add_financial_cards(doc, field_values)
                self._add_annual_financial_table(doc, annual_financial_table)
                return
            return super()._add_submodule_table(
                doc, module_name, submodule_name, rows,
                interface_responses=interface_responses, field_values=field_values,
                annual_financial_table=annual_financial_table,
            )

        resolved_rows = _filter_displayable_field_rows(
            rows,
            [(item, self._resolve_field_value(item, interface_responses, field_values)) for item in rows],
        )
        if not resolved_rows and module_name == "财务分析":
            return
        has_change_field = any(item.name == "变更记录" for item in rows)
        change_records = self._get_table_records("change_records") or _change_records(interface_responses.get("9000000004")) if has_change_field else []
        has_bid_field = any(item.name == "企业招投标信息" for item in rows)
        bid_records = self._get_table_records("bid_records") or _bid_records(interface_responses) if has_bid_field else []
        bid_field_names = {"企业招投标信息"} if bid_records else set()
        change_field_names = {"变更记录"} if change_records else set()
        table_mortgage_names = {"动产抵押"}
        table_pledge_names = {"股权出质"}
        table_financing_names = {"融资情况"}
        table_execution_risk_names = {"执行风险"}
        table_consumption_limit_names = {"限制高消费"}
        table_discredit_executor_names = {"失信执行人"}
        exclude_names = bid_field_names | change_field_names | table_mortgage_names | table_pledge_names | table_financing_names | table_execution_risk_names | table_consumption_limit_names | table_discredit_executor_names
        table_rows = [(item, value) for item, value in resolved_rows if item.name not in exclude_names]

        has_mortgage_field = any(item.name == "动产抵押" for item in rows)
        has_pledge_field = any(item.name == "股权出质" for item in rows)
        has_financing_field = any(item.name == "融资情况" for item in rows)
        has_execution_risk_field = any(item.name in table_execution_risk_names for item in rows)
        has_consumption_limit_field = any(item.name in table_consumption_limit_names for item in rows)
        has_discredit_executor_field = any(item.name in table_discredit_executor_names for item in rows)

        if not table_rows and not any((change_records, bid_records)):
            has_detail = any([
                has_mortgage_field and self._get_table_records("mortgage_records"),
                has_pledge_field and self._get_table_records("pledge_records"),
                has_financing_field and self._get_table_records("financing_records"),
                has_execution_risk_field and self._get_table_records("execution_risk"),
                has_consumption_limit_field and self._get_table_records("consumption_limit"),
                has_discredit_executor_field and self._get_table_records("discredit_executor"),
            ])
            if not has_detail:
                return

        if not _should_hide_submodule_heading(module_name, submodule_name):
            self._add_submodule_heading(doc, submodule_name)

        if table_rows:
            table_row_count = _layout_row_count(table_rows)
            table = doc.add_table(rows=max(table_row_count, 1), cols=4)
            _style_table(table, widths=[1.3, 2.0, 1.3, 2.0], header=False)
            row_index = 0
            pending_pair: list = []
            for item, value in table_rows:
                if _is_long_text_field(item, value):
                    if pending_pair:
                        self._write_pair_row(table.rows[row_index].cells, pending_pair)
                        row_index += 1
                        pending_pair = []
                    self._write_long_row(table.rows[row_index].cells, item, value)
                    row_index += 1
                else:
                    pending_pair.append((item, value))
                    if len(pending_pair) == 2:
                        self._write_pair_row(table.rows[row_index].cells, pending_pair)
                        row_index += 1
                        pending_pair = []
            if pending_pair:
                self._write_pair_row(table.rows[row_index].cells, pending_pair)

        if change_records:
            self._add_change_detail_table(doc, change_records[:3], total_count=len(change_records))
        if bid_records:
            self._add_bid_detail_table(doc, bid_records[:3], total_count=len(bid_records))

        if has_mortgage_field:
            mortgage_records = self._get_table_records("mortgage_records")
            if mortgage_records:
                self._add_mortgage_detail_table(doc, mortgage_records[:3], total_count=len(mortgage_records))
        if has_pledge_field:
            pledge_records = self._get_table_records("pledge_records")
            if pledge_records:
                self._add_pledge_detail_table(doc, pledge_records[:3], total_count=len(pledge_records))
        if has_financing_field:
            financing_records = self._get_table_records("financing_records")
            if financing_records:
                self._add_financing_detail_table(doc, financing_records[:3], total_count=len(financing_records))

        if has_execution_risk_field:
            execution_risk_records = self._get_table_records("execution_risk")
            if execution_risk_records:
                self._add_execution_risk_detail_table(doc, execution_risk_records[:3], total_count=len(execution_risk_records))

        if has_consumption_limit_field:
            consumption_limit_records = self._get_table_records("consumption_limit")
            if consumption_limit_records:
                self._add_consumption_limit_detail_table(doc, consumption_limit_records[:3], total_count=len(consumption_limit_records))

        if has_discredit_executor_field:
            discredit_executor_records = self._get_table_records("discredit_executor")
            if discredit_executor_records:
                self._add_discredit_executor_detail_table(doc, discredit_executor_records[:3], total_count=len(discredit_executor_records))

    def _write_pair_row(self, cells, pairs):
        for pair_index in range(2):
            label_cell = cells[pair_index * 2]
            value_cell = cells[pair_index * 2 + 1]
            _set_cell_shading(label_cell, CARD_BLUE_LIGHT)
            if pair_index < len(pairs):
                item, value = pairs[pair_index]
                _cell_text(label_cell, item.name, bold=True, size=FONT_SIZE_LABEL, color=TEXT_DARK)
                _cell_text(value_cell, value, size=FONT_SIZE_TABLE, color=TEXT_DARK)
            else:
                _cell_text(label_cell, "", bold=True, size=FONT_SIZE_LABEL, color=TEXT_DARK)
                _cell_text(value_cell, "", size=FONT_SIZE_TABLE, color=TEXT_DARK)

    def _write_long_row(self, cells, item, value):
        _set_cell_shading(cells[0], CARD_BLUE_LIGHT)
        _cell_text(cells[0], item.name, bold=True, size=FONT_SIZE_LABEL, color=TEXT_DARK)
        merged = cells[1].merge(cells[3])
        _cell_text(merged, value, size=FONT_SIZE_TABLE, color=TEXT_DARK)

    def _has_equity_detail_data(self, interface_responses: Mapping[str, Any]) -> bool:
        return bool(
            self._get_table_records("investor_detail")
            or _investor_detail_records(interface_responses)
            or _equity_change_detail_records(interface_responses)
            or self._get_table_records("outbound_investment")
            or _outbound_investment_detail_records(interface_responses)
        )

    def _add_equity_detail_section(self, doc: Document, interface_responses: Mapping[str, Any]) -> None:
        self._add_section_heading(doc, "股权关联", "-", SECTION_TONES.get("股权关联", "315C7A"))
        self._add_investor_detail_table(
            doc,
            self._get_table_records("investor_detail") or _investor_detail_records(interface_responses),
        )
        self._add_equity_change_detail_table(doc, _equity_change_detail_records(interface_responses))
        self._add_outbound_investment_detail_table(
            doc,
            self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses),
        )
        doc.add_paragraph()

    def render(
        self, fields, output_path, *,
        company_name="企业客户", interface_responses=None, field_values=None,
        annual_financial_table=None, internal_info_section=None,
        risk_detail_tables=None, table_data=None,
        product_recommendation_path=DEFAULT_PRODUCT_RECOMMENDATION_PATH,
        disable_all_llm=False, generated_at=None,
        explicit_module_filter=False,
    ) -> Path:
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        generated_at = generated_at or datetime.now()
        interface_responses = interface_responses or {}
        field_values = field_values or {}
        self._table_data = table_data or {}
        self._current_company_name = company_name
        self._relationship_graph_path = output.with_name(f"{output.stem}_relationship_graph.png")
        self._disable_all_llm = bool(disable_all_llm)

        doc = Document()
        self._setup_document(doc)
        self._setup_styles(doc)
        self._add_header(doc, company_name, generated_at)
        self._add_ai_insight_card(doc, field_values)
        self._add_identity_cards(doc, field_values, generated_at)

        field_module_names = {f.module for f in fields}

        grouped_fields = _group_fields(fields)

        for module_name, submodules in grouped_fields.items():
            has_data = self._module_has_displayable_data(
                module_name, submodules,
                interface_responses=interface_responses, field_values=field_values,
                annual_financial_table=annual_financial_table,
                risk_detail_tables=risk_detail_tables or (),
            )
            if not has_data:
                continue

            section_title = module_name
            if module_name == "风险提示":
                section_title = _title_with_warning(module_name, _overall_risk_warning_level(risk_detail_tables or (), interface_responses))
            self._add_section_heading(doc, section_title, self._source_count(module_name, submodules), SECTION_TONES.get(module_name, "315C7A"))

            if module_name == "客户全貌":
                self._add_tag_bar(doc, field_values)

            for submodule_name, rows in submodules.items():
                self._add_submodule_table(
                    doc, module_name, submodule_name, rows,
                    interface_responses=interface_responses, field_values=field_values,
                    annual_financial_table=annual_financial_table,
                )
            if module_name == "行内合作情况" and internal_info_section and internal_info_section.get("internal_groups"):
                self._add_internal_info(doc, internal_info_section)
            if module_name == "风险提示":
                if risk_detail_tables:
                    _INLINE_RISK_CODES = {"9000000041", "9000000039", "9000000040"}
                    filtered = [t for t in risk_detail_tables if not set(t.interface_codes) & _INLINE_RISK_CODES]
                    if filtered:
                        self._add_risk_detail_tables(doc, filtered, interface_responses)
            doc.add_paragraph()

        if "股权关联" not in grouped_fields and self._has_equity_detail_data(interface_responses):
            self._add_equity_detail_section(doc, interface_responses)

        if not explicit_module_filter and internal_info_section and internal_info_section.get("internal_groups") and "行内合作情况" not in field_module_names:
            self._add_section_heading(doc, "行内合作情况", internal_info_section.get("source_count") or "CRM 数据", SECTION_TONES.get("行内合作情况", "315C7A"))
            self._add_internal_info(doc, internal_info_section)
            doc.add_paragraph()

        if not explicit_module_filter and (
            "AI综合智能分析建议" not in field_module_names
            and "访谈脚本" not in field_module_names
            and "营销方案" not in field_module_names
        ):
            self._add_ai_analysis_section(
                doc, company_name, fields,
                interface_responses, field_values, annual_financial_table,
                risk_detail_tables or (), product_recommendation_path,
            )
        doc.save(output)
        return output

    def _add_ai_analysis_section(
        self, doc, company_name, fields, interface_responses, field_values,
        annual_financial_table, risk_detail_tables, product_recommendation_path,
    ) -> None:
        recommendation = generate_marketing_recommendation(
            company_name, fields, interface_responses,
            field_values=field_values, annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            product_recommendation_path=product_recommendation_path,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )
        items = recommendation.get("products") or []
        if not items and recommendation.get("product"):
            items = [{"product": recommendation["product"], "reason": recommendation.get("reason") or ""}]

        analysis_items = generate_ai_analysis_items(
            company_name, fields, interface_responses,
            field_values=field_values, annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )
        if not items and not analysis_items:
            return

        if items:
            self._add_section_heading(doc, "营销方案", "-", SECTION_TONES.get("营销方案", "315C7A"))
            self._add_ai_marketing_cards(
                doc, company_name, fields, interface_responses, field_values,
                annual_financial_table, risk_detail_tables, product_recommendation_path,
            )

        if analysis_items:
            self._add_section_heading(doc, "AI 智能分析", "-", SECTION_TONES.get("AI综合智能分析建议", "315C7A"))
            self._add_company_overall_evaluation(
                doc, company_name, fields, interface_responses, field_values,
                annual_financial_table, risk_detail_tables, items=analysis_items,
            )

    def _add_ai_interview_table(self, doc: Document) -> None:
        table = doc.add_table(rows=1, cols=2)
        _style_table(table, widths=[1.2, CONTENT_WIDTH_IN - 1.2], header=False)
        _set_cell_shading(table.rows[0].cells[0], "F5F6F8")
        _cell_text(table.rows[0].cells[0], "访谈脚本", bold=True, size=8.5)
        _cell_text(table.rows[0].cells[1], "待补充：后续可结合AI智能分析、风险提示和营销方案生成客户经理现场访谈问题清单。", size=8.5)

    def _add_ai_marketing_cards(
        self, doc, company_name, fields, interface_responses, field_values,
        annual_financial_table, risk_detail_tables, product_recommendation_path,
    ) -> None:
        recommendation = generate_marketing_recommendation(
            company_name, fields, interface_responses,
            field_values=field_values, annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            product_recommendation_path=product_recommendation_path,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )
        items = recommendation.get("products") or []
        if not items and recommendation.get("product"):
            items = [{"product": recommendation["product"], "reason": recommendation.get("reason") or "", "source": recommendation.get("source") or "fallback"}]

        for index, item in enumerate(items[:5], start=1):
            product = item.get("product")
            if not product:
                continue
            card = doc.add_table(rows=3, cols=2)
            _style_table(card, widths=[1.08, CONTENT_WIDTH_IN - 1.08], header=False)
            _set_cell_shading(card.rows[0].cells[0], "F5F6F8")
            _cell_text(card.rows[0].cells[0], f"推荐产品 {index}", bold=True, size=FONT_SIZE_LABEL)
            _cell_text(card.rows[0].cells[1], product.product_name, bold=True, size=FONT_SIZE_BODY, color=TEXT_DARK)
            _set_cell_shading(card.rows[1].cells[0], "F5F6F8")
            _cell_text(card.rows[1].cells[0], "适用客群", bold=True, size=FONT_SIZE_SMALL)
            _cell_text(card.rows[1].cells[1], product.customer_feature or "-", size=FONT_SIZE_TABLE)
            _set_cell_shading(card.rows[2].cells[0], "F5F6F8")
            _cell_text(card.rows[2].cells[0], "推荐理由", bold=True, size=FONT_SIZE_SMALL)
            _cell_text(card.rows[2].cells[1], item.get("reason") or "-", size=FONT_SIZE_TABLE)

    def _add_company_overall_evaluation(
        self, doc, company_name, fields, interface_responses, field_values,
        annual_financial_table, risk_detail_tables, *, items=None,
    ) -> None:
        analysis_items = items or generate_ai_analysis_items(
            company_name, fields, interface_responses,
            field_values=field_values, annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )
        if not analysis_items:
            return

        self._add_submodule_heading(doc, "经营判断")
        for index, item in enumerate(analysis_items[:3], start=1):
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = Pt(0)
            para.paragraph_format.space_after = Pt(5)
            para.paragraph_format.line_spacing = 1.2
            prefix = para.add_run(f"{index}. 【{item['title']}】")
            prefix.bold = True
            prefix.font.name = FONT_NAME
            prefix.font.size = Pt(FONT_SIZE_LABEL)
            prefix.font.color.rgb = RGBColor.from_string(TEXT_DARK)
            _set_run_east_asia_font(prefix, FONT_NAME)
            body = para.add_run(f"{item['summary']}")
            body.font.name = FONT_NAME
            body.font.size = Pt(FONT_SIZE_TABLE)
            body.font.color.rgb = RGBColor.from_string(TEXT_DARK)
            _set_run_east_asia_font(body, FONT_NAME)
            if item.get("attention"):
                note = doc.add_paragraph()
                note.paragraph_format.space_after = Pt(4)
                note.paragraph_format.left_indent = Pt(12)
                run = note.add_run(f"关注：{item['attention']}")
                run.font.name = FONT_NAME
                run.font.size = Pt(FONT_SIZE_SMALL)
                run.font.color.rgb = RGBColor.from_string(TEXT_MUTED)
                _set_run_east_asia_font(run, FONT_NAME)

    # --- detail table methods (compact, matching sample) ---

    def _add_annual_financial_table(self, doc, table_data):
        columns = [col for col in (table_data.get("columns") or []) if isinstance(col, Mapping)][:3]
        groups = [group for group in (table_data.get("groups") or []) if isinstance(group, Mapping)]
        if not columns or not groups:
            return
        total_rows = 1 + sum(1 + len(group.get("rows") or []) for group in groups)
        table = doc.add_table(rows=total_rows, cols=len(columns) + 1)
        widths = [1.3] + [(CONTENT_WIDTH_IN - 1.3) / len(columns)] * len(columns)
        _style_table(table, widths=widths, header=False)
        for cell in table.rows[0].cells:
            _set_cell_shading(cell, TEXT_DARK)
        _cell_text(table.rows[0].cells[0], "期间", bold=True, size=FONT_SIZE_SMALL, color="FFFFFF")
        for index, col in enumerate(columns, start=1):
            _cell_text(table.rows[0].cells[index], col.get("label") or "-", bold=True, size=FONT_SIZE_SMALL, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER)
        row_index = 1
        for group in groups:
            cells = table.rows[row_index].cells
            cells[0].merge(cells[-1])
            _set_cell_shading(cells[0], "FFFFFF")
            _cell_text(cells[0], group.get("title") or "-", bold=True, size=7.8, color=ACCENT_RED)
            row_index += 1
            for item in group.get("rows") or []:
                if not isinstance(item, Mapping):
                    continue
                fill = "F4F8FB" if row_index % 2 == 0 else "FFFFFF"
                for cell in table.rows[row_index].cells:
                    _set_cell_shading(cell, fill)
                _cell_text(table.rows[row_index].cells[0], item.get("label") or "-", size=7.5, color=TEXT_DARK)
                values = item.get("values") or {}
                for col_index, col in enumerate(columns, start=1):
                    key = col.get("key")
                    value = values.get(key) if key in values else "-"
                    _cell_text(table.rows[row_index].cells[col_index], value, size=7.5, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER, blank_as_dash=False)
                row_index += 1

    def _add_risk_detail_table(self, doc, table_def, records):
        self._add_submodule_heading(doc, table_def.title)
        headers = table_def.headers[:6]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, _risk_record_value(record, header), size=7.0, color=TEXT_DARK)

    def _add_bid_detail_table(self, doc, records, *, total_count=None):
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"企业招投标信息（共{count}条）")
        headers = BID_DETAIL_HEADERS
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, _bid_record_value(record, header), size=7.0, color=TEXT_DARK)

    def _add_mortgage_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"动产抵押（共{count}条）")
        headers = ["登记编号", "被担保债权数额", "被担保债券种类", "担保范围", "履行债务结束日期", "债务数额", "抵押物名称", "抵押权人名称"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_pledge_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"股权出质（共{count}条）")
        headers = ["登记编号", "出质人名称", "出质状态", "出质股权数额", "出质登记日期"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_financing_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"历史融资金额（共{count}条）")
        headers = ["融资事件全称", "投资金额", "币种中文", "投资金额级别", "融资方名称全称", "投资方名称全称", "投资时间"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_execution_risk_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"执行风险（共{count}条）")
        headers = ["被执行人名称", "案件状态", "案号", "申请执行标的（元）", "执行法院", "立案时间"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_consumption_limit_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"限制高消费（共{count}条）")
        headers = ["判决案号", "发布日期", "立案日期", "被执行人", "执行法院", "组织机构代码", "是否有效"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_discredit_executor_detail_table(self, doc, records, *, total_count=None):
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        self._add_submodule_heading(doc, f"失信执行人（共{count}条）")
        headers = ["案号", "立案时间", "被执行人的履行情况", "执行法院", "案件状态"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, header in zip(table.rows[row_index].cells, headers):
                _set_cell_shading(cell, fill)
                _cell_text(cell, record.get(header) or "-", size=7.0, color=TEXT_DARK)

    def _add_manager_detail_table(self, doc, records):
        if not records:
            return
        self._add_submodule_heading(doc, f"主要管理人员（共{len(records)}人）")
        headers = ("姓名", "职务")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN * 0.25, CONTENT_WIDTH_IN * 0.30], header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell in table.rows[row_index].cells:
                _set_cell_shading(cell, fill)
            _cell_text(table.rows[row_index].cells[0], record.get("姓名") or "-", size=7.0, color=TEXT_DARK)
            _cell_text(table.rows[row_index].cells[1], record.get("职务") or "-", size=7.0, color=TEXT_DARK)

    def _add_investor_detail_table(self, doc, records):
        if not records:
            return
        self._add_submodule_heading(doc, f"股东基础信息（共{len(records)}条）")
        headers = ("股东名称", "出资比例", "股东类型", "认缴出资额度", "出资日期")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN * 0.28, CONTENT_WIDTH_IN * 0.14, CONTENT_WIDTH_IN * 0.14, CONTENT_WIDTH_IN * 0.24, CONTENT_WIDTH_IN * 0.20], header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell in table.rows[row_index].cells:
                _set_cell_shading(cell, fill)
            _cell_text(table.rows[row_index].cells[0], record.get("股东名称") or "-", size=7.0, color=TEXT_DARK)
            _cell_text(table.rows[row_index].cells[1], record.get("出资比例") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[2], record.get("股东类型") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[3], record.get("认缴出资额度") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[4], record.get("出资日期") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _add_equity_change_detail_table(self, doc, records):
        if not records:
            return
        self._add_submodule_heading(doc, f"股权变更记录（共{len(records)}条）")
        headers = ("变更日期", "股东名称", "转让前股权比例", "转让后股权比例")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN * 0.18, CONTENT_WIDTH_IN * 0.38, CONTENT_WIDTH_IN * 0.22, CONTENT_WIDTH_IN * 0.22], header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, header, bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell in table.rows[row_index].cells:
                _set_cell_shading(cell, fill)
            _cell_text(table.rows[row_index].cells[0], record.get("变更日期") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[1], record.get("股东名称") or "-", size=7.0, color=TEXT_DARK)
            _cell_text(table.rows[row_index].cells[2], record.get("转让前股权比例") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[3], record.get("转让后股权比例") or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _add_reserved_key_people_submodule(self, doc):
        self._add_submodule_heading(doc, "受益所有人")
        table = doc.add_table(rows=1, cols=2)
        _style_table(table, widths=[1.25, CONTENT_WIDTH_IN - 1.25], header=False)
        _set_cell_shading(table.rows[0].cells[0], CARD_BLUE_LIGHT)
        _cell_text(table.rows[0].cells[0], "预留字段", bold=True, size=7.8, color=TEXT_DARK)
        _cell_text(table.rows[0].cells[1], "-", size=7.8, color=TEXT_DARK)

    def _add_internal_entry_table(self, doc, entry):
        columns = [col for col in (entry.get("columns") or []) if isinstance(col, Mapping)]
        rows = [row for row in (entry.get("rows") or []) if isinstance(row, Mapping)]
        if not columns or not rows:
            return
        self._add_submodule_heading(doc, entry.get("title") or entry.get("entityInfo") or "明细")
        table = doc.add_table(rows=len(rows) + 1, cols=len(columns))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(columns)] * len(columns), header=True)
        for cell, col in zip(table.rows[0].cells, columns):
            _set_cell_shading(cell, CARD_BLUE_LIGHT)
            _cell_text(cell, col.get("name") or col.get("field") or "-", bold=True, size=7.2, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, item in enumerate(rows, start=1):
            fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
            for cell, col in zip(table.rows[row_index].cells, columns):
                _set_cell_shading(cell, fill)
                _cell_text(cell, item.get(col.get("field")) or "-", size=7.0, color=TEXT_DARK, align=WD_ALIGN_PARAGRAPH.CENTER, blank_as_dash=False)

    def _render_relationship_graph_if_needed(self, rows, interface_responses):
        return ()

    def _add_relationship_graph_image(self, doc, image_path):
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run().add_picture(str(image_path), width=Inches(CONTENT_WIDTH_IN))


def _format_card_value(value: str) -> str:
    if not value or value == "-":
        return "-"
    if "万" in value or "亿" in value or "%" in value:
        return value
    clean = value.replace(",", "").replace("，", "").replace("人民币", "")
    try:
        n = float(clean)
        if abs(n) >= 100000000:
            return f"{n / 100000000:.1f} 亿"
        elif abs(n) >= 10000:
            return f"{n / 10000:,.0f} 万"
        return str(int(n))
    except (ValueError, AttributeError):
        return value[:12]
