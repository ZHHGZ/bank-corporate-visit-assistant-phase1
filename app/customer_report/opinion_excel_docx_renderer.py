"""Render the opinion-collection Excel structure into a customer report DOCX."""
from __future__ import annotations

import json
import re
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime

from pathlib import Path
from typing import Any, Mapping, Sequence

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.customer_report.annual_financial_table import build_zhongshu_annual_financial_table
from app.customer_report.customer_report_docx_renderer import (
    CONTENT_WIDTH_IN,
    FONT_NAME,
    SECTION_SOURCE_COUNTS,
    SECTION_TONES,
    _cell_text,
    _set_east_asia_font,
    _set_run_east_asia_font,
    _set_cell_shading,
    _style_table,
)
from app.customer_report.report_summary_helpers import normalize_text
from app.parsers._format_utils import (
    _billboard_summary_for_catalog,
    _change_records,
    _change_summary,
    _collateral_summary,
    _count_records,
    _count_text,
    _credit_rating_summary_for_catalog,
    _customs_credit_summary,
    _customs_record_summary,
    _date_sort_key,
    _enterprise_certification_summary,
    _equity_change_summary,
    _extract_record_dicts,
    _find_first_scalar_by_keys,
    _find_value_by_label,
    _financing_status_summary,
    _first_record_by_keys,
    _first_value,
    _format_amount_for_display,
    _format_capital_without_decimal,
    _format_ratio_for_display,
    _green_industry_summary,
    _invent_patent_records,
    _invent_patent_summary,
    _is_amount_label,
    _is_success_response,
    _join_counts,
    _join_present,
    _legal_person_office_summary,
    _looks_like_detail_record,
    _normalize_label,
    _outbound_investment_summary,
    _records_by_key,
    _records_by_path,
    _registered_paid_capital_value,
    _shareholder_summary,
    _small_micro_catalog_summary,
    _tax_ab_level_summary,
    _tax_rating_summary,
    _to_decimal,
    _truncate_chinese_text,
    _year_sort_key,
    _zhongzhishutong_qualification_summary,
)



DEFAULT_SUBMODULE_NAME = "综合信息"
HIDDEN_DEFAULT_SUBMODULE_MODULES = {"关键人", "关键人物", "触达路径", "资金需求", "风险提示", "股权关联", "财务分析", "行内合作情况", "访谈脚本", "营销方案", "AI综合智能分析建议"}
INTERFACE_CODE_RE = re.compile(r"(?<!\d)\d{10}(?!\d)")
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PRODUCT_RECOMMENDATION_PATH = PROJECT_ROOT / "develop_docs" / "【汇总】AI“访客助手”对公产品推荐内容.xlsx"
LONG_TEXT_FIELD_NAMES = {
    "经营范围",
    "主营范围",
    "国标行业",
    "变更记录",
    "有效专利（发明专利）数量",
    "企业认证",
    "税务评级",
    "股权结构图",
    "穿透查询股东关系",
    "股权变更（含变更频次、时间）",
    "企业对外投资",
    "股东旗下其他关联企业信息",
    "股东基础信息（含股东名称、股东类型、投资比例、投资数额、出资时间）",
    "股权穿透查询",
    "主要管理人员",
    "涉诉及各类风险信息",
    "法人、实控人跨企业任职信息",
    "名下关联企业、注册资本、经营状态",
    "关联任职图谱",
    "企业所属行业",
    "近一年贷款增量数据",
    "担保状态",
    "法律诉讼信息",
}


@dataclass
class OpinionReportField:
    """One field row from the opinion Excel."""

    module: str
    submodule: str
    name: str
    interface_source: str = ""
    remark: str = ""
    interface_codes: list[str] = field(default_factory=list)


@dataclass
class RiskDetailTable:
    """One optional risk-detail table from Sheet2."""

    category: str
    title: str
    headers: list[str]
    interface_codes: list[str] = field(default_factory=list)
    warning_levels: dict[str, str] = field(default_factory=dict)


@dataclass
class ProductRecommendation:
    """One product recommendation row from the visitor-assistant Excel."""

    department: str
    customer_feature: str
    product_name: str
    recommendation_points: str


USER_VALUE_KEYS = (("value",), ("text",), ("result",), ("content",))

_FALLBACK_TEXT_MUTED = "6B778C"


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
                r.font.color.rgb = RGBColor.from_string(_FALLBACK_TEXT_MUTED)
                _set_run_east_asia_font(r, FONT_NAME)
            break


class OpinionExcelDOCXRenderer:
    """Build a DOCX report from the opinion-collection Excel template.

    The Excel structure is expected to use:
    A=功能模块, B=二级模块, C=具体字段、内容, D=外数接口, E=备注.
    """

    def render_from_excel(
        self,
        excel_path: str | Path,
        output_path: str | Path,
        *,
        sheet_name: str = "Sheet1",
        company_name: str = "企业客户",
        interface_responses: Mapping[str, Any] | None = None,
        field_values: Mapping[str, Any] | None = None,
        annual_financial_table: Mapping[str, Any] | None = None,
        internal_info_section: Mapping[str, Any] | None = None,
        risk_detail_sheet_name: str = "Sheet2",
        product_recommendation_path: str | Path | None = DEFAULT_PRODUCT_RECOMMENDATION_PATH,
        disable_all_llm: bool = False,
        generated_at: datetime | None = None,
    ) -> Path:
        fields = self.load_fields(excel_path, sheet_name=sheet_name)
        risk_tables = self.load_risk_detail_tables(excel_path, sheet_name=risk_detail_sheet_name)
        return self.render(
            fields,
            output_path,
            company_name=company_name,
            interface_responses=interface_responses,
            field_values=field_values,
            annual_financial_table=annual_financial_table,
            internal_info_section=internal_info_section,
            risk_detail_tables=risk_tables,
            product_recommendation_path=product_recommendation_path,
            disable_all_llm=disable_all_llm,
            generated_at=generated_at,
        )

    def load_fields(self, excel_path: str | Path, *, sheet_name: str = "Sheet1") -> list[OpinionReportField]:
        workbook = load_workbook(excel_path, data_only=True)
        worksheet = workbook[sheet_name]
        fields: list[OpinionReportField] = []
        current_module = ""
        current_submodule_by_module: dict[str, str] = {}

        for row_index in range(2, worksheet.max_row + 1):
            module = normalize_text(_merged_value(worksheet, row_index, 1))
            submodule = normalize_text(_merged_value(worksheet, row_index, 2))
            field_name = normalize_text(_merged_value(worksheet, row_index, 3))
            interface_source = normalize_text(_merged_value(worksheet, row_index, 4))
            remark = normalize_text(_merged_value(worksheet, row_index, 5))

            if module:
                current_module = module
                current_submodule_by_module.setdefault(current_module, DEFAULT_SUBMODULE_NAME)
            if submodule and current_module:
                current_submodule_by_module[current_module] = submodule
            if not field_name:
                continue

            resolved_submodule = submodule or current_submodule_by_module.get(current_module) or DEFAULT_SUBMODULE_NAME
            fields.append(
                OpinionReportField(
                    module=current_module or "未分组",
                    submodule=resolved_submodule,
                    name=field_name,
                    interface_source=interface_source,
                    remark=remark,
                    interface_codes=INTERFACE_CODE_RE.findall(interface_source),
                )
            )
        return fields

    def load_risk_detail_tables(self, excel_path: str | Path, *, sheet_name: str = "Sheet2") -> list[RiskDetailTable]:
        workbook = load_workbook(excel_path, data_only=True)
        if sheet_name not in workbook.sheetnames:
            return []
        worksheet = workbook[sheet_name]
        tables: list[RiskDetailTable] = []
        category = ""
        row_index = 1
        while row_index <= worksheet.max_row:
            first_cell = normalize_text(_merged_value(worksheet, row_index, 1))
            row_values = [
                normalize_text(_merged_value(worksheet, row_index, column))
                for column in range(1, worksheet.max_column + 1)
            ]
            non_blank = [value for value in row_values if value]
            if not non_blank:
                row_index += 1
                continue
            if _is_risk_category_title(first_cell):
                category = first_cell
                row_index += 1
                continue
            if _is_risk_table_title(first_cell):
                header_row = _next_non_blank_row(worksheet, row_index + 1)
                if header_row:
                    headers = [
                        normalize_text(_merged_value(worksheet, header_row, column))
                        for column in range(1, worksheet.max_column + 1)
                    ]
                    headers = [header for header in headers if header]
                    if headers:
                        tables.append(
                            RiskDetailTable(
                                category=category,
                                title=first_cell,
                                headers=headers,
                                interface_codes=INTERFACE_CODE_RE.findall(first_cell),
                            )
                        )
                row_index += 1
                continue
            row_index += 1
        return tables

    def render(
        self,
        fields: Sequence[OpinionReportField],
        output_path: str | Path,
        *,
        company_name: str = "企业客户",
        interface_responses: Mapping[str, Any] | None = None,
        field_values: Mapping[str, Any] | None = None,
        annual_financial_table: Mapping[str, Any] | None = None,
        internal_info_section: Mapping[str, Any] | None = None,
        risk_detail_tables: Sequence[RiskDetailTable] | None = None,
        table_data: Mapping[str, list[dict[str, str]]] | None = None,
        product_recommendation_path: str | Path | None = DEFAULT_PRODUCT_RECOMMENDATION_PATH,
        disable_all_llm: bool = False,
        generated_at: datetime | None = None,
        explicit_module_filter: bool = False,
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

        field_module_names = {f.module for f in fields}

        for module_name, submodules in _group_fields(fields).items():
            section_title = module_name
            if module_name == "风险提示":
                section_title = _title_with_warning(module_name, _overall_risk_warning_level(risk_detail_tables or (), interface_responses))
            if not self._module_has_displayable_data(
                module_name,
                submodules,
                interface_responses=interface_responses,
                field_values=field_values,
                annual_financial_table=annual_financial_table,
                risk_detail_tables=risk_detail_tables or (),
            ):
                section_title = f"{section_title}（无数据）"
            source_count = self._source_count(module_name, submodules)
            self._add_section_heading(
                doc,
                section_title,
                source_count,
                SECTION_TONES.get(module_name, "315C7A"),
            )
            if section_title.endswith("（无数据）"):
                _add_fallback_paragraph_for_module(doc, module_name)
                if module_name == "行内合作情况" and internal_info_section and internal_info_section.get("internal_groups"):
                    self._add_internal_info(doc, internal_info_section)
                    doc.add_paragraph()
                elif module_name == "股权关联":
                    self._add_investor_detail_table(doc, self._get_table_records("investor_detail") or _investor_detail_records(interface_responses))
                    self._add_equity_change_detail_table(doc, _equity_change_detail_records(interface_responses))
                    self._add_outbound_investment_detail_table(doc, self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses))
                doc.add_paragraph()
                continue
            equity_management_details_added = False
            for submodule_name, rows in submodules.items():
                self._add_submodule_table(
                    doc,
                    module_name,
                    submodule_name,
                    rows,
                    interface_responses=interface_responses,
                    field_values=field_values,
                    annual_financial_table=annual_financial_table,
                )
                if module_name == "客户全貌" and submodule_name == "股权管理" and "股权关联" not in field_module_names:
                    self._add_investor_detail_table(doc, self._get_table_records("investor_detail") or _investor_detail_records(interface_responses))
                    self._add_equity_change_detail_table(doc, _equity_change_detail_records(interface_responses))
                    self._add_outbound_investment_detail_table(doc, self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses))
                    equity_management_details_added = True
                elif module_name == "股权关联":
                    self._add_investor_detail_table(doc, self._get_table_records("investor_detail") or _investor_detail_records(interface_responses))
                    self._add_equity_change_detail_table(doc, _equity_change_detail_records(interface_responses))
                    self._add_outbound_investment_detail_table(doc, self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses))
                    equity_management_details_added = True
            if module_name == "客户全貌" and not equity_management_details_added and "股权关联" not in field_module_names:
                self._add_investor_detail_table(doc, self._get_table_records("investor_detail") or _investor_detail_records(interface_responses))
                self._add_equity_change_detail_table(doc, _equity_change_detail_records(interface_responses))
                self._add_outbound_investment_detail_table(doc, self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses))
            if module_name == "行内合作情况" and internal_info_section and internal_info_section.get("internal_groups"):
                self._add_internal_info(doc, internal_info_section)
            if module_name == "风险提示":
                if risk_detail_tables:
                    self._add_risk_detail_tables(doc, risk_detail_tables, interface_responses)
            doc.add_paragraph()

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
                doc,
                company_name,
                fields,
                interface_responses,
                field_values,
                annual_financial_table,
                risk_detail_tables or (),
                product_recommendation_path,
            )

        doc.save(output)
        return output

    def _setup_document(self, doc: Document) -> None:
        section = doc.sections[0]
        section.top_margin = section.right_margin = section.bottom_margin = section.left_margin = Inches(1)

    def _setup_styles(self, doc: Document) -> None:
        normal = doc.styles["Normal"]
        normal.font.name = FONT_NAME
        normal.font.size = Pt(10.5)
        _set_east_asia_font(normal, FONT_NAME)
        normal.paragraph_format.space_after = Pt(5)
        normal.paragraph_format.line_spacing = 1.10

        for style_name, size, color in (
            ("Heading 1", 16, "2E74B5"),
            ("Heading 2", 12, "1F4D78"),
        ):
            style = doc.styles[style_name]
            style.font.name = FONT_NAME
            style.font.size = Pt(size)
            style.font.color.rgb = RGBColor.from_string(color)
            _set_east_asia_font(style, FONT_NAME)

    def _add_header(self, doc: Document, company_name: str, generated_at: datetime) -> None:
        title = doc.add_paragraph()
        title.paragraph_format.space_after = Pt(4)
        run = title.add_run(f"{company_name} 客户画像报告")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(20)
        run.font.color.rgb = RGBColor.from_string("16202A")
        _set_run_east_asia_font(run, FONT_NAME)

        meta = doc.add_table(rows=1, cols=2)
        _style_table(meta, widths=[1.2, 5.3], header=False)
        _cell_text(meta.rows[0].cells[0], "报告日期", bold=True, color="657180")
        _cell_text(meta.rows[0].cells[1], f"{generated_at:%Y-%m-%d}", color="16202A")
        doc.add_paragraph()

    def _add_section_heading(self, doc: Document, title: str, source_count: str, tone: str) -> None:
        table = doc.add_table(rows=1, cols=2)
        _style_table(table, widths=[0.12, 6.38], header=False)
        _set_cell_shading(table.rows[0].cells[0], tone)
        _set_cell_shading(table.rows[0].cells[1], "EEF3F7")
        _section_title_cell(table.rows[0].cells[1], title, source_count)

    def _add_profile_tags_rule(self, doc: Document) -> None:
        table = doc.add_table(rows=1, cols=1)
        _style_table(table, widths=[CONTENT_WIDTH_IN], header=False)
        cell = table.rows[0].cells[0]
        _set_cell_shading(cell, "E6DFA6")
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 0.2

    def _add_submodule_heading(self, doc: Document, text: str) -> None:
        heading = doc.add_paragraph()
        heading.paragraph_format.space_before = Pt(5)
        heading.paragraph_format.space_after = Pt(3)
        _add_colored_warning_runs(
            heading,
            normalize_text(text),
            base_color="1F4D78",
            size=11,
            bold=True,
        )

    def _add_submodule_table(
        self,
        doc: Document,
        module_name: str,
        submodule_name: str,
        rows: Sequence[OpinionReportField],
        *,
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
        annual_financial_table: Mapping[str, Any] | None,
    ) -> None:
        if not _should_hide_submodule_heading(module_name, submodule_name):
            self._add_submodule_heading(doc, submodule_name)

        if module_name == "经营能力" and submodule_name == "营收与财报":
            table_data = annual_financial_table or build_zhongshu_annual_financial_table(
                interface_responses,
                display_limit=3,
                fill_missing_years=False,
            )
            if table_data:
                self._add_annual_financial_table(doc, table_data)
                return

        resolved_rows = _filter_displayable_field_rows(
            rows,
            [
                (item, self._resolve_field_value(item, interface_responses, field_values))
                for item in rows
            ],
        )
        relationship_graph_paths = self._render_relationship_graph_if_needed(rows, interface_responses)
        has_change_field = any(item.name == "变更记录" for item in rows)
        if has_change_field:
            change_records = self._get_table_records("change_records") or _change_records(interface_responses.get("9000000004"))
        else:
            change_records = []
        has_bid_field = any(item.name == "企业招投标信息" for item in rows)
        if has_bid_field:
            bid_records = self._get_table_records("bid_records") or _bid_records(interface_responses)
        else:
            bid_records = []
        bid_field_names = {"企业招投标信息"} if bid_records else set()
        change_field_names = {"变更记录"} if change_records else set()
        graph_field_names = {"股权结构图"} if relationship_graph_paths else set()
        table_exclude_names = bid_field_names | change_field_names | graph_field_names | {"动产抵押", "股权出质", "融资情况"}
        table_rows = [
            (item, value)
            for item, value in resolved_rows
            if item.name not in table_exclude_names
        ]

        if table_rows:
            table_row_count = _layout_row_count(table_rows)
            table = doc.add_table(rows=max(table_row_count, 1), cols=4)
            _style_table(table, widths=[1.25, 2.0, 1.25, 2.0], header=False)

            row_index = 0
            pending_pair: list[tuple[OpinionReportField, str]] = []
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

        for image_path in relationship_graph_paths:
            self._add_relationship_graph_image(doc, image_path)

        if change_records:
                    self._add_change_detail_table(doc, change_records[:3], total_count=len(change_records))
        if bid_records:
            self._add_bid_detail_table(doc, bid_records[:3], total_count=len(bid_records))

        has_mortgage_field = any(item.name == "动产抵押" for item in rows)
        if has_mortgage_field:
            mortgage_records = self._get_table_records("mortgage_records")
            if mortgage_records:
                self._add_mortgage_detail_table(doc, mortgage_records[:3], total_count=len(mortgage_records))
        has_pledge_field = any(item.name == "股权出质" for item in rows)
        if has_pledge_field:
            pledge_records = self._get_table_records("pledge_records")
            if pledge_records:
                self._add_pledge_detail_table(doc, pledge_records[:3], total_count=len(pledge_records))
        has_financing_field = any(item.name == "融资情况" for item in rows)
        if has_financing_field:
            financing_records = self._get_table_records("financing_records")
            if financing_records:
                self._add_financing_detail_table(doc, financing_records[:3], total_count=len(financing_records))

    def _add_annual_financial_table(self, doc: Document, table_data: Mapping[str, Any]) -> None:
        columns = [col for col in (table_data.get("columns") or []) if isinstance(col, Mapping)][:3]
        groups = [group for group in (table_data.get("groups") or []) if isinstance(group, Mapping)]
        if not columns or not groups:
            return

        total_rows = 1 + sum(1 + len(group.get("rows") or []) for group in groups)
        table = doc.add_table(rows=total_rows, cols=len(columns) + 1)
        widths = [1.35] + [(CONTENT_WIDTH_IN - 1.35) / len(columns)] * len(columns)
        _style_table(table, widths=widths, header=False)

        _set_cell_shading(table.rows[0].cells[0], "EEF3F7")
        _cell_text(table.rows[0].cells[0], "", size=8)
        for index, col in enumerate(columns, start=1):
            _set_cell_shading(table.rows[0].cells[index], "EEF3F7")
            label = normalize_text(col.get("label"))
            unit = normalize_text(col.get("unit") or "")
            _cell_text(
                table.rows[0].cells[index],
                f"{label}\n{unit}" if unit else label,
                size=8.5,
                align=WD_ALIGN_PARAGRAPH.CENTER,
            )

        row_index = 1
        for group in groups:
            cells = table.rows[row_index].cells
            cells[0].merge(cells[-1])
            _cell_text(cells[0], group.get("title") or "-", bold=True, size=8.8)
            _set_cell_shading(cells[0], "FFFFFF")
            row_index += 1

            for item in group.get("rows") or []:
                if not isinstance(item, Mapping):
                    continue
                cells = table.rows[row_index].cells
                fill = "F7FAFC" if row_index % 2 == 0 else "FFFFFF"
                for cell in cells:
                    _set_cell_shading(cell, fill)
                _cell_text(cells[0], item.get("label") or "-", size=8.2)
                values = item.get("values") or {}
                for col_index, col in enumerate(columns, start=1):
                    key = col.get("key")
                    value = values.get(key) if key in values else "-"
                    _cell_text(
                        cells[col_index],
                        value,
                        size=8.2,
                        align=WD_ALIGN_PARAGRAPH.CENTER,
                        blank_as_dash=False,
                    )
                row_index += 1

    def _add_internal_info(self, doc: Document, data: Mapping[str, Any]) -> None:
        for group in data.get("internal_groups") or []:
            if not isinstance(group, Mapping):
                continue
            heading = doc.add_paragraph()
            heading.paragraph_format.space_before = Pt(5)
            heading.paragraph_format.space_after = Pt(3)
            run = heading.add_run(normalize_text(group.get("title") or "行内合作情况"))
            run.bold = True
            run.font.name = FONT_NAME
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor.from_string("1F4D78")
            _set_run_east_asia_font(run, FONT_NAME)
            for entry in group.get("items") or []:
                if isinstance(entry, Mapping):
                    self._add_internal_entry_table(doc, entry)

    def _add_internal_entry_table(self, doc: Document, entry: Mapping[str, Any]) -> None:
        columns = [col for col in (entry.get("columns") or []) if isinstance(col, Mapping)]
        rows = [row for row in (entry.get("rows") or []) if isinstance(row, Mapping)]
        if not columns or not rows:
            return
        self._add_submodule_heading(doc, entry.get("title") or entry.get("entityInfo") or "明细")
        table = doc.add_table(rows=len(rows) + 1, cols=len(columns))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(columns)] * len(columns), header=True)
        for cell, col in zip(table.rows[0].cells, columns):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, col.get("name") or col.get("field") or "-", bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, item in enumerate(rows, start=1):
            for cell, col in zip(table.rows[row_index].cells, columns):
                fill = "FFFFFF" if row_index % 2 == 1 else "F7FAFC"
                _set_cell_shading(cell, fill)
                _cell_text(
                    cell,
                    item.get(col.get("field")) or "-",
                    size=7.2,
                    color="26313C",
                    align=WD_ALIGN_PARAGRAPH.CENTER,
                    blank_as_dash=False,
                )

    def _write_pair_row(self, cells, pairs: Sequence[tuple[OpinionReportField, str]]) -> None:
        for pair_index in range(2):
            label_cell = cells[pair_index * 2]
            value_cell = cells[pair_index * 2 + 1]
            if pair_index < len(pairs):
                item, value = pairs[pair_index]
                _set_cell_shading(label_cell, "F5F6F8")
                _cell_text(label_cell, item.name, bold=True, size=8.5)
                _cell_text(value_cell, value, size=8.3)
            else:
                _set_cell_shading(label_cell, "F5F6F8")
                _cell_text(label_cell, "-", bold=True, size=8.5)
                _cell_text(value_cell, "-", size=8.3)

    def _write_long_row(self, cells, item: OpinionReportField, value: str) -> None:
        _set_cell_shading(cells[0], "F5F6F8")
        _cell_text(cells[0], item.name, bold=True, size=8.5)
        merged = cells[1].merge(cells[3])
        _cell_text(merged, value, size=8.2)

    def _render_relationship_graph_if_needed(
        self,
        rows: Sequence[OpinionReportField],
        interface_responses: Mapping[str, Any],
    ) -> tuple[Path, ...]:
        return ()

    def _add_relationship_graph_image(self, doc: Document, image_path: Path) -> None:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(4)
        paragraph.paragraph_format.space_after = Pt(4)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.add_run().add_picture(str(image_path), width=Inches(CONTENT_WIDTH_IN))

    def _add_risk_detail_tables(
        self,
        doc: Document,
        tables: Sequence[RiskDetailTable],
        interface_responses: Mapping[str, Any],
    ) -> None:
        current_category = ""
        visible_index = 0
        for table_def in tables:
            records = _risk_records_for_table(table_def, interface_responses, self._table_data)
            if not records:
                continue
            if table_def.category != current_category:
                current_category = table_def.category
                visible_index = 0
                paragraph = doc.add_paragraph()
                paragraph.paragraph_format.space_before = Pt(8)
                paragraph.paragraph_format.space_after = Pt(4)
                run = paragraph.add_run(_strip_number_prefix(current_category))
                run.bold = True
                run.font.name = FONT_NAME
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor.from_string("A43C3C")
                _set_run_east_asia_font(run, FONT_NAME)
            visible_index += 1
            display_table = RiskDetailTable(
                category=table_def.category,
                title=_title_with_warning(
                    _renumber_risk_table_title(table_def.title, visible_index),
                    _risk_table_warning_level(table_def, interface_responses),
                ),
                headers=table_def.headers,
                interface_codes=table_def.interface_codes,
                warning_levels=table_def.warning_levels,
            )
            self._add_risk_detail_table(doc, display_table, records[:3])

    def _add_risk_detail_table(
        self,
        doc: Document,
        table_def: RiskDetailTable,
        records: Sequence[Mapping[str, Any]],
    ) -> None:
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        _add_colored_warning_runs(
            title,
            table_def.title,
            base_color="26313C",
            size=9.5,
            bold=True,
        )

        headers = table_def.headers[:6]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            for cell, header in zip(table.rows[row_index].cells, headers):
                _cell_text(cell, _risk_record_value(record, header), size=7.1)

    def _add_bid_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, Any]],
        *,
        total_count: int | None = None,
    ) -> None:
        count = total_count if total_count is not None else len(records)
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"企业招投标信息（共{count}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = BID_DETAIL_HEADERS
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            for cell, header in zip(table.rows[row_index].cells, headers):
                _cell_text(cell, _bid_record_value(record, header), size=7.1)

    def _add_mortgage_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, Any]],
        *,
        total_count: int | None = None,
    ) -> None:
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"动产抵押（共{count}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)
        headers = ["登记编号", "被担保债权数额", "被担保债券种类", "担保范围", "履行债务结束日期", "债务数额", "抵押物名称", "抵押权人名称"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            for cell, header in zip(table.rows[row_index].cells, headers):
                _cell_text(cell, record.get(header) or "-", size=7.1)

    def _add_pledge_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, Any]],
        *,
        total_count: int | None = None,
    ) -> None:
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"股权出质（共{count}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)
        headers = ["登记编号", "出质人名称", "出质状态", "出质股权数额", "出质登记日期"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            for cell, header in zip(table.rows[row_index].cells, headers):
                _cell_text(cell, record.get(header) or "-", size=7.1)

    def _add_financing_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, Any]],
        *,
        total_count: int | None = None,
    ) -> None:
        if not records:
            return
        count = total_count if total_count is not None else len(records)
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"历史融资金额（共{count}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)
        headers = ["融资事件全称", "投资金额", "币种中文", "投资金额级别", "融资方名称全称", "投资方名称全称", "投资时间"]
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN / len(headers)] * len(headers), header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            for cell, header in zip(table.rows[row_index].cells, headers):
                _cell_text(cell, record.get(header) or "-", size=7.1)

    def _add_change_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, str]],
        *,
        total_count: int,
    ) -> None:
        if not records:
            return
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"共{total_count}条变更记录（展示最近3条变更记录）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = ("变更日期", "变更事项", "变更前内容", "变更后内容")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(
            table,
            widths=[CONTENT_WIDTH_IN * 0.18, CONTENT_WIDTH_IN * 0.20, CONTENT_WIDTH_IN * 0.31, CONTENT_WIDTH_IN * 0.31],
            header=True,
        )
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            _cell_text(table.rows[row_index].cells[0], record.get("变更日期") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[1], record.get("变更事项") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[2], record.get("变更前内容") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[3], record.get("变更后内容") or "-", size=7.1)

    def _add_manager_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, str]],
    ) -> None:
        if not records:
            return
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"主要管理人员（共{len(records)}人）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = ("姓名", "职务")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN * 0.25, CONTENT_WIDTH_IN * 0.30], header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B")
        for row_index, record in enumerate(records, start=1):
            _cell_text(table.rows[row_index].cells[0], record.get("姓名") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[1], record.get("职务") or "-", size=7.1)

    def _add_reserved_key_people_submodule(self, doc: Document) -> None:
        heading = doc.add_paragraph()
        heading.paragraph_format.space_before = Pt(5)
        heading.paragraph_format.space_after = Pt(3)
        run = heading.add_run("受益所有人")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("1F4D78")
        _set_run_east_asia_font(run, FONT_NAME)

        table = doc.add_table(rows=1, cols=2)
        _style_table(table, widths=[1.25, CONTENT_WIDTH_IN - 1.25], header=False)
        _set_cell_shading(table.rows[0].cells[0], "F5F6F8")
        _cell_text(table.rows[0].cells[0], "预留字段", bold=True, size=8.5)
        _cell_text(table.rows[0].cells[1], "-", size=8.3)

    def _add_investor_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, str]],
    ) -> None:
        if not records:
            return
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"股东基础信息（共{len(records)}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = ("股东名称", "出资比例", "股东类型", "认缴出资额度", "出资日期")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(
            table,
            widths=[
                CONTENT_WIDTH_IN * 0.28,
                CONTENT_WIDTH_IN * 0.14,
                CONTENT_WIDTH_IN * 0.14,
                CONTENT_WIDTH_IN * 0.24,
                CONTENT_WIDTH_IN * 0.20,
            ],
            header=True,
        )
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            _cell_text(table.rows[row_index].cells[0], record.get("股东名称") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[1], record.get("出资比例") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[2], record.get("股东类型") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[3], record.get("认缴出资额度") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[4], record.get("出资日期") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _add_equity_change_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, str]],
    ) -> None:
        if not records:
            return
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"股权变更记录（共{len(records)}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = ("变更日期", "股东名称", "转让前股权比例", "转让后股权比例")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(
            table,
            widths=[CONTENT_WIDTH_IN * 0.18, CONTENT_WIDTH_IN * 0.38, CONTENT_WIDTH_IN * 0.22, CONTENT_WIDTH_IN * 0.22],
            header=True,
        )
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            _cell_text(table.rows[row_index].cells[0], record.get("变更日期") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[1], record.get("股东名称") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[2], record.get("转让前股权比例") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[3], record.get("转让后股权比例") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _add_outbound_investment_detail_table(
        self,
        doc: Document,
        records: Sequence[Mapping[str, str]],
    ) -> None:
        if not records:
            return
        title = doc.add_paragraph()
        title.paragraph_format.space_before = Pt(4)
        title.paragraph_format.space_after = Pt(2)
        run = title.add_run(f"对外投资（共{len(records)}条）")
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor.from_string("26313C")
        _set_run_east_asia_font(run, FONT_NAME)

        headers = ("公司名称", "占比", "认缴金额", "企业状态")
        table = doc.add_table(rows=len(records) + 1, cols=len(headers))
        _style_table(table, widths=[CONTENT_WIDTH_IN * 0.44, CONTENT_WIDTH_IN * 0.16, CONTENT_WIDTH_IN * 0.22, CONTENT_WIDTH_IN * 0.18], header=True)
        for cell, header in zip(table.rows[0].cells, headers):
            _set_cell_shading(cell, "EAF1F7")
            _cell_text(cell, header, bold=True, size=7.4, color="1F5F8B", align=WD_ALIGN_PARAGRAPH.CENTER)
        for row_index, record in enumerate(records, start=1):
            _cell_text(table.rows[row_index].cells[0], record.get("公司名称") or "-", size=7.1)
            _cell_text(table.rows[row_index].cells[1], record.get("占比") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[2], record.get("认缴金额") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            _cell_text(table.rows[row_index].cells[3], record.get("企业状态") or "-", size=7.1, align=WD_ALIGN_PARAGRAPH.CENTER)

    def _add_ai_analysis_section(
        self,
        doc: Document,
        company_name: str,
        fields: Sequence[OpinionReportField],
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
        annual_financial_table: Mapping[str, Any] | None,
        risk_detail_tables: Sequence[RiskDetailTable],
        product_recommendation_path: str | Path | None,
    ) -> None:
        self._add_section_heading(doc, "访谈脚本", "-", SECTION_TONES.get("访谈脚本", "315C7A"))
        self._add_ai_interview_table(doc)
        self._add_section_heading(doc, "营销方案", "-", SECTION_TONES.get("营销方案", "315C7A"))
        self._add_ai_marketing_table(
            doc,
            company_name,
            fields,
            interface_responses,
            field_values,
            annual_financial_table,
            risk_detail_tables,
            product_recommendation_path,
        )
        self._add_section_heading(doc, "AI智能分析", "-", SECTION_TONES.get("AI综合智能分析建议", "315C7A"))
        self._add_company_overall_evaluation(
            doc,
            company_name,
            fields,
            interface_responses,
            field_values,
            annual_financial_table,
            risk_detail_tables,
        )

    def _add_company_overall_evaluation(
        self,
        doc: Document,
        company_name: str,
        fields: Sequence[OpinionReportField],
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
        annual_financial_table: Mapping[str, Any] | None,
        risk_detail_tables: Sequence[RiskDetailTable],
    ) -> None:
        items = generate_ai_analysis_items(
            company_name,
            fields,
            interface_responses,
            field_values=field_values,
            annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )
        if not items:
            items = [
                {
                    "title": "数据支撑不足",
                    "summary": "现有接口数据不足以形成稳定的公司整体判断。",
                    "evidence": "基础、经营、财务及风险类接口未形成可交叉验证的核心信息。",
                    "attention": "客户经理需在拜访现场补充核实企业经营状态、主营业务、资金需求和风险事项。",
                }
            ]

        table = doc.add_table(rows=len(items[:6]), cols=1)
        _style_table(table, widths=[CONTENT_WIDTH_IN], header=False)
        for index, item in enumerate(items[:6], start=1):
            cell = table.rows[index - 1].cells[0]
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.left_indent = Pt(0)
            paragraph.paragraph_format.first_line_indent = Pt(0)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.paragraph_format.line_spacing = 1.08

            prefix = paragraph.add_run(f"{index}. 【{item['title']}】：")
            prefix.bold = True
            prefix.font.name = FONT_NAME
            prefix.font.size = Pt(8.8)
            prefix.font.color.rgb = RGBColor.from_string("26313C")
            _set_run_east_asia_font(prefix, FONT_NAME)

            body = paragraph.add_run(
                f"{item['summary']} 数据依据：{item['evidence']} 拜访关注：{item['attention']}"
            )
            body.font.name = FONT_NAME
            body.font.size = Pt(8.5)
            body.font.color.rgb = RGBColor.from_string("26313C")
            _set_run_east_asia_font(body, FONT_NAME)
        _add_table_bottom_space(doc)

    def _add_ai_interview_table(self, doc: Document) -> None:
        table = doc.add_table(rows=1, cols=2)
        _style_table(table, widths=[1.25, 5.25], header=False)
        _set_cell_shading(table.rows[0].cells[0], "F5F6F8")
        _cell_text(table.rows[0].cells[0], "访谈脚本", bold=True, size=8.5)
        _cell_text(
            table.rows[0].cells[1],
            "待补充：后续可结合AI智能分析、风险提示和营销方案生成客户经理现场访谈问题清单。",
            size=8.5,
        )
        _add_table_bottom_space(doc)

    def _add_ai_marketing_table(
        self,
        doc: Document,
        company_name: str,
        fields: Sequence[OpinionReportField],
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
        annual_financial_table: Mapping[str, Any] | None,
        risk_detail_tables: Sequence[RiskDetailTable],
        product_recommendation_path: str | Path | None,
    ) -> None:
        recommendation = generate_marketing_recommendation(
            company_name,
            fields,
            interface_responses,
            field_values=field_values,
            annual_financial_table=annual_financial_table,
            risk_detail_tables=risk_detail_tables,
            product_recommendation_path=product_recommendation_path,
            disable_all_llm=getattr(self, "_disable_all_llm", False),
        )

        items = recommendation.get("products") or []
        if not items and recommendation.get("product"):
            items = [
                {
                    "product": recommendation["product"],
                    "reason": recommendation.get("reason") or "",
                    "source": recommendation.get("source") or "fallback",
                }
            ]

        rendered_count = 0
        for index, item in enumerate(items[:5], start=1):
            product = item.get("product")
            if not product:
                continue
            rendered_count += 1
            table = doc.add_table(rows=2, cols=4)
            _style_table(table, widths=[1.08, 1.72, 1.08, 2.62], header=False)
            _set_cell_shading(table.rows[0].cells[0], "F5F6F8")
            _cell_text(table.rows[0].cells[0], f"推荐产品{index}", bold=True, size=8.5)
            _cell_text(table.rows[0].cells[1], product.product_name, size=8.5)
            _set_cell_shading(table.rows[0].cells[2], "F5F6F8")
            _cell_text(table.rows[0].cells[2], "适用客群", bold=True, size=8.5)
            _cell_text(table.rows[0].cells[3], product.customer_feature or "-", size=8.5)
            _set_cell_shading(table.rows[1].cells[0], "F5F6F8")
            _cell_text(table.rows[1].cells[0], "推荐理由", bold=True, size=8.5)
            merged = table.rows[1].cells[1].merge(table.rows[1].cells[3])
            _cell_text(merged, item.get("reason") or "-", size=8.5)
        if rendered_count:
            _add_table_bottom_space(doc)

    def _resolve_field_value(
        self,
        item: OpinionReportField,
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
    ) -> str:
        explicit = _explicit_field_value(item, field_values)
        if explicit:
            return explicit

        derived = _derived_field_value(item, interface_responses)
        if derived:
            return derived

        for code in item.interface_codes:
            response = interface_responses.get(code)
            value = _find_value_by_label(response, item.name)
            if value:
                return _format_amount_for_display(value, item.name)
            count = _count_records(response)
            if count:
                return f"已获取 {count} 条记录"

        if item.interface_codes:
            if any(_is_success_response(response) for response in (interface_responses.get(code) for code in item.interface_codes) if isinstance(response, Mapping)):
                return "-"
            return "-"
        return "-"

    def _module_has_displayable_data(
        self,
        module_name: str,
        submodules: Mapping[str, Sequence[OpinionReportField]],
        *,
        interface_responses: Mapping[str, Any],
        field_values: Mapping[str, Any],
        annual_financial_table: Mapping[str, Any] | None,
        risk_detail_tables: Sequence[RiskDetailTable],
    ) -> bool:
        if module_name == "经营能力":
            table_data = annual_financial_table or build_zhongshu_annual_financial_table(
                interface_responses,
                display_limit=3,
                fill_missing_years=False,
            )
            if _has_annual_financial_data(table_data):
                return True
            if any(self._get_table_records(key) for key in ("bid_records", "mortgage_records", "pledge_records", "financing_records")):
                return True
        if module_name == "风险提示" and any(
            _risk_records_for_table(table_def, interface_responses)
            for table_def in risk_detail_tables
        ):
            return True

        if module_name == "股权关联":
            has_investor = bool(self._get_table_records("investor_detail") or _investor_detail_records(interface_responses))
            has_equity_change = bool(_equity_change_detail_records(interface_responses))
            has_outbound = bool(self._get_table_records("outbound_investment") or _outbound_investment_detail_records(interface_responses))
            if has_investor or has_equity_change or has_outbound:
                return True

        for rows in submodules.values():
            if any(item.name == "企业招投标信息" for item in rows) and _bid_records(interface_responses):
                return True
            for item in rows:
                value = self._resolve_field_value(item, interface_responses, field_values)
                if _is_displayable_value(value):
                    return True
        return False

    def _source_count(self, module_name: str, submodules: Mapping[str, Sequence[OpinionReportField]]) -> str:
        codes = {
            code
            for rows in submodules.values()
            for item in rows
            for code in item.interface_codes
        }
        if codes:
            return f"{len(codes)} 类接口"
        return SECTION_SOURCE_COUNTS.get(module_name, "-")

    def _get_table_records(self, key: str) -> list[dict[str, str]] | None:
        records = (self._table_data or {}).get(key)
        return records if records else None


def _group_fields(fields: Sequence[OpinionReportField]) -> OrderedDict[str, OrderedDict[str, list[OpinionReportField]]]:
    grouped: OrderedDict[str, OrderedDict[str, list[OpinionReportField]]] = OrderedDict()
    for item in fields:
        grouped.setdefault(item.module, OrderedDict()).setdefault(item.submodule, []).append(item)
    return grouped


def _should_hide_submodule_heading(module_name: str, submodule_name: str) -> bool:
    return module_name in HIDDEN_DEFAULT_SUBMODULE_MODULES and submodule_name == DEFAULT_SUBMODULE_NAME


def _merged_value(worksheet: Worksheet, row: int, column: int) -> Any:
    cell = worksheet.cell(row=row, column=column)
    if cell.value not in (None, ""):
        return cell.value
    for merged_range in worksheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return worksheet.cell(merged_range.min_row, merged_range.min_col).value
    return cell.value


def _next_non_blank_row(worksheet: Worksheet, start_row: int) -> int | None:
    for row_index in range(start_row, worksheet.max_row + 1):
        values = [
            normalize_text(_merged_value(worksheet, row_index, column))
            for column in range(1, worksheet.max_column + 1)
        ]
        if any(values):
            return row_index
    return None


def _is_risk_category_title(text: str) -> bool:
    return bool(re.match(r"^\d+、", text))


def _is_risk_table_title(text: str) -> bool:
    return bool(re.match(r"^\d+[）)]", text))


def _section_title_cell(cell, title: Any, source_count: Any) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.05
    _add_colored_warning_runs(
        paragraph,
        normalize_text(title) or "-",
        base_color="16202A",
        size=13,
        bold=True,
    )


def _add_small_block_heading(doc: Document, title: str) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(title)
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor.from_string("A43C3C")
    _set_run_east_asia_font(run, FONT_NAME)


WARNING_LABELS = {
    "green": "绿色预警",
    "yellow": "黄色预警",
    "red": "红色预警",
}
WARNING_COLORS = {
    "green": "178A45",
    "yellow": "D89A00",
    "red": "C62828",
}
WARNING_RANK = {"green": 1, "yellow": 2, "red": 3}


def _add_colored_warning_runs(
    paragraph,
    text: str,
    *,
    base_color: str,
    size: float,
    bold: bool,
) -> None:
    text = normalize_text(text) or "-"
    match = re.search(r"（(绿色预警|黄色预警|红色预警)）$", text)
    parts: list[tuple[str, str]] = []
    if match:
        parts.append((text[: match.start()], base_color))
        warning_text = match.group(0)
        warning_level = _warning_level_from_label(match.group(1))
        parts.append((warning_text, WARNING_COLORS.get(warning_level, base_color)))
    else:
        parts.append((text, base_color))

    for part, color in parts:
        if not part:
            continue
        run = paragraph.add_run(part)
        run.bold = bold
        run.font.name = FONT_NAME
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
        _set_run_east_asia_font(run, FONT_NAME)


def _title_with_warning(title: str, level: str) -> str:
    label = WARNING_LABELS.get(normalize_text(level))
    return f"{normalize_text(title)}（{label}）" if label else normalize_text(title)


def _warning_level_from_label(label: str) -> str:
    for level, text in WARNING_LABELS.items():
        if text == label:
            return level
    return ""


def _max_warning_level(levels: Sequence[str]) -> str:
    best = ""
    best_rank = 0
    for level in levels:
        normalized = normalize_text(level)
        rank = WARNING_RANK.get(normalized, 0)
        if rank > best_rank:
            best = normalized
            best_rank = rank
    return best


def _explicit_field_value(item: OpinionReportField, field_values: Mapping[str, Any]) -> str:
    keys = (
        item.name,
        f"{item.module}.{item.submodule}.{item.name}",
        f"{item.module}|{item.submodule}|{item.name}",
    )
    for key in keys:
        value = field_values.get(key)
        if value not in (None, ""):
            return normalize_text(value)
    return ""


def _derived_field_value(item: OpinionReportField, interface_responses: Mapping[str, Any]) -> str:
    name = normalize_text(item.name)
    profile = _basic_profile(interface_responses)
    base_record = _first_record_by_keys(interface_responses.get("9000000001"), ("EntpBscInfArry", "EntpBsicInfArry"))

    simple_values = {
        "企业名称": _first_value(base_record, ("EntpNm", "企业名称")),
        "统一社会信用代码": _first_value(base_record, ("USCC", "统一社会信用代码")),
        "法定代表人": _first_value(base_record, ("EntpLglRprsNm", "LglRprsNm", "法定代表人")),
        "注册 / 实缴资本": _registered_paid_capital_value(base_record),
        "注册 / 实缴资本（万元）": _registered_paid_capital_value(base_record),
        "工商核准日期": _first_value(base_record, ("EntpAprvDt", "核准日期")),
        "营业期限自": _first_value(base_record, ("CrpnTrmBeg", "营业期限自")),
        "营业期限至": _first_value(base_record, ("CrpnTrmTo", "营业期限至")),
        "注册地址": _first_value(base_record, ("Addr", "RgstAddrDstcNo", "注册地址")),
        "企业类型": _first_value(base_record, ("EntpTp", "企业类型")),
        "企业经营状态": _first_value(base_record, ("EntpCrpnSt", "企业经营状态")),
        "经营范围": _first_value(base_record, ("CrpnBsnScop", "经营范围")),
        "国标行业": _join_present(
            "",
            _first_value(base_record, ("IdyNm", "行业名称")),
            f"（{_first_value(base_record, ('IdyCd', '行业代码'))}）"
            if _first_value(base_record, ("IdyCd", "行业代码"))
            else "",
        ),
        "主营范围": _first_value(base_record, ("CrpnBsnScop", "经营范围")),
        "员工人数": _first_value(base_record, ("StfPplCnt", "员工人数")),
        "员工人数（人）": _first_value(base_record, ("StfPplCnt", "员工人数")),
        "企业规模": _find_value_by_label(interface_responses.get("9000000003"), "EntpScl") or _find_value_by_label(interface_responses.get("9000000003"), "企业规模"),
    }
    if name in simple_values and simple_values[name]:
        return simple_values[name]

    if name in {"社保人数", "社保人数（人）"}:
        value = _find_first_scalar_by_keys(interface_responses.get("9000000002"), ("TownStfBscPnsnInsunPIPC", "StfBscMdclInsunPIPC"))
        return value if value else ""
    if name == "近三年员工数量变化":
        staff = profile.get("员工人数")
        return f"员工规模保持稳定，当前约{staff}人" if staff else ""
    if name == "变更记录":
        return _change_summary(interface_responses.get("9000000004"))
    if name == "有效专利（发明专利）数量":
        return _invent_patent_summary(interface_responses.get("9000000008"))
    if name in {"中知数通科技力评分", "中知数通科技力评分/分层结果"}:
        score = _find_first_scalar_by_keys(interface_responses, ("Scor",))
        level = _find_first_scalar_by_keys(interface_responses, ("StrtfcnRslt",))
        if score and level:
            return f"中知数通科技力评分：{score}/{level}"
        if score:
            return f"中知数通科技力评分：{score}"
        if level:
            return f"中知数通科技力评分：{level}"
        return "-"
    if name == "科技资质情况":
        qualification_text = _zhongzhishutong_qualification_summary(interface_responses)
        if qualification_text:
            return qualification_text
        return "-"
    if name == "企业认证":
        certification_text = _enterprise_certification_summary(interface_responses.get("9000000003"))
        if certification_text:
            return certification_text
        return "-"
    if name == "是否绿色产业":
        green_industry_text = _green_industry_summary(interface_responses.get("9000000006"))
        if green_industry_text:
            return green_industry_text
        return "-"
    if name == "是否纳入小微企业目录":
        small_micro_text = _small_micro_catalog_summary(interface_responses.get("9000000006"))
        if small_micro_text:
            return small_micro_text
        return "-"
    if name == "信用评级":
        credit_rating_text = _credit_rating_summary_for_catalog(interface_responses.get("9000000006"))
        if credit_rating_text:
            return credit_rating_text
        return "-"
    if name == "荣誉榜单":
        billboard_text = _billboard_summary_for_catalog(interface_responses.get("9000000006"))
        if billboard_text:
            return billboard_text
        return "-"
    if name in {"产业赛道排名", "战略性新兴产业分类"}:
        return "-"
    if name == "新增新拓/存量/白名单/高风险等客户分类标签":
        return "样例客户：存量客户；风险标签需结合行内系统确认"
    if name in {"股权结构图", "穿透查询股东关系", "股权穿透查询"}:
        shareholder = _shareholder_summary(interface_responses.get("9000000011"))
        relation_count = _count_records(interface_responses.get("9000000012")) + _count_records(interface_responses.get("9000000013"))
        if shareholder and relation_count:
            return f"{shareholder}；关联/控制路径记录{relation_count}条"
        return shareholder
    if name == "股权变更（含变更频次、时间）":
        return _equity_change_summary(interface_responses.get("9000000015"))
    if name == "企业对外投资":
        return _outbound_investment_summary(interface_responses.get("9000000031"))
    if name == "股东旗下其他关联企业信息":
        count = _count_records(interface_responses.get("9000000032")) or _count_records(interface_responses.get("9000000034"))
        return f"关联企业/多节点关联记录{count}条" if count else ""
    if name == "主要管理人员":
        manager = _find_first_scalar_by_keys(interface_responses.get("9000000010"), ("Nm", "MbrNm", "PrsnNm"))
        position = _find_first_scalar_by_keys(interface_responses.get("9000000010"), ("Pos", "PosNm"))
        return _join_present("，", manager, position) or _count_text(interface_responses.get("9000000010"), "管理人员")
    if name == "股东基础信息（含股东名称、股东类型、投资比例、投资数额、出资时间）":
        return _shareholder_summary(interface_responses.get("9000000011"))
    if name == "涉诉及各类风险信息":
        counts = _risk_summary_counts(interface_responses)
        return "、".join(f"{label}{count}条" for label, count in counts.items()) if counts else ""
    if name == "法人、实控人跨企业任职信息":
        return _legal_person_office_summary(interface_responses)
    if name == "名下关联企业、注册资本、经营状态":
        return _legal_person_office_summary(interface_responses) or _outbound_investment_summary(interface_responses.get("9000000014"))
    if name == "关联任职图谱":
        count = _count_records(interface_responses.get("9000000016")) + _count_records(interface_responses.get("9000000017"))
        return f"关联任职/受益所有人图谱记录{count}条" if count else ""
    if name == "税务评级":
        return _tax_rating_summary(interface_responses.get("9000000005"))
    if name == "海关信用等级":
        return _customs_credit_summary(interface_responses.get("9000000007"))
    if name == "海关备案":
        return _customs_record_summary(interface_responses.get("9000000007"))
    if name == "纳税 A/B 级":
        return _tax_ab_level_summary(interface_responses.get("9000000005"))
    if name == "企业所属行业":
        return profile.get("行业") or ""
    if name == "企业招投标信息":
        return _count_text(interface_responses.get("9000000019"), "招投标")
    if name == "企业最高科技资质" or name == "中知数通专利、科技评分":
        patent_count = _count_records(interface_responses.get("9000000008"))
        return f"专利记录{patent_count}条" if patent_count else ""
    if name == "近一年贷款增量数据":
        return "-"
    if name == "动产抵押":
        return _collateral_summary(interface_responses, "动产抵押", ("9000000023", "9000000024", "9000000025", "9000000026", "9000000027", "9000000028"))
    if name == "股权出质":
        return _collateral_summary(interface_responses, "股权出质", ("9000000020", "9000000021", "9000000022"))
    if name == "融资情况":
        return _financing_status_summary(interface_responses.get("9000000029"))
    if name == "担保状态":
        count = _count_records(interface_responses.get("9000000018"))
        return f"年报担保/资产状况相关记录{count}条" if count else ""
    if name == "历史融资金额":
        amount = _find_first_scalar_by_keys(interface_responses.get("9000000029"), ("InvsAmt", "QtAmt", "InvsQtAmt"))
        return f"样例融资金额{_format_amount_for_display(amount, '融资金额')}" if amount else _count_text(interface_responses.get("9000000029"), "投融资")
    if name == "企业被执行人":
        count = _count_records(interface_responses.get("9000000039")) + _count_records(interface_responses.get("9000000040"))
        return f"被执行/限制消费相关记录{count}条" if count else ""
    if name == "异常经营":
        return _count_text(interface_responses.get("9000000051"), "经营异常")
    if name == "法律诉讼信息":
        counts = _risk_summary_counts(interface_responses)
        return "、".join(f"{label}{count}条" for label, count in counts.items()) if counts else ""
    if name == "案件进度":
        status = _find_first_scalar_by_keys(interface_responses, ("CaseSt", "PrfmSt", "AlrdyPrfmSt"))
        return status or ""
    if name == "外部失信":
        count = _risk_summary_counts(interface_responses).get("失信记录", 0)
        return f"失信相关记录{count}条" if count else ""
    if name == "行内逾期欠息":
        return "-"
    return ""


def _is_displayable_value(value: Any) -> bool:
    text = normalize_text(value)
    if not text or text == "-":
        return False
    if text in {"暂无", "暂无数据", "暂无记录", "无", "N/A", "n/a", "None"}:
        return False
    return True


def _filter_displayable_field_rows(
    rows: Sequence[OpinionReportField],
    resolved_rows: Sequence[tuple[OpinionReportField, str]],
) -> list[tuple[OpinionReportField, str]]:
    """去掉无值字段；未接入接口（无 interface_codes）且无数据的模板项也不展示."""
    filtered: list[tuple[OpinionReportField, str]] = []
    for item, value in resolved_rows:
        if not item.interface_codes and not _is_displayable_value(value):
            continue
        if not _is_displayable_value(value):
            continue
        filtered.append((item, value))
    return filtered


def _has_annual_financial_data(table_data: Mapping[str, Any] | None) -> bool:
    if not table_data:
        return False
    for group in table_data.get("groups") or []:
        if not isinstance(group, Mapping):
            continue
        for item in group.get("rows") or []:
            if not isinstance(item, Mapping):
                continue
            values = item.get("values") or {}
            if isinstance(values, Mapping) and any(_is_displayable_value(value) for value in values.values()):
                return True
    return False


def build_company_overall_evaluation_items(
    company_name: str,
    interface_responses: Mapping[str, Any],
    *,
    annual_financial_table: Mapping[str, Any] | None = None,
) -> list[dict[str, str]]:
    """Build a concise, evidence-backed business summary without calling an LLM."""
    items: list[dict[str, str]] = []
    profile = _basic_profile(interface_responses)
    financial_table = annual_financial_table or build_zhongshu_annual_financial_table(interface_responses)
    risk_counts = _risk_summary_counts(interface_responses)

    if profile:
        status = profile.get("经营状态") or "-"
        years = profile.get("成立年限")
        industry = profile.get("行业") or "-"
        capital = profile.get("注册资本")
        staff = profile.get("员工人数")
        scope = profile.get("经营范围") or ""
        evidence_parts = [
            f"经营状态为{status}",
            f"所属行业为{industry}",
        ]
        if years:
            evidence_parts.append(f"成立年限约{years}年")
        if capital:
            evidence_parts.append(f"注册资本{capital}")
        if staff:
            evidence_parts.append(f"员工人数{staff}人")
        if scope:
            evidence_parts.append(f"经营范围包含{_shorten(scope, 36)}")
        items.append(
            {
                "title": "主体经营基础较清晰" if status not in {"-", "吊销", "注销"} else "主体状态需核实",
                "summary": f"{company_name}基础身份信息可支撑初步主体判断。",
                "evidence": "、".join(evidence_parts) + "。",
                "attention": "拜访时重点核实当前实际经营地址、主营产品与工商登记信息是否一致。",
            }
        )

    financial_item = _financial_evaluation_item(financial_table)
    if financial_item:
        items.append(financial_item)

    change_count = _count_records(interface_responses.get("9000000004"))
    equity_change_count = _count_records(interface_responses.get("9000000015"))
    if change_count or equity_change_count:
        evidence = []
        if change_count:
            evidence.append(f"工商变更记录{change_count}条")
        if equity_change_count:
            evidence.append(f"股权变更记录{equity_change_count}条")
        items.append(
            {
                "title": "变更监测存在",
                "summary": "系统已识别企业工商或股权变更线索，需结合变更事项判断经营稳定性。",
                "evidence": "、".join(evidence) + "。",
                "attention": "客户经理需关注近期是否存在法定代表人、股东、注册资本、经营范围等关键事项变化。",
            }
        )

    shareholder_count = _count_records(interface_responses.get("9000000011"))
    outbound_invest_count = _count_records(interface_responses.get("9000000031"))
    legal_invest_count = _count_records(interface_responses.get("9000000014"))
    if shareholder_count or outbound_invest_count or legal_invest_count:
        evidence = []
        if shareholder_count:
            evidence.append(f"股东信息{shareholder_count}条")
        if outbound_invest_count:
            evidence.append(f"企业对外投资{outbound_invest_count}条")
        if legal_invest_count:
            evidence.append(f"法人对外投资或任职{legal_invest_count}条")
        items.append(
            {
                "title": "关联关系具备延展线索",
                "summary": "股东、法人任职和对外投资信息可作为客户关系挖掘入口。",
                "evidence": "、".join(evidence) + "。",
                "attention": "拜访时可围绕股东背景、关联企业结算往来、上下游协同和集团授信边界进行确认。",
            }
        )

    funding_counts = _funding_summary_counts(interface_responses)
    if funding_counts:
        evidence = "、".join(f"{name}{count}条" for name, count in funding_counts.items()) + "。"
        items.append(
            {
                "title": "资金需求存在可跟进信号",
                "summary": "抵质押、投融资或担保类记录提示企业可能存在融资安排或资产担保需求。",
                "evidence": evidence,
                "attention": "客户经理可进一步了解现有融资余额、担保物状态、融资用途及新增授信需求。",
            }
        )

    risk_total = sum(risk_counts.values())
    if risk_total:
        evidence = "、".join(f"{name}{count}条" for name, count in risk_counts.items() if count) + "。"
        items.append(
            {
                "title": "外部风险需重点核验",
                "summary": "涉诉、失信、处罚或经营异常类记录已形成风险提示，需要在营销前完成事实核验。",
                "evidence": evidence,
                "attention": "拜访前建议核实案件状态、是否已履行、处罚整改情况以及是否影响准入和授信审批。",
            }
        )

    if len(items) < 3:
        positive_counts = _positive_operation_counts(interface_responses)
        if positive_counts:
            evidence = "、".join(f"{name}{count}条" for name, count in positive_counts.items()) + "。"
            items.append(
                {
                    "title": "经营活跃度具备补充证据",
                    "summary": "招投标、专利、商标、许可等记录可补充说明企业经营活动和技术资产情况。",
                    "evidence": evidence,
                    "attention": "拜访时可进一步确认核心产品、订单来源、技术壁垒及未来投入计划。",
                }
            )

    return items[:6]


def generate_ai_analysis_items(
    company_name: str,
    fields: Sequence[OpinionReportField],
    interface_responses: Mapping[str, Any],
    *,
    field_values: Mapping[str, Any] | None = None,
    annual_financial_table: Mapping[str, Any] | None = None,
    risk_detail_tables: Sequence[RiskDetailTable] = (),
    disable_all_llm: bool = False,
) -> list[dict[str, str]]:
    payload = build_ai_analysis_payload(
        company_name,
        fields,
        interface_responses,
        field_values=field_values or {},
        annual_financial_table=annual_financial_table,
        risk_detail_tables=risk_detail_tables,
    )
    items = [] if disable_all_llm else generate_ai_analysis_items_with_llm(company_name, payload)
    if items:
        return items
    return build_company_overall_evaluation_items(
        company_name,
        interface_responses,
        annual_financial_table=annual_financial_table,
    )


def build_ai_analysis_payload(
    company_name: str,
    fields: Sequence[OpinionReportField],
    interface_responses: Mapping[str, Any],
    *,
    field_values: Mapping[str, Any] | None = None,
    annual_financial_table: Mapping[str, Any] | None = None,
    risk_detail_tables: Sequence[RiskDetailTable] = (),
) -> dict[str, Any]:
    field_values = field_values or {}
    table_data = annual_financial_table or build_zhongshu_annual_financial_table(
        interface_responses,
        display_limit=3,
        fill_missing_years=False,
    )
    shown_fields, data_gaps = _compact_shown_fields(fields, interface_responses, field_values)
    return {
        "company_name": company_name,
        "profile": _compact_mapping(_basic_profile(interface_responses), max_items=8),
        "shown_fields": shown_fields,
        "financial": _compact_financial_table(table_data),
        "risk_summary": dict(_risk_summary_counts(interface_responses)),
        "risk_samples": _compact_risk_samples(risk_detail_tables, interface_responses),
        "funding_summary": dict(_funding_summary_counts(interface_responses)),
        "positive_operation_summary": dict(_positive_operation_counts(interface_responses)),
        "data_gaps": data_gaps[:20],
        "note": "本payload仅包含当前报告已展示或摘要化的数据，不包含CRM行内信息和原始接口大报文。",
    }


def generate_ai_analysis_items_with_llm(company_name: str, payload: Mapping[str, Any]) -> list[dict[str, str]]:
    try:
        from app.customer_report.llm_section_summarizer import CustomerReportLLMConfig, CustomerReportLLMSummarizer

        config = CustomerReportLLMConfig()
        if not config.api_key and config.provider != "private":
            print("[WARN] AI analysis skipped: LLM api_key is empty and provider is not private")
            return []
        summarizer = CustomerReportLLMSummarizer(config)
        summarizer.save_debug_payload(company_name, "AI智能分析", payload)
        messages = [
            {
                "role": "system",
                "content": (
                    "你是银行客户经理访前报告分析助手。请只基于用户提供的JSON数据判断，"
                    "不要编造未出现的数据。输出必须是JSON数组，长度3到6。"
                ),
            },
            {
                "role": "user",
                "content": (
                    "请基于以下紧凑报告数据，输出3-6条公司AI智能分析。"
                    "每条包含title、summary、evidence、attention四个字段。"
                    "title不超过18字；summary是一句话判断；evidence必须引用JSON中的具体数据；"
                    "attention写客户经理拜访关注点。只返回JSON数组，不要Markdown。"
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, separators=(",", ":"))},
        ]
        response = summarizer.client.chat.completions.create(
            model=config.model,
            messages=messages,
            temperature=config.temperature,
            max_tokens=config.max_tokens,
        )
        content = response.choices[0].message.content or ""
        items = _parse_ai_analysis_json_items(content)
        if items:
            summarizer.append_trace(company_name, "AI智能分析", payload, json.dumps(items, ensure_ascii=False, indent=2))
        else:
            print("[WARN] AI analysis returned unparsable JSON or fewer than 3 items; fallback to rule analysis")
        return items
    except Exception as exc:
        print(f"[WARN] AI analysis failed: {exc}")
        return []


def _parse_ai_analysis_json_items(content: str) -> list[dict[str, str]]:
    text = normalize_text(content)
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    start = text.find("[")
    end = text.rfind("]")
    if start < 0 or end <= start:
        return []
    try:
        parsed = json.loads(text[start : end + 1])
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []
    items: list[dict[str, str]] = []
    for raw in parsed:
        if not isinstance(raw, Mapping):
            continue
        item = {
            "title": _shorten(raw.get("title") or raw.get("标题") or "综合判断", 18),
            "summary": _shorten(raw.get("summary") or raw.get("判断") or "", 90),
            "evidence": _shorten(raw.get("evidence") or raw.get("数据依据") or "", 160),
            "attention": _shorten(raw.get("attention") or raw.get("拜访关注") or "", 140),
        }
        if item["summary"] and item["evidence"] and item["attention"]:
            items.append(item)
        if len(items) >= 6:
            break
    return items if len(items) >= 3 else []


def choose_product_recommendation(
    interface_responses: Mapping[str, Any],
    *,
    product_recommendation_path: str | Path | None = DEFAULT_PRODUCT_RECOMMENDATION_PATH,
) -> ProductRecommendation:
    products = choose_product_recommendations(
        interface_responses,
        product_recommendation_path=product_recommendation_path,
    )
    if products:
        return products[0]
    return ProductRecommendation(
        department="公司金融部（投资银行部）",
        customer_feature="各类客户适用",
        product_name="公司存款产品",
        recommendation_points="按照客户资金周转情况匹配单位活期存款、通知存款、协定存款、单位定期存款或单位大额存单。",
    )


def choose_product_recommendations(
    interface_responses: Mapping[str, Any],
    *,
    product_recommendation_path: str | Path | None = DEFAULT_PRODUCT_RECOMMENDATION_PATH,
) -> list[ProductRecommendation]:
    products = load_product_recommendations(product_recommendation_path)
    profile = _basic_profile(interface_responses)
    industry = normalize_text(profile.get("行业"))
    scope = normalize_text(profile.get("经营范围"))
    text = f"{industry} {scope}"
    funding_counts = _funding_summary_counts(interface_responses)
    positive_counts = _positive_operation_counts(interface_responses)
    risk_counts = _risk_summary_counts(interface_responses)

    preferred_names: list[str] = []
    if "科技" in text or positive_counts.get("专利"):
        preferred_names.extend(["广科e贷", "科技型企业技术改造贷款", "公司存款产品"])
    if any(word in text for word in ("制造", "设备", "智能装备", "工业", "研发")):
        preferred_names.extend(["科技型企业技术改造贷款", "南网电费证", "广园贷", "广科e贷"])
    if any(word in text for word in ("电费", "用电", "电力", "能源")):
        preferred_names.extend(["南网电费证", "公司存款产品"])
    if funding_counts:
        preferred_names.extend(["法人账户透支业务", "公司存款产品"])
    if risk_counts:
        preferred_names.extend(["公司存款产品", "公司理财产品"])
    preferred_names.extend(["公司存款产品", "公司理财产品"])

    selected: list[ProductRecommendation] = []
    seen: set[str] = set()
    for name in preferred_names:
        product = _find_product_by_name(products, name)
        if product and product.product_name not in seen:
            selected.append(product)
            seen.add(product.product_name)
        if len(selected) >= 5:
            return selected[:5]

    if len(selected) < 3:
        ranked_products = sorted(
            products,
            key=lambda product: _score_product_recommendation(product, text, funding_counts, positive_counts, risk_counts),
            reverse=True,
        )
        for product in ranked_products:
            if product.product_name in seen:
                continue
            selected.append(product)
            seen.add(product.product_name)
            if len(selected) >= 5:
                break

    if len(selected) < 3:
        for product in products:
            if product.product_name in seen:
                continue
            selected.append(product)
            seen.add(product.product_name)
            if len(selected) >= 3:
                break

    return selected[:5]


def load_product_recommendations(product_recommendation_path: str | Path | None) -> list[ProductRecommendation]:
    if not product_recommendation_path:
        return []
    path = Path(product_recommendation_path)
    if not path.exists():
        return []
    workbook = load_workbook(path, data_only=True)
    worksheet = workbook.active
    products: list[ProductRecommendation] = []
    for row_index in range(3, worksheet.max_row + 1):
        department = normalize_text(worksheet.cell(row_index, 2).value)
        customer_feature = normalize_text(worksheet.cell(row_index, 3).value)
        product_name = normalize_text(worksheet.cell(row_index, 4).value)
        recommendation_points = normalize_text(worksheet.cell(row_index, 5).value)
        if not product_name or product_name == "重点产品名称":
            continue
        products.append(
            ProductRecommendation(
                department=department,
                customer_feature=customer_feature,
                product_name=product_name,
                recommendation_points=recommendation_points,
            )
        )
    return products


def _score_product_recommendation(
    product: ProductRecommendation,
    company_text: str,
    funding_counts: Mapping[str, int],
    positive_counts: Mapping[str, int],
    risk_counts: Mapping[str, int],
) -> int:
    product_text = f"{product.product_name} {product.customer_feature} {product.recommendation_points}"
    score = 0
    if any(word in company_text for word in ("科技", "研发", "专利", "软件")) and any(word in product_text for word in ("科技", "专利", "信用融资")):
        score += 4
    if any(word in company_text for word in ("制造", "设备", "工业", "智能装备")) and any(word in product_text for word in ("设备", "技改", "制造", "电费")):
        score += 4
    if funding_counts and any(word in product_text for word in ("透支", "融资", "贷款", "授信", "周转")):
        score += 3
    if positive_counts and any(word in product_text for word in ("经营", "结算", "线上", "信用")):
        score += 2
    if risk_counts and any(word in product_text for word in ("存款", "理财", "结算", "监管")):
        score += 3
    if "各类客户适用" in product.customer_feature:
        score += 1
    return score


def generate_marketing_recommendation(
    company_name: str,
    fields: Sequence[OpinionReportField],
    interface_responses: Mapping[str, Any],
    *,
    field_values: Mapping[str, Any] | None = None,
    annual_financial_table: Mapping[str, Any] | None = None,
    risk_detail_tables: Sequence[RiskDetailTable] = (),
    product_recommendation_path: str | Path | None = DEFAULT_PRODUCT_RECOMMENDATION_PATH,
    disable_all_llm: bool = False,
) -> dict[str, Any]:
    products = load_product_recommendations(product_recommendation_path)
    report_payload = build_ai_analysis_payload(
        company_name,
        fields,
        interface_responses,
        field_values=field_values or {},
        annual_financial_table=annual_financial_table,
        risk_detail_tables=risk_detail_tables,
    )
    llm_result = None if disable_all_llm else generate_marketing_recommendation_with_llm(
        company_name,
        report_payload,
        products,
    )
    if llm_result:
        return {
            "products": llm_result,
            "product": llm_result[0]["product"],
            "reason": llm_result[0]["reason"],
            "source": "llm",
        }

    selected_products = choose_product_recommendations(
        interface_responses,
        product_recommendation_path=product_recommendation_path,
    )
    items = [
        {
            "product": product,
            "reason": build_marketing_recommendation_text(
                company_name,
                interface_responses,
                annual_financial_table=annual_financial_table,
                product=product,
            ),
            "source": "fallback",
        }
        for product in selected_products
    ]
    return {
        "products": items,
        "product": items[0]["product"] if items else ProductRecommendation(
            department="公司金融部（投资银行部）",
            customer_feature="各类客户适用",
            product_name="公司存款产品",
            recommendation_points="按照客户资金周转情况匹配单位活期存款、通知存款、协定存款、单位定期存款或单位大额存单。",
        ),
        "reason": items[0]["reason"] if items else "",
        "source": "fallback",
    }


def generate_marketing_recommendation_with_llm(
    company_name: str,
    report_payload: Mapping[str, Any],
    products: Sequence[ProductRecommendation],
) -> list[dict[str, Any]] | None:
    if not products:
        print("[WARN] AI marketing recommendation skipped: product list is empty")
        return None
    try:
        from app.customer_report.llm_section_summarizer import CustomerReportLLMConfig, CustomerReportLLMSummarizer

        config = CustomerReportLLMConfig()
        if not config.api_key and config.provider != "private":
            print("[WARN] AI marketing recommendation skipped: LLM api_key is empty and provider is not private")
            return None
        payload = {
            "company_name": company_name,
            "report_payload": report_payload,
            "product_list": _product_list_payload(products),
        }
        summarizer = CustomerReportLLMSummarizer(config)
        summarizer.save_debug_payload(company_name, "营销方案", payload)
        messages = [
            {
                "role": "system",
                "content": (
                    "你是银行对公客户经理的营销方案助手。请只基于用户提供的企业报告JSON和产品清单JSON选择产品，"
                    "不得编造产品名称，不得引用输入中不存在的数据。"
                ),
            },
            {
                "role": "user",
                "content": (
                    "请从product_list中选择3-5个最适合该企业的产品，并按匹配度从高到低排序。"
                    "每个产品都要生成50-100个中文字符的推荐理由。"
                    "选择时结合企业行业、经营范围、科技属性、专利资质、财报表现、资金需求和风险情况。"
                    "如果企业存在明显风险，推荐理由需体现先核实风险或审慎推荐。"
                    "只返回JSON数组，数组长度3到5。每个元素格式为："
                    "{\"recommended_product\":\"产品名称\",\"reason\":\"50-100字推荐理由\"}"
                ),
            },
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False, separators=(",", ":"))},
        ]
        response = summarizer.client.chat.completions.create(
            model=config.model,
            messages=messages,
            temperature=config.temperature,
            max_tokens=min(config.max_tokens, 500),
        )
        content = response.choices[0].message.content or ""
        parsed = _parse_marketing_recommendation_json(content, products)
        if parsed:
            summarizer.append_trace(company_name, "营销方案", payload, json.dumps(parsed, ensure_ascii=False, indent=2))
            return parsed
        print("[WARN] AI marketing recommendation returned unparsable or unmatched product; fallback to rule recommendation")
        return None
    except Exception as exc:
        print(f"[WARN] AI marketing recommendation failed: {exc}")
        return None


def _product_list_payload(products: Sequence[ProductRecommendation]) -> list[dict[str, str]]:
    return [
        {
            "product_name": product.product_name,
            "customer_feature": _compact_text(product.customer_feature, limit=180),
            "recommendation_points": _compact_text(product.recommendation_points, limit=220),
        }
        for product in products
        if product.product_name
    ]


def _parse_marketing_recommendation_json(
    content: str,
    products: Sequence[ProductRecommendation],
) -> list[dict[str, Any]] | None:
    text = normalize_text(content)
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    start = text.find("[")
    end = text.rfind("]")
    if start < 0 or end <= start:
        return None
    try:
        parsed = json.loads(text[start : end + 1])
    except json.JSONDecodeError:
        return None
    if not isinstance(parsed, list):
        return None
    items: list[dict[str, Any]] = []
    seen: set[str] = set()
    for raw in parsed:
        if not isinstance(raw, Mapping):
            continue
        product_name = normalize_text(raw.get("recommended_product") or raw.get("product") or raw.get("产品名称"))
        reason = normalize_text(raw.get("reason") or raw.get("推荐理由"))
        if not product_name or not reason:
            continue
        product = _find_product_by_name(products, product_name)
        if not product or product.product_name in seen:
            continue
        items.append(
            {
                "product": product,
                "reason": _fit_length(reason, min_len=50, max_len=100),
                "source": "llm",
            }
        )
        seen.add(product.product_name)
        if len(items) >= 5:
            break
    return items if len(items) >= 3 else None


def build_marketing_recommendation_text(
    company_name: str,
    interface_responses: Mapping[str, Any],
    *,
    annual_financial_table: Mapping[str, Any] | None = None,
    product: ProductRecommendation | None = None,
) -> str:
    profile = _basic_profile(interface_responses)
    financial_table = annual_financial_table or build_zhongshu_annual_financial_table(interface_responses)
    product = product or choose_product_recommendation(interface_responses)
    industry = normalize_text(profile.get("行业")) or "中小微企业"
    revenue = _latest_metric_value(_financial_row_values(financial_table or {}, "营业总收入"))
    funding_counts = _funding_summary_counts(interface_responses)
    risk_counts = _risk_summary_counts(interface_responses)

    basis = f"{company_name}属于{industry}"
    if revenue:
        basis += f"，{revenue[0]}营业收入{revenue[1]}"
    if funding_counts:
        basis += "，存在抵质押或融资相关记录"
    elif risk_counts:
        basis += "，需先核实外部风险后审慎营销"
    else:
        basis += "，可从日常结算和资金管理切入"

    if product.product_name == "科技型企业技术改造贷款":
        text = f"{basis}。建议推荐{product.product_name}，围绕设备更新、数字化改造或绿色技改资金需求切入，匹配中长期授信方案。"
    elif product.product_name == "广科e贷":
        text = f"{basis}。建议推荐{product.product_name}，以科技资质、专利和经营流水为切入点，提供线上测额、快捷申请的信用融资方案。"
    elif product.product_name == "南网电费证":
        text = f"{basis}。建议推荐{product.product_name}，围绕电费缴纳和现金流缓释需求，提供表外开证、最长一年延付的低成本融资方案。"
    elif product.product_name == "法人账户透支业务":
        text = f"{basis}。建议推荐{product.product_name}，满足资金频繁周转和临时垫款需求，支持企业网银随借随还。"
    else:
        text = f"{basis}。建议推荐{product.product_name}，结合资金周转节奏配置活期、通知、协定或定期类产品，沉淀结算关系。"
    return _fit_length(text, min_len=50, max_len=100)


def _basic_profile(interface_responses: Mapping[str, Any]) -> dict[str, Any]:
    response = interface_responses.get("9000000001")
    record = _first_record_by_keys(response, ("EntpBscInfArry", "EntpBsicInfArry"))
    if not record:
        return {}
    return {
        "经营状态": _first_value(record, ("EntpCrpnSt", "CrpnSt", "经营状态")),
        "行业": _first_value(record, ("IdyNm", "行业名称", "国标行业")),
        "注册资本": _first_value(record, ("RgstCpitl", "PclupCpitl", "注册资本")),
        "员工人数": _first_value(record, ("StfPplCnt", "员工人数")),
        "成立年限": _company_age(_first_value(record, ("FoundDt", "成立日期"))),
        "经营范围": _first_value(record, ("CrpnBsnScop", "经营范围", "主营范围")),
    }


def _financial_evaluation_item(table_data: Mapping[str, Any] | None) -> dict[str, str] | None:
    if not table_data:
        return None
    revenue = _financial_row_values(table_data, "营业总收入")
    profit = _financial_row_values(table_data, "净利润")
    liability_rate = _financial_row_values(table_data, "资产负债率(%)")
    latest_revenue = _latest_metric_value(revenue)
    latest_profit = _latest_metric_value(profit)
    latest_liability_rate = _latest_metric_value(liability_rate)
    if not any((latest_revenue, latest_profit, latest_liability_rate)):
        return None

    evidence = []
    if latest_revenue:
        evidence.append(f"{latest_revenue[0]}营业总收入{latest_revenue[1]}")
    if latest_profit:
        evidence.append(f"{latest_profit[0]}净利润{latest_profit[1]}")
    if latest_liability_rate:
        evidence.append(f"{latest_liability_rate[0]}资产负债率{latest_liability_rate[1]}%")
    title = "经营财务有数据支撑"
    summary = "年报或财报图片解析结果已提供收入、利润或偿债能力指标，可用于判断经营质量。"
    if latest_profit and _to_float(latest_profit[1]) is not None and _to_float(latest_profit[1]) < 0:
        title = "盈利表现需重点关注"
        summary = "最新可用财务数据出现亏损信号，需结合收入变化和成本结构判断持续盈利能力。"
    return {
        "title": title,
        "summary": summary,
        "evidence": "、".join(evidence) + "。",
        "attention": "拜访时建议核实财报口径、收入确认方式、主要成本变化及近一年订单情况。",
    }


def _financial_row_values(table_data: Mapping[str, Any], row_label: str) -> Mapping[str, Any]:
    for group in table_data.get("groups") or []:
        if not isinstance(group, Mapping):
            continue
        for row in group.get("rows") or []:
            if isinstance(row, Mapping) and normalize_text(row.get("label")) == row_label:
                values = row.get("values") or {}
                return values if isinstance(values, Mapping) else {}
    return {}


def _latest_metric_value(values: Mapping[str, Any]) -> tuple[str, str] | None:
    for year in sorted((str(key) for key in values), reverse=True):
        value = normalize_text(values.get(year))
        if value and value != "-":
            return f"{year}年", value
    return None


def _compact_shown_fields(
    fields: Sequence[OpinionReportField],
    interface_responses: Mapping[str, Any],
    field_values: Mapping[str, Any],
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    shown: dict[str, dict[str, Any]] = OrderedDict()
    gaps: list[str] = []
    resolver = OpinionExcelDOCXRenderer()
    for item in fields:
        value = resolver._resolve_field_value(item, interface_responses, field_values)
        text = normalize_text(value)
        if not text or text == "-":
            continue
        key = f"{item.module}.{item.submodule}.{item.name}"
        if text in {"接口无数据", "接口有数据，样例无对应字段"}:
            gaps.append(key)
            continue
        module = shown.setdefault(item.module, OrderedDict())
        submodule = module.setdefault(item.submodule, OrderedDict())
        submodule[item.name] = _compact_text(text, limit=260 if _is_long_text_field(item, text) else 120)
    return shown, gaps


def _compact_financial_table(table_data: Mapping[str, Any] | None) -> dict[str, Any]:
    if not table_data:
        return {}
    columns = [col for col in (table_data.get("columns") or []) if isinstance(col, Mapping)][:3]
    if not columns:
        return {}
    wanted_rows = {
        "营业总收入",
        "同比(%)",
        "净利润",
        "资产总计",
        "负债总计",
        "所有者权益总计",
        "负债和股东权益总计",
        "净利率(%)",
        "资产负债率(%)",
    }
    output: dict[str, Any] = {
        "years": [col.get("label") or col.get("key") for col in columns],
        "metrics": OrderedDict(),
    }
    for group in table_data.get("groups") or []:
        if not isinstance(group, Mapping):
            continue
        for row in group.get("rows") or []:
            if not isinstance(row, Mapping):
                continue
            label = normalize_text(row.get("label"))
            if label not in wanted_rows:
                continue
            values = row.get("values") or {}
            metric_values = OrderedDict()
            for col in columns:
                key = col.get("key")
                value = normalize_text(values.get(key)) if isinstance(values, Mapping) else ""
                if value and value != "-":
                    metric_values[str(col.get("label") or key)] = value
            if metric_values:
                output["metrics"][label] = metric_values
    return output


def _compact_risk_samples(
    risk_detail_tables: Sequence[RiskDetailTable],
    interface_responses: Mapping[str, Any],
) -> dict[str, list[dict[str, Any]]]:
    output: dict[str, list[dict[str, Any]]] = OrderedDict()
    visible_index_by_category: dict[str, int] = {}
    for table_def in risk_detail_tables:
        records = _risk_records_for_table(table_def, interface_responses)
        if not records:
            continue
        category = _strip_number_prefix(table_def.category)
        visible_index_by_category[category] = visible_index_by_category.get(category, 0) + 1
        title = _renumber_risk_table_title(table_def.title, visible_index_by_category[category])
        samples = []
        for record in records[:3]:
            item = OrderedDict()
            for header in table_def.headers[:6]:
                value = _risk_record_value(record, header)
                if value and value != "-":
                    item[header] = _compact_text(value, limit=80)
            if item:
                samples.append(item)
        if samples:
            output.setdefault(category, []).append(
                {
                    "title": title,
                    "total_count": len(records),
                    "samples": samples,
                }
            )
    return output


def _compact_mapping(data: Mapping[str, Any], *, max_items: int = 12) -> dict[str, Any]:
    output = OrderedDict()
    for key, value in data.items():
        text = normalize_text(value)
        if not text or text == "-":
            continue
        output[str(key)] = _compact_text(text, limit=120)
        if len(output) >= max_items:
            break
    return output


def _compact_text(value: Any, *, limit: int) -> str:
    text = normalize_text(value)
    if len(text) <= limit:
        return text
    return text[:limit].rstrip("，,、；;。 ") + "...[已截断]"


def _risk_summary_counts(interface_responses: Mapping[str, Any]) -> OrderedDict[str, int]:
    groups = OrderedDict(
        (
            ("涉诉执行", ("9000000036", "9000000037", "9000000060", "9000000038", "9000000039", "9000000041", "9000000055")),
            ("失信记录", ("9000000040", "9000000042", "9000000043", "9000000044", "9000000045", "9000000046")),
            ("行政处罚", ("9000000047", "9000000048", "9000000049", "9000000050", "9000000054")),
            ("经营异常", ("9000000051",)),
            ("金融风险", ("9000000053",)),
        )
    )
    return OrderedDict(
        (name, sum(_count_records(interface_responses.get(code)) for code in codes))
        for name, codes in groups.items()
        if sum(_count_records(interface_responses.get(code)) for code in codes)
    )


def _funding_summary_counts(interface_responses: Mapping[str, Any]) -> OrderedDict[str, int]:
    groups = OrderedDict(
        (
            ("股权出质", ("9000000020", "9000000021", "9000000022")),
            ("动产抵押", ("9000000023", "9000000024", "9000000025", "9000000026", "9000000027", "9000000028")),
            ("投融资", ("9000000029",)),
        )
    )
    return OrderedDict(
        (name, sum(_count_records(interface_responses.get(code)) for code in codes))
        for name, codes in groups.items()
        if sum(_count_records(interface_responses.get(code)) for code in codes)
    )


def _positive_operation_counts(interface_responses: Mapping[str, Any]) -> OrderedDict[str, int]:
    groups = OrderedDict(
        (
            ("招投标", ("9000000019",)),
            ("专利", ("9000000008",)),
            ("商标", ("9000000072",)),
            ("公示许可", ("9000000070",)),
        )
    )
    return OrderedDict(
        (name, sum(_count_records(interface_responses.get(code)) for code in codes))
        for name, codes in groups.items()
        if sum(_count_records(interface_responses.get(code)) for code in codes)
    )


def _manager_detail_records(interface_responses: Mapping[str, Any]) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for record in _records_by_path(interface_responses.get("9000000010"), "BODY.QryRsltArry[].MainMgtPrsnArry[]"):
        name = _first_value(record, ("PrsnNm", "Nm", "MbrNm"))
        position = _first_value(record, ("Psitn", "Pos", "PosNm"))
        if not name and not position:
            continue
        records.append({"姓名": name or "-", "职务": position or "-"})
    return records


def _investor_detail_records(interface_responses: Mapping[str, Any]) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for record in _records_by_path(interface_responses.get("9000000011"), "BODY.QryRsltArry[].ShrhlrInfArry[]"):
        name = _first_value(record, ("ShrhlrNm", "InvsrNm", "EntpNm", "公司名称"))
        ratio = _format_ratio_for_display(_first_value(record, ("CtbRto", "InvsRto", "ShrhldgRto", "HoldRto")))
        shareholder_type = _first_value(record, ("ShrhlrTp", "股东类型"))
        amount = _format_amount_for_display(_first_value(record, ("SbcbCtbLmt", "InvsAmt", "SubConAm", "认缴出资额度", "投资数额")), "认缴出资额度")
        date = _first_value(record, ("CtbDt", "InvsDt", "出资日期", "出资时间"))
        if not name and not ratio and not shareholder_type and not amount and not date:
            continue
        records.append(
            {
                "股东名称": name or "-",
                "出资比例": ratio or "-",
                "股东类型": shareholder_type or "-",
                "认缴出资额度": amount or "-",
                "出资日期": date or "-",
            }
        )
    return records


def _equity_change_detail_records(interface_responses: Mapping[str, Any]) -> list[dict[str, str]]:
    payload = interface_responses.get("9000000015")
    if isinstance(payload, Mapping) and not _is_success_response(payload):
        return []
    records: list[dict[str, str]] = []
    for record in _records_by_path(payload, "BODY.QryRsltArry[].StkRghtMdfInfArry[]"):
        date = _first_value(record, ("StkRghtMdfDt",))
        shareholder_name = _first_value(record, ("ShrhlrNm",))
        before_ratio = _format_ratio_for_display(_first_value(record, ("TfrPrevsStkRghtRto",)))
        after_ratio = _format_ratio_for_display(_first_value(record, ("TfrAfStkRghtRto",)))
        if not any((date, shareholder_name, before_ratio, after_ratio)):
            continue
        records.append(
            {
                "变更日期": date or "-",
                "股东名称": shareholder_name or "-",
                "转让前股权比例": before_ratio or "-",
                "转让后股权比例": after_ratio or "-",
            }
        )
    records.sort(key=lambda item: _date_sort_key(item.get("变更日期", "")), reverse=True)
    return records[:3]


def _outbound_investment_detail_records(interface_responses: Mapping[str, Any]) -> list[dict[str, str]]:
    payload = interface_responses.get("9000000031")
    if isinstance(payload, Mapping) and not _is_success_response(payload):
        return []
    records: list[dict[str, str]] = []
    for record in _records_by_key(payload, "EntpExtInvsInfArry"):
        name = _first_value(record, ("EntpNm", "EntpBrchNm", "公司名称"))
        ratio = _format_ratio_for_display(_first_value(record, ("CtbRto", "InvsRto", "HoldRto", "占比")))
        amount = _format_amount_for_display(_first_value(record, ("SbcbCtbLmt", "SubConAm", "认缴出资额度", "认缴金额")), "认缴金额")
        status = _first_value(record, ("EntpSt", "企业状态"))
        if not name and not ratio and not amount and not status:
            continue
        records.append({"公司名称": name or "-", "占比": ratio or "-", "认缴金额": amount or "-", "企业状态": status or "-"})
    return records


def _financing_record_fields(record_type: str, record: Mapping[str, Any]) -> list[tuple[str, str]]:
    if record_type == "投资事件":
        return [
            ("披露日期", _first_value(record, ("InvsPublcDt", "PublcDt"))),
            ("融资事件全称", _first_value(record, ("InvsEvFullNm", "FincngEvFullNm"))),
            ("投资方名称", _first_value(record, ("InvsInvsrNmFullNm", "InvsrNmFullNm", "InvsInvsrNmAbbr", "InvsrNmAbbr"))),
            ("融资方全称", _first_value(record, ("InvsFincrNmFullNm", "FincrNmFullNm"))),
            ("投资时间", _first_value(record, ("InvsInvsTm", "InvsTm"))),
            ("投资金额", _format_amount_for_display(_first_value(record, ("InvsInvsAmt", "InvsAmt")), "投资金额")),
            ("币种", _first_value(record, ("InvsCcyCN", "CcyCN"))),
            ("投资金额级别", _first_value(record, ("InvsAmtLvl", "InvsInvsAmtLvl"))),
        ]
    return [
        ("披露日期", _first_value(record, ("PublcDt",))),
        ("融资事件全称", _first_value(record, ("FincngEvFullNm",))),
        ("投资方名称", _first_value(record, ("InvsrNmFullNm", "InvsrNmAbbr"))),
        ("融资方全称", _first_value(record, ("FincrNmFullNm", "FincrNmAbbr"))),
        ("投资时间", _first_value(record, ("InvsTm",))),
        ("投资金额", _format_amount_for_display(_first_value(record, ("InvsAmt",)), "投资金额")),
        ("币种", _first_value(record, ("CcyCN",))),
        ("投资金额级别", _first_value(record, ("InvsAmtLvl",))),
    ]


def _cn_number(value: int) -> str:
    return {
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
        10: "十",
    }.get(value, str(value))


def _company_age(found_date: str) -> int | None:
    match = re.search(r"(\d{4})", normalize_text(found_date))
    if not match:
        return None
    return max(0, datetime.now().year - int(match.group(1)))


def _shorten(text: str, limit: int) -> str:
    text = normalize_text(text)
    if len(text) <= limit:
        return text.rstrip("，,、；;。.")
    return text[:limit].rstrip("，,、；;。") + "..."


def _to_float(value: Any) -> float | None:
    text = normalize_text(value).replace(",", "").replace("%", "")
    try:
        return float(text)
    except ValueError:
        return None


def _find_product_by_name(products: Sequence[ProductRecommendation], name: str) -> ProductRecommendation | None:
    target = _normalize_label(name)
    for product in products:
        if _normalize_label(product.product_name) == target:
            return product
    for product in products:
        if target in _normalize_label(product.product_name):
            return product
    return None


def _fit_length(text: str, *, min_len: int, max_len: int) -> str:
    text = normalize_text(text)
    if len(text) <= max_len:
        return text
    trimmed = text[:max_len].rstrip("，,、；;。.")
    for delimiter in ("。", "；", "，", "、"):
        index = trimmed.rfind(delimiter)
        if index >= min_len:
            return trimmed[: index + 1].rstrip("，,、；;") + "。"
    return trimmed.rstrip("，,、；;") + "。"


def _overall_risk_warning_level(
    risk_detail_tables: Sequence[RiskDetailTable],
    interface_responses: Mapping[str, Any],
) -> str:
    hit_levels = _risk_hit_warning_levels(risk_detail_tables, interface_responses)
    if any(level == "red" for level in hit_levels):
        return "red"
    green_count = sum(1 for level in hit_levels if level == "green")
    if green_count >= 5:
        return "yellow"
    if green_count:
        return "green"
    return ""


def _risk_table_warning_level(
    table_def: RiskDetailTable,
    interface_responses: Mapping[str, Any],
) -> str:
    levels = [
        table_def.warning_levels.get(code, "")
        for code in _risk_hit_codes_for_table(table_def, interface_responses)
    ]
    return _max_warning_level(levels)


def _risk_hit_warning_levels(
    risk_detail_tables: Sequence[RiskDetailTable],
    interface_responses: Mapping[str, Any],
) -> list[str]:
    hit_levels_by_code: OrderedDict[str, str] = OrderedDict()
    for table_def in risk_detail_tables:
        for code in _risk_hit_codes_for_table(table_def, interface_responses):
            level = table_def.warning_levels.get(code)
            if level:
                hit_levels_by_code[code] = level
    return list(hit_levels_by_code.values())


def _risk_hit_codes_for_table(
    table_def: RiskDetailTable,
    interface_responses: Mapping[str, Any],
) -> list[str]:
    return [
        code
        for code in table_def.interface_codes
        if _risk_records_for_codes((code,), table_def.headers, interface_responses)
    ]


def _risk_records_for_table(
    table_def: RiskDetailTable,
    interface_responses: Mapping[str, Any],
    table_data: Mapping[str, list[dict[str, str]]] | None = None,
) -> list[Mapping[str, Any]]:
    if table_data:
        for code in table_def.interface_codes:
            key = _risk_interface_to_table_key(code)
            rows = table_data.get(key) if key else None
            if rows:
                return [dict(r) for r in rows]
    return _risk_records_for_codes(table_def.interface_codes, table_def.headers, interface_responses)


def _risk_interface_to_table_key(interface_code: str) -> str:
    _MAP: dict[str, str] = {
        "9000000041": "execution_risk",
        "9000000039": "consumption_limit",
        "9000000040": "discredit_executor",
    }
    return _MAP.get(interface_code, "")


def _risk_records_for_codes(
    interface_codes: Sequence[str],
    headers: Sequence[str],
    interface_responses: Mapping[str, Any],
) -> list[Mapping[str, Any]]:
    records: list[Mapping[str, Any]] = []
    for code in interface_codes:
        response = interface_responses.get(code)
        if not response:
            continue
        if isinstance(response, Mapping) and not _is_success_response(response):
            continue
        records.extend(_extract_record_dicts(response))
    return [_compact_record(record) for record in records if _record_has_any_value(record, headers)]


def _compact_record(record: Mapping[str, Any]) -> Mapping[str, Any]:
    return {
        str(key): value
        for key, value in record.items()
        if value not in (None, "") and not isinstance(value, (list, dict))
    }


def _record_has_any_value(record: Mapping[str, Any], headers: Sequence[str]) -> bool:
    return any(_risk_record_value(record, header) != "-" for header in headers)


def _risk_record_value(record: Mapping[str, Any], header: str) -> str:
    keys = RISK_HEADER_KEYS.get(_normalize_label(header), ())
    for key in keys:
        value = record.get(key)
        if value not in (None, ""):
            return _format_amount_for_display(value, header)
    found = _find_value_by_label(record, header)
    return _format_amount_for_display(found, header) if found else "-"


BID_DETAIL_HEADERS = ["标题", "公告类型", "招投标分类", "发布时间", "采购人", "招标代理机构"]


def _bid_records(interface_responses: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    response = interface_responses.get("9000000019")
    if not response:
        return []
    if isinstance(response, Mapping) and not _is_success_response(response):
        return []
    records = [_compact_record(record) for record in _extract_record_dicts(response)]
    return [record for record in records if _record_has_any_bid_value(record)]


def _record_has_any_bid_value(record: Mapping[str, Any]) -> bool:
    return any(_bid_record_value(record, header) != "-" for header in BID_DETAIL_HEADERS)


def _bid_record_value(record: Mapping[str, Any], header: str) -> str:
    keys = BID_HEADER_KEYS.get(_normalize_label(header), ())
    for key in keys:
        value = record.get(key)
        if value not in (None, ""):
            return normalize_text(value)
    return _find_value_by_label(record, header) or "-"


RISK_HEADER_KEYS: dict[str, tuple[str, ...]] = {
    _normalize_label("被执行人名称"): ("BeExecPrsnNm", "BeExecPrsnNm1", "BeExecNm", "EntpNm", "PlcFldNm"),
    _normalize_label("案件状态"): ("CaseSt", "PrfmSt", "AlrdyPrfmSt"),
    _normalize_label("案号"): ("ExecCaseNo", "ExecCaseNo1", "JdgntCaseNo"),
    _normalize_label("申请执行标的（元）"): ("AppExecSbjMtr", "BnExecTrgt"),
    _normalize_label("执行法院"): ("ExecCrtNm", "ExecCrt", "ExecCrtCd"),
    _normalize_label("立案时间"): ("CrtRgstDt", "FlngTm", "RelsDt"),
    _normalize_label("案由"): ("RgstrRsn", "Gnd"),
    _normalize_label("开庭日期"): ("BegJdgDt",),
    _normalize_label("当事人"): ("PartyNm", "AcusrNm", "DfedntNm"),
    _normalize_label("公告日期"): ("HitBidAncmDt", "RelsDt"),
    _normalize_label("公告类型"): ("AncmTp",),
    _normalize_label("公告人"): ("AncmPrsn",),
    _normalize_label("判决案号"): ("JdgntCaseNo", "ExecCaseNo"),
    _normalize_label("发布日期"): ("RelsDt",),
    _normalize_label("被执行人"): ("BeExecPrsnNm", "BeExecNm", "PlcFldNm"),
    _normalize_label("组织机构代码"): ("OIBC", "IdCardNoOrOIBC", "OrgBrchId"),
    _normalize_label("是否有效"): ("WhthrVld", "VldFlg", "RmveIndr"),
    _normalize_label("未履行金额"): ("NotPrfmAmt",),
    _normalize_label("终本日期"): ("ClsDt",),
    _normalize_label("被执行人的履行情况"): ("AlrdyPrfmSt", "NotPrfmSt", "PrfmSt"),
    _normalize_label("列入机关"): ("ColInDtrmInst", "ExecBssMkUnit"),
    _normalize_label("列入文号"): ("ExecBssPrefNo",),
    _normalize_label("数据日期"): ("DataDt",),
    _normalize_label("列入原因"): ("ColInRsn", "LsCrdtBhvrCdtn"),
    _normalize_label("企业中文名称"): ("EntpNm", "PlcFldNm"),
    _normalize_label("注册地址"): ("Addr",),
    _normalize_label("法人代表"): ("LglRprs", "LPNm"),
    _normalize_label("统一信用代码"): ("USCC", "IdCardNoOrOIBC"),
    _normalize_label("首次注册日期"): ("CrtRgstDt",),
    _normalize_label("企业名称"): ("EntpNm", "PlcFldNm", "BeExecPrsnNm"),
    _normalize_label("处罚日期"): ("PnlDt", "RelsDt", "DataDt"),
    _normalize_label("处罚依据"): ("PnlBss",),
    _normalize_label("单位"): ("PnlDcsnInst", "ExecBssMkUnit"),
    _normalize_label("处罚结果"): ("PnlRslt", "PnlCntnt", "PnlDcsn"),
    _normalize_label("纳税人名称"): ("TaxpyNm", "EntpNm"),
    _normalize_label("纳税人识别码"): ("TaxpyIdentNo", "USCC"),
    _normalize_label("违法事实及处罚情况"): ("LglInstmtCrntBsn", "PnlCntnt", "RskDtlExpln"),
    _normalize_label("认定日期或所属日期"): ("RelsDt", "DataDt"),
    _normalize_label("黑名单类型"): ("RskTp",),
    _normalize_label("管理或征管机关"): ("ExecBssMkUnit", "PnlDcsnInst"),
    _normalize_label("纳税人登记名称"): ("TaxpyNm", "EntpNm"),
    _normalize_label("纳税人识别号"): ("TaxpyIdentNo", "USCC"),
    _normalize_label("案件性质"): ("CaseNtr", "RskTp"),
    _normalize_label("处罚决定文书"): ("PnlDcsnDocNo", "ExecCaseNo"),
    _normalize_label("处罚状态"): ("CaseSt", "PnlSt"),
    _normalize_label("处罚机关"): ("PnlDcsnInst", "ExecBssMkUnit"),
    _normalize_label("决定书号"): ("PnlDcsnDocNo", "ExecCaseNo"),
    _normalize_label("处罚原因"): ("PnlRsn", "ColInRsn", "RskDtlExpln"),
    _normalize_label("清算负责人"): ("LqdtnPic", "LglRprs"),
    _normalize_label("清算组成员"): ("LqdtnMmbr",),
    _normalize_label("股权数额"): ("EqtyAmt",),
    _normalize_label("协助公示通知书文号"): ("AsstPubNtcNo",),
    _normalize_label("股权冻结状态"): ("EqtyFrzSt",),
    _normalize_label("币种"): ("Ccy",),
    _normalize_label("列入日期"): ("ColInDt",),
    _normalize_label("有效标志"): ("VldFlg",),
    _normalize_label("单位名称"): ("EntpNm", "PlcFldNm"),
    _normalize_label("纳入理由"): ("ColInRsn", "RskDtlExpln"),
    _normalize_label("失信行为简况"): ("LsCrdtBhvrCdtn", "RskDtlExpln"),
    _normalize_label("文书编号"): ("DocNo", "PnlDcsnDocNo", "ExecCaseNo"),
    _normalize_label("处罚机构"): ("PnlDcsnInst", "ExecBssMkUnit"),
    _normalize_label("被处罚单位"): ("EntpNm", "PlcFldNm"),
    _normalize_label("处罚事由"): ("PnlRsn", "RskDtlExpln"),
    _normalize_label("处罚金额（万元）"): ("PnlAmt",),
    _normalize_label("登记时间"): ("CrtRgstDt", "DataDt"),
    _normalize_label("公布时间"): ("RelsDt",),
    _normalize_label("罪名"): ("Chrg", "RskTp"),
    _normalize_label("判决作出机构"): ("JdgInst", "ExecCrtNm", "ExecBssMkUnit"),
}


BID_HEADER_KEYS: dict[str, tuple[str, ...]] = {
    _normalize_label("标题"): ("BidTtl", "BidNm", "BidName", "ProjNm", "Ttl", "Title"),
    _normalize_label("公告类型"): ("AncmTp", "AncmTpEncd", "公告类型"),
    _normalize_label("招投标分类"): ("BidClsf", "BidTp", "BidType", "招投标分类"),
    _normalize_label("发布时间"): ("IssuTm", "RelsDt", "PubDt", "发布时间", "发布日期"),
    _normalize_label("采购人"): ("PrchsPrsnNm", "PrchsPrsn", "Purchaser", "采购人"),
    _normalize_label("招标代理机构"): ("CallBidAgncBrch", "BidAgncNm", "Agency", "招标代理机构"),
}


def _strip_number_prefix(text: str) -> str:
    return re.sub(r"^\d+、\s*", "", normalize_text(text))


def _clean_risk_table_title(text: str) -> str:
    title = re.sub(r"^\d+[）)]\s*", "", normalize_text(text))
    title = re.sub(r"[（(][^）)]*\d{10}[^）)]*[）)]", "", title)
    return title.strip() or normalize_text(text)


def _renumber_risk_table_title(text: str, index: int) -> str:
    title = _clean_risk_table_title(text)
    return f"{index}）{title}"


def _add_table_bottom_space(doc: Document, points: float = 8) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(points)
    paragraph.paragraph_format.line_spacing = 1
    run = paragraph.add_run("")
    run.font.size = Pt(1)


def _is_long_text_field(item: OpinionReportField, value: str) -> bool:
    label = normalize_text(item.name)
    return label in LONG_TEXT_FIELD_NAMES or len(normalize_text(value)) > 80


def _layout_row_count(rows: Sequence[tuple[OpinionReportField, str]]) -> int:
    count = 0
    pending = 0
    for item, value in rows:
        if _is_long_text_field(item, value):
            if pending:
                count += 1
                pending = 0
            count += 1
            continue
        pending += 1
        if pending == 2:
            count += 1
            pending = 0
    if pending:
        count += 1
    return count
