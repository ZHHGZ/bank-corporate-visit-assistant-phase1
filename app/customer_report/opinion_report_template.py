"""Structured customer report template loading."""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Mapping

from app.customer_report.opinion_excel_docx_renderer import (
    INTERFACE_CODE_RE,
    OpinionExcelDOCXRenderer,
    OpinionReportField,
    RiskDetailTable,
)


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE_CONFIG_PATH = PROJECT_ROOT / "configs" / "customer_report_opinion_template.json"


@dataclass(frozen=True)
class OpinionReportTemplate:
    """Report structure that used to be read from the opinion Excel workbook."""

    fields: tuple[OpinionReportField, ...]
    risk_detail_tables: tuple[RiskDetailTable, ...]
    version: str = ""
    description: str = ""


def load_opinion_report_template(path: str | Path | None = None) -> OpinionReportTemplate:
    """Load the structured report template from JSON config."""
    config_path = Path(path) if path else DEFAULT_TEMPLATE_CONFIG_PATH
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(payload, Mapping):
        raise ValueError(f"Invalid opinion report template config: {config_path}")

    fields = tuple(_field_from_mapping(item) for item in _list_value(payload, "fields"))
    warning_levels = {
        str(code): str(level)
        for code, level in (payload.get("risk_warning_levels") or {}).items()
    } if isinstance(payload.get("risk_warning_levels"), Mapping) else {}
    risk_detail_tables = tuple(_risk_table_from_mapping(item, warning_levels) for item in _list_value(payload, "risk_detail_tables"))
    if not fields:
        raise ValueError(f"Opinion report template has no fields: {config_path}")

    return OpinionReportTemplate(
        fields=fields,
        risk_detail_tables=risk_detail_tables,
        version=str(payload.get("version") or ""),
        description=str(payload.get("description") or ""),
    )


def export_opinion_report_template_from_excel(
    excel_path: str | Path,
    output_path: str | Path = DEFAULT_TEMPLATE_CONFIG_PATH,
    *,
    version: str = "",
    risk_warning_levels: Mapping[str, str] | None = None,
) -> Path:
    """Regenerate JSON config from the source Excel when the field template changes."""
    renderer = OpinionExcelDOCXRenderer()
    fields = renderer.load_fields(excel_path)
    risk_detail_tables = renderer.load_risk_detail_tables(excel_path)
    payload = {
        "version": version,
        "source": Path(excel_path).name,
        "description": "客户画像报告字段和风险明细结构配置，由意见征集分析模板结构化生成。",
        "fields": [asdict(item) for item in fields],
        "risk_detail_tables": [asdict(item) for item in risk_detail_tables],
        "risk_warning_levels": dict(risk_warning_levels or {}),
    }
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output


def _field_from_mapping(item: Any) -> OpinionReportField:
    if not isinstance(item, Mapping):
        raise ValueError("Opinion report field config item must be an object")
    interface_source = str(item.get("interface_source") or "")
    interface_codes = item.get("interface_codes")
    if not isinstance(interface_codes, list):
        interface_codes = INTERFACE_CODE_RE.findall(interface_source)
    return OpinionReportField(
        module=str(item.get("module") or "未分组"),
        submodule=str(item.get("submodule") or "综合信息"),
        name=str(item.get("name") or ""),
        interface_source=interface_source,
        remark=str(item.get("remark") or ""),
        interface_codes=[str(code) for code in interface_codes if str(code)],
    )


def _risk_table_from_mapping(item: Any, warning_levels: Mapping[str, str] | None = None) -> RiskDetailTable:
    if not isinstance(item, Mapping):
        raise ValueError("Risk detail table config item must be an object")
    title = str(item.get("title") or "")
    interface_codes = item.get("interface_codes")
    if not isinstance(interface_codes, list):
        interface_codes = INTERFACE_CODE_RE.findall(title)
    headers = item.get("headers")
    resolved_codes = [str(code) for code in interface_codes if str(code)]
    return RiskDetailTable(
        category=str(item.get("category") or ""),
        title=title,
        headers=[str(header) for header in headers if str(header)] if isinstance(headers, list) else [],
        interface_codes=resolved_codes,
        warning_levels={
            code: str((warning_levels or {}).get(code) or "")
            for code in resolved_codes
            if (warning_levels or {}).get(code)
        },
    )


def _list_value(payload: Mapping[str, Any], key: str) -> list[Any]:
    value = payload.get(key)
    return value if isinstance(value, list) else []
