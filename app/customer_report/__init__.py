"""从 llm-standalone 等价迁移的正式客户画像报告渲染链路.

本包整体复用原项目 `src/customer_report` 的字段解析、模板加载和 DOCX 排版逻辑，
仅调整了 import 路径与环境变量读取方式，未改动字段映射与版式，确保生成的报告
结构、字段和样式与原项目完全一致。
"""
from app.customer_report.opinion_excel_card_docx_renderer import (
    OpinionExcelCardDOCXRenderer,
)
from app.customer_report.opinion_excel_docx_renderer import (
    OpinionExcelDOCXRenderer,
    OpinionReportField,
    RiskDetailTable,
)
from app.customer_report.opinion_report_template import (
    OpinionReportTemplate,
    load_opinion_report_template,
)

__all__ = [
    "OpinionExcelCardDOCXRenderer",
    "OpinionExcelDOCXRenderer",
    "OpinionReportField",
    "RiskDetailTable",
    "OpinionReportTemplate",
    "load_opinion_report_template",
]
