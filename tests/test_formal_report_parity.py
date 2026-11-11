"""正式报告等价性验收：迁移渲染器与 llm-standalone 原渲染器 DOCX manifest 零差异.

已废弃：新版架构 parser/renderer 分层后，旧版 Word 逐字对齐不再适用作为验收标准。
改为验证新版报告能否稳定生成、字段是否正确（见其他 field/layout 测试）。
"""
from __future__ import annotations

import pytest

pytest.skip("旧版 Word 对齐测试已废弃，新版架构不再适用", allow_module_level=True)

import sys
import zipfile
from datetime import datetime
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = APP_ROOT.parent
LLM_STANDALONE_ROOT = REPO_ROOT / "llm-standalone"

for path in (str(APP_ROOT), str(LLM_STANDALONE_ROOT)):
    if path not in sys.path:
        sys.path.insert(0, path)

FIXED_COMPANY = "等价性验收样例企业"
FIXED_GENERATED_AT = datetime(2026, 1, 2, 3, 4, 5)


def _required_codes() -> list[str]:
    from app.core.esb_client import report_required_codes

    return report_required_codes()


def build_fixture_responses() -> dict[str, object]:
    """构造固定输入：对所有正式报告交易码使用同一套 mock 响应."""
    from app.core.mock_data import build_mock_response

    codes = _required_codes()
    responses: dict[str, object] = {
        code: build_mock_response(code, FIXED_COMPANY) for code in codes
    }
    # 9000000060 走链式明细，固定为空列表以保证确定性
    responses["9000000060"] = []
    return responses


def _render(renderer_cls, template, output_path: Path, responses: dict[str, object]) -> Path:
    from app.customer_report.annual_financial_table import (
        build_zhongshu_annual_financial_table,
    )

    return renderer_cls().render(
        template.fields,
        output_path,
        company_name=FIXED_COMPANY,
        interface_responses=responses,
        field_values={},
        annual_financial_table=build_zhongshu_annual_financial_table(
            responses, display_limit=3, fill_missing_years=False
        ),
        internal_info_section=None,
        risk_detail_tables=template.risk_detail_tables,
        disable_all_llm=True,
        generated_at=FIXED_GENERATED_AT,
    )


def _document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def _structural_manifest(docx_path: Path) -> dict[str, object]:
    """提取与库版本无关的结构清单：段落文本序列 + 每张表的行列与单元格文本."""
    from docx import Document

    doc = Document(str(docx_path))
    paragraphs = [p.text for p in doc.paragraphs]
    tables = []
    for table in doc.tables:
        tables.append(
            {
                "rows": len(table.rows),
                "cols": len(table.columns),
                "cells": [[cell.text for cell in row.cells] for row in table.rows],
            }
        )
    return {"paragraphs": paragraphs, "tables": tables}


def _first_manifest_diff(a: dict[str, object], b: dict[str, object]) -> str:
    pa, pb = a["paragraphs"], b["paragraphs"]  # type: ignore[index]
    if pa != pb:
        for idx, (x, y) in enumerate(zip(pa, pb)):
            if x != y:
                return f"段落[{idx}] 不一致: app={x!r} src={y!r}"
        return f"段落数量不一致: app={len(pa)} src={len(pb)}"
    ta, tb = a["tables"], b["tables"]  # type: ignore[index]
    if len(ta) != len(tb):
        return f"表格数量不一致: app={len(ta)} src={len(tb)}"
    for idx, (x, y) in enumerate(zip(ta, tb)):
        if x != y:
            return f"表格[{idx}] 不一致: app={x} src={y}"
    return "无差异"


def _generate_pair(tmp_dir: Path) -> tuple[Path, Path]:
    from app.customer_report import (
        OpinionExcelCardDOCXRenderer as AppRenderer,
        load_opinion_report_template as app_load_template,
    )
    from src.customer_report.opinion_excel_card_docx_renderer import (
        OpinionExcelCardDOCXRenderer as SrcRenderer,
    )
    from src.customer_report.opinion_report_template import (
        load_opinion_report_template as src_load_template,
    )

    responses = build_fixture_responses()
    tmp_dir.mkdir(parents=True, exist_ok=True)
    app_out = _render(AppRenderer, app_load_template(), tmp_dir / "app_report.docx", responses)
    src_out = _render(SrcRenderer, src_load_template(), tmp_dir / "src_report.docx", responses)
    return app_out, src_out


def test_formal_report_manifest_zero_diff(tmp_path: Path | None = None) -> None:
    tmp_dir = Path(tmp_path) if tmp_path is not None else APP_ROOT / "temp" / "_parity"
    app_out, src_out = _generate_pair(tmp_dir)

    app_manifest = _structural_manifest(app_out)
    src_manifest = _structural_manifest(src_out)
    assert app_manifest == src_manifest, _first_manifest_diff(app_manifest, src_manifest)

    # 正文 XML 逐字节相同（同版本 python-docx、固定输入下应完全一致）
    assert _document_xml(app_out) == _document_xml(src_out), "word/document.xml 正文不一致"


if __name__ == "__main__":
    out_dir = APP_ROOT / "temp" / "_parity"
    a, b = _generate_pair(out_dir)
    manifest_a = _structural_manifest(a)
    manifest_b = _structural_manifest(b)
    same_manifest = manifest_a == manifest_b
    same_xml = _document_xml(a) == _document_xml(b)
    print(f"app docx: {a}")
    print(f"src docx: {b}")
    print(f"段落数: app={len(manifest_a['paragraphs'])} src={len(manifest_b['paragraphs'])}")
    print(f"表格数: app={len(manifest_a['tables'])} src={len(manifest_b['tables'])}")
    print(f"结构 manifest 相同: {same_manifest}")
    print(f"正文 document.xml 相同: {same_xml}")
    if not same_manifest:
        print("首个差异:", _first_manifest_diff(manifest_a, manifest_b))
    if same_manifest and same_xml:
        print("PARITY OK: 两端报告 manifest 零差异")
    else:
        raise SystemExit("PARITY FAILED")
