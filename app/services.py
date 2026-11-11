"""所有业务服务逻辑 - 精简统一版本."""
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from app.core.esb_client import (
    call_esb_batch,
    is_mock,
    is_success,
)
from app.core.llm_client import chat_completion
from app.parsers.basic_profile import parse_basic_profile
from app.parsers.business import parse_business_profile
from app.parsers.contact import parse_contact_profile
from app.parsers.equity import parse_equity_profile
from app.parsers.funding import parse_funding_profile
from app.parsers.risk import RISK_ARRAYS, parse_risk_profile
from app.aggregator import (
    build_field_values_from_all_data,
    build_table_data_from_all_data,
    build_annual_financial_table_from_all_data,
)
from app.utils.logger import get_logger


logger = get_logger(__name__)


# ========== 接口分组定义 ==========

INTERFACE_GROUPS = {
    "basic": ["9000000001", "9000000002", "9000000003", "9000000004", "9000000005", "9000000006", "9000000007", "9000000008", "9000000009"],
    "key_person": [
        "9000000010", "9000000011", "9000000012", "9000000013",
        "9000000014", "9000000015", "9000000016", "9000000017",
    ],
    "business": ["9000000018", "9000000019", "9000000008", "9000000015", "9000000020", "9000000021", "9000000022", "9000000023", "9000000024", "9000000025", "9000000026", "9000000027", "9000000028", "9000000029"],
    "contact": [
        "9000000001", "9000000030",
        "9000000031", "9000000032", "9000000033", "9000000034",
    ],
    "funding": ["9000000023", "9000000024", "9000000025", "9000000020", "9000000021", "9000000022", "9000000029"],
    "risk": [
        "9000000035", "9000000036", "9000000037", "9000000038", "9000000039", "9000000040",
        "9000000041", "9000000042", "9000000043", "9000000044", "9000000045", "9000000046",
        "9000000047", "9000000048", "9000000049", "9000000050", "9000000051", "9000000052",
        "9000000053", "9000000054", "9000000055", "9000000056",
    ],
}

# RISK_ARRAYS 已迁移至 app.parsers.risk，此处保留别名供旧代码引用


def _classify_responses(responses: dict[str, Any]) -> tuple[list[str], list[str], list[str]]:
    """拆分真实成功 / 失败 / mock 兜底的交易码."""
    success, failed, mocked = [], [], []
    for code, resp in responses.items():
        if is_mock(resp):
            mocked.append(code)
            success.append(code)
        elif is_success(resp):
            success.append(code)
        else:
            failed.append(code)
    return success, failed, mocked


def _meta(success: list[str], failed: list[str], mocked: list[str]) -> dict[str, Any]:
    return {
        "sourceInterfaces": success,
        "failedInterfaces": failed,
        "mockedInterfaces": mocked,
        "timestamp": datetime.now().isoformat(),
    }
def _first_record_by_key(data: Any, struct_key: str) -> dict[str, Any]:
    """递归查找指定结构体名下的第一条记录."""
    if isinstance(data, dict):
        found = data.get(struct_key)
        if isinstance(found, dict):
            return found
        if isinstance(found, list) and found:
            first = found[0]
            return first if isinstance(first, dict) else {}
        for child in data.values():
            result = _first_record_by_key(child, struct_key)
            if result:
                return result
    elif isinstance(data, list):
        for item in data:
            result = _first_record_by_key(item, struct_key)
            if result:
                return result
    return {}

def _join_term(begin: str, end: str) -> str:
    if begin and end:
        return f"{begin} 至 {end}"
    return begin or end


async def get_basic_info(company_name: str) -> dict[str, Any]:
    """API-01: 获取企业基础信息."""
    logger.info(f"[API-01] 获取基础信息: {company_name}")
    responses = await call_esb_batch(INTERFACE_GROUPS["basic"], company_name)
    success, failed, mocked = _classify_responses(responses)

    profile = parse_basic_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def get_key_person(company_name: str) -> dict[str, Any]:
    """API-02: 获取关键人信息（字段对齐 HiAgent 方案）."""
    logger.info(f"[API-02] 获取关键人: {company_name}")
    responses = await call_esb_batch(
        INTERFACE_GROUPS["key_person"],
        company_name,
        extras={"9000000012": {"RelaTp": "manager,invest", "FltrLvl": "1"}},
    )
    success, failed, mocked = _classify_responses(responses)
    profile = parse_equity_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def get_business(company_name: str) -> dict[str, Any]:
    """API-03: 获取经营能力."""
    logger.info(f"[API-03] 获取经营能力: {company_name}")
    responses = await call_esb_batch(INTERFACE_GROUPS["business"], company_name)
    success, failed, mocked = _classify_responses(responses)
    profile = parse_business_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def get_contact(company_name: str) -> dict[str, Any]:
    """API-04: 获取触达路径."""
    logger.info(f"[API-04] 获取触达路径: {company_name}")
    responses = await call_esb_batch(INTERFACE_GROUPS["contact"], company_name)
    success, failed, mocked = _classify_responses(responses)
    profile = parse_contact_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def get_funding(company_name: str) -> dict[str, Any]:
    """API-05: 获取资金需求."""
    logger.info(f"[API-05] 获取资金需求: {company_name}")
    responses = await call_esb_batch(INTERFACE_GROUPS["funding"], company_name)
    success, failed, mocked = _classify_responses(responses)
    profile = parse_funding_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def get_risk(company_name: str) -> dict[str, Any]:
    """API-06: 获取风险信息."""
    logger.info(f"[API-06] 获取风险信息: {company_name}")
    responses = await call_esb_batch(INTERFACE_GROUPS["risk"], company_name)
    success, failed, mocked = _classify_responses(responses)
    profile = parse_risk_profile(responses)
    return {**profile, **_meta(success, failed, mocked)}


async def generate_analysis(company_name: str, all_data: dict[str, Any]) -> dict[str, Any]:
    """API-07: 生成AI智能分析."""
    logger.info(f"[API-07] 生成AI分析: {company_name}")

    # 生成客户全貌摘要
    basic = all_data.get("basic_info") or {}
    basic_info = basic.get("basicInfo") or {}
    prompt = f"""请用100字概括这家企业：
企业名称：{basic_info.get('companyName', company_name)}
行业：{basic_info.get('industry', '未知')}
注册资本：{basic_info.get('registeredCapital', '未知')}
经营状态：{basic_info.get('status', '未知')}"""

    summary = await chat_completion(prompt)

    # 生成SWOT分析
    risk = all_data.get("risk") or {}
    risk_summary = risk.get("riskSummary") or {}
    swot_prompt = f"""基于以下已知信息生成SWOT分析：
企业：{company_name}
风险等级：{risk_summary.get('riskLevel', '未知')}
风险数量：{risk_summary.get('totalRiskCount', 0)}

要求：
1. 只能使用已知信息，不得虚构市场份额、客户、行业地位等事实；
2. strengths、weaknesses、opportunities、threats 各输出2条，每条不超过50字；
3. 只输出一个完整合法的JSON对象，不要Markdown代码块、前言或结语。"""

    swot_text = await chat_completion(swot_prompt)

    return {
        "summaries": {"客户全貌": summary},
        "aiAnalysis": {"rawText": swot_text},
        "marketingRecommendation": {},
        "timestamp": datetime.now().isoformat(),
    }


async def generate_report(
    company_name: str,
    all_data: dict[str, Any],
    output_format: str,
    *,
    option_payload: dict[str, Any] | None = None,
    bsn_data: dict[str, Any] | None = None,
    modules: list[str] | None = None,
) -> dict[str, Any]:
    """API-08: 生成报告.

    - JSON：沿用现状，直接落盘聚合结果（HiAgent 友好）。
    - DOCX：优先使用 API-01~07 传入的聚合数据填充报告，ESB 缓存数据作为补充回退。
    - modules：指定要渲染的模块列表，不传则渲染全部。
    """
    logger.info(f"[API-08] 生成报告: {company_name}, format={output_format}")

    from app.config import get_settings
    import json

    cfg = get_settings()
    output_dir = cfg.resolve_path(cfg.report_output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 顶层 company_name 无效（空 / "string" 占位）时，回退用基础信息里的企业名
    basic = all_data.get("basic_info") or {}
    basic_info = basic.get("basicInfo") or {}
    display_name = company_name
    if not display_name or display_name.strip().lower() == "string":
        display_name = basic_info.get("companyName") or company_name or "未命名企业"

    generated_at = datetime.now()
    date_str = generated_at.strftime("%Y%m%d")
    file_stamp = generated_at.strftime("%H%M%S_%f")
    report_id = uuid.uuid4().hex
    company_dir = output_dir / f"{date_str}_{display_name}"
    company_dir.mkdir(parents=True, exist_ok=True)

    if output_format == "json":
        # JSON格式：直接保存所有聚合数据
        json_path = company_dir / f"{display_name}_report_{file_stamp}_{report_id}.json"
        json_path.write_text(json.dumps(all_data, ensure_ascii=False, indent=2), encoding="utf-8")
        return {
            "reportId": report_id,
            "filePath": str(json_path),
            "fileName": json_path.name,
            "fileSize": f"{json_path.stat().st_size / 1024:.1f}KB",
            "generatedAt": datetime.now().isoformat(),
            "reportSummary": {"format": "json", "sections": list(all_data.keys())},
        }

    docx_path = company_dir / f"{display_name}_report_{file_stamp}_{report_id}.docx"
    modules = await _render_formal_docx(
        display_name,
        docx_path,
        generated_at,
        option_payload=option_payload,
        bsn_data=bsn_data,
        all_data=all_data,
        modules=modules,
    )
    return {
        "reportId": report_id,
        "filePath": str(docx_path),
        "fileName": docx_path.name,
        "fileSize": f"{docx_path.stat().st_size / 1024:.1f}KB",
        "generatedAt": datetime.now().isoformat(),
        "reportSummary": {"format": "docx", "sections": modules},
    }


def resolve_report_file(report_id: str) -> Path | None:
    """按报告 ID 定位已生成文件，并严格限制在报告输出目录内."""
    if not re.fullmatch(r"[0-9a-f]{32}", report_id):
        return None

    from app.config import get_settings

    cfg = get_settings()
    output_dir = cfg.resolve_path(cfg.report_output_dir).resolve()
    if not output_dir.is_dir():
        return None

    for suffix in (".docx", ".json"):
        for candidate in output_dir.rglob(f"*_{report_id}{suffix}"):
            resolved = candidate.resolve()
            try:
                resolved.relative_to(output_dir)
            except ValueError:
                continue
            if resolved.is_file():
                return resolved
    return None

async def _render_formal_docx(
    company_name: str,
    output_path: Any,
    generated_at: datetime,
    *,
    option_payload: dict[str, Any] | None = None,
    bsn_data: dict[str, Any] | None = None,
    all_data: dict[str, Any] | None = None,
    modules: list[str] | None = None,
) -> list[str]:
    """组装正式报告输入并调用官方 DOCX 渲染器，返回报告模块顺序."""
    from app.config import get_settings
    from app.core.esb_client import ensure_report_responses
    from app.customer_report import (
        OpinionExcelCardDOCXRenderer,
        load_opinion_report_template,
    )
    from app.customer_report.annual_financial_table import (
        build_zhongshu_annual_financial_table,
    )
    from app.module_registry import (
        normalize_module_list,
        remap_fields_to_new_modules,
        sort_fields_by_module_order,
    )

    cfg = get_settings()

    field_values = build_field_values_from_all_data(all_data or {})
    if isinstance(option_payload, dict):
        field_values.update(option_payload)

    try:
        if cfg.report_esb_fallback:
            interface_responses = await ensure_report_responses(company_name)
        else:
            logger.info("[API-08] REPORT_ESB_FALLBACK=false，跳过 ESB 二次查询，仅使用上游 JSON")
            interface_responses = {}
    except Exception:
        logger.warning(f"[API-08] ESB 原始报文获取失败，将仅使用 API-01~07 传入的聚合数据生成报告")
        interface_responses = {}

    interface_responses = {k: v for k, v in interface_responses.items() if not is_mock(v)}

    table_data = build_table_data_from_all_data(all_data or {})

    template = load_opinion_report_template()
    raw_fields = template.fields

    raw_fields = remap_fields_to_new_modules(raw_fields)
    raw_fields = sort_fields_by_module_order(raw_fields)

    allowed_modules = normalize_module_list(modules)
    if allowed_modules:
        fields = [f for f in raw_fields if f.module in allowed_modules]
    else:
        fields = list(raw_fields)

    annual_financial_table = build_zhongshu_annual_financial_table(
        interface_responses,
        display_limit=3,
        fill_missing_years=False,
    ) or build_annual_financial_table_from_all_data(all_data or {})

    internal_info_section = bsn_data if isinstance(bsn_data, dict) and bsn_data.get("internal_groups") else None

    OpinionExcelCardDOCXRenderer().render(
        fields,
        output_path,
        company_name=company_name,
        interface_responses=interface_responses,
        field_values=field_values,
        annual_financial_table=annual_financial_table,
        internal_info_section=internal_info_section,
        risk_detail_tables=template.risk_detail_tables,
        table_data=table_data,
        disable_all_llm=not cfg.llm_enabled,
        generated_at=generated_at,
        explicit_module_filter=bool(modules),
    )

    result_modules: list[str] = []
    for item in fields:
        if item.module not in result_modules:
            result_modules.append(item.module)
    return result_modules
