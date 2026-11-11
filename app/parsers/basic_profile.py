"""客户全貌（API-01）统一解析：单一数据源，产出 basicInfo / display / tables."""
from __future__ import annotations

from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.customer_report.opinion_excel_docx_renderer import (
    _billboard_summary_for_catalog,
    _change_records,
    _credit_rating_summary_for_catalog,
    _customs_credit_summary,
    _customs_record_summary,
    _enterprise_certification_summary,
    _find_first_scalar_by_keys,
    _green_industry_summary,
    _invent_patent_summary,
    _registered_paid_capital_value,
    _small_micro_catalog_summary,
    _tax_ab_level_summary,
    _tax_rating_summary,
    _zhongzhishutong_qualification_summary,
)
from app.parsers._esb_records import find_first_scalar, first_record, first_struct

CHANGE_RECORD_LIMIT = 3


def parse_basic_info_record(body: Mapping[str, Any] | None) -> dict[str, Any]:
    """从 9000000001 照面解析基础信息."""
    if not body:
        return {}

    info = first_record(body, "EntpBscInfArry", "EntpBsicInfArry")
    if not info:
        return {}

    term_begin = pick(info, "CrpnTrmBeg")
    term_end = pick(info, "CrpnTrmTo")
    industry_name = pick(info, "IdyNm")
    industry_code = pick(info, "IdyCd")
    industry = industry_name
    if industry_name and industry_code:
        industry = f"{industry_name}（{industry_code}）"

    org_code = pick(info, "OIBC", "OrgBrchId", "IdCardNoOrOIBC")

    return {
        "companyName": pick(info, "EntpNm"),
        "creditCode": pick(info, "USCC"),
        "orgCode": org_code,
        "legalPerson": pick(info, "EntpLglRprsNm", "LglRprsNm"),
        "registeredCapital": pick(info, "RgstCpitl"),
        "paidInCapital": pick(info, "PclupCpitl"),
        "registeredPaidCapital": _registered_paid_capital_value(info),
        "establishDate": pick(info, "FoundDt"),
        "approvalDate": pick(info, "EntpAprvDt"),
        "businessTermBegin": term_begin,
        "businessTermEnd": term_end,
        "businessTerm": " 至 ".join(part for part in (term_begin, term_end) if part) if term_begin or term_end else "",
        "address": pick(info, "Addr", "RgstAddrDstcNo"),
        "companyType": pick(info, "EntpTp"),
        "status": pick(info, "EntpCrpnSt"),
        "businessScope": pick(info, "CrpnBsnScop"),
        "industry": industry,
        "industryCode": industry_code,
        "employeeCount": pick(info, "StfPplCnt"),
    }


def normalize_change_records(payload: Any, *, limit: int = CHANGE_RECORD_LIMIT) -> list[dict[str, str]]:
    rows = _change_records(payload)
    result: list[dict[str, str]] = []
    for row in rows[:limit]:
        if not isinstance(row, Mapping):
            continue
        result.append(
            {
                "changeDate": str(row.get("变更日期") or ""),
                "changeItem": str(row.get("变更事项") or ""),
                "beforeChange": str(row.get("变更前内容") or ""),
                "afterChange": str(row.get("变更后内容") or ""),
            }
        )
    return result


def build_tech_score_text(resp_4012: Any) -> str:
    """从 9000000009 BODY 中提取科技力评分."""
    if not resp_4012 or not is_success(resp_4012):
        return ""
    body = get_body(resp_4012)
    score = _find_first_scalar_by_keys(body, ("Scor",))
    level = _find_first_scalar_by_keys(body, ("StrtfcnRslt",))
    if score and level:
        return f"中知数通科技力评分：{score}/{level}"
    if score:
        return f"中知数通科技力评分：{score}"
    if level:
        return f"中知数通科技力评分：{level}"
    return ""


def build_certification_display(responses: Mapping[str, Any]) -> dict[str, str]:
    """评级及标签：报告展示文案（renderer 口径）."""
    text: dict[str, str] = {}

    resp_4012 = responses.get("9000000009")

    tech_score = build_tech_score_text(resp_4012)
    if tech_score:
        text["中知数通科技力评分/分层结果"] = tech_score

    resp_267 = responses.get("9000000003")
    resp_245 = responses.get("9000000005")
    resp_506 = responses.get("9000000006")
    resp_507 = responses.get("9000000007")
    resp_509 = responses.get("9000000008")

    patent_text = _invent_patent_summary(resp_509)
    if patent_text:
        text["有效专利（发明专利）数量"] = patent_text

    qualification_text = _zhongzhishutong_qualification_summary(get_body(resp_4012) if resp_4012 and is_success(resp_4012) else {})
    if qualification_text:
        text["科技资质情况"] = qualification_text

    certification_text = _enterprise_certification_summary(resp_267)
    if certification_text:
        text["企业认证"] = certification_text

    green_text = _green_industry_summary(resp_506)
    if green_text:
        text["是否绿色产业"] = green_text

    micro_text = _small_micro_catalog_summary(resp_506)
    if micro_text:
        text["是否纳入小微企业目录"] = micro_text

    credit_text = _credit_rating_summary_for_catalog(resp_506)
    if credit_text:
        text["信用评级"] = credit_text

    honor_text = _billboard_summary_for_catalog(resp_506)
    if honor_text:
        text["荣誉榜单"] = honor_text

    tax_text = _tax_rating_summary(resp_245)
    if tax_text:
        text["税务评级"] = tax_text

    tax_ab = _tax_ab_level_summary(resp_245)
    if tax_ab:
        text["纳税A/B级"] = tax_ab

    customs_credit = _customs_credit_summary(resp_507)
    if customs_credit:
        text["海关信用等级"] = customs_credit

    customs_record = _customs_record_summary(resp_507)
    if customs_record:
        text["海关备案"] = customs_record

    return text


def build_basic_display(basic_info: Mapping[str, Any]) -> dict[str, str]:
    """基础信息：模板中文字段名 → 展示值."""
    display: dict[str, str] = {}
    mapping = {
        "企业名称": basic_info.get("companyName"),
        "统一社会信用代码": basic_info.get("creditCode"),
        "组织机构代码": basic_info.get("orgCode"),
        "法定代表人": basic_info.get("legalPerson"),
        "注册 / 实缴资本（万元）": basic_info.get("registeredPaidCapital") or basic_info.get("registeredCapital"),
        "工商核准日期": basic_info.get("approvalDate") or basic_info.get("establishDate"),
        "注册地址": basic_info.get("address"),
        "企业类型": basic_info.get("companyType"),
        "企业经营状态": basic_info.get("status"),
        "经营范围": basic_info.get("businessScope"),
        "国标行业": basic_info.get("industry"),
        "员工人数（人）": basic_info.get("employeeCount"),
        "社保人数（人）": basic_info.get("socialSecurityCount"),
        "企业规模": basic_info.get("companyScale"),
    }
    for label, value in mapping.items():
        if value not in (None, "", "None"):
            display[label] = str(value)

    term_begin = basic_info.get("businessTermBegin")
    term_end = basic_info.get("businessTermEnd")
    if term_begin:
        display["营业期限自"] = str(term_begin)
    if term_end:
        display["营业期限至"] = str(term_end)
    if term_begin or term_end:
        display["营业期限"] = str(basic_info.get("businessTerm") or "")

    return display


def parse_basic_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    """解析客户全貌完整结构（API-01 唯一取数入口）."""
    basic_info: dict[str, Any] = {}

    if responses.get("9000000001") and is_success(responses["9000000001"]):
        body = get_body(responses["9000000001"])
        parsed = parse_basic_info_record(body)
        if parsed:
            basic_info.update(parsed)

    if responses.get("9000000002") and is_success(responses["9000000002"]):
        body = get_body(responses["9000000002"])
        social = find_first_scalar(body, "TownStfBscPnsnInsunPIPC")
        if not social:
            social = find_first_scalar(body, "StfBscMdclInsunPIPC")
        if social:
            basic_info["socialSecurityCount"] = social

    if responses.get("9000000003") and is_success(responses["9000000003"]):
        body = get_body(responses["9000000003"])
        scale = first_struct(body, "A02RspStruct")
        if scale:
            basic_info["companyScale"] = pick(scale, "EntpScl")
        if not basic_info.get("companyScale"):
            scale_text = find_first_scalar(body, "EntpScl")
            if scale_text:
                basic_info["companyScale"] = scale_text

    change_records: list[dict[str, str]] = []
    if responses.get("9000000004") and is_success(responses["9000000004"]):
        change_records = normalize_change_records(responses["9000000004"])

    display = build_basic_display(basic_info)
    display.update(build_certification_display(responses))

    return {
        "basicInfo": basic_info,
        "display": display,
        "tables": {"changeRecords": change_records},
        # 兼容旧 HiAgent 入参：顶层 changeRecords
        "changeRecords": change_records,
    }


# 兼容旧 import 路径
build_certification_report_text = build_certification_display
