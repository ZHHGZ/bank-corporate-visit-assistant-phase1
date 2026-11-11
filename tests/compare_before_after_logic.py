"""改前 vs 改后逻辑对比（同一套 mock ESB 响应）."""
from __future__ import annotations

import asyncio
from typing import Any
from unittest.mock import AsyncMock, patch

from app.core.esb_client import get_body, pick
from app.core.mock_data import build_mock_response
from app.parsers._esb_records import find_first_scalar as _find_first_scalar, first_record as _first_record, first_struct as _first_struct, records as _records
from app.parsers.basic_profile import parse_basic_profile
from app.services import (
    INTERFACE_GROUPS,
    _classify_responses,
    _first_record_by_key,
    _join_term,
    _meta,
    get_basic_info,
)
from app.aggregator import build_field_values_from_all_data, build_table_data_from_all_data, _join_list
from app.customer_report.basic_profile_builders import (
    build_certification_report_text,
    normalize_change_records,
    parse_basic_info_record,
)


def get_basic_info_before(responses: dict[str, Any]) -> dict[str, Any]:
    """阶段 4 改动前（commit 2f2b9ff）API-01 逻辑复刻."""
    from app.core.esb_client import is_success

    success, failed, mocked = _classify_responses(responses)
    data: dict[str, Any] = {"basicInfo": {}, "certifications": {}, "changeRecords": []}

    if "9000000001" in responses and is_success(responses["9000000001"]):
        body = get_body(responses["9000000001"])
        info = _first_record(body, "EntpBscInfArry")
        data["basicInfo"] = {
            "companyName": pick(info, "EntpNm"),
            "creditCode": pick(info, "USCC"),
            "legalPerson": pick(info, "EntpLglRprsNm"),
            "registeredCapital": pick(info, "RgstCpitl"),
            "establishDate": pick(info, "FoundDt"),
            "approvalDate": pick(info, "EntpAprvDt"),
            "businessTerm": _join_term(pick(info, "CrpnTrmBeg"), pick(info, "CrpnTrmTo")),
            "address": pick(info, "Addr"),
            "companyType": pick(info, "EntpTp"),
            "status": pick(info, "EntpCrpnSt"),
            "businessScope": pick(info, "CrpnBsnScop"),
            "industry": pick(info, "IdyNm"),
            "employeeCount": pick(info, "StfPplCnt"),
        }

    if "9000000002" in responses and is_success(responses["9000000002"]):
        body = get_body(responses["9000000002"])
        rec = _first_record(body, "DataArry")
        if rec:
            data["basicInfo"]["socialSecurityCount"] = pick(rec, "TownStfBscPnsnInsunPIPC")

    if "9000000003" in responses and is_success(responses["9000000003"]):
        body = get_body(responses["9000000003"])
        scale = _first_struct(body, "A02RspStruct")
        if scale:
            data["basicInfo"]["companyScale"] = pick(scale, "EntpScl")
        cert_names = []
        cert_mappings = [
            ("EntpAcmlnAcqrngcncaAC", "CNCA认证"),
            ("EntpAcmlnAcqrngHgntchLEAC", "高新企业认证"),
            ("EntpAcmlnAcqrngNLNTAC", "无线网电信认证"),
            ("EntpAcmlnAcqrngEnvPrtctAC", "环保认证"),
            ("EntpAcmlnAcqrngFinAC", "金融认证"),
            ("EntpAcmlnAcqrngFdAC", "食品认证"),
            ("EntpAcmlnAcqrngScrRelAC", "涉密认证"),
            ("EntpAcmlnAcqrngDlSftAC", "双软认证"),
            ("EntpAcmlnAcqrngCstmsAC", "海关认证"),
            ("EntpAcmlnAcqrngTenBgAC", "十大认证"),
        ]
        for field, label in cert_mappings:
            count = _find_first_scalar(body, field)
            if count and count not in ("0", "None", None, ""):
                cert_names.append(f"{label}{count}项")
        if cert_names:
            data["certifications"]["certifications"] = cert_names

    if "9000000008" in responses and is_success(responses["9000000008"]):
        body = get_body(responses["9000000008"])
        data["certifications"]["patentCount"] = len(_records(body, "RsltArry"))

    if "9000000005" in responses and is_success(responses["9000000005"]):
        body = get_body(responses["9000000005"])
        ratings = _records(body, "TaxnEvltArry")
        if ratings:
            latest = max(ratings, key=lambda r: pick(r, "RtgYr"))
            data["certifications"]["taxRating"] = pick(latest, "AutLvl")

    if "9000000004" in responses and is_success(responses["9000000004"]):
        data["changeRecords"] = normalize_change_records(responses["9000000004"])

    if "9000000006" in responses and is_success(responses["9000000006"]):
        body = get_body(responses["9000000006"])
        ratings = _records(body, "CrdtEvltListArry")
        if ratings:
            latest_rating = max(ratings, key=lambda r: pick(r, "EvltAnul"), default={})
            data["certifications"]["creditRating"] = pick(latest_rating, "CrdtLvl") or pick(
                latest_rating, "EvltRslt"
            )
        green = pick(body, "WhthrGreenIndus")
        if green:
            data["certifications"]["isGreenIndustry"] = green
        micro = pick(body, "WhthrAbsrbSmlMcrListEntpCtlg")
        if micro:
            data["certifications"]["isSmallMicro"] = micro
        honors = _records(body, "BllbrdListArry")
        if honors:
            data["certifications"]["honors"] = [pick(h, "BllbrdNm") for h in honors if pick(h, "BllbrdNm")]

    if "9000000007" in responses and is_success(responses["9000000007"]):
        body = get_body(responses["9000000007"])
        from app.services import _first_record_by_key

        credit_rec = _first_record_by_key(body, "CstmsCrdtLvlStruct")
        if credit_rec:
            data["certifications"]["customsCreditLevel"] = pick(credit_rec, "CstmsCrdtLvlCrdtLvlNm") or pick(
                credit_rec, "CstmsCrdtLvlCrdtLvlEncd"
            )
        record_rec = _first_record_by_key(body, "CstmsRcdInfStruct")
        if record_rec:
            data["certifications"]["customsRecord"] = pick(record_rec, "CstmsRcdEncd") or pick(
                record_rec, "CstmsRcdNo"
            )

    report_text = build_certification_report_text(responses)
    if report_text:
        data["certifications"]["reportText"] = report_text

    return {**data, **_meta(success, failed, mocked)}


def build_field_values_before(all_data: dict[str, Any]) -> dict[str, str]:
    """阶段 4 改动前 _build_field_values 客户全貌部分复刻."""
    result: dict[str, str] = {}
    basic = all_data.get("basic_info") or {}
    basic_info = basic.get("basicInfo") or {}
    certs = basic.get("certifications") or {}
    report_text = certs.get("reportText") or {}

    basic_mapping = {
        "企业名称": basic_info.get("companyName"),
        "统一社会信用代码": basic_info.get("creditCode"),
        "法定代表人": basic_info.get("legalPerson"),
        "注册 / 实缴资本（万元）": basic_info.get("registeredCapital"),
        "工商核准日期": basic_info.get("approvalDate") or basic_info.get("establishDate"),
        "注册地址": basic_info.get("address"),
        "企业类型": basic_info.get("companyType"),
        "企业经营状态": basic_info.get("status"),
        "经营范围": basic_info.get("businessScope"),
        "国标行业": basic_info.get("industry"),
        "员工人数（人）": basic_info.get("employeeCount"),
        "社保人数（人）": basic_info.get("socialSecurityCount"),
        "企业规模": basic_info.get("companyScale"),
        "有效专利（发明专利）数量": report_text.get("有效专利（发明专利）数量") or certs.get("patentCount"),
        "税务评级": report_text.get("税务评级") or certs.get("taxRating"),
        "信用评级": report_text.get("信用评级") or certs.get("creditRating"),
        "企业认证": report_text.get("企业认证") or _join_list(certs.get("certifications") or []),
        "是否绿色产业": report_text.get("是否绿色产业") or str(certs.get("isGreenIndustry", "") or ""),
        "是否纳入小微企业目录": report_text.get("是否纳入小微企业目录") or str(certs.get("isSmallMicro", "") or ""),
        "海关信用等级": report_text.get("海关信用等级") or certs.get("customsCreditLevel"),
        "海关备案": report_text.get("海关备案") or ("已备案" if certs.get("customsRecord") else ""),
        "荣誉榜单": report_text.get("荣誉榜单") or _join_list(certs.get("honors") or [], limit=3),
        "中知数通科技力评分/分层结果": report_text.get("中知数通科技力评分/分层结果"),
        "科技资质情况": report_text.get("科技资质情况"),
        "纳税A/B级": report_text.get("纳税A/B级"),
    }
    for label, value in basic_mapping.items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    bus_term = basic_info.get("businessTerm", "")
    if bus_term and "至" in bus_term:
        parts = bus_term.split("至", 1)
        if parts[0]:
            result["营业期限自"] = parts[0].strip()
        if len(parts) > 1 and parts[1]:
            result["营业期限至"] = parts[1].strip()
    return result


def main() -> None:
    company = "逻辑对比企业"
    responses = {code: build_mock_response(code, company) for code in INTERFACE_GROUPS["basic"]}
    body = get_body(responses["9000000001"])
    body["QryRsltArry"][0]["EntpBscInfArry"][0].update(
        {"PclupCpitl": "3000", "IdyCd": "I65", "EntpAprvDt": "20100201"}
    )

    before = get_basic_info_before(responses)
    after = parse_basic_profile(responses)

    fv_before = build_field_values_before({"basic_info": before})
    fv_after = build_field_values_from_all_data({"basic_info": after})

    basic_keys = [
        "企业名称", "统一社会信用代码", "法定代表人", "注册 / 实缴资本（万元）",
        "工商核准日期", "注册地址", "企业类型", "企业经营状态", "经营范围", "国标行业",
        "员工人数（人）", "社保人数（人）", "企业规模", "营业期限自", "营业期限至",
        "有效专利（发明专利）数量", "税务评级", "信用评级", "企业认证",
        "是否绿色产业", "是否纳入小微企业目录", "海关信用等级", "海关备案", "荣誉榜单",
        "科技资质情况", "纳税A/B级",
    ]

    print("=== API-01 JSON 结构 ===")
    print("before keys:", sorted(k for k in before if k not in ("sourceInterfaces", "failedInterfaces", "mockedInterfaces", "timestamp")))
    print("after keys:", sorted(k for k in after if k not in ("sourceInterfaces", "failedInterfaces", "mockedInterfaces", "timestamp")))

    print("\n=== field_values 客户全貌字段对比 ===")
    same, diff, only_before, only_after = [], [], [], []
    all_labels = sorted(set(basic_keys) | set(fv_before) | set(fv_after))
    for label in all_labels:
        if label not in basic_keys:
            continue
        b, a = fv_before.get(label), fv_after.get(label)
        if b == a:
            if b is not None:
                same.append(label)
        elif b and a:
            diff.append((label, b, a))
        elif b and not a:
            only_before.append((label, b))
        elif a and not b:
            only_after.append((label, a))

    print(f"一致: {len(same)}")
    for x in same:
        print(f"  [OK] {x}")
    print(f"\n不一致: {len(diff)}")
    for label, b, a in diff:
        print(f"  [DIFF] {label}")
        print(f"    before: {b[:80]!r}")
        print(f"    after:  {a[:80]!r}")
    print(f"\n仅改前有值: {len(only_before)}")
    for label, b in only_before:
        print(f"  [BEFORE_ONLY] {label}: {b[:80]!r}")
    print(f"\n仅改后有值: {len(only_after)}")
    for label, a in only_after:
        print(f"  [AFTER_ONLY] {label}: {a[:80]!r}")

    td_b = build_table_data_from_all_data({"basic_info": before})
    td_a = build_table_data_from_all_data({"basic_info": after})
    print("\n=== change_records 表 ===")
    print("before rows:", len(td_b.get("change_records", [])))
    print("after rows:", len(td_a.get("change_records", [])))
    if td_b.get("change_records") == td_a.get("change_records"):
        print("[OK] 变更表一致")
    else:
        print("[DIFF] 变更表不同")
        print(" before:", td_b.get("change_records"))
        print(" after:", td_a.get("change_records"))


if __name__ == "__main__":
    main()
