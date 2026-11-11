"""客户全貌（API-01）字段解析与聚合测试."""
from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, patch

from app.core.esb_client import get_body
from app.core.mock_data import build_mock_response
from app.parsers.basic_profile import (
    build_certification_display,
    normalize_change_records,
    parse_basic_info_record,
    parse_basic_profile,
)
from app.services import INTERFACE_GROUPS, get_basic_info
from app.aggregator import build_field_values_from_all_data, _resolve_basic_display


def _success_response(body: dict) -> dict:
    return {
        "SysHead": {"RspCd": "200"},
        "BODY": {"BsnRetCd": "200", "RspCd": "200", **body},
    }


def test_parse_basic_info_record_includes_industry_code_and_paid_capital() -> None:
    body = {
        "QryRsltArry": [{
            "EntpBscInfArry": [{
                "EntpNm": "示例企业",
                "USCC": "91440101MA5XXXXX0X",
                "EntpLglRprsNm": "张三",
                "RgstCpitl": "5000",
                "PclupCpitl": "3000",
                "FoundDt": "20100115",
                "EntpAprvDt": "20100201",
                "CrpnTrmBeg": "20100115",
                "CrpnTrmTo": "长期",
                "Addr": "广州市天河区示例路88号",
                "EntpTp": "有限责任公司",
                "EntpCrpnSt": "存续",
                "CrpnBsnScop": "软件开发",
                "IdyNm": "软件和信息技术服务业",
                "IdyCd": "I65",
                "StfPplCnt": "120",
            }],
        }],
    }
    parsed = parse_basic_info_record(body)
    assert parsed["industry"] == "软件和信息技术服务业（I65）"
    assert parsed["registeredPaidCapital"] == "5000 / 3000"


def test_normalize_change_records_sorts_and_limits_to_three() -> None:
    payload = _success_response({
        "QryRsltArry": [{
            "MdfInfArry": [
                {"MdfDt": "20220101", "MdfItm": "法定代表人", "MdfPrevsCntnt": "李四", "MdfAfCntnt": "王五"},
                {"MdfDt": "20240301", "MdfItm": "注册资本", "MdfPrevsCntnt": "3000万元", "MdfAfCntnt": "5000万元"},
                {"MdfDt": "20230101", "MdfItm": "经营范围", "MdfPrevsCntnt": "软件开发", "MdfAfCntnt": "软件开发；信息技术咨询"},
                {"MdfDt": "20210101", "MdfItm": "地址", "MdfPrevsCntnt": "旧地址", "MdfAfCntnt": "新地址"},
            ],
        }],
    })
    rows = normalize_change_records(payload)
    assert len(rows) == 3
    assert rows[0]["changeDate"] == "20240301"


def test_build_certification_display_includes_patent_and_tax() -> None:
    responses = {
        "9000000008": build_mock_response("9000000008", "示例企业"),
        "9000000005": build_mock_response("9000000005", "示例企业"),
    }
    text = build_certification_display(responses)
    assert "有效专利（发明专利）数量" in text
    assert "税务评级" in text


def test_build_field_values_uses_display() -> None:
    all_data = {
        "basic_info": {
            "basicInfo": {"companyName": "示例企业"},
            "display": {
                "企业名称": "示例企业",
                "注册 / 实缴资本（万元）": "5000 / 3000",
                "税务评级": "共1条纳税评级记录(最近3条)：\n1、认证级别A（2024年）",
            },
            "tables": {
                "changeRecords": [
                    {
                        "changeDate": "20240301",
                        "changeItem": "注册资本",
                        "beforeChange": "3000万元",
                        "afterChange": "5000万元",
                    }
                ],
            },
            "changeRecords": [
                {
                    "changeDate": "20240301",
                    "changeItem": "注册资本",
                    "beforeChange": "3000万元",
                    "afterChange": "5000万元",
                }
            ],
        },
    }
    fields = build_field_values_from_all_data(all_data)
    assert fields["企业名称"] == "示例企业"
    assert fields["注册 / 实缴资本（万元）"] == "5000 / 3000"
    assert "纳税评级记录" in fields["税务评级"]


def test_get_basic_info_returns_display_and_tables() -> None:
    company = "客户全貌测试企业"
    responses = {code: build_mock_response(code, company) for code in INTERFACE_GROUPS["basic"]}
    body = get_body(responses["9000000001"])
    body["QryRsltArry"][0]["EntpBscInfArry"][0].update({
        "PclupCpitl": "3000",
        "IdyCd": "I65",
        "EntpAprvDt": "20100201",
    })

    async def _run() -> dict:
        with patch("app.services.call_esb_batch", new=AsyncMock(return_value=responses)):
            return await get_basic_info(company)

    result = asyncio.run(_run())
    assert result["basicInfo"]["registeredPaidCapital"] == "5000 / 3000"
    assert "display" in result
    assert result["display"]["企业名称"] == company
    assert "税务评级" in result["display"]
    assert result["tables"]["changeRecords"]
    assert result["changeRecords"] == result["tables"]["changeRecords"]
    assert "certifications" not in result


def test_resolve_basic_display_legacy_report_text_keeps_basic_fields() -> None:
    """旧 HiAgent 入参：仅有 basicInfo + certifications.reportText，无 display."""
    basic = {
        "basicInfo": {
            "companyName": "示例企业",
            "creditCode": "91440101MA5XXXXX0X",
            "registeredCapital": "5000",
            "industry": "软件和信息技术服务业",
        },
        "certifications": {
            "reportText": {
                "税务评级": "共1条纳税评级记录(最近3条)：\n1、认证级别A（2024年）",
            },
        },
    }
    display = _resolve_basic_display(basic)
    assert display["企业名称"] == "示例企业"
    assert display["统一社会信用代码"] == "91440101MA5XXXXX0X"
    assert "纳税评级记录" in display["税务评级"]


def test_parse_basic_profile_unified() -> None:
    company = "统一解析企业"
    responses = {code: build_mock_response(code, company) for code in INTERFACE_GROUPS["basic"]}
    profile = parse_basic_profile(responses)
    assert profile["basicInfo"]["companyName"] == company
    assert profile["display"]["企业名称"] == company
    assert isinstance(profile["tables"]["changeRecords"], list)
