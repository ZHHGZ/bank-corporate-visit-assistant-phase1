"""客户全貌改前/改后逻辑对比测试."""
from __future__ import annotations

from app.core.esb_client import get_body
from app.core.mock_data import build_mock_response
from app.parsers.basic_profile import parse_basic_profile
from app.services import INTERFACE_GROUPS
from app.aggregator import build_field_values_from_all_data, build_table_data_from_all_data, _resolve_basic_display
from tests.compare_before_after_logic import (
    build_field_values_before,
    get_basic_info_before,
)


def _mock_responses(company: str = "逻辑对比企业") -> dict:
    responses = {code: build_mock_response(code, company) for code in INTERFACE_GROUPS["basic"]}
    body = get_body(responses["9000000001"])
    body["QryRsltArry"][0]["EntpBscInfArry"][0].update(
        {"PclupCpitl": "3000", "IdyCd": "I65", "EntpAprvDt": "20100201"}
    )
    return responses


BASIC_FIELD_KEYS = {
    "企业名称", "统一社会信用代码", "组织机构代码", "法定代表人",
    "注册 / 实缴资本（万元）", "工商核准日期", "注册地址", "企业类型",
    "企业经营状态", "经营范围", "国标行业", "员工人数（人）", "社保人数（人）",
    "企业规模", "营业期限自", "营业期限至", "有效专利（发明专利）数量",
    "税务评级", "信用评级", "企业认证", "是否绿色产业", "是否纳入小微企业目录",
    "海关信用等级", "海关备案", "荣誉榜单", "科技资质情况", "纳税A/B级",
    "中知数通科技力评分/分层结果",
}


def _pick_basic_fields(field_values: dict[str, str]) -> dict[str, str]:
    return {k: v for k, v in field_values.items() if k in BASIC_FIELD_KEYS}


def test_legacy_report_text_path_keeps_basic_fields() -> None:
    before = get_basic_info_before(_mock_responses())
    legacy = {
        "basicInfo": before["basicInfo"],
        "certifications": before["certifications"],
        "changeRecords": before["changeRecords"],
    }
    fv_before = _pick_basic_fields(build_field_values_before({"basic_info": before}))
    fv_legacy = _pick_basic_fields(build_field_values_from_all_data({"basic_info": legacy}))
    assert fv_before == fv_legacy


def test_field_values_legacy_path_matches_stage4_before() -> None:
    responses = _mock_responses()
    before = get_basic_info_before(responses)
    after = parse_basic_profile(responses)
    fv_before = build_field_values_before({"basic_info": before})
    fv_after = build_field_values_from_all_data({"basic_info": after})

    allowed_diffs = {
        "国标行业": ("软件和信息技术服务业", "软件和信息技术服务业（I65）"),
        "注册 / 实缴资本（万元）": ("5000", "5000 / 3000"),
    }
    for key in fv_before:
        if key in allowed_diffs:
            assert fv_after.get(key) == allowed_diffs[key][1]
        else:
            assert fv_after.get(key) == fv_before[key], key

    # 改后新增展示字段（不改旧字段语义）
    if after.get("display", {}).get("营业期限"):
        assert fv_after.get("营业期限") == after["display"]["营业期限"]


def test_change_records_table_unchanged() -> None:
    responses = _mock_responses()
    before = get_basic_info_before(responses)
    after = parse_basic_profile(responses)
    td_before = build_table_data_from_all_data({"basic_info": before})
    td_after = build_table_data_from_all_data({"basic_info": after})
    assert td_before.get("change_records") == td_after.get("change_records")


def test_resolve_basic_display_includes_org_code_in_legacy_path() -> None:
    display = _resolve_basic_display({
        "basicInfo": {"companyName": "示例", "orgCode": "123456789"},
    })
    assert display["组织机构代码"] == "123456789"
