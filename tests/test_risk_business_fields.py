"""风险提示 / 经营能力逻辑对比测试."""
from __future__ import annotations

from app.core.mock_data import build_mock_response
from app.parsers.business import parse_business_profile
from app.parsers.risk import parse_risk_profile
from app.services import (
    INTERFACE_GROUPS,
)
from app.aggregator import build_field_values_from_all_data, build_table_data_from_all_data, _resolve_risk_display


def _structured_only(profile: dict, *, skip: set[str] | None = None) -> dict:
    omitted = skip or {"display", "tables", "sourceInterfaces", "failedInterfaces", "mockedInterfaces", "timestamp"}
    return {k: v for k, v in profile.items() if k not in omitted}


def test_risk_display_matches_legacy_field_values() -> None:
    responses = {code: build_mock_response(code, "风险测试企业") for code in INTERFACE_GROUPS["risk"]}
    profile = parse_risk_profile(responses)
    legacy = _structured_only(profile)
    fv_legacy = build_field_values_from_all_data({"risk": legacy})
    fv_new = build_field_values_from_all_data({"risk": profile})
    for key in ("企业被执行人", "异常经营", "法律诉讼信息"):
        assert fv_legacy.get(key) == fv_new.get(key), key


def test_resolve_risk_display_legacy_path() -> None:
    risk = {
        "dishonestInfo": {"count": 2},
        "businessAbnormality": {"count": 1},
        "litigationInfo": {"count": 3},
        "penaltyInfo": {"count": 0},
    }
    display = _resolve_risk_display(risk)
    assert display["企业被执行人"] == "被执行/限消记录2条"
    assert display["异常经营"] == "经营异常1条"
    assert "涉诉3条" in display["法律诉讼信息"]


def test_business_tables_match_legacy_table_builder() -> None:
    responses = {code: build_mock_response(code, "经营测试企业") for code in INTERFACE_GROUPS["business"]}
    profile = parse_business_profile(responses)
    aggregated = build_table_data_from_all_data({"business": profile})
    assert aggregated.get("bid_records") == profile.get("tables", {}).get("bid_records")
    assert aggregated.get("mortgage_records") == profile.get("tables", {}).get("mortgage_records")
    assert aggregated.get("pledge_records") == profile.get("tables", {}).get("pledge_records")
    assert aggregated.get("financing_records") == profile.get("tables", {}).get("financing_records")


def test_business_display_revenue_field() -> None:
    responses = {code: build_mock_response(code, "经营测试企业") for code in INTERFACE_GROUPS["business"]}
    profile = parse_business_profile(responses)
    fields = build_field_values_from_all_data({"business": profile})
    if profile.get("financialData", {}).get("revenue"):
        assert fields.get("近 1-3 年营收数据及变动情况") == str(profile["financialData"]["revenue"])
