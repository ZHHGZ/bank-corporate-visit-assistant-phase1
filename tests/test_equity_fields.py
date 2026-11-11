"""股权关联（API-02）解析与逻辑对比测试."""
from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, patch

from app.core.mock_data import build_mock_response
from app.parsers.equity import build_equity_tables, parse_equity_profile
from app.services import (
    INTERFACE_GROUPS,
    get_key_person,
)
from app.aggregator import build_field_values_from_all_data, build_table_data_from_all_data


def _mock_responses(company: str = "股权测试企业") -> dict:
    return {code: build_mock_response(code, company) for code in INTERFACE_GROUPS["key_person"]}


def _structured_only(profile: dict) -> dict:
    skip = {"display", "tables", "sourceInterfaces", "failedInterfaces", "mockedInterfaces", "timestamp"}
    return {k: v for k, v in profile.items() if k not in skip}


def test_parse_equity_profile_has_display_and_tables() -> None:
    profile = parse_equity_profile(_mock_responses())
    assert "display" in profile
    assert "tables" in profile
    assert isinstance(profile["tables"].get("investor_detail"), list)


def test_equity_tables_match_legacy_table_builder() -> None:
    profile = parse_equity_profile(_mock_responses())
    legacy_tables = build_table_data_from_all_data({"key_person": _structured_only(profile)})
    assert profile["tables"].get("investor_detail") == legacy_tables.get("investor_detail")
    assert profile["tables"].get("manager_detail") == legacy_tables.get("manager_detail")
    assert profile["tables"].get("equity_change_records") == legacy_tables.get("equity_change_records")


def test_field_values_include_equity_display() -> None:
    profile = parse_equity_profile(_mock_responses())
    fields = build_field_values_from_all_data({"key_person": profile})
    if profile["display"].get("股东基础信息（含股东名称、股东类型、投资比例、投资数额、出资时间）"):
        assert "股东基础信息（含股东名称、股东类型、投资比例、投资数额、出资时间）" in fields


def test_get_key_person_returns_display_and_tables() -> None:
    company = "股权API测试企业"
    responses = _mock_responses(company)

    async def _run() -> dict:
        with patch("app.services.call_esb_batch", new=AsyncMock(return_value=responses)):
            return await get_key_person(company)

    result = asyncio.run(_run())
    assert "display" in result
    assert "tables" in result
    assert result["shareholders"] == parse_equity_profile(responses)["shareholders"]
