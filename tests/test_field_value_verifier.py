"""阶段四-小块二：字段取值验证工具测试."""
from __future__ import annotations

import json
from pathlib import Path

from app.customer_report.field_mapping import load_field_mapping
from app.customer_report.field_value_verifier import (
  build_mock_esb_responses,
  format_report,
  load_esb_responses_from_json,
  summarize_results,
  verify_all,
  verify_field,
)


def test_verify_basic_fields_hit_with_mock() -> None:
  mapping = load_field_mapping()
  responses = build_mock_esb_responses("验证测试企业", ["9000000001", "9000000002"])
  company_name = mapping.get_field_by_key("company_name")
  assert company_name is not None
  result = verify_field(company_name, responses)
  assert result.status == "hit"
  assert result.sample_value == "验证测试企业"


def test_verify_patent_count_hits_mock_recursive_count() -> None:
  mapping = load_field_mapping()
  responses = build_mock_esb_responses("验证测试企业", ["9000000008"])
  patent = mapping.get_field_by_key("patent_count")
  assert patent is not None
  result = verify_field(patent, responses)
  assert result.status == "hit"
  assert result.sample_value == "2"


def test_verify_financing_is_skipped_as_derived() -> None:
  mapping = load_field_mapping()
  responses = build_mock_esb_responses("验证测试企业", ["9000000029"])
  financing = mapping.get_field_by_key("financing")
  assert financing is not None
  result = verify_field(financing, responses)
  assert result.status == "skipped"


def test_verify_missing_interface_returns_miss() -> None:
  mapping = load_field_mapping()
  company_name = mapping.get_field_by_key("company_name")
  assert company_name is not None
  result = verify_field(company_name, {})
  assert result.status == "miss"
  assert "接口响应缺失" in result.detail


def test_load_esb_responses_from_json_wrapper_format(tmp_path: Path) -> None:
  payload = {
    "company_name": "样例企业",
    "responses": {
      "9000000001": {
        "BODY": {
          "BsnRetCd": "200",
          "QryRsltArry": [{"EntpBscInfArry": [{"EntpNm": "样例企业"}]}],
        }
      }
    },
  }
  path = tmp_path / "responses.json"
  path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
  responses = load_esb_responses_from_json(path)
  mapping = load_field_mapping()
  company_name = mapping.get_field_by_key("company_name")
  assert company_name is not None
  result = verify_field(company_name, responses)
  assert result.status == "hit"


def test_verify_basic_module_report_has_summary() -> None:
  mapping = load_field_mapping()
  responses = build_mock_esb_responses("验证测试企业")
  results = verify_all(mapping, responses, module_id="basic")
  summary = summarize_results(results)
  assert summary["hit"] >= 10
  report = format_report(results)
  assert "字段取值验证报告" in report
  assert "命中" in report
