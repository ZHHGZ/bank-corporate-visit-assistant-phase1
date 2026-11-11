"""阶段三：通过 API-08 验证 9 大模块结构与模块开关."""
from __future__ import annotations

import json
from pathlib import Path

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app


def _report_payload(company_name: str = "模块开关测试企业") -> dict:
    return {
        "companyName": company_name,
        "basicInfo": {"basicInfo": {"companyName": company_name, "industry": "科技推广和应用服务业"}},
        "keyPerson": {"shareholders": []},
        "business": {"financialData": {}},
        "contact": {"contactInfo": {}},
        "funding": {"financing": {}},
        "risk": {"riskSummary": {}},
        "analysis": {"summaries": {}},
        "outputFormat": "docx",
    }


def test_report_without_modules_returns_all_available() -> None:
    """不传 modules 时，返回所有可用模块."""
    client = TestClient(app)
    response = client.post("/api/v1/enterprise/report", json=_report_payload())
    assert response.status_code == 200
    result = response.json()["data"]
    sections = result.get("reportSummary", {}).get("sections", [])
    assert len(sections) >= 4, f"期望至少 4 个模块，实际 {len(sections)}: {sections}"
    assert any("客户全貌" in s for s in sections)
    assert any("风险提示" in s for s in sections)


def test_report_with_modules_basic_only() -> None:
    """传 modules=["basic"] 只返回客户全貌."""
    client = TestClient(app)
    payload = _report_payload()
    payload["modules"] = ["basic"]
    response = client.post("/api/v1/enterprise/report", json=payload)
    assert response.status_code == 200
    result = response.json()["data"]
    sections = result.get("reportSummary", {}).get("sections", [])
    assert "客户全貌" in sections, f"sections 应包含客户全貌: {sections}"
    assert "风险提示" not in sections, f"sections 不应包含风险提示: {sections}"
    assert "营销方案" not in sections, f"sections 不应包含营销方案: {sections}"
    assert "访谈脚本" not in sections, f"sections 不应包含访谈脚本: {sections}"


def test_report_with_old_chinese_module_name() -> None:
    """旧中文模块名 '风险合规' 映射为 '风险提示'."""
    client = TestClient(app)
    payload = _report_payload()
    payload["modules"] = ["风险合规", "客户全貌"]
    response = client.post("/api/v1/enterprise/report", json=payload)
    assert response.status_code == 200
    result = response.json()["data"]
    sections = result.get("reportSummary", {}).get("sections", [])
    assert "风险提示" in sections
    assert "客户全貌" in sections


def test_report_with_nonexistent_module_name() -> None:
    """不存在的模块名原样保留，不影响正常模块."""
    client = TestClient(app)
    payload = _report_payload()
    payload["modules"] = ["basic", "不存在的模块"]
    response = client.post("/api/v1/enterprise/report", json=payload)
    assert response.status_code == 200
    result = response.json()["data"]
    sections = result.get("reportSummary", {}).get("sections", [])
    assert "客户全貌" in sections


def test_report_generates_report_id_and_download_url() -> None:
    """报告生成后返回 reportId 和 downloadUrl."""
    client = TestClient(app)
    response = client.post("/api/v1/enterprise/report", json=_report_payload())
    assert response.status_code == 200
    data = response.json()["data"]
    assert len(data["reportId"]) == 32
    assert data["downloadUrl"].endswith(f"/api/v1/enterprise/reports/{data['reportId']}/download")
