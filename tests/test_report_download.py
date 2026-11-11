"""API-08 reportId 下载链路测试."""
from __future__ import annotations

import json
import shutil
from pathlib import Path

from fastapi.testclient import TestClient

from app.config import get_settings
from app.main import app


def test_generate_json_report_and_download_by_report_id() -> None:
    client = TestClient(app)
    company_name = "下载接口测试企业"
    payload = {
        "companyName": company_name,
        "basicInfo": {"basicInfo": {"companyName": company_name}},
        "keyPerson": {"shareholders": []},
        "business": {"financialData": {}},
        "contact": {"contactInfo": {}},
        "funding": {"financing": {}},
        "risk": {"riskSummary": {}},
        "analysis": {"summaries": {}},
        "outputFormat": "json",
    }

    response = client.post("/api/v1/enterprise/report", json=payload)
    assert response.status_code == 200
    result = response.json()["data"]
    assert len(result["reportId"]) == 32
    assert result["downloadUrl"].endswith(
        f"/api/v1/enterprise/reports/{result['reportId']}/download"
    )

    download = client.get(result["downloadUrl"])
    assert download.status_code == 200
    assert download.headers["content-type"].startswith("application/json")
    downloaded = json.loads(download.content)
    assert downloaded["basic_info"]["basicInfo"]["companyName"] == company_name
    assert downloaded["key_person"] == {"shareholders": []}
    assert downloaded["business"] == {"financialData": {}}
    assert downloaded["analysis"] == {"summaries": {}}

    output_dir = get_settings().resolve_path(get_settings().report_output_dir)
    generated_file = Path(result["filePath"])
    generated_dir = generated_file.parent
    if generated_dir.is_relative_to(output_dir) and generated_dir.name.endswith(company_name):
        shutil.rmtree(generated_dir)


def test_download_rejects_invalid_or_missing_report_id() -> None:
    client = TestClient(app)

    invalid = client.get("/api/v1/enterprise/reports/not-a-report-id/download")
    assert invalid.status_code == 404

    missing = client.get(
        "/api/v1/enterprise/reports/00000000000000000000000000000000/download"
    )
    assert missing.status_code == 404
