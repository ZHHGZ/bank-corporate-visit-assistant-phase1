"""端到端报告生成测试：mock → parser → aggregator → DOCX."""
from __future__ import annotations

import asyncio
import os
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.services import (
    INTERFACE_GROUPS,
    get_basic_info,
    get_key_person,
    get_business,
    get_contact,
    get_risk,
    generate_report,
)
from app.config import get_settings


async def test_e2e_report():
    company_name = "示例科技有限公司"
    print(f"开始端到端报告测试: {company_name}")
    print("=" * 60)

    # 1. 调用 API-01~06
    print("[API-01] 获取基础信息...")
    basic_info = await get_basic_info(company_name)
    print(f"  产出 keys: {list(basic_info.keys())}")
    print(f"  display: {list(basic_info.get('display', {}).keys())}")
    tables = basic_info.get("tables", {})
    print(f"  tables: {dict((k, len(v)) for k, v in tables.items())}")

    print("[API-02] 获取关键人...")
    key_person = await get_key_person(company_name)
    print(f"  产出 keys: {list(key_person.keys())}")
    print(f"  display: {list(key_person.get('display', {}).keys())}")
    tables_kp = key_person.get("tables", {})
    print(f"  tables: {dict((k, len(v)) for k, v in tables_kp.items())}")

    print("[API-03] 获取经营能力...")
    business = await get_business(company_name)
    print(f"  产出 keys: {list(business.keys())}")
    print(f"  display: {list(business.get('display', {}).keys())}")
    tables_biz = business.get("tables", {})
    print(f"  tables: {dict((k, len(v)) for k, v in tables_biz.items())}")
    for k, v in tables_biz.items():
        if v:
            print(f"    [{k}] 首行: {dict((ck, cv[:30] if isinstance(cv, str) and len(cv)>30 else cv) for ck, cv in v[0].items())}")

    print("[API-04] 获取触达路径...")
    contact = await get_contact(company_name)
    print(f"  产出 keys: {list(contact.keys())}")

    print("[API-06] 获取风险信息...")
    risk = await get_risk(company_name)
    print(f"  产出 keys: {list(risk.keys())}")
    print(f"  riskSummary: {risk.get('riskSummary', {})}")
    tables_risk = risk.get("tables", {})
    print(f"  tables: {dict((k, len(v)) for k, v in tables_risk.items())}")
    for k, v in tables_risk.items():
        if v:
            print(f"    [{k}] 首行: {v[0]}")

    # 2. 组装 all_data
    all_data = {
        "basic_info": basic_info,
        "key_person": key_person,
        "business": business,
        "contact": contact,
        "funding": {},
        "risk": risk,
    }

    # 3. 生成 DOCX 报告
    print("\n[API-08] 生成 DOCX 报告...")
    result = await generate_report(
        company_name,
        all_data,
        "docx",
    )
    print(f"  报告ID: {result.get('reportId')}")
    print(f"  文件路径: {result.get('filePath')}")
    print(f"  文件大小: {result.get('fileSize')}")
    print(f"  模块列表: {result.get('reportSummary', {}).get('sections')}")

    assert Path(result["filePath"]).is_file(), f"报告文件不存在: {result['filePath']}"
    print("\n端到端测试完成！")


if __name__ == "__main__":
    cfg = get_settings()
    print(f"ESB mock 回退: {cfg.esb_mock_fallback}")
    asyncio.run(test_e2e_report())
