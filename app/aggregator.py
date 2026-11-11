"""API-08 聚合逻辑：将各模块数据摊平为 renderer 可消费的 field_values / table_data 字典."""
from __future__ import annotations

from typing import Any


def _join_list(items: list[str], *, limit: int = 5) -> str:
    if not items:
        return ""
    shown = items[:limit]
    return "、".join(str(s) for s in shown if s)





# ═══════════════════════════════════════════════════════════════════
# 各模块 display 解析（兼容新旧 HiAgent 数据格式）
# ═══════════════════════════════════════════════════════════════════

def _resolve_basic_display(basic: dict[str, Any]) -> dict[str, str]:
    new_display = basic.get("display")
    if new_display:
        display = dict(new_display)
        report_text = (basic.get("certifications") or {}).get("reportText") or {}
        for label, value in report_text.items():
            if label not in display and value not in (None, "", "None"):
                display[label] = str(value)
        return display

    basic_info = basic.get("basicInfo") or {}
    certs = basic.get("certifications") or {}
    report_text = certs.get("reportText") or {}
    display: dict[str, str] = {}
    legacy = {
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
    for label, value in legacy.items():
        if value not in (None, "", "None"):
            display[label] = str(value)

    term_begin = basic_info.get("businessTermBegin")
    term_end = basic_info.get("businessTermEnd")
    if term_begin:
        display.setdefault("营业期限自", str(term_begin))
    if term_end:
        display.setdefault("营业期限至", str(term_end))
    if "营业期限自" not in display or "营业期限至" not in display:
        bus_term = basic_info.get("businessTerm", "")
        if bus_term and "至" in bus_term:
            parts = bus_term.split("至", 1)
            if "营业期限自" not in display and parts[0]:
                display["营业期限自"] = parts[0].strip()
            if "营业期限至" not in display and len(parts) > 1 and parts[1]:
                display["营业期限至"] = parts[1].strip()

    return display


def _resolve_equity_display(key_person: dict[str, Any]) -> dict[str, str]:
    new_display = key_person.get("display")
    if new_display:
        return dict(new_display)

    display: dict[str, str] = {}
    controller = key_person.get("actualController") or {}
    if controller.get("name"):
        parts = [f"实控人：{controller['name']}"]
        if controller.get("controlPath"):
            parts.append(f"控制路径：{controller['controlPath']}")
        display["实际控制人"] = "；".join(parts)
    return display


def _resolve_business_display(business: dict[str, Any]) -> dict[str, str]:
    if business.get("display"):
        return dict(business["display"])
    fin_data = business.get("financialData") or {}
    display: dict[str, str] = {}
    revenue = fin_data.get("revenue")
    if revenue not in (None, "", "None"):
        display["近 1-3 年营收数据及变动情况"] = str(revenue)
    return display


def _resolve_risk_display(risk: dict[str, Any]) -> dict[str, str]:
    if risk.get("display"):
        return dict(risk["display"])

    dishonest = risk.get("dishonestInfo") or {}
    business_abnormal = risk.get("businessAbnormality") or {}
    litigation = risk.get("litigationInfo") or {}
    penalty = risk.get("penaltyInfo") or {}
    dc = dishonest.get("count", 0)
    bc = business_abnormal.get("count", 0)
    lc = litigation.get("count", 0)
    pc = penalty.get("count", 0)
    litigation_parts = []
    if lc:
        litigation_parts.append(f"涉诉{lc}条")
    if dc:
        litigation_parts.append(f"失信{dc}条")
    if pc:
        litigation_parts.append(f"处罚{pc}条")
    if bc:
        litigation_parts.append(f"经营异常{bc}条")
    return {
        "企业被执行人": f"被执行/限消记录{dc}条" if dc else "被执行/限消记录0条",
        "异常经营": f"经营异常{bc}条" if bc else "经营异常0条",
        "法律诉讼信息": "、".join(litigation_parts) if litigation_parts else "暂无诉讼记录",
    }


# ═══════════════════════════════════════════════════════════════════
# 主入口：field_values / table_data 构建
# ═══════════════════════════════════════════════════════════════════

def build_field_values_from_all_data(all_data: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}

    basic = all_data.get("basic_info") or {}
    basic_info = basic.get("basicInfo") or {}

    for label, value in _resolve_basic_display(basic).items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    key_person = all_data.get("key_person") or {}
    for label, value in _resolve_equity_display(key_person).items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    business = all_data.get("business") or {}
    for label, value in _resolve_business_display(business).items():
        if value not in (None, "", "None"):
            result[label] = str(value)
    industry = (all_data.get("basic_info") or {}).get("basicInfo") or {}
    if industry.get("industry") not in (None, "", "None"):
        result["企业所属行业"] = str(industry.get("industry"))

    contact = all_data.get("contact") or {}
    contact_info = contact.get("contactInfo") or {}
    contact_mapping = {
        "联系人电话": contact_info.get("phone"),
        "邮箱": contact_info.get("email"),
    }
    for label, value in contact_mapping.items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    funding = all_data.get("funding") or {}
    mortgage = funding.get("mortgageInfo") or {}
    pledge = funding.get("equityPledge") or {}
    fin_events = funding.get("financing") or {}
    mc = mortgage.get("count", 0)
    pc = pledge.get("count", 0)
    fr = fin_events.get("rounds", 0)
    latest_event = ""
    evts = fin_events.get("events") or []
    if evts:
        e = evts[0]
        latest_event = f"，最新：{e.get('eventName', '')}{' ' + str(e.get('amount','')) + '万' if e.get('amount') else ''}"
    funding_values = {
        "动产抵押": f"动产抵押{mc}条" if mc else "动产抵押0条",
        "股权出质": f"股权出质{pc}条" if pc else "股权出质0条",
        "融资情况": f"融资{fr}轮{latest_event}" if fr else "暂无融资记录",
    }
    for label, value in funding_values.items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    risk = all_data.get("risk") or {}
    for label, value in _resolve_risk_display(risk).items():
        if value not in (None, "", "None"):
            result[label] = str(value)

    return result


def build_table_data_from_all_data(all_data: dict[str, Any]) -> dict[str, list[dict[str, str]]]:
    result: dict[str, list[dict[str, str]]] = {}

    basic = all_data.get("basic_info") or {}
    tables = basic.get("tables") or {}
    changes = tables.get("changeRecords") or basic.get("changeRecords") or []
    if changes:
        result["change_records"] = [
            {"变更日期": r.get("changeDate", "") or "", "变更事项": r.get("changeItem", "") or "",
             "变更前内容": r.get("beforeChange", "") or "", "变更后内容": r.get("afterChange", "") or ""}
            for r in changes
        ]

    key_person = all_data.get("key_person") or {}
    equity_tables = key_person.get("tables") or {}

    if equity_tables.get("investor_detail"):
        result["investor_detail"] = equity_tables["investor_detail"]
    else:
        shareholders = key_person.get("shareholders") or []
        if shareholders:
            result["investor_detail"] = [
                {"股东名称": r.get("name", "") or "", "出资比例": r.get("shareholdingRatio", "") or "",
                 "股东类型": r.get("type", "") or "", "认缴出资额度": r.get("subscriptionAmount", "") or "",
                 "出资日期": "-"}
                for r in shareholders
            ]

    if equity_tables.get("outbound_investment"):
        result["outbound_investment"] = equity_tables["outbound_investment"]
    else:
        investments = key_person.get("investmentInfo") or []
        if investments:
            result["outbound_investment"] = [
                {"公司名称": r.get("investedCompany", "") or "", "占比": r.get("investmentRatio", "") or "",
                 "认缴金额": "-", "企业状态": "-"}
                for r in investments
            ]

    if equity_tables.get("manager_detail"):
        result["manager_detail"] = equity_tables["manager_detail"]
    else:
        executives = key_person.get("executives") or []
        if executives:
            result["manager_detail"] = [
                {"姓名": r.get("name", "") or "", "职务": r.get("position", "") or ""}
                for r in executives
            ]

    if equity_tables.get("equity_change_records"):
        result["equity_change_records"] = equity_tables["equity_change_records"]
    else:
        equity_changes = key_person.get("equityChangeRecords") or []
        if equity_changes:
            result["equity_change_records"] = [
                {"变更日期": r.get("changeDate", "") or "", "股东名称": r.get("shareholderName", "") or "",
                 "转让前股权比例": r.get("beforeRatio", "") or "", "转让后股权比例": r.get("afterRatio", "") or ""}
                for r in equity_changes
            ]

    business = all_data.get("business") or {}
    business_tables = business.get("tables") or {}
    for key in ("bid_records", "mortgage_records", "pledge_records", "financing_records"):
        rows = business_tables.get(key)
        if rows:
            result[key] = rows

    risk = all_data.get("risk") or {}
    risk_tables = risk.get("tables") or {}
    for key in ("execution_risk", "consumption_limit", "discredit_executor"):
        rows = risk_tables.get(key)
        if rows:
            result[key] = rows

    return result


# ═══════════════════════════════════════════════════════════════════
# 年报财务表（API-03 financialData → renderer 格式）
# ═══════════════════════════════════════════════════════════════════

def _to_number(value: Any) -> float | None:
    if value in (None, "", "None"):
        return None
    try:
        return float(str(value).replace(",", "").replace("，", ""))
    except (ValueError, TypeError):
        return None


def _fmt(value: Any) -> str:
    n = _to_number(value)
    if n is None:
        return "-"
    return f"{n:,.2f}"


def build_annual_financial_table_from_all_data(all_data: dict[str, Any]) -> dict[str, Any] | None:
    business = all_data.get("business") or {}
    fin = business.get("financialData") or {}
    if not fin or not fin.get("year"):
        return None

    year = str(fin.get("year"))
    revenue = _to_number(fin.get("revenue"))
    profit = _to_number(fin.get("profit"))
    assets = _to_number(fin.get("assets"))
    liabilities = _to_number(fin.get("liabilities"))
    equity = _to_number(fin.get("equity"))

    columns = [{"key": year, "label": f"{year}年", "unit": ""}]

    def row(label, value):
        return {"label": label, "values": {year: _fmt(value)}, "source": "API-03 financialData", "has_source": True}

    def derived_row(label, compute_fn):
        v = compute_fn()
        return {"label": label, "values": {year: _fmt(v)}, "source": "API-03 financialData (computed)", "has_source": v is not None}

    debt_equity_total = None
    if liabilities is not None and equity is not None:
        debt_equity_total = liabilities + equity

    net_margin = None
    if profit is not None and revenue is not None and revenue != 0:
        net_margin = profit / revenue * 100

    debt_ratio = None
    if liabilities is not None and assets is not None and assets != 0:
        debt_ratio = liabilities / assets * 100

    groups = [
        {
            "title": "利润表",
            "rows": [
                row("营业总收入", revenue),
                {"label": "同比(%)", "values": {year: "-"}, "source": "", "has_source": False},
                row("净利润", profit),
                {"label": "同比(%)", "values": {year: "-"}, "source": "", "has_source": False},
            ],
        },
        {
            "title": "资产负债表",
            "rows": [
                row("资产总计", assets),
                row("负债合计", liabilities),
                {"label": "所有者权益", "values": {"": ""}, "source": "", "has_source": True},
                row("所有者权益总计", equity),
                derived_row("负债和股东权益总计", lambda: debt_equity_total),
            ],
        },
        {
            "title": "盈利能力",
            "rows": [
                derived_row("净利率(%)", lambda: net_margin),
            ],
        },
        {
            "title": "偿还能力",
            "rows": [
                derived_row("资产负债率(%)", lambda: debt_ratio),
            ],
        },
    ]

    return {
        "title": "年报财务摘要",
        "source_interface": "",
        "source_interface_name": "API-03 经营能力 financialData",
        "columns": columns,
        "groups": groups,
    }
