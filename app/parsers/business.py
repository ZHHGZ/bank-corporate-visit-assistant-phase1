"""经营能力（API-03）统一解析：结构化数据 + display + tables（4个二级模块表格）."""
from __future__ import annotations

from collections import defaultdict
from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.parsers._esb_records import records
from app.parsers._format_utils import _count_text


def parse_business_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    data: dict[str, Any] = {"financialData": {}}

    if "9000000018" in responses and is_success(responses["9000000018"]):
        body = get_body(responses["9000000018"])
        report_rows = records(body, "EntpAnulRprtBscInfArry")
        assets = records(body, "AstSituInfArry")
        latest_report = max(report_rows, key=lambda r: pick(r, "AnulRprtYr"), default={})
        latest_key = pick(latest_report, "BscInfAnulRprtMjrKey")
        asset = next(
            (row for row in assets if pick(row, "AstSituAnulRprtMjrKey") == latest_key),
            assets[0] if assets else {},
        )
        data["financialData"] = {
            "year": pick(latest_report, "AnulRprtYr"),
            "revenue": pick(asset, "AstSituBsnsTotIncm"),
            "profit": pick(asset, "AstSituNetPrft"),
            "assets": pick(asset, "AstSituAstTotAmt"),
            "liabilities": pick(asset, "AstSituLbyTotAmt"),
            "equity": pick(asset, "AstSituOwnrEqtTot"),
        }

    data["display"] = build_business_display(data, responses)
    data["tables"] = build_business_tables(responses)
    return data


# ═══════════════════════════════════════════════════════════════════
# display（展示文案）
# ═══════════════════════════════════════════════════════════════════

def build_business_display(data: Mapping[str, Any], responses: Mapping[str, Any]) -> dict[str, str]:
    text: dict[str, str] = {}
    fin_data = data.get("financialData") or {}
    revenue = fin_data.get("revenue")
    if revenue not in (None, "", "None"):
        text["近 1-3 年营收数据及变动情况"] = str(revenue)

    bidding_text = _count_text(responses.get("9000000019"), "招投标")
    if bidding_text:
        text["企业招投标信息"] = bidding_text

    mortgage_text = _count_text(responses.get("9000000023"), "动产抵押")
    if mortgage_text:
        text["动产抵押"] = mortgage_text

    pledge_text = _count_text(responses.get("9000000020"), "股权出质")
    if pledge_text:
        text["股权出质"] = pledge_text

    fin_text = _count_text(responses.get("9000000029"), "历史融资")
    if fin_text:
        text["融资情况"] = fin_text

    return text


# ═══════════════════════════════════════════════════════════════════
# 4 个二级模块表格
# ═══════════════════════════════════════════════════════════════════

def _build_bidding_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """企业招投标信息 — 9000000019."""
    if "9000000019" not in responses or not is_success(responses["9000000019"]):
        return []
    body = get_body(responses["9000000019"])
    biddings = records(body, "BidInfArry")
    result: list[dict[str, str]] = []
    for item in biddings:
        record: dict[str, str] = {}
        if pick(item, "BidTtl"):
            record["标题"] = pick(item, "BidTtl")
        if pick(item, "BidClsf"):
            record["招投标分类"] = pick(item, "BidClsf")
        if pick(item, "AncmTp"):
            record["公告类型"] = pick(item, "AncmTp")
        if pick(item, "PrchsPrsnNm"):
            record["采购人"] = pick(item, "PrchsPrsnNm")
        if pick(item, "CallBidAgncBrch"):
            record["招标代理机构"] = pick(item, "CallBidAgncBrch")
        if pick(item, "IssuTm"):
            record["发布时间"] = pick(item, "IssuTm")
        if record:
            result.append(record)
    return result


def _build_mortgage_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """动产抵押 — 0233/0236/0237/0238 按 RgstNo 聚合."""
    reg_map: dict[str, dict[str, str]] = defaultdict(dict)

    if "9000000023" in responses and is_success(responses["9000000023"]):
        for item in records(get_body(responses["9000000023"]), "ChtlMrtgRgstInfArry"):
            reg_no = pick(item, "RgstNo")
            if not reg_no:
                continue
            reg_map[reg_no].update({
                "登记编号": pick(item, "RgstNo"),
                "被担保债权数额": pick(item, "BeGrntOblgtryCntAmt"),
                "被担保债券种类": pick(item, "BeGrntBondKind"),
                "担保范围": pick(item, "GrntScop"),
                "履行债务结束日期": pick(item, "PrfmDbtEndDt"),
            })

    if "9000000026" in responses and is_success(responses["9000000026"]):
        for item in records(get_body(responses["9000000026"]), "CMBeGrntMainOblgtryArry"):
            reg_no = pick(item, "RgstNo")
            if not reg_no:
                continue
            reg_map[reg_no]["债务数额"] = pick(item, "PrfmCntAmt")

    if "9000000027" in responses and is_success(responses["9000000027"]):
        for item in records(get_body(responses["9000000027"]), "ChtlMrtgCltlInfArry"):
            reg_no = pick(item, "RgstNo")
            if not reg_no:
                continue
            reg_map[reg_no]["抵押物名称"] = pick(item, "CltlNm")

    if "9000000028" in responses and is_success(responses["9000000028"]):
        for item in records(get_body(responses["9000000028"]), "ChtlMrtgPldgrInfArry"):
            reg_no = pick(item, "RgstNo")
            if not reg_no:
                continue
            reg_map[reg_no]["抵押权人名称"] = pick(item, "MrtgPldgrNm")

    column_order = ["登记编号", "被担保债权数额", "被担保债券种类", "担保范围",
                    "履行债务结束日期", "债务数额", "抵押物名称", "抵押权人名称"]
    result: list[dict[str, str]] = []
    for reg_no in reg_map:
        record: dict[str, str] = {col: reg_map[reg_no].get(col, "") for col in column_order}
        result.append(record)
    return result


def _build_pledge_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """股权出质 — 9000000020."""
    if "9000000020" not in responses or not is_success(responses["9000000020"]):
        return []
    body = get_body(responses["9000000020"])
    pledges = records(body, "StkRghtPlgBscInfArry")
    result: list[dict[str, str]] = []
    for item in pledges:
        record: dict[str, str] = {}
        if pick(item, "RgstNo"):
            record["登记编号"] = pick(item, "RgstNo")
        if pick(item, "PlgPrsnNm"):
            record["出质人名称"] = pick(item, "PlgPrsnNm")
        if pick(item, "StkRghtPlgSt"):
            record["出质状态"] = pick(item, "StkRghtPlgSt")
        if pick(item, "PlgStkRghtCntAmt"):
            record["出质股权数额"] = pick(item, "PlgStkRghtCntAmt")
        if pick(item, "RgstDt"):
            record["出质登记日期"] = pick(item, "RgstDt")
        if record:
            result.append(record)
    return result


def _build_financing_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """历史融资金额 — 9000000029."""
    if "9000000029" not in responses or not is_success(responses["9000000029"]):
        return []
    body = get_body(responses["9000000029"])
    financings = records(body, "FincngEvArry")
    result: list[dict[str, str]] = []
    for item in financings:
        record: dict[str, str] = {}
        if pick(item, "FincngEvFullNm"):
            record["融资事件全称"] = pick(item, "FincngEvFullNm")
        if pick(item, "InvsAmt"):
            record["投资金额"] = pick(item, "InvsAmt")
        if pick(item, "CcyCN"):
            record["币种中文"] = pick(item, "CcyCN")
        if pick(item, "InvsAmtLvl"):
            record["投资金额级别"] = pick(item, "InvsAmtLvl")
        if pick(item, "FincrNmFullNm"):
            record["融资方名称全称"] = pick(item, "FincrNmFullNm")
        if pick(item, "InvsrNmFullNm"):
            record["投资方名称全称"] = pick(item, "InvsrNmFullNm")
        if pick(item, "InvsTm"):
            record["投资时间"] = pick(item, "InvsTm")
        if record:
            result.append(record)
    return result


def build_business_tables(responses: Mapping[str, Any]) -> dict[str, list[dict[str, str]]]:
    tables: dict[str, list[dict[str, str]]] = {}

    bidding = _build_bidding_table(responses)
    if bidding:
        tables["bid_records"] = bidding

    mortgage = _build_mortgage_table(responses)
    if mortgage:
        tables["mortgage_records"] = mortgage

    pledge = _build_pledge_table(responses)
    if pledge:
        tables["pledge_records"] = pledge

    financing = _build_financing_table(responses)
    if financing:
        tables["financing_records"] = financing

    return tables
