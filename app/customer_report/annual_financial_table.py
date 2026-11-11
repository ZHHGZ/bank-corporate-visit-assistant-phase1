"""Build Zhongshu annual financial statement tables from ESB responses."""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Callable, Mapping


ZHONGSHU_ANNUAL_REPORT_TXN = "9000000018"

ValueFunc = Callable[[dict[str, Any], list[dict[str, Any]], int], Any]


def build_zhongshu_annual_financial_table(
    responses: Mapping[str, Any],
    *,
    display_limit: int = 5,
    fill_missing_years: bool = True,
) -> dict[str, Any] | None:
    """Return a PDF/front-end friendly annual financial table.

    Only Zhongshu interface EDS0608/9000000018 is used.
    """
    response = responses.get(ZHONGSHU_ANNUAL_REPORT_TXN)
    if not isinstance(response, Mapping):
        return None

    records = _annual_records(response)
    if not records:
        return None
    display_records = _display_records(records, limit=display_limit, fill_missing_years=fill_missing_years)

    columns = [
        {"key": item["year"], "label": f"{item['year']}年", "unit": ""}
        for item in display_records
    ]
    groups = []
    for group in _row_groups():
        rows = []
        for row in group["rows"]:
            label = row["label"] if row["has_source"] else f"{row['label']}（无数据来源）"
            rows.append(
                {
                    "label": label,
                    "values": {
                        item["year"]: _format_value(row["value"](item, display_records, index), row["kind"])
                        for index, item in enumerate(display_records)
                    },
                    "source": row["source"],
                    "has_source": row["has_source"],
                }
            )
        groups.append({"title": group["title"], "rows": rows})

    return {
        "title": "中数年报财务摘要",
        "source_interface": ZHONGSHU_ANNUAL_REPORT_TXN,
        "source_interface_name": "中数智汇-企业查询年报信息(大报文)",
        "columns": columns,
        "groups": groups,
    }


def _annual_records(response: Mapping[str, Any]) -> list[dict[str, Any]]:
    base_rows = [
        row for row in path_values(response, "BODY.RsltStruct.EntpAnulRprtBscInfArry[]")
        if isinstance(row, Mapping)
    ]
    asset_rows = [
        row for row in path_values(response, "BODY.RsltStruct.AstSituInfArry[]")
        if isinstance(row, Mapping)
    ]
    year_by_key = {
        str(row.get("BscInfAnulRprtMjrKey") or ""): _normalize_year(row.get("AnulRprtYr"))
        for row in base_rows
        if row.get("BscInfAnulRprtMjrKey") and _normalize_year(row.get("AnulRprtYr"))
    }

    records: list[dict[str, Any]] = []
    for index, asset_row in enumerate(asset_rows):
        key = str(asset_row.get("AstSituAnulRprtMjrKey") or "")
        year = year_by_key.get(key)
        if not year and index < len(base_rows):
            year = _normalize_year(base_rows[index].get("AnulRprtYr"))
        if not year:
            continue
        records.append({"year": year, "asset": dict(asset_row), "base_key": key})

    deduped: dict[str, dict[str, Any]] = {}
    for record in records:
        deduped.setdefault(record["year"], record)
    return sorted(deduped.values(), key=lambda item: _year_sort_key(item["year"]), reverse=True)


def _display_records(records: list[dict[str, Any]], *, limit: int = 5, fill_missing_years: bool = True) -> list[dict[str, Any]]:
    if not fill_missing_years:
        return records[:limit]
    by_year = {record["year"]: record for record in records}
    numeric_years = [_year_sort_key(record["year"]) for record in records if _year_sort_key(record["year"])]
    if numeric_years:
        start_year = max(numeric_years)
        return [
            by_year.get(str(year), {"year": str(year), "asset": {}, "base_key": ""})
            for year in range(start_year, start_year - limit, -1)
        ]
    return records[:limit]


def _row_groups() -> list[dict[str, Any]]:
    return [
        {
            "title": "利润表",
            "rows": [
                _source_row("营业总收入", "amount", "AstSituBsnsTotIncm", "BODY.RsltStruct.AstSituInfArry[].AstSituBsnsTotIncm"),
                _derived_row("同比(%)", "percent", lambda item, records, index: _yoy(records, index, "AstSituBsnsTotIncm"), "营业总收入同比"),
                _source_row("净利润", "amount", "AstSituNetPrft", "BODY.RsltStruct.AstSituInfArry[].AstSituNetPrft"),
                _derived_row("同比(%)", "percent", lambda item, records, index: _yoy(records, index, "AstSituNetPrft"), "净利润同比"),
            ],
        },
        {
            "title": "资产负债表",
            "rows": [
                _source_row("资产总计", "amount", "AstSituAstTotAmt", "BODY.RsltStruct.AstSituInfArry[].AstSituAstTotAmt"),
                _source_row("负债总计", "amount", "AstSituLbyTotAmt", "BODY.RsltStruct.AstSituInfArry[].AstSituLbyTotAmt"),
                _heading_row("所有者权益"),
                _source_row("所有者权益总计", "amount", "AstSituOwnrEqtTot", "BODY.RsltStruct.AstSituInfArry[].AstSituOwnrEqtTot"),
                _derived_row("负债和股东权益总计", "amount", lambda item, records, index: _sum_fields(item, "AstSituLbyTotAmt", "AstSituOwnrEqtTot"), "AstSituLbyTotAmt+AstSituOwnrEqtTot"),
            ],
        },
        {
            "title": "盈利能力",
            "rows": [
                _derived_row("净利率(%)", "percent", lambda item, records, index: _ratio(item, "AstSituNetPrft", "AstSituBsnsTotIncm"), "净利润/营业总收入"),
            ],
        },
        {
            "title": "偿还能力",
            "rows": [
                _derived_row("资产负债率(%)", "percent", lambda item, records, index: _ratio(item, "AstSituLbyTotAmt", "AstSituAstTotAmt"), "负债总计/资产合计"),
            ],
        },
    ]


def _source_row(label: str, kind: str, field: str, source: str) -> dict[str, Any]:
    return {
        "label": label,
        "kind": kind,
        "source": source,
        "has_source": True,
        "value": lambda item, records, index: item["asset"].get(field),
    }


def _derived_row(label: str, kind: str, value: ValueFunc, source: str) -> dict[str, Any]:
    return {"label": label, "kind": kind, "source": source, "has_source": True, "value": value}


def _heading_row(label: str) -> dict[str, Any]:
    return {"label": label, "kind": "heading", "source": "", "has_source": True, "value": lambda item, records, index: ""}


def _sum_fields(item: dict[str, Any], *fields: str) -> Decimal | None:
    values = [_decimal(item["asset"].get(field)) for field in fields]
    if any(value is None for value in values):
        return None
    return sum((value for value in values if value is not None), Decimal("0"))


def _yoy(records: list[dict[str, Any]], index: int, field: str) -> Decimal | None:
    current = _decimal(records[index]["asset"].get(field))
    if current is None:
        return None
    for previous in records[index + 1:]:
        previous_value = _decimal(previous["asset"].get(field))
        if previous_value not in (None, Decimal("0")):
            return (current - previous_value) / previous_value * Decimal("100")
    return None


def _ratio(item: dict[str, Any], numerator_field: str, denominator_field: str) -> Decimal | None:
    numerator = _decimal(item["asset"].get(numerator_field))
    denominator = _decimal(item["asset"].get(denominator_field))
    if numerator is None or denominator in (None, Decimal("0")):
        return None
    return numerator / denominator * Decimal("100")


def _decimal(value: Any) -> Decimal | None:
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, ValueError):
        return None


def _format_value(value: Any, kind: str) -> str:
    if kind == "heading":
        return ""
    number = _decimal(value)
    if number is None:
        return "-"
    if kind in {"amount", "percent"}:
        return f"{number.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):,}"
    return str(value)


def _year_sort_key(year: str) -> int:
    try:
        return int(year)
    except ValueError:
        return 0


def _normalize_year(value: Any) -> str:
    text = str(value or "").strip()
    match = re.search(r"(?:19|20)\d{2}", text)
    return match.group(0) if match else ""


def path_values(data: Any, path: str) -> list[Any]:
    if not path:
        return []
    parts = path.split(".")
    values = [data]
    for part in parts:
        is_array = part.endswith("[]")
        key = part[:-2] if is_array else part
        next_values = []
        for value in values:
            if isinstance(value, Mapping) and key in value:
                child = value[key]
                if is_array:
                    if isinstance(child, list):
                        next_values.extend(child)
                    elif isinstance(child, Mapping):
                        next_values.append(child)
                else:
                    next_values.append(child)
        values = next_values
        if not values:
            return []
    return values
