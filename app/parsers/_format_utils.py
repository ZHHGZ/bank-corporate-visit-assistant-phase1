"""纯数据工具函数：ESB 响应提取、记录遍历、字段拼装。

从 customer_report/opinion_excel_docx_renderer.py 迁出，消除 parser → renderer 的反向依赖。
本模块不依赖任何 docx/openpyxl 库，parser 和 renderer 均可安全 import。
"""
from __future__ import annotations

import json
import re
from decimal import Decimal, InvalidOperation
from typing import Any, Mapping, Sequence

from app.utils.text import normalize_text

# ═══════════════════════════════════════════════════════════════════
# 底层工具：递归遍历、键值提取
# ═══════════════════════════════════════════════════════════════════

def _is_success_response(payload: Mapping[str, Any]) -> bool:
    code = _find_first_scalar_by_keys(
        payload,
        ("BsnRetCd", "BsnRetCode", "BsnRspCd", "RspCd", "StCd", "MfcrRetStEncd", "RqstRetEncd"),
    )
    if not code:
        return True
    return code == "200"


def _records_by_key(payload: Any, target_key: str) -> list[Mapping[str, Any]]:
    records: list[Mapping[str, Any]] = []
    if isinstance(payload, Mapping):
        for key, value in payload.items():
            if key == target_key:
                if isinstance(value, list):
                    records.extend(item for item in value if isinstance(item, Mapping))
                elif isinstance(value, Mapping):
                    records.append(value)
            else:
                records.extend(_records_by_key(value, target_key))
    elif isinstance(payload, Sequence) and not isinstance(payload, (str, bytes, bytearray)):
        for item in payload:
            records.extend(_records_by_key(item, target_key))
    return records


def _first_record_by_keys(payload: Any, keys: Sequence[str]) -> Mapping[str, Any]:
    if isinstance(payload, Mapping):
        for key in keys:
            value = payload.get(key)
            if isinstance(value, list) and value and isinstance(value[0], Mapping):
                return value[0]
            if isinstance(value, Mapping):
                return value
        for value in payload.values():
            found = _first_record_by_keys(value, keys)
            if found:
                return found
    elif isinstance(payload, Sequence) and not isinstance(payload, (str, bytes, bytearray)):
        for item in payload:
            found = _first_record_by_keys(item, keys)
            if found:
                return found
    return {}


def _first_value(payload: Mapping[str, Any], keys: Sequence[str]) -> str:
    for key in keys:
        value = payload.get(key)
        if value not in (None, "") and not isinstance(value, (list, dict)):
            return normalize_text(value)
    return ""


def _find_first_scalar_by_keys(payload: Any, keys: Sequence[str]) -> str:
    if isinstance(payload, Mapping):
        for key in keys:
            value = payload.get(key)
            if value not in (None, "") and not isinstance(value, (list, dict)):
                return normalize_text(value)
        for value in payload.values():
            found = _find_first_scalar_by_keys(value, keys)
            if found:
                return found
    elif isinstance(payload, Sequence) and not isinstance(payload, (str, bytes, bytearray)):
        for item in payload:
            found = _find_first_scalar_by_keys(item, keys)
            if found:
                return found
    return ""


def _find_value_by_label(payload: Any, label: str) -> str:
    if payload in (None, ""):
        return ""
    if isinstance(payload, Mapping) and not _is_success_response(payload):
        return ""
    if isinstance(payload, Mapping):
        for key, value in payload.items():
            key_text = normalize_text(key)
            if key_text == label or _normalize_label(key_text) == _normalize_label(label):
                if isinstance(value, Mapping):
                    return normalize_text(value.get("value") or value.get("text") or value.get("name"))
                if not isinstance(value, (list, tuple, dict)):
                    return normalize_text(value)
            if key in {"label", "name", "title", "fieldName", "field_name"} and _normalize_label(value) == _normalize_label(label):
                for value_key in ("value", "text", "result", "content"):
                    found = payload.get(value_key)
                    if found not in (None, "") and not isinstance(found, (list, tuple, dict)):
                        return normalize_text(found)
        for value in payload.values():
            found = _find_value_by_label(value, label)
            if found:
                return found
    elif isinstance(payload, Sequence) and not isinstance(payload, (str, bytes, bytearray)):
        for item in payload:
            found = _find_value_by_label(item, label)
            if found:
                return found
    return ""


def _records_by_path(payload: Any, path: str) -> list[Mapping[str, Any]]:
    current = payload
    parts = [part for part in path.split(".") if part]
    for part_index, raw_part in enumerate(parts):
        is_array = raw_part.endswith("[]")
        key = raw_part[:-2] if is_array else raw_part
        if isinstance(current, list):
            current = current[0] if current else None
        if not isinstance(current, Mapping):
            return []
        current = current.get(key)
        if is_array and part_index < len(parts) - 1:
            if not isinstance(current, list):
                return []
            current = current[0] if current else None
    if isinstance(current, list):
        return [item for item in current if isinstance(item, Mapping)]
    return []


def _extract_record_dicts(payload: Any) -> list[Mapping[str, Any]]:
    records: list[Mapping[str, Any]] = []
    if isinstance(payload, Mapping):
        nested_records: list[Mapping[str, Any]] = []
        for value in payload.values():
            if isinstance(value, (list, Mapping)):
                nested_records.extend(_extract_record_dicts(value))
        if nested_records:
            return nested_records
        if _looks_like_detail_record(payload):
            records.append(payload)
    elif isinstance(payload, Sequence) and not isinstance(payload, (str, bytes, bytearray)):
        for item in payload:
            records.extend(_extract_record_dicts(item))
    return records


def _looks_like_detail_record(value: Mapping[str, Any]) -> bool:
    scalar_count = sum(1 for item in value.values() if item not in (None, "") and not isinstance(item, (list, dict)))
    return scalar_count >= 1


def _count_records(payload: Any) -> int:
    if isinstance(payload, list):
        return len(payload)
    if isinstance(payload, Mapping):
        if not _is_success_response(payload):
            return 0
        counts = []
        for value in payload.values():
            if isinstance(value, list):
                counts.append(len(value))
            elif isinstance(value, Mapping):
                nested = _count_records(value)
                if nested:
                    counts.append(nested)
        return max(counts) if counts else 0
    return 0


# ═══════════════════════════════════════════════════════════════════
# 排序 / 格式化工具
# ═══════════════════════════════════════════════════════════════════

def _year_sort_key(value: str) -> int:
    normalized = normalize_text(value)
    if normalized.isdigit():
        return int(normalized)
    return -1


def _date_sort_key(value: str) -> str:
    text = normalize_text(value)
    match = re.search(r"(\d{4})[-/]?(\d{2})[-/]?(\d{2})", text)
    if match:
        return "".join(match.groups())
    return ""


def _normalize_label(value: Any) -> str:
    return re.sub(r"[\s/／、，,（）()：:]+", "", normalize_text(value)).lower()


def _join_present(separator: str, *values: Any) -> str:
    parts = [normalize_text(value) for value in values if normalize_text(value)]
    return separator.join(parts)


def _join_counts(items: Sequence[tuple[str, Any]]) -> str:
    parts = []
    for label, payload in items:
        count = _count_records(payload)
        if count:
            parts.append(f"{label}{count}条")
    return "、".join(parts)


def _count_text(payload: Any, label: str) -> str:
    count = _count_records(payload)
    return f"{label}{count}条" if count else ""


def _truncate_chinese_text(text: str, limit: int) -> str:
    value = normalize_text(text)
    if not value:
        return ""
    if len(value) <= limit:
        return value
    return value[:limit] + "..."


# ═══════════════════════════════════════════════════════════════════
# 金额格式化
# ═══════════════════════════════════════════════════════════════════

def _to_decimal(value: Any) -> Decimal | None:
    text = normalize_text(value).strip()
    if not text:
        return None
    numeric = re.sub(r"(人民币元|人民币|美元|港元|万元|万|元)$", "", text.replace(",", "")).strip()
    try:
        return Decimal(numeric)
    except (InvalidOperation, ValueError):
        return None


def _format_capital_without_decimal(value: Any) -> str:
    text = normalize_text(value).strip()
    if not text:
        return ""
    decimal = _to_decimal(text)
    if decimal is None:
        return text
    if decimal == decimal.to_integral_value():
        return str(decimal.quantize(Decimal("1")))
    normalized = decimal.normalize()
    return format(normalized, "f").rstrip("0").rstrip(".")


def _is_amount_label(label: str) -> bool:
    normalized = normalize_text(label)
    if "注册 / 实缴资本" in normalized:
        return False
    return any(word in normalized for word in ("金额", "资本", "出资", "数额", "主债权", "标的", "罚款", "罚没", "Amt", "Cpitl", "Lmt", "CntAmt"))


def _format_amount_for_display(value: Any, label: str = "") -> str:
    text = normalize_text(value).strip()
    if not text:
        return ""
    if not _is_amount_label(label):
        return text
    number = _format_capital_without_decimal(text)
    if not number:
        return ""
    if number == text and _to_decimal(text) is None:
        return text
    unit = "元" if "元" in normalize_text(label) and "万元" not in normalize_text(label) else "万"
    return f"{number}{unit}"


def _format_ratio_for_display(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    if "%" in text:
        return text
    decimal = _to_decimal(text)
    if decimal is None:
        return text
    if Decimal("0") < decimal <= Decimal("1"):
        decimal *= Decimal("100")
    if decimal == decimal.to_integral_value():
        return f"{decimal.quantize(Decimal('1'))}%"
    return f"{decimal.quantize(Decimal('0.01')).normalize()}%"


# ═══════════════════════════════════════════════════════════════════
# 字段级拼装函数（parser 直接调用）
# ═══════════════════════════════════════════════════════════════════

def _registered_paid_capital_value(record: Mapping[str, Any]) -> str:
    registered = _format_capital_without_decimal(_first_value(record, ("RgstCpitl", "注册资本")))
    paid = _format_capital_without_decimal(_first_value(record, ("PclupCpitl", "实缴资本")))
    return f"{registered or '-'} / {paid or '-'}"


# ── 客户全貌 ──

def _tax_rating_summary(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    records = _records_by_key(payload, "TaxnEvltArry")
    if not records:
        return ""
    records = sorted(
        records,
        key=lambda record: _year_sort_key(_first_value(record, ("RtgYr",))),
        reverse=True,
    )
    total = _find_first_scalar_by_keys(payload, ("TotCnt",))
    total_count = total or str(len(records))
    lines = [f"共{total_count}条纳税评级记录(最近3条)："]
    for index, record in enumerate(records[:3], start=1):
        year = _first_value(record, ("RtgYr",))
        level = _first_value(record, ("AutLvl", "RtgLvl", "TaxnLvl", "RatingLvl"))
        if year and level:
            lines.append(f"{index}、认证级别{level}（{year}年）")
        elif level:
            lines.append(f"{index}、认证级别{level}")
        elif year:
            lines.append(f"{index}、{year}年")
    return "\n".join(lines)


def _tax_ab_level_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("TaxnEvltArry",))
    if not record:
        return ""
    year = _first_value(record, ("RtgYr",))
    level = _first_value(record, ("AutLvl", "RtgLvl", "TaxnLvl", "RatingLvl"))
    note = _first_value(record, ("RskDtlExpln", "RtgExpln", "RtgCmnt", "RtgDesc"))
    return _join_present("，", f"{year}" if year else "", level, note)


def _customs_credit_summary(payload: Any) -> str:
    rating_record = _first_record_by_keys(payload, ("CstmsCrdtLvlStruct",))
    rating = _first_value(rating_record, ("CstmsCrdtLvlCrdtLvlNm", "CstmsCrdtLvlCrdtLvlEncd"))
    date = _first_value(rating_record, ("CstmsCrdtLvlAfrmTm",))
    certificate = _first_value(rating_record, ("CstmsCrdtLvlAfrmCtfNo",))
    if rating and date:
        return f"{rating}（认定时间：{date}）"
    if rating or date or certificate:
        return _join_present("，", rating, f"认定时间{date}" if date else "", f"证书编号{certificate}" if certificate else "")
    return _find_first_scalar_by_keys(payload, ("CrdtLvl", "CreditLevel", "信用等级"))


def _customs_record_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("CstmsRcdInfStruct",))
    if not record:
        return _find_first_scalar_by_keys(payload, ("CstmsRcdEncd", "LoCstmsNm", "CrpnCgyNm"))
    registration = _first_value(record, ("CstmsRcdEncd", "CstmsRcdNo"))
    customs = _first_value(record, ("LoCstmsNm", "CstmsNm"))
    industry = _first_value(record, ("CrpnCgyNm", "IdyKindNm", "CrpnCgyEncd"))
    valid_date = _first_value(record, ("ClncGdsVldty",))
    return _join_present("，", registration, customs, industry, f"有效期{valid_date}" if valid_date else "")


def _change_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("MdfInfArry", "DataArry", "CrnPgDataArry"))
    if not record:
        return ""
    item = _first_value(record, ("MdfItm", "MdfTp", "变更事项"))
    date = _first_value(record, ("MdfDt", "BtchDt", "变更日期"))
    before = _first_value(record, ("MdfPrevsCntnt", "MdfPrevsObj"))
    after = _first_value(record, ("MdfAfCntnt", "MdfAfObj"))
    summary = _join_present("，", item, date)
    if before or after:
        summary += f"；{before or '-'} -> {after or '-'}"
    count = _count_records(payload)
    return f"变更记录{count}条，{summary}" if summary else f"变更记录{count}条"


def _change_records(payload: Any) -> list[Mapping[str, str]]:
    records = _records_by_key(payload, "MdfInfArry")
    if not records:
        records = _records_by_key(payload, "DataArry")
    if not records:
        records = _records_by_key(payload, "CrnPgDataArry")
    if not records:
        return []

    normalized: list[dict[str, str]] = []
    for record in records:
        date = _first_value(record, ("MdfDt", "BtchDt", "变更日期"))
        item = _first_value(record, ("MdfItm", "MdfTp", "变更事项"))
        before = _truncate_chinese_text(_first_value(record, ("MdfPrevsCntnt", "MdfPrevsObj")), 20)
        after = _truncate_chinese_text(_first_value(record, ("MdfAfCntnt", "MdfAfObj")), 20)
        if not any((date, item, before, after)):
            continue
        normalized.append(
            {
                "变更日期": date,
                "变更事项": item,
                "变更前内容": before,
                "变更后内容": after,
            }
        )
    normalized.sort(key=lambda item: _date_sort_key(item.get("变更日期", "")), reverse=True)
    return normalized


def _zhongzhishutong_qualification_summary(payload: Any) -> str:
    records = _records_by_key(payload, "ScnsTechQualfArry")
    lines: list[str] = []
    for record in records:
        tag = _first_value(record, ("EntpTag",))
        title = _first_value(record, ("ScnsTechQualfTtl",))
        publish_date = _first_value(record, ("ScnsTechQualfPublhDt",))
        if not (tag or title or publish_date):
            continue
        prefix = f"{tag}资质：{title}" if tag or title else ""
        suffix = f"（公布日期：{publish_date}）" if publish_date else ""
        lines.append(f"{prefix} {suffix}".strip())
    return "\n".join(lines)


def _enterprise_certification_summary(payload: Any) -> str:
    field_mappings = (
        ("EntpAcmlnAcqrngcncaAC", "企业累计获得cnca认证数量"),
        ("EntpAcmlnAcqrngHgntchLEAC", "企业累计获得高新企业认证数量"),
        ("EntpAcmlnAcqrngNLNTAC", "企业累计获得无线网电信认证数量"),
        ("EntpAcmlnAcqrngEnvPrtctAC", "企业累计获得环保认证数量"),
        ("EntpAcmlnAcqrngFinAC", "企业累计获得金融认证数量"),
        ("EntpAcmlnAcqrngFdAC", "企业累计获得食品认证数量"),
        ("EntpAcmlnAcqrngScrRelAC", "企业累计获得涉密认证数量"),
        ("EntpAcmlnAcqrngDlSftAC", "企业累计获得双软认证数量"),
        ("EntpAcmlnAcqrngCstmsAC", "企业累计获得海关认证数量"),
        ("EntpAcmlnAcqrngTenBgAC", "企业累计获得十大认证数量"),
        ("EntpNrOneYrAcqrngWebRcdCnt", "企业近一年获得网站备案数量"),
        ("EntpNrDblYrAcqrngWebRcdCnt", "企业近两年获得网站备案数量"),
        ("EntpNrThreeYrAcqrngWebRcdCnt", "企业近三年获得网站备案数量"),
        ("EntpAcmlnAcqrngWebRcdInf", "企业累计获得网站备案信息"),
        ("ROWebRcdTmDstnSttcTmPntDrtn", "最近一次网站备案时间距统计时点时长"),
        ("ROTaxPayLvl", "最近一次纳税等级"),
        ("ROCstmsRtg", "最近一次海关评级"),
        ("RORtdsGzllEntpYr", "最近一次被评为瞪羚企业的年份"),
        ("RORtdsUncrnEntpYr", "最近一次被评为独角兽企业的年份"),
        ("RORtdsSeedUncrnEntpYr", "最近一次被评为种子独角兽企业的年份"),
        ("RORtdsPtnlUncrnEntpYr", "最近一次被评为潜在独角兽企业的年份"),
        ("RORtdsCltvtGzllEntpYr", "最近一次被评为培育瞪羚企业的年份"),
        ("WhthrInclLtlGntEntpList", "是否列入小巨人企业名单"),
    )
    lines: list[str] = []
    for field_name, label in field_mappings:
        value = _find_first_scalar_by_keys(payload, (field_name,))
        if not value:
            continue
        lines.append(f"{len(lines) + 1}、{label}：{value}")
    return "\n".join(lines)


def _green_industry_summary(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    flag = _find_first_scalar_by_keys(payload, ("WhthrGreenIndus",))
    if not flag:
        return ""
    if flag == "否":
        return "否"
    if flag == "是":
        first_record = _first_record_by_keys(payload, ("GreenListEntpAfrmListArry",))
        catalog = "-".join(
            part
            for part in (
                _first_value(first_record, ("GreenIndusDrctnCtlgOneLvl",)),
                _first_value(first_record, ("GreenIndusDrctnCtlgSecdy",)),
                _first_value(first_record, ("GreenIndusDrctnCtlgThreeLvl",)),
            )
            if part
        )
        if catalog:
            return f"是，推荐绿色产业目录：{catalog}"
        return "是"
    return flag


def _small_micro_catalog_summary(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    return _find_first_scalar_by_keys(payload, ("WhthrAbsrbSmlMcrListEntpCtlg",))


def _credit_rating_summary_for_catalog(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    records = _records_by_key(payload, "CrdtEvltListArry")
    if not records:
        return ""
    records = sorted(
        records,
        key=lambda record: (
            _year_sort_key(_first_value(record, ("EvltAnul",))),
            _date_sort_key(_first_value(record, ("EvltDt",))),
        ),
        reverse=True,
    )
    lines: list[str] = []
    lines.append(f"共{len(records)}条评级记录（展示最近3条评级记录）")
    for index, record in enumerate(records[:3], start=1):
        year = _first_value(record, ("EvltAnul",))
        qualf_name = _first_value(record, ("QualfNm",))
        credit_level = _first_value(record, ("CrdtLvl",)) or _first_value(record, ("EvltRslt",))
        eval_date = _first_value(record, ("EvltDt",))
        parts = []
        if year:
            parts.append(f"{year}年")
        if qualf_name:
            parts.append(qualf_name)
        if credit_level:
            parts.append(credit_level)
        main_text = "-".join(parts) if parts else "-"
        if eval_date:
            lines.append(f"{index}、{main_text}（评价日期：{eval_date}）")
        else:
            lines.append(f"{index}、{main_text}")
    return "\n".join(lines)


def _billboard_summary_for_catalog(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    records = _records_by_key(payload, "BllbrdListArry")
    if not records:
        return ""
    records = sorted(
        records,
        key=lambda record: (
            _year_sort_key(_first_value(record, ("BllbrdYr",))),
            _date_sort_key(_first_value(record, ("RelsDt",))),
        ),
        reverse=True,
    )
    lines: list[str] = []
    lines.append(f"共{len(records)}条榜单记录（展示最近3条榜单记录）")
    for index, record in enumerate(records[:3], start=1):
        name = _first_value(record, ("BllbrdNm",)) or "-"
        release_date = _first_value(record, ("RelsDt",))
        if release_date:
            lines.append(f"{index}、{name}（发布日期：{release_date}）")
        else:
            lines.append(f"{index}、{name}")
    return "\n".join(lines)


def _invent_patent_records(payload: Any) -> list[Mapping[str, Any]]:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return []
    records = _records_by_key(payload, "PtntBscInfStruct")
    if records:
        return records
    return _extract_record_dicts(payload)


def _invent_patent_summary(payload: Any) -> str:
    records = _invent_patent_records(payload)
    if not records:
        count = _count_records(payload)
        return f"共{count}条有效发明专利" if count else ""

    records = sorted(records, key=lambda record: _date_sort_key(_first_value(record, ("PtntBscInfAppDt", "REGDATE", "AppDt"))), reverse=True)
    latest = records[:3]
    lines = [f"共{len(records)}条有效发明专利，", "最新3条专利："]
    for index, record in enumerate(latest, start=1):
        name = _first_value(record, ("PtntBscInfPtntNm", "MtchNmHghlght", "PtntBscInfOpnNo")) or "-"
        app_date = _first_value(record, ("PtntBscInfAppDt", "REGDATE", "AppDt")) or "-"
        lines.append(f"{index}、{name}（申请日期：{app_date}）")
    return "\n".join(lines)


# ── 股权关联 ──

def _equity_change_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("StkRghtMdfInfArry",))
    if not record:
        return ""
    name = _first_value(record, ("ShrhlrNm",))
    date = _first_value(record, ("StkRghtMdfDt",))
    before = _first_value(record, ("TfrPrevsStkRghtRto",))
    after = _first_value(record, ("TfrAfStkRghtRto",))
    count = _count_records(payload)
    return f"股权变更{count}条，{name}于{date}持股比例{before}->{after}"


def _shareholder_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("ShrhlrInfArry",))
    if not record:
        return ""
    name = _first_value(record, ("ShrhlrNm",))
    shareholder_type = _first_value(record, ("ShrhlrTp", "EntpRgstNo"))
    ratio = _first_value(record, ("CtbRto",))
    amount = _format_amount_for_display(_first_value(record, ("SbcbCtbLmt",)), "投资数额")
    date = _first_value(record, ("CtbDt",))
    count = _count_records(payload)
    detail = _join_present("，", name, shareholder_type, f"投资比例{ratio}" if ratio else "", f"投资数额{amount}" if amount else "", f"出资时间{date}" if date else "")
    return f"股东信息{count}条，{detail}" if detail else f"股东信息{count}条"


def _outbound_investment_summary(payload: Any) -> str:
    record = _first_record_by_keys(payload, ("EntpExtInvsInfArry", "LPExtInvsArry"))
    if not record:
        return ""
    name = _first_value(record, ("EntpNm",))
    status = _first_value(record, ("EntpSt",))
    capital = _format_amount_for_display(_first_value(record, ("RgstCpitl", "RgstdCptl")), "注册资本")
    ratio = _first_value(record, ("CtbRto",))
    count = _count_records(payload)
    detail = _join_present("，", name, f"经营状态{status}" if status else "", f"注册资本{capital}" if capital else "", f"持股比例{ratio}" if ratio else "")
    return f"对外投资{count}条，{detail}" if detail else f"对外投资{count}条"


def _legal_person_office_summary(interface_responses: Mapping[str, Any]) -> str:
    record = _first_record_by_keys(interface_responses.get("9000000016"), ("LglRprsOthrCoTakeOfcArry",))
    if not record:
        return _outbound_investment_summary(interface_responses.get("9000000014"))
    name = _first_value(record, ("EntpBrchNm",))
    status = _first_value(record, ("EntpSt",))
    position = _first_value(record, ("Pos",))
    capital = _format_amount_for_display(_first_value(record, ("RgstdCptl",)), "注册资本")
    legal_person = _first_value(record, ("LglRprsNm", "WhthrLglRprs"))
    count = _count_records(interface_responses.get("9000000016"))
    detail = _join_present("，", name, f"经营状态{status}" if status else "", f"注册资本{capital}" if capital else "", f"任职{position}" if position else "", legal_person)
    return f"法人跨企业任职{count}条，{detail}" if detail else f"法人跨企业任职{count}条"


# ── 经营能力 / 资金需求（collateral / financing 拼装）──

def _collateral_summary(interface_responses: Mapping[str, Any], label: str, codes: Sequence[str]) -> str:
    count = sum(_count_records(interface_responses.get(code)) for code in codes)
    if not count:
        return ""
    if label == "股权出质":
        record = _first_record_by_keys(interface_responses.get("9000000020"), ("StkRghtPlgBscInfArry",))
        pledgee = _first_value(record, ("PlgeNm",))
        amount = _format_amount_for_display(_first_value(record, ("PlgStkRghtCntAmt",)), "出质金额")
        status = _first_value(record, ("StkRghtPlgSt",))
        return f"{label}{count}条，质权人{pledgee or '-'}，出质金额{amount or '-'}，状态{status or '-'}"
    record = _first_record_by_keys(interface_responses.get("9000000023"), ("ChtlMrtgRgstInfArry",))
    amount = _format_amount_for_display(_first_value(record, ("BeGrntOblgtryCntAmt",)), "被担保债权金额")
    date = _first_value(record, ("RgstDt",))
    return f"{label}{count}条，被担保债权金额{amount or '-'}，登记日期{date or '-'}"


def _financing_status_summary(payload: Any) -> str:
    if not payload or (isinstance(payload, Mapping) and not _is_success_response(payload)):
        return ""
    records = _records_by_key(payload, "FincngEvArry")
    if not records:
        return ""
    records = sorted(records, key=lambda record: _date_sort_key(_first_value(record, ("PublcDt", "InvsTm"))), reverse=True)
    latest = records[:3]
    lines = [f"共{len(records)}条融资事件，", "最新3条记录："]
    for index, record in enumerate(latest, start=1):
        event = _first_value(record, ("FincngEvFullNm",))
        date = _first_value(record, ("PublcDt", "InvsTm"))
        amount = _format_amount_for_display(_first_value(record, ("InvsAmt",)), "投资金额")
        lines.append(f"{index}、{event or '-'}（日期：{date or '-'}，金额：{amount or '-'}）")
    return "\n".join(lines)
