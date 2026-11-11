"""风险提示（API-06）统一解析：结构化数据 + display + tables（3张二级模块表格）."""
from __future__ import annotations

from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.parsers._esb_records import records

RISK_ARRAYS: dict[str, tuple[str, ...]] = {
    "9000000036": ("EntpJdgntWritInfArry",),
    "9000000037": ("OpnCrtSnAncmArry",),
    "9000000038": ("CrtAncmArry",),
    "9000000039": ("RstrctCnsmnBeExecPrsn",),
    "9000000040": ("BeExecPrsnInfArry",),
    "9000000041": ("ExecInfArry",),
    "9000000042": ("SvrIlglLsCrdtEntpListArry",),
    "9000000043": ("CstmsLsCrdtEntpListArry",),
    "9000000044": ("GovtLsCrdtEntpListArry",),
    "9000000045": ("DZLsCrdtEntpListArry",),
    "9000000046": ("AnlDataArry",),
    "9000000047": ("TaxnAdmntvPnltyArry",),
    "9000000048": ("MjrRevIlglArry",),
    "9000000049": ("IdyCmrcAdmntvPnltyArry",),
    "9000000050": ("EnvPrtctAdmntvPnltyArry",),
    "9000000051": ("CrpnAbnArry",),
    "9000000052": ("JdctrHlpArry",),
    "9000000053": ("RfrFinBlklistArry",),
    "9000000054": ("PnltyEntpInfArry",),
    "9000000055": ("ClsCaseArray",),
    "9000000056": (
        "CrtAncmListArry", "JdctrCaseListArry", "JdgntWritListArry",
        "BnkrPcdrListArry", "ClrgInfArry", "LandMrtgArry",
        "StkRghtPlgArry", "AdmntvPnltyArry", "LosCrdtBlklistArry",
    ),
}


def parse_risk_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    total_risk = 0
    red_flags = 0

    for code, array_names in RISK_ARRAYS.items():
        resp = responses.get(code)
        if not resp or not is_success(resp):
            continue
        count = sum(len(records(get_body(resp), name)) for name in array_names)
        total_risk += count
        if code in ("9000000040", "9000000042", "9000000048", "9000000051"):
            red_flags += count

    risk_level = "高风险" if red_flags > 0 or total_risk > 10 else ("中风险" if total_risk > 5 else "低风险")

    litigation = records(get_body(responses.get("9000000036", {})), "EntpJdgntWritInfArry")
    dishonest = records(get_body(responses.get("9000000040", {})), "BeExecPrsnInfArry")
    penalties = records(get_body(responses.get("9000000049", {})), "IdyCmrcAdmntvPnltyArry")
    abnormalities = records(get_body(responses.get("9000000051", {})), "CrpnAbnArry")

    data = {
        "riskSummary": {
            "totalRiskCount": total_risk,
            "riskLevel": risk_level,
            "redFlags": red_flags,
        },
        "litigationInfo": {
            "count": len(litigation),
            "records": [
                {
                    "caseNumber": pick(item, "ExecCaseNo"),
                    "caseReason": pick(item, "RgstrRsn"),
                    "caseType": pick(item, "CaseTpNm"),
                    "role": pick(item, "RoleNm"),
                    "closeTime": pick(item, "AudEndTm"),
                }
                for item in litigation[:10]
            ],
        },
        "dishonestInfo": {"count": len(dishonest)},
        "penaltyInfo": {
            "count": len(penalties),
            "records": [
                {
                    "documentNumber": pick(item, "PnlDcsnDocNo"),
                    "reason": pick(item, "RskDtlExpln") or pick(item, "PnlRsn"),
                    "content": pick(item, "PnlCntnt") or pick(item, "RskDtlExpln"),
                    "authority": pick(item, "PnlDcsnInst") or pick(item, "ExecBssMkUnit"),
                    "date": pick(item, "RelsDt") or pick(item, "DataDt"),
                }
                for item in penalties[:10]
            ],
        },
        "businessAbnormality": {"count": len(abnormalities)},
    }
    data["display"] = build_risk_display(data, responses)
    data["tables"] = build_risk_tables(responses)
    return data


def build_risk_display(data: Mapping[str, Any], responses: Mapping[str, Any]) -> dict[str, str]:
    _ = responses
    dishonest = data.get("dishonestInfo") or {}
    business_abnormal = data.get("businessAbnormality") or {}
    litigation = data.get("litigationInfo") or {}
    penalty = data.get("penaltyInfo") or {}

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

    exec_count = _count_risk_array(responses, "9000000041", "ExecInfArry")
    limit_count = _count_risk_array(responses, "9000000039", "RstrctCnsmnBeExecPrsn")
    discredit_count = _count_risk_array(responses, "9000000040", "BeExecPrsnInfArry")

    return {
        "企业被执行人": f"被执行/限消记录{dc}条" if dc else "被执行/限消记录0条",
        "异常经营": f"经营异常{bc}条" if bc else "经营异常0条",
        "法律诉讼信息": "、".join(litigation_parts) if litigation_parts else "暂无诉讼记录",
        "执行风险": f"执行风险（共{exec_count}条）" if exec_count else "执行风险（共0条）",
        "限制高消费": f"限制高消费（共{limit_count}条）" if limit_count else "限制高消费（共0条）",
        "失信执行人": f"失信执行人（共{discredit_count}条）" if discredit_count else "失信执行人（共0条）",
    }


def _count_risk_array(responses: Mapping[str, Any], code: str, array_name: str) -> int:
    from app.core.esb_client import get_body, is_success
    from app.parsers._esb_records import records
    if code not in responses or not is_success(responses[code]):
        return 0
    return len(records(get_body(responses[code]), array_name))


# ═══════════════════════════════════════════════════════════════════
# 3 张二级模块表格
# ═══════════════════════════════════════════════════════════════════

def _build_execution_risk_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """执行风险 — 9000000041."""
    if "9000000041" not in responses or not is_success(responses["9000000041"]):
        return []
    body = get_body(responses["9000000041"])
    items = records(body, "ExecInfArry")
    result: list[dict[str, str]] = []
    for item in items:
        record: dict[str, str] = {}
        if pick(item, "BeExecPrsnNm"):
            record["被执行人名称"] = pick(item, "BeExecPrsnNm")
        if pick(item, "CaseSt"):
            record["案件状态"] = pick(item, "CaseSt")
        if pick(item, "ExecCaseNo"):
            record["案号"] = pick(item, "ExecCaseNo")
        if pick(item, "AppExecSbjMtr"):
            record["申请执行标的（元）"] = pick(item, "AppExecSbjMtr")
        if pick(item, "ExecCrtNm"):
            record["执行法院"] = pick(item, "ExecCrtNm")
        if pick(item, "CrtRgstDt"):
            record["立案时间"] = pick(item, "CrtRgstDt")
        if record:
            result.append(record)
    return result


def _build_consumption_limit_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """限制高消费 — 9000000039."""
    if "9000000039" not in responses or not is_success(responses["9000000039"]):
        return []
    body = get_body(responses["9000000039"])
    items = records(body, "RstrctCnsmnBeExecPrsn")
    result: list[dict[str, str]] = []
    for item in items:
        record: dict[str, str] = {}
        if pick(item, "JdgntCaseNo"):
            record["判决案号"] = pick(item, "JdgntCaseNo")
        if pick(item, "RelsDt"):
            record["发布日期"] = pick(item, "RelsDt")
        if pick(item, "CrtRgstDt"):
            record["立案日期"] = pick(item, "CrtRgstDt")
        if pick(item, "BeExecPrsnNm"):
            record["被执行人"] = pick(item, "BeExecPrsnNm")
        if pick(item, "ExecCrtCd"):
            record["执行法院"] = pick(item, "ExecCrtCd")
        if pick(item, "OIBC"):
            record["组织机构代码"] = pick(item, "OIBC")
        if pick(item, "WhthrVld"):
            record["是否有效"] = pick(item, "WhthrVld")
        if record:
            result.append(record)
    return result


def _build_discredit_executor_table(responses: Mapping[str, Any]) -> list[dict[str, str]]:
    """失信执行人 — 9000000040."""
    if "9000000040" not in responses or not is_success(responses["9000000040"]):
        return []
    body = get_body(responses["9000000040"])
    items = records(body, "BeExecPrsnInfArry")
    result: list[dict[str, str]] = []
    for item in items:
        record: dict[str, str] = {}
        if pick(item, "ExecCaseNo"):
            record["案号"] = pick(item, "ExecCaseNo")
        if pick(item, "CrtRgstDt"):
            record["立案时间"] = pick(item, "CrtRgstDt")
        if pick(item, "PrfmSt"):
            record["被执行人的履行情况"] = pick(item, "PrfmSt")
        if pick(item, "ExecCrtNm"):
            record["执行法院"] = pick(item, "ExecCrtNm")
        if pick(item, "CaseSt"):
            record["案件状态"] = pick(item, "CaseSt")
        if record:
            result.append(record)
    return result


def build_risk_tables(responses: Mapping[str, Any]) -> dict[str, list[dict[str, str]]]:
    tables: dict[str, list[dict[str, str]]] = {}

    execution_risk = _build_execution_risk_table(responses)
    if execution_risk:
        tables["execution_risk"] = execution_risk

    consumption_limit = _build_consumption_limit_table(responses)
    if consumption_limit:
        tables["consumption_limit"] = consumption_limit

    discredit_executor = _build_discredit_executor_table(responses)
    if discredit_executor:
        tables["discredit_executor"] = discredit_executor

    return tables
