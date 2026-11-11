"""股权关联（API-02）统一解析：结构化数据 + display + tables."""
from __future__ import annotations

from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.parsers._format_utils import (
    _count_records,
    _equity_change_summary,
    _find_first_scalar_by_keys,
    _join_present,
    _legal_person_office_summary,
    _outbound_investment_summary,
    _shareholder_summary,
)
from app.parsers._esb_records import records


def _count_text(payload: Any, label: str) -> str:
    count = _count_records(payload)
    return f"{label}{count}条" if count else ""


def parse_equity_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    """解析 API-02 结构化数据（逻辑与改前 get_key_person 一致）."""
    data: dict[str, Any] = {
        "shareholders": [],
        "executives": [],
        "actualController": {},
        "relationshipGraph": {"nodes": [], "edges": []},
        "investmentInfo": [],
        "positionInfo": [],
        "legalPersonInvestments": [],
        "equityChangeRecords": [],
        "legalPersonPositions": [],
        "beneficialOwners": [],
    }

    if "9000000011" in responses and is_success(responses["9000000011"]):
        body = get_body(responses["9000000011"])
        shareholders = records(body, "ShrhlrInfArry")
        data["shareholders"] = [
            {
                "name": pick(s, "ShrhlrNm"),
                "type": pick(s, "ShrhlrTp"),
                "shareholdingRatio": pick(s, "CtbRto"),
                "subscriptionAmount": pick(s, "SbcbCtbLmt"),
            }
            for s in shareholders
        ]

    if "9000000010" in responses and is_success(responses["9000000010"]):
        body = get_body(responses["9000000010"])
        personnel = records(body, "MainMgtPrsnArry")
        data["executives"] = [
            {
                "name": pick(p, "PrsnNm"),
                "position": pick(p, "Psitn"),
                "gender": pick(p, "Gnd"),
            }
            for p in personnel
        ]

    if "9000000012" in responses and is_success(responses["9000000012"]):
        body = get_body(responses["9000000012"])
        nodes_raw = records(body, "NodeArry")
        edges_raw = records(body, "RelaArry")
        nodes, edges = [], []
        for n in nodes_raw:
            nid = pick(n, "NodeId")
            name = pick(n, "NodeNm")
            ntype = {"1": "company", "2": "person", "3": "phone", "4": "address"}.get(
                pick(n, "NodeTp"), pick(n, "NodeTp")
            )
            if name:
                nodes.append({"id": nid or name, "name": name, "type": ntype})
        for e in edges_raw:
            attr = e.get("RelaAttrStruct") if isinstance(e.get("RelaAttrStruct"), dict) else e
            source = pick(e, "SrcNodeId")
            target = pick(e, "TrgtNodeId")
            relation_code = pick(e, "RelaTp")
            relation = {"1": "投资", "2": "任职", "3": "疑似电话", "4": "疑似地址"}.get(
                relation_code, relation_code
            )
            ratio = pick(attr, "CtbRto", "ShrhldngRto")
            if source and target:
                edges.append({"source": source, "target": target, "relation": relation, "ratio": ratio})
        data["relationshipGraph"] = {"nodes": nodes, "edges": edges}

        node_names = {n["id"]: n["name"] for n in nodes}
        data["investmentInfo"] = [
            {
                "investor": node_names.get(e["source"], e["source"]),
                "investedCompany": node_names.get(e["target"], e["target"]),
                "investmentRatio": e["ratio"],
            }
            for e in edges
            if e["relation"] == "投资"
        ]
        data["positionInfo"] = []
        for raw in edges_raw:
            if pick(raw, "RelaTp") != "2":
                continue
            attr = raw.get("RelaAttrStruct") if isinstance(raw.get("RelaAttrStruct"), dict) else {}
            source, target = pick(raw, "SrcNodeId"), pick(raw, "TrgtNodeId")
            data["positionInfo"].append({
                "person": node_names.get(source, source),
                "company": node_names.get(target, target),
                "position": pick(attr, "PosDesc", "Pos"),
            })

    if "9000000013" in responses and is_success(responses["9000000013"]):
        body = get_body(responses["9000000013"])
        ctrl_nodes = records(body, "NodeArry")
        ctrl_relations = records(body, "RelaArry")
        controller_name = ""
        node_names = {pick(node, "NodeId"): pick(node, "NodeNm") for node in ctrl_nodes}
        controller_node_id = ""
        for node in ctrl_nodes:
            attr = node.get("NodeAttrStrct") if isinstance(node.get("NodeAttrStrct"), dict) else {}
            controller_name = controller_name or pick(attr, "FnlHldgShrhlr")
            controller_node_id = controller_node_id or pick(attr, "ActlCtrlrNode")
        if not controller_name and controller_node_id:
            controller_name = node_names.get(controller_node_id, "")
        paths = [
            pick(r.get("RelaAttrStruct") or {}, "CtrlPthRela")
            for r in ctrl_relations
            if isinstance(r, dict)
        ]
        paths = [path for path in paths if path]
        if controller_name or paths:
            data["actualController"] = {
                "name": controller_name,
                "controlPath": "；".join(dict.fromkeys(paths)),
            }

    if "9000000014" in responses and is_success(responses["9000000014"]):
        body = get_body(responses["9000000014"])
        investments = records(body, "EntpExtInvsInfArry")
        data["legalPersonInvestments"] = [
            {
                "investedCompany": pick(item, "EntpNm"),
                "investmentRatio": pick(item, "CtbRto"),
                "registeredCapital": pick(item, "RgstCpitl"),
                "status": pick(item, "EntpSt"),
            }
            for item in investments[:20]
        ]

    if "9000000015" in responses and is_success(responses["9000000015"]):
        body = get_body(responses["9000000015"])
        changes = records(body, "StkRghtMdfInfArry")
        data["equityChangeRecords"] = [
            {
                "changeDate": pick(c, "StkRghtMdfDt"),
                "shareholderName": pick(c, "ShrhlrNm"),
                "beforeRatio": pick(c, "TfrPrevsStkRghtRto"),
                "afterRatio": pick(c, "TfrAfStkRghtRto"),
            }
            for c in changes[:10]
        ]

    if "9000000016" in responses and is_success(responses["9000000016"]):
        body = get_body(responses["9000000016"])
        positions = records(body, "LglRprsOthrCoTakeOfcArry")
        data["legalPersonPositions"] = [
            {
                "companyName": pick(p, "EntpBrchNm"),
                "position": pick(p, "Pos"),
                "status": pick(p, "EntpSt"),
            }
            for p in positions[:10]
        ]

    if "9000000017" in responses and is_success(responses["9000000017"]):
        body = get_body(responses["9000000017"])
        owners = records(body, "UltmBenfOwrArry")
        data["beneficialOwners"] = [
            {
                "name": pick(o, "BenfOwrNm"),
                "ratio": pick(o, "CtbRto"),
                "type": pick(o, "BenfTp"),
            }
            for o in owners[:10]
        ]

    display = build_equity_display(responses, data)
    tables = build_equity_tables(data)
    return {**data, "display": display, "tables": tables}


def build_equity_display(
    responses: Mapping[str, Any],
    data: Mapping[str, Any],
) -> dict[str, str]:
    """展示文案：与 renderer _derived_field_value 股权相关字段口径一致."""
    text: dict[str, str] = {}

    shareholder = _shareholder_summary(responses.get("9000000011"))
    relation_count = _count_records(responses.get("9000000012")) + _count_records(responses.get("9000000013"))
    if shareholder and relation_count:
        penetration = f"{shareholder}；关联/控制路径记录{relation_count}条"
    else:
        penetration = shareholder
    if penetration:
        for label in ("股权结构图", "穿透查询股东关系", "股权穿透查询"):
            text[label] = penetration

    equity_change = _equity_change_summary(responses.get("9000000015"))
    if equity_change:
        text["股权变更（含变更频次、时间）"] = equity_change

    outbound = _outbound_investment_summary(responses.get("9000000031"))
    if outbound:
        text["企业对外投资"] = outbound

    related_count = _count_records(responses.get("9000000032")) or _count_records(responses.get("9000000034"))
    if related_count:
        text["股东旗下其他关联企业信息"] = f"关联企业/多节点关联记录{related_count}条"

    manager = _find_first_scalar_by_keys(responses.get("9000000010"), ("Nm", "MbrNm", "PrsnNm"))
    position = _find_first_scalar_by_keys(responses.get("9000000010"), ("Pos", "PosNm"))
    manager_text = _join_present("，", manager, position) or _count_text(responses.get("9000000010"), "管理人员")
    if manager_text:
        text["主要管理人员"] = manager_text

    if shareholder:
        text["股东基础信息（含股东名称、股东类型、投资比例、投资数额、出资时间）"] = shareholder

    office_text = _legal_person_office_summary(responses)
    if office_text:
        text["法人、实控人跨企业任职信息"] = office_text
        text["名下关联企业、注册资本、经营状态"] = office_text

    graph_count = _count_records(responses.get("9000000016")) + _count_records(responses.get("9000000017"))
    if graph_count:
        text["关联任职图谱"] = f"关联任职/受益所有人图谱记录{graph_count}条"

    controller = data.get("actualController") or {}
    if controller.get("name"):
        parts = [f"实控人：{controller['name']}"]
        if controller.get("controlPath"):
            parts.append(f"控制路径：{controller['controlPath']}")
        text["实际控制人"] = "；".join(parts)

    return text


def build_equity_tables(data: Mapping[str, Any]) -> dict[str, list[dict[str, str]]]:
    """表格数据：与改前 _build_table_data_from_all_data 中 key_person 分支一致."""
    tables: dict[str, list[dict[str, str]]] = {}

    shareholders = data.get("shareholders") or []
    if shareholders:
        tables["investor_detail"] = [
            {
                "股东名称": r.get("name", "") or "",
                "出资比例": r.get("shareholdingRatio", "") or "",
                "股东类型": r.get("type", "") or "",
                "认缴出资额度": r.get("subscriptionAmount", "") or "",
                "出资日期": "-",
            }
            for r in shareholders
        ]

    investments = data.get("investmentInfo") or []
    if investments:
        tables["outbound_investment"] = [
            {
                "公司名称": r.get("investedCompany", "") or "",
                "占比": r.get("investmentRatio", "") or "",
                "认缴金额": "-",
                "企业状态": "-",
            }
            for r in investments
        ]

    executives = data.get("executives") or []
    if executives:
        tables["manager_detail"] = [
            {"姓名": r.get("name", "") or "", "职务": r.get("position", "") or ""}
            for r in executives
        ]

    equity_changes = data.get("equityChangeRecords") or []
    if equity_changes:
        tables["equity_change_records"] = [
            {
                "变更日期": r.get("changeDate", "") or "",
                "股东名称": r.get("shareholderName", "") or "",
                "转让前股权比例": r.get("beforeRatio", "") or "",
                "转让后股权比例": r.get("afterRatio", "") or "",
            }
            for r in equity_changes
        ]

    return tables
