"""ESB HTTP客户端 - 完全独立实现，不依赖旧项目."""
from __future__ import annotations

import json
import uuid
from datetime import datetime
from typing import Any

import httpx
import yaml

from app.config import PROJECT_ROOT, get_settings
from app.utils.logger import get_logger


logger = get_logger(__name__)

# ─────────────────────────────────────────────
# 接口元数据：交易码 → (service_code, scene_code, name, is_big_message)
#
# 权威来源为 `configs/esb/interface_index.yaml`（与 llm-standalone 同源），启动时
# 加载并覆盖下方内置兜底表；大报文标记沿用原项目规则：service_code 以 401320 开头
# 或接口名包含“大报文”。内置表仅在 YAML 缺失时兜底。
# ─────────────────────────────────────────────
_FALLBACK_INTERFACES: dict[str, dict[str, Any]] = {
    # ── 客户全貌 ──
    "9000000001": {"svc": "DEMO-SVC-01", "scn": "01", "name": "Demo ESB interface",          "big": False},
    "9000000002": {"svc": "DEMO-SVC-02", "scn": "08", "name": "Demo ESB interface",          "big": False},
    "9000000003": {"svc": "DEMO-SVC-03", "scn": "43", "name": "Demo ESB interface",             "big": False},
    "9000000004": {"svc": "DEMO-SVC-01", "scn": "08", "name": "Demo ESB interface",          "big": False},
    "9000000005": {"svc": "DEMO-SVC-04", "scn": "02", "name": "Demo ESB interface",             "big": False},
    "9000000006": {"svc": "DEMO-SVC-03", "scn": "52", "name": "Demo ESB interface",          "big": False},
    "9000000007": {"svc": "DEMO-SVC-05", "scn": "01", "name": "Demo ESB interface",          "big": True},
    "9000000008": {"svc": "DEMO-SVC-06", "scn": "02", "name": "Demo ESB interface",             "big": True},
    "9000000009": {"svc": "DEMO-SVC-07", "scn": "22", "name": "Demo ESB interface", "big": False},
    # ── 关键人（0207=管理人员，0208=股东，与中数智汇字段映射一致）──
    "9000000010": {"svc": "DEMO-SVC-01", "scn": "02", "name": "Demo ESB interface",         "big": False},
    "9000000011": {"svc": "DEMO-SVC-01", "scn": "03", "name": "Demo ESB interface",             "big": False},
    "9000000012": {"svc": "DEMO-SVC-08", "scn": "01", "name": "Demo ESB interface",     "big": False},
    "9000000013": {"svc": "DEMO-SVC-08", "scn": "02", "name": "Demo ESB interface",             "big": False},
    # ── 经营能力 ──
    "9000000018": {"svc": "DEMO-SVC-09", "scn": "02", "name": "Demo ESB interface",          "big": True},
    "9000000019": {"svc": "DEMO-SVC-10", "scn": "10", "name": "Demo ESB interface",           "big": False},
    # ── 触达路径 ──
    "9000000030": {"svc": "DEMO-SVC-01", "scn": "06", "name": "Demo ESB interface",             "big": False},
    # ── 资金需求 ──
    "9000000023": {"svc": "DEMO-SVC-11", "scn": "04", "name": "Demo ESB interface",          "big": False},
    "9000000024": {"svc": "DEMO-SVC-11", "scn": "05", "name": "Demo ESB interface",          "big": False},
    "9000000025": {"svc": "DEMO-SVC-11", "scn": "06", "name": "Demo ESB interface",          "big": False},
    "9000000026": {"svc": "DEMO-SVC-11", "scn": "07", "name": "Demo ESB interface",   "big": False},
    "9000000027": {"svc": "DEMO-SVC-11", "scn": "08", "name": "Demo ESB interface",        "big": False},
    "9000000028": {"svc": "DEMO-SVC-11", "scn": "09", "name": "Demo ESB interface",        "big": False},
    "9000000020": {"svc": "DEMO-SVC-11", "scn": "01", "name": "Demo ESB interface",       "big": False},
    "9000000021": {"svc": "DEMO-SVC-11", "scn": "02", "name": "Demo ESB interface",          "big": False},
    "9000000022": {"svc": "DEMO-SVC-11", "scn": "03", "name": "Demo ESB interface",          "big": False},
    "9000000029": {"svc": "DEMO-SVC-12", "scn": "04", "name": "Demo ESB interface",            "big": True},
    # ── 风险合规 ──
    "9000000035": {"svc": "DEMO-SVC-13", "scn": "18", "name": "Demo ESB interface",          "big": False},
    "9000000036": {"svc": "DEMO-SVC-13", "scn": "01", "name": "Demo ESB interface",             "big": False},
    "9000000037": {"svc": "DEMO-SVC-13", "scn": "11", "name": "Demo ESB interface",             "big": False},
    "9000000038": {"svc": "DEMO-SVC-13", "scn": "13", "name": "Demo ESB interface",             "big": False},
    "9000000039": {"svc": "DEMO-SVC-13", "scn": "17", "name": "Demo ESB interface",       "big": False},
    "9000000040": {"svc": "DEMO-SVC-14", "scn": "04", "name": "Demo ESB interface",          "big": False},
    "9000000041": {"svc": "DEMO-SVC-14", "scn": "06", "name": "Demo ESB interface",             "big": False},
    "9000000042": {"svc": "DEMO-SVC-14", "scn": "07", "name": "Demo ESB interface",       "big": False},
    "9000000043": {"svc": "DEMO-SVC-14", "scn": "08", "name": "Demo ESB interface",             "big": False},
    "9000000044": {"svc": "DEMO-SVC-14", "scn": "09", "name": "Demo ESB interface",   "big": False},
    "9000000045": {"svc": "DEMO-SVC-14", "scn": "10", "name": "Demo ESB interface",          "big": False},
    "9000000046": {"svc": "DEMO-SVC-14", "scn": "11", "name": "Demo ESB interface",             "big": False},
    "9000000047": {"svc": "DEMO-SVC-15", "scn": "01", "name": "Demo ESB interface",          "big": False},
    "9000000048": {"svc": "DEMO-SVC-15", "scn": "02", "name": "Demo ESB interface",          "big": False},
    "9000000049": {"svc": "DEMO-SVC-15", "scn": "03", "name": "Demo ESB interface",          "big": False},
    "9000000050": {"svc": "DEMO-SVC-15", "scn": "05", "name": "Demo ESB interface",          "big": False},
    "9000000051": {"svc": "DEMO-SVC-16", "scn": "02", "name": "Demo ESB interface",             "big": False},
    "9000000052": {"svc": "DEMO-SVC-16", "scn": "05", "name": "Demo ESB interface",             "big": False},
    "9000000053": {"svc": "DEMO-SVC-17", "scn": "01", "name": "Demo ESB interface",          "big": False},
    "9000000054": {"svc": "DEMO-SVC-18", "scn": "02", "name": "Demo ESB interface",             "big": False},
    "9000000055": {"svc": "DEMO-SVC-13", "scn": "20", "name": "Demo ESB interface",             "big": False},
    "9000000056": {"svc": "DEMO-SVC-12", "scn": "01", "name": "Demo ESB interface",        "big": True},
}

# 交易码走大报文通道的判定：与原项目 ESBEndpoint.is_big_message 保持一致。
_BIG_SVC_PREFIX = "DEMO-LARGE"


def _load_interface_index() -> dict[str, dict[str, Any]]:
    """从权威 interface_index.yaml 加载全部接口端点元数据."""
    path = PROJECT_ROOT / "configs" / "esb" / "interface_index.yaml"
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:  # pragma: no cover - 配置缺失兜底
        logger.warning(f"加载 interface_index.yaml 失败，使用内置兜底表: {exc}")
        return {}
    result: dict[str, dict[str, Any]] = {}
    for code, item in (data.get("interfaces") or {}).items():
        if not isinstance(item, dict):
            continue
        svc = str(item.get("service_code") or "")
        scn = str(item.get("scene_code") or "")
        name = str(item.get("interface_name") or code)
        result[str(code)] = {
            "svc": svc,
            "scn": scn,
            "name": name,
            "big": svc.startswith(_BIG_SVC_PREFIX) or "大报文" in name,
        }
    return result


# 权威 YAML 覆盖内置兜底表
ESB_INTERFACES: dict[str, dict[str, Any]] = {**_FALLBACK_INTERFACES, **_load_interface_index()}


def _serial_no() -> str:
    cfg = get_settings()
    now = datetime.now()
    return f"{cfg.esb_consumer_sys_id}{now.strftime('%Y%m%d%H%M%S')}{uuid.uuid4().hex[:8]}"


def _build_request(txn_code: str, body: dict[str, Any]) -> dict[str, Any]:
    """组装ESB请求报文."""
    cfg = get_settings()
    iface = ESB_INTERFACES[txn_code]
    now = datetime.now()
    sn = _serial_no()
    return {
        "SysHead": {
            "SvcCd": iface["svc"],
            "ScnCd": iface["scn"],
            "CnsmrSysID": cfg.esb_consumer_sys_id,
            "CnsmrSrlNo": sn,
            "GlblSrlNo": sn,
            "RelGlblSrlNo": "",
            "TxnDt": now.strftime("%Y%m%d"),
            "TxnTm": now.strftime("%H%M%S"),
            "OrgnlCnsmrSysID": cfg.esb_original_consumer_sys_id,
            "TxnMd": "ONLINE",
            "OrgnlTmlIdNo": "",
            "TmlIdNo": "",
            "OrgnlCnsmrSvcNo": "",
            "CnsmrSvcNo": "",
            "UsrLng": "",
            "FileFlg": "",
            "MACVal": "",
            "MACFctr": "",
            "PINSd": "",
            "SvcVerNo": "",
            "PrtyLvl": "",
            "VerfFlg": "",
            "SysRsrv": "",
            "Company": cfg.esb_company,
        },
        "AppHead": {
            "UsrLvl": "",
            "UsrTp": "",
            "TlrSrlNo": "",
            "CnlDtlTp": "",
            "CnlTp": cfg.esb_channel_type,
            "InstId": cfg.esb_inst_id,
            "UsrNo": cfg.esb_user_no,
            "CnsmrTxnCd": txn_code,
        },
        "BODY": {
            "IntfTxnCd": txn_code,
            "DeptNo": cfg.esb_dept_no,
            "SvcNo": cfg.esb_service_no,
            "SysNo": cfg.esb_system_no,
            "CnterNo": cfg.esb_center_no,
            **body,
        },
    }


def _target_url(txn_code: str) -> str:
    cfg = get_settings()
    return cfg.esb_big_message_url if ESB_INTERFACES[txn_code]["big"] else cfg.esb_url


def _std_body(company_name: str, extra: dict[str, Any] | None = None) -> dict[str, Any]:
    """标准查询入参（固定按企业名称查询，QryTp=32）."""
    body: dict[str, Any] = {
        "QryCntnt": company_name,
        "QryTp": "32",
        "PerPgRcrdCnt": "100",
        "CrnPgCnt": "1",
    }
    if extra:
        body.update(extra)
    return body


def _clean_body(body: dict[str, Any]) -> dict[str, Any]:
    """去掉原项目 query body 中用于占位的空值，避免发送未知证件字段."""
    return {k: v for k, v in body.items() if v not in (None, "", [], {})}


def _company_identity_body(
    company_name: str,
    *,
    name_field: str,
    credit_field: str,
    key_field: str | None = None,
    register_field: str | None = None,
) -> dict[str, Any]:
    """按原项目大报文身份字段命名构造企业查询体；当前外部入参只有企业名."""
    body: dict[str, Any] = {name_field: company_name, credit_field: None}
    if key_field:
        body[key_field] = None
    if register_field:
        body[register_field] = None
    body["OIBC"] = None
    return _clean_body(body)


_REPORT_SPECIAL_BODY_TXNS: set[str] = {
    "9000000056",
    "9000000006",
    "9000000007",
    "9000000008",
    "9000000018",
    "9000000033",
    "9000000029",
    "9000000017",
    "9000000034",
    "9000000057",
    "9000000058",
    "9000000059",
}


def _body_for_txn(
    txn_code: str,
    company_name: str,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """构造交易请求 BODY；特殊正式报告接口不套用通用 QryCntnt/QryTp."""
    report_body = _report_extra_body(txn_code, company_name)
    merged_extra = {**report_body, **(extra or {})}
    if txn_code in _REPORT_SPECIAL_BODY_TXNS:
        return _clean_body(merged_extra)
    return _std_body(company_name, merged_extra)


async def call_esb(txn_code: str, body: dict[str, Any]) -> dict[str, Any]:
    """调用单个ESB接口，返回原始响应或错误字典."""
    cfg = get_settings()
    payload = json.dumps(_build_request(txn_code, body), ensure_ascii=False)
    url = _target_url(txn_code)
    
    try:
        async with httpx.AsyncClient(timeout=cfg.esb_timeout) as client:
            resp = await client.post(
                url,
                content=payload.encode("utf-8"),
                headers={"Content-Type": "application/json; charset=utf-8"},
            )
            resp.raise_for_status()
            return resp.json()
    except httpx.TimeoutException:
        logger.warning(f"ESB超时: {txn_code}")
        return {"_error": "timeout", "_txn": txn_code}
    except httpx.HTTPStatusError as e:
        logger.warning(f"ESB HTTP错误: {txn_code} status={e.response.status_code}")
        return {"_error": f"http_{e.response.status_code}", "_txn": txn_code}
    except Exception as e:
        logger.warning(f"ESB调用失败: {txn_code} err={e}")
        return {"_error": str(e), "_txn": txn_code}


async def call_esb_batch(
    txn_codes: list[str],
    company_name: str,
    extras: dict[str, dict[str, Any]] | None = None,
    *,
    persist: bool = True,
) -> dict[str, Any]:
    """
    并发调用多个ESB接口（按企业名称查询），返回 {txn_code: response}.

    当 `ESB_MOCK_FALLBACK=true` 且真实调用失败时，自动替换为 mock 响应（带 `_mock` 标记）。
    调用成功或 mock 兜底后，会把 `企业名称 + 交易码 + 完整响应` 写入原始响应上下文缓存，
    供 API-08 生成正式报告时复用（`persist=False` 可关闭写入）。
    """
    import asyncio

    from app.core.mock_data import build_mock_response

    cfg = get_settings()

    async def _call(code: str) -> tuple[str, Any]:
        body = _body_for_txn(code, company_name, (extras or {}).get(code, {}))
        result = await call_esb(code, body)
        if not is_success(result) and cfg.esb_mock_fallback:
            name = ESB_INTERFACES.get(code, {}).get("name", code)
            err = result.get("_error") if isinstance(result, dict) else "unknown"
            logger.warning(f"ESB失败，启用 mock 兜底: {code} ({name}) reason={err}")
            result = build_mock_response(code, company_name)
        return code, result

    pairs = await asyncio.gather(*[_call(code) for code in txn_codes])
    responses = dict(pairs)
    if persist:
        _persist_responses(company_name, responses)
    return responses


def _persist_responses(company_name: str, responses: dict[str, Any]) -> None:
    """将原始响应写入上下文缓存，失败不影响主流程."""
    try:
        from app.core import raw_context

        raw_context.save_responses(company_name, responses)
    except Exception as exc:  # pragma: no cover - 缓存写入不阻断业务
        logger.warning(f"写入原始响应上下文失败: {exc}")


def is_mock(response: Any) -> bool:
    """是否为 mock 兜底响应."""
    return isinstance(response, dict) and response.get("_mock") is True


def is_success(response: Any) -> bool:
    """判断ESB响应是否业务成功（含 mock 成功响应）."""
    if not isinstance(response, dict):
        return False
    if "_error" in response:
        return False
    if is_mock(response):
        return True
    sys_head = response.get("SysHead")
    if isinstance(sys_head, dict) and str(sys_head.get("RetSt", "")).upper() == "F":
        return False
    body = get_body(response)
    # 业务状态字段均位于 BODY 顶层；不递归扫描业务明细，避免误判。
    for code_key in ("BsnRetCd", "RspCd", "StCd", "MfcrRetStEncd", "RqstRetEncd"):
        code = body.get(code_key)
        if code is not None and str(code) not in ("", "200"):
            return False
    return True


def get_body(response: dict[str, Any]) -> dict[str, Any]:
    """从ESB响应中提取BODY."""
    return response.get("BODY") or {}


def pick(data: dict[str, Any], *keys: str) -> str:
    """从字典中取第一个非空值，返回字符串."""
    for k in keys:
        v = data.get(k)
        if v not in (None, "", [], {}):
            return str(v)
    return ""


# ─────────────────────────────────────────────
# 正式报告（API-08）原始响应通道：补齐交易码 + 特殊入参 + 111→112 链式查询
# ─────────────────────────────────────────────

# 正式渲染器在模板字段之外还会直接引用的交易码（与 opinion_excel_docx_renderer 对齐）。
RENDERER_TXN_CODES: set[str] = {
    "9000000001", "9000000010", "9000000011", "9000000030", "9000000031",
    "9000000004", "9000000015", "9000000019", "9000000020", "9000000021",
    "9000000022", "9000000023", "9000000005", "9000000002", "9000000003",
    "9000000012", "9000000013", "9000000056", "9000000006", "9000000007",
    "9000000008", "9000000018", "9000000029",
    "9000000009",
}


def _report_extra_body(txn_code: str, company_name: str) -> dict[str, Any]:
    """正式报告相关接口的特殊入参（对齐原项目 _build_extra_body 的关键项）."""
    if txn_code == "9000000002":
        return {"SnglPgDataCnt": "100", "PgngPgCnt": "1"}
    if txn_code == "9000000012":
        return {"RelaTp": "manager,invest", "FltrLvl": "1"}
    if txn_code == "9000000013":
        return {"RelaTp": "entOutInvest,entShareholder,personShareholder,control", "FltrLvl": "3"}
    if txn_code == "9000000056":
        return {
            **_company_identity_body(
                company_name,
                name_field="EntpNm",
                credit_field="USCC",
                key_field="ListEntpMjrKey",
                register_field="RgstNo",
            ),
            "RskTpEncd": None,
            "PgCd": "1",
            "PgCnt": "100",
        }
    if txn_code == "9000000006":
        return _company_identity_body(
            company_name,
            name_field="MainBdyNm",
            credit_field="USCC",
            key_field="ListEntpMjrKey",
            register_field="RgstNo",
        )
    if txn_code == "9000000007":
        return _company_identity_body(
            company_name,
            name_field="EntpNm",
            credit_field="USCC",
            key_field="ListEntpMjrKey",
            register_field="RgstNo",
        )
    if txn_code == "9000000008":
        return {
            **_company_identity_body(
                company_name,
                name_field="EntpNm",
                credit_field="USCC",
                key_field="ListEntpMjrKey",
                register_field="RgstNo",
            ),
            "EntpNm": company_name,
            "QryKywrd": company_name,
            "WhthrPtntTpSttcCnt": "false",
            "WhthrInvntPtntAndAuthDdplc": "0",
            "WhthrPlg": "0",
            "PgCd": "1",
            "PgCnt": "100",
            "PtntVldCdVal": "2",
            "PtntTpEncd": "11",
        }
    if txn_code == "9000000018":
        return _company_identity_body(
            company_name,
            name_field="EntpNm",
            credit_field="UnfdCrdtCd",
            key_field="ZsEntpMjrKey",
            register_field="EntpRgstNo",
        )
    if txn_code == "9000000033":
        return {
            **_company_identity_body(
                company_name,
                name_field="EntpNm",
                credit_field="USCC",
                key_field="ListEntpMjrKey",
                register_field="EntpRgstNo",
            ),
            "EntpNm": company_name,
            "WhthrAll": "0",
            "PgCd": "1",
            "PgCnt": "100",
        }
    if txn_code == "9000000029":
        return _company_identity_body(
            company_name,
            name_field="ListEntpNmOrAbbrOrUsedNm",
            credit_field="USCC",
            key_field="ListEntpMjrKey",
            register_field="EntpRgstNo",
        )
    if txn_code == "9000000017":
        return {"ListEntpIndr": company_name}
    if txn_code == "9000000034":
        return {"ListEntpIndr": json.dumps([company_name], ensure_ascii=False)}
    if txn_code == "9000000057":
        return {"KeyWrdArryFld": json.dumps([company_name], ensure_ascii=False)}
    if txn_code == "9000000058":
        return {"EntpNm": company_name}
    if txn_code == "9000000059":
        return {"EntpNm": company_name, "PgCd": "1", "PgCnt": "100"}
    iface = ESB_INTERFACES.get(txn_code)
    if iface and iface.get("big"):
        # 大报文类接口以企业名主体字段查询
        return {"EntpNm": company_name}
    return {}


def _find_arrays(data: Any, array_key: str) -> list[dict[str, Any]]:
    """在嵌套结构中递归定位第一个匹配的业务数组."""
    if isinstance(data, dict):
        value = data.get(array_key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        for nested in data.values():
            found = _find_arrays(nested, array_key)
            if found:
                return found
    elif isinstance(data, list):
        for item in data:
            found = _find_arrays(item, array_key)
            if found:
                return found
    return []


def _extract_opening_notice_ids(response: Any) -> list[str]:
    """从 9000000037 响应中提取开庭公告 ID（最多 10 条）."""
    ids: list[str] = []
    for row in _find_arrays(response, "OpnCrtSnAncmArry"):
        value = row.get("OpnCrtSnAncmId")
        if value not in (None, ""):
            ids.append(str(value))
    return ids[:10]


async def _call_opening_notice_details(company_name: str, resp_111: Any) -> list[Any] | None:
    """按 9000000037 返回的公告 ID 链式查询 9000000060（QryTp=999），返回明细列表."""
    import asyncio

    from app.core.mock_data import build_mock_response

    if "9000000060" not in ESB_INTERFACES:
        return None
    ids = _extract_opening_notice_ids(resp_111)
    if not ids:
        logger.info("9000000037 未返回 OpnCrtSnAncmId，跳过 9000000060 明细查询")
        return None
    cfg = get_settings()

    async def _one(notice_id: str) -> Any:
        body = {"QryCntnt": notice_id, "QryTp": "999", "PerPgRcrdCnt": "100", "CrnPgCnt": "1"}
        result = await call_esb("9000000060", body)
        if not is_success(result) and cfg.esb_mock_fallback:
            result = build_mock_response("9000000060", company_name)
        return result

    return list(await asyncio.gather(*[_one(nid) for nid in ids]))


def report_required_codes() -> list[str]:
    """正式报告所需交易码 = 模板字段/风险表引用 ∪ 渲染器引用，过滤到有端点元数据的."""
    codes: set[str] = set(RENDERER_TXN_CODES)
    try:
        from app.customer_report import load_opinion_report_template

        template = load_opinion_report_template()
        for field in template.fields:
            codes.update(field.interface_codes)
        for table in template.risk_detail_tables:
            codes.update(table.interface_codes)
    except Exception as exc:  # pragma: no cover - 模板缺失兜底
        logger.warning(f"加载正式报告模板交易码失败: {exc}")
    return sorted(code for code in codes if code in ESB_INTERFACES)


async def ensure_report_responses(company_name: str, *, use_cache: bool = True) -> dict[str, Any]:
    """为正式报告准备完整原始响应：优先复用上下文缓存，缺失的交易码自动补调.

    返回 `{交易码: 完整原始响应}`，其中 9000000060 为明细列表（可能为空/缺省）。
    """
    from app.core import raw_context

    required = report_required_codes()
    context = raw_context.load_responses(company_name) if use_cache else {}
    present = {code for code, value in context.items() if value is not None}

    missing_standard = [c for c in required if c not in present and c != "9000000060"]
    fetched: dict[str, Any] = {}
    if missing_standard:
        logger.info(f"[API-08] 补调缺失交易码 {len(missing_standard)} 个: {missing_standard}")
        fetched = await call_esb_batch(missing_standard, company_name)

    merged: dict[str, Any] = {**context, **fetched}

    if "9000000060" in required and "9000000060" not in merged:
        details = await _call_opening_notice_details(company_name, merged.get("9000000037"))
        if details is not None:
            merged["9000000060"] = details
            _persist_responses(company_name, {"9000000060": details})

    return merged
