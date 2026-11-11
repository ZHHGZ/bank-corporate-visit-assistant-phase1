from __future__ import annotations

import json

from app.core.esb_client import _body_for_txn


COMPANY_NAME = "测试集团有限公司"


def test_standard_report_txn_keeps_qrycntnt_body() -> None:
    body = _body_for_txn("9000000012", COMPANY_NAME)

    assert body == {
        "QryCntnt": COMPANY_NAME,
        "QryTp": "32",
        "PerPgRcrdCnt": "100",
        "CrnPgCnt": "1",
        "RelaTp": "manager,invest",
        "FltrLvl": "1",
    }


def test_9000000505_uses_risk_big_message_fields_without_generic_query() -> None:
    body = _body_for_txn("9000000056", COMPANY_NAME)

    assert body == {
        "EntpNm": COMPANY_NAME,
        "PgCd": "1",
        "PgCnt": "100",
    }
    assert "QryCntnt" not in body
    assert "QryTp" not in body
    assert "USCC" not in body
    assert "ListEntpMjrKey" not in body
    assert "RgstNo" not in body


def test_identity_big_message_txns_use_reference_name_fields() -> None:
    assert _body_for_txn("9000000006", COMPANY_NAME) == {"MainBdyNm": COMPANY_NAME}
    assert _body_for_txn("9000000007", COMPANY_NAME) == {"EntpNm": COMPANY_NAME}
    assert _body_for_txn("9000000018", COMPANY_NAME) == {"EntpNm": COMPANY_NAME}
    assert _body_for_txn("9000000029", COMPANY_NAME) == {
        "ListEntpNmOrAbbrOrUsedNm": COMPANY_NAME
    }


def test_9000000509_preserves_patent_filters_without_generic_query() -> None:
    body = _body_for_txn("9000000008", COMPANY_NAME)

    assert body["EntpNm"] == COMPANY_NAME
    assert body["QryKywrd"] == COMPANY_NAME
    assert body["WhthrPtntTpSttcCnt"] == "false"
    assert body["WhthrInvntPtntAndAuthDdplc"] == "0"
    assert body["WhthrPlg"] == "0"
    assert body["PgCd"] == "1"
    assert body["PgCnt"] == "100"
    assert body["PtntVldCdVal"] == "2"
    assert body["PtntTpEncd"] == "11"
    assert "QryCntnt" not in body
    assert "QryTp" not in body


def test_special_report_body_allows_explicit_identity_extras() -> None:
    body = _body_for_txn("9000000056", COMPANY_NAME, {"USCC": "91310000MA1K000000"})

    assert body["EntpNm"] == COMPANY_NAME
    assert body["USCC"] == "91310000MA1K000000"
    assert "QryCntnt" not in body


def test_related_identity_txns_match_reference_payload_shape() -> None:
    assert _body_for_txn("9000000033", COMPANY_NAME) == {
        "EntpNm": COMPANY_NAME,
        "WhthrAll": "0",
        "PgCd": "1",
        "PgCnt": "100",
    }
    assert _body_for_txn("9000000034", COMPANY_NAME) == {
        "ListEntpIndr": json.dumps([COMPANY_NAME], ensure_ascii=False)
    }
