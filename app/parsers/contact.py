"""触达路径（API-04）解析：联系方式、分支机构、对外投资、关联方、参股控股、多节点关联."""
from __future__ import annotations

from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.parsers._esb_records import records


def parse_contact_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    data: dict[str, Any] = {
        "contactInfo": {},
        "branches": [],
        "outboundInvestments": [],
        "relatedParties": [],
        "holdingCompanies": [],
        "multiNodeRelations": [],
    }

    if responses.get("9000000001") and is_success(responses["9000000001"]):
        body = get_body(responses["9000000001"])
        info = records(body, "EntpBscInfArry")
        if info:
            info = info[0]
        else:
            info = {}
        data["contactInfo"] = {
            "phone": pick(info, "CrspTel"),
            "email": pick(info, "Email"),
            "address": pick(info, "Addr"),
        }

    if responses.get("9000000030") and is_success(responses["9000000030"]):
        body = get_body(responses["9000000030"])
        branches = records(body, "BrnchBrchArry")
        data["branches"] = [
            {
                "name": pick(b, "BrnchBrchNm"),
                "creditCode": pick(b, "BrnchBrchUSCC"),
                "registrationAuthority": pick(b, "BrnchBrchRgstInst"),
                "registrationNumber": pick(b, "BrnchBrchEntpRgstNo"),
            }
            for b in branches
        ]

    if responses.get("9000000031") and is_success(responses["9000000031"]):
        body = get_body(responses["9000000031"])
        investments = records(body, "EntpExtInvsInfArry")
        data["outboundInvestments"] = [
            {
                "companyName": pick(i, "EntpNm"),
                "ratio": pick(i, "CtbRto"),
                "investmentAmount": pick(i, "SbcbCtbLmt"),
                "status": pick(i, "EntpSt"),
            }
            for i in investments[:20]
        ]

    if responses.get("9000000032") and is_success(responses["9000000032"]):
        body = get_body(responses["9000000032"])
        parties = records(body, "RelaPartyArry")
        data["relatedParties"] = [
            {
                "name": pick(p, "RelaPartyNm"),
                "relation": pick(p, "RelaTp"),
                "relationDesc": pick(p, "RelaDesc"),
            }
            for p in parties[:10]
        ]

    if responses.get("9000000033") and is_success(responses["9000000033"]):
        body = get_body(responses["9000000033"])
        holdings = records(body, "PtcHldgInfArry")
        data["holdingCompanies"] = [
            {
                "name": pick(h, "EntpNm"),
                "ratio": pick(h, "ShrhldngRto"),
                "status": pick(h, "EntpSt"),
            }
            for h in holdings[:20]
        ]

    if responses.get("9000000034") and is_success(responses["9000000034"]):
        body = get_body(responses["9000000034"])
        relas = records(body, "RelaArry")
        data["multiNodeRelations"] = [
            {
                "source": pick(r, "SrcNodeNm"),
                "target": pick(r, "TrgtNodeNm"),
                "relation": pick(r, "RelaTp"),
            }
            for r in relas[:10]
        ]

    return data
