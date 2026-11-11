"""资金需求（API-05）解析：动产抵押、股权出质、投融资事件."""
from __future__ import annotations

from typing import Any, Mapping

from app.core.esb_client import get_body, is_success, pick
from app.parsers._esb_records import records


def parse_funding_profile(responses: Mapping[str, Any]) -> dict[str, Any]:
    data: dict[str, Any] = {
        "mortgageInfo": {"count": 0},
        "equityPledge": {"count": 0},
        "financing": {},
    }

    mortgage_arrays = {
        "9000000023": "ChtlMrtgRgstInfArry",
        "9000000024": "ChtlMrtgMdfInfArry",
        "9000000025": "ChtlMrtgLgtInfArry",
    }
    data["mortgageInfo"]["count"] = sum(
        len(records(get_body(responses[code]), array_name))
        for code, array_name in mortgage_arrays.items()
        if code in responses and is_success(responses[code])
    )

    pledge_arrays = {
        "9000000020": "StkRghtPlgBscInfArry",
        "9000000021": "StkRghtPlgMdfInfArry",
        "9000000022": "StkRghtPlgLgtInfArry",
    }
    data["equityPledge"]["count"] = sum(
        len(records(get_body(responses[code]), array_name))
        for code, array_name in pledge_arrays.items()
        if code in responses and is_success(responses[code])
    )

    if responses.get("9000000029") and is_success(responses["9000000029"]):
        body = get_body(responses["9000000029"])
        financing = records(body, "FincngEvArry")
        data["financing"] = {
            "rounds": len(financing),
            "events": [
                {
                    "eventName": pick(item, "FincngEvFullNm"),
                    "round": pick(item, "InvsChrctrstcRndCN", "InvsChrctrstcRnd"),
                    "amount": pick(item, "InvsAmt"),
                    "date": pick(item, "InvsTm", "PublcDt"),
                    "investor": pick(item, "InvsrNmFullNm", "InvsrNmAbbr"),
                }
                for item in financing[:10]
            ],
        }

    return data
