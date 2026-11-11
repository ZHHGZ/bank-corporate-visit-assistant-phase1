"""ESB 失败时的 mock 响应，字段结构与 services 解析逻辑对齐."""
from __future__ import annotations

from typing import Any


def build_mock_response(txn_code: str, company_name: str) -> dict[str, Any]:
    """构造形如真实 ESB 成功响应的 mock，并打上 _mock 标记."""
    body = _mock_body(txn_code, company_name, "91440101MA5XXXXX0X")
    return {
        "_mock": True,
        "SysHead": {"RspCd": "200"},
        "BODY": {
            "BsnRetCd": "200",
            "RspCd": "200",
            **body,
        },
    }


def _mock_body(txn_code: str, company_name: str, credit_code: str) -> dict[str, Any]:
    """按 llm-standalone 官方 ESB YAML 中的 source/path 构造样例报文."""
    builders = {
        "9000000001": lambda: {
            "QryRsltArry": [{"EntpBscInfArry": [{
                "EntpNm": company_name, "USCC": credit_code,
                "EntpLglRprsNm": "张三", "RgstCpitl": "5000",
                "FoundDt": "20100115", "CrpnTrmBeg": "20100115",
                "CrpnTrmTo": "长期", "Addr": "广州市天河区示例路88号",
                "EntpTp": "有限责任公司", "EntpCrpnSt": "存续",
                "CrpnBsnScop": "软件开发；信息技术咨询服务",
                "IdyNm": "软件和信息技术服务业", "StfPplCnt": "120",
                "CrspTel": "020-88886666", "Email": "contact@example.com",
            }]}]
        },
        "9000000002": lambda: {"QryRsltArry": [{"DataArry": [
            {"AnulRprtId": "2024", "TownStfBscPnsnInsunPIPC": "95"}
        ]}]},
        "9000000003": lambda: {"QryRsltArry": [{"A02RspStruct": {"EntpScl": "中型"}}]},
        "9000000004": lambda: {
            "QryRsltArry": [{"MdfInfArry": [{
                "MdfDt": "20240301", "MdfItm": "注册资本",
                "MdfPrevsCntnt": "3000万元", "MdfAfCntnt": "5000万元",
            }]}]
        },
        "9000000005": lambda: {"QryRsltArry": [{"TaxnEvltArry": [{"RtgYr": "2024", "AutLvl": "A"}]}]},
        "9000000006": lambda: {"RetDataStruct": {"QualfHnrStruct": {"BllbrdListArry": []}}},
        "9000000007": lambda: {"RsltArry": [{"CstmsRgstEncdInfListArry": []}]},
        "9000000008": lambda: {
            "RsltArry": [
                {"PtntBscInfStruct": {"PtntNm": "示例发明专利A", "PtntTp": "发明"}},
                {"PtntBscInfStruct": {"PtntNm": "示例实用新型B", "PtntTp": "实用新型"}},
            ]
        },
        "9000000010": lambda: {
            "QryRsltArry": [{"MainMgtPrsnArry": [
                {"PrsnNm": "张三", "Psitn": "执行董事", "Gnd": "男"},
                {"PrsnNm": "王五", "Psitn": "财务负责人", "Gnd": "女"},
            ]}]
        },
        "9000000011": lambda: {
            "QryRsltArry": [{"ShrhlrInfArry": [
                {"ShrhlrNm": "李四", "ShrhlrTp": "自然人", "CtbRto": "60", "SbcbCtbLmt": "3000"},
                {"ShrhlrNm": "示例投资有限公司", "ShrhlrTp": "企业", "CtbRto": "40", "SbcbCtbLmt": "2000"},
            ]}]
        },
        "9000000012": lambda: {
            "QryRsltArry": [{"EntpFmlDataArry": [{
                "NodeArry": [
                    {"NodeId": "c1", "NodeNm": company_name, "NodeTp": "1", "NodeAttrStrct": {"EntpNm": company_name}},
                    {"NodeId": "p1", "NodeNm": "李四", "NodeTp": "2", "NodeAttrStrct": {}},
                    {"NodeId": "c2", "NodeNm": "示例科技子公司", "NodeTp": "1", "NodeAttrStrct": {}},
                ],
                "RelaArry": [
                    {"SrcNodeId": "p1", "TrgtNodeId": "c1", "RelaTp": "1", "RelaAttrStruct": {"CtbRto": "60"}},
                    {"SrcNodeId": "c1", "TrgtNodeId": "c2", "RelaTp": "1", "RelaAttrStruct": {"CtbRto": "100"}},
                    {"SrcNodeId": "p1", "TrgtNodeId": "c1", "RelaTp": "2", "RelaAttrStruct": {"PosDesc": "执行董事"}},
                ],
            }]}]
        },
        "9000000013": lambda: {
            "QryRsltArry": [{"CtrlPthDataArry": [{
                "NodeArry": [
                    {"NodeId": "p1", "NodeNm": "李四", "NodeTp": "2",
                     "NodeAttrStrct": {"ActlCtrlrNode": "p1", "FnlHldgShrhlr": "李四"}},
                    {"NodeId": "c1", "NodeNm": company_name, "NodeTp": "1", "NodeAttrStrct": {}},
                ],
                "RelaArry": [{"SrcNodeId": "p1", "TrgtNodeId": "c1", "RelaTp": "12",
                              "RelaAttrStruct": {"CtrlPthRela": f"李四 → {company_name}"}}],
            }]}]
        },
        "9000000018": lambda: {
            "RsltStruct": {
                "EntpAnulRprtBscInfArry": [{
                    "BscInfAnulRprtMjrKey": "MOCK-2024",
                    "AnulRprtYr": "2024", "EntpNm": company_name,
                }],
                "AstSituInfArry": [{
                    "AstSituAnulRprtMjrKey": "MOCK-2024",
                    "AstSituAstTotAmt": "18000", "AstSituBsnsTotIncm": "12000",
                    "AstSituNetPrft": "800", "AstSituLbyTotAmt": "6000",
                    "AstSituOwnrEqtTot": "12000",
                }],
            }
        },
        "9000000019": lambda: {
            "QryRsltArry": [{"BidInfArry": [
                {"BidTtl": "某信息化建设项目", "BidClsf": "政府采购", "AncmTp": "中标公告", "IssuTm": "20240520", "PrchsPrsnNm": "某采购单位", "CallBidAgncBrch": "某招标代理"},
                {"BidTtl": "某运维服务采购", "BidClsf": "企业采购", "AncmTp": "采购公告", "IssuTm": "20231108", "PrchsPrsnNm": "某采购单位", "CallBidAgncBrch": "某招标代理"},
            ]}]
        },
        "9000000030": lambda: {
            "QryRsltArry": [{"BrnchBrchArry": [{
                "BrnchBrchNm": f"{company_name}深圳分公司",
                "BrnchBrchUSCC": "91440300MA5XXXXX0X",
                "BrnchBrchRgstInst": "深圳市市场监督管理局",
                "BrnchBrchEntpRgstNo": "440300000000001",
            }]}]
        },
        "9000000023": lambda: {"QryRsltArry": [{"ChtlMrtgRgstInfArry": [{
            "RgstNo": "MOCK-MTG-001", "BeGrntOblgtryCntAmt": "500",
            "BeGrntBondKind": "借贷合同", "GrntScop": "全部资产抵押", "PrfmDbtEndDt": "20250630",
        }]}]},
        "9000000024": lambda: {"QryRsltArry": [{"ChtlMrtgMdfInfArry": []}]},
        "9000000025": lambda: {"QryRsltArry": [{"ChtlMrtgLgtInfArry": []}]},
        "9000000026": lambda: {"QryRsltArry": [{"CMBeGrntMainOblgtryArry": [{
            "RgstNo": "MOCK-MTG-001", "PrfmCntAmt": "500",
        }]}]},
        "9000000027": lambda: {"QryRsltArry": [{"ChtlMrtgCltlInfArry": [{
            "RgstNo": "MOCK-MTG-001", "CltlNm": "生产设备一套",
        }]}]},
        "9000000028": lambda: {"QryRsltArry": [{"ChtlMrtgPldgrInfArry": [{
            "RgstNo": "MOCK-MTG-001", "MrtgPldgrNm": "示例银行广州分行",
        }]}]},
        "9000000020": lambda: {"QryRsltArry": [{"StkRghtPlgBscInfArry": [{
            "RgstNo": "MOCK-PLG-001", "PlgPrsnNm": "李四",
            "StkRghtPlgSt": "有效", "PlgStkRghtCntAmt": "1500", "RgstDt": "20230115",
        }]}]},
        "9000000021": lambda: {"QryRsltArry": [{"StkRghtPlgMdfInfArry": []}]},
        "9000000022": lambda: {"QryRsltArry": [{"StkRghtPlgLgtInfArry": []}]},
        "9000000029": lambda: {
            "RetDataStruct": {"FincngEvArry": [
                {"FincngEvFullNm": "A轮融资", "InvsAmt": "3000", "CcyCN": "人民币",
                 "InvsAmtLvl": "1000-5000万", "FincrNmFullNm": company_name,
                 "InvsrNmFullNm": "示例投资有限公司", "InvsTm": "20220601"},
                {"FincngEvFullNm": "天使轮融资", "InvsAmt": "500", "CcyCN": "人民币",
                 "InvsAmtLvl": "500-1000万", "FincrNmFullNm": company_name,
                 "InvsrNmFullNm": "某天使基金", "InvsTm": "20210315"},
            ]}
        },
        "9000000035": lambda: {"QryRsltArry": []},
        "9000000036": lambda: {
            "QryRsltArry": [{"EntpJdgntWritInfArry": [
                {"ExecCaseNo": "（2023）粤01民初0001号", "RgstrRsn": "合同纠纷", "RoleNm": "被告"}
            ]}]
        },
        "9000000037": lambda: {"QryRsltArry": [{"OpnCrtSnAncmArry": []}]},
        "9000000038": lambda: {"QryRsltArry": [{"CrtAncmArry": []}]},
        "9000000039": lambda: {"QryRsltArry": [{"RstrctCnsmnBeExecPrsn": [{
            "JdgntCaseNo": "（2024）粤01执0001号", "CrtRgstDt": "20240115",
            "BeExecPrsnNm": "张三", "ExecCrtCd": "440100",
            "OIBC": "MA5XXXXX0", "WhthrVld": "1",
        }]}]},
        "9000000040": lambda: {"QryRsltArry": [{"BeExecPrsnInfArry": [{
            "ExecCaseNo": "（2023）粤01执0500号", "CrtRgstDt": "20230601",
            "PrfmSt": "全部未履行", "ExecCrtNm": "广州市中级人民法院", "CaseSt": "执行中",
        }]}]},
        "9000000041": lambda: {"QryRsltArry": [{"ExecInfArry": [{
            "BeExecPrsnNm": company_name, "CaseSt": "执行中",
            "ExecCaseNo": "（2023）粤01执0500号", "AppExecSbjMtr": "2000000",
            "ExecCrtNm": "广州市中级人民法院", "CrtRgstDt": "20230601",
        }]}]},
        "9000000042": lambda: {"QryRsltArry": [{"SvrIlglLsCrdtEntpListArry": []}]},
        "9000000043": lambda: {"QryRsltArry": [{"CstmsLsCrdtEntpListArry": []}]},
        "9000000044": lambda: {"QryRsltArry": [{"GovtLsCrdtEntpListArry": []}]},
        "9000000045": lambda: {"QryRsltArry": [{"DZLsCrdtEntpListArry": []}]},
        "9000000046": lambda: {"QryRsltArry": [{"AnlDataArry": []}]},
        "9000000047": lambda: {"QryRsltArry": [{"TaxnAdmntvPnltyArry": []}]},
        "9000000048": lambda: {"QryRsltArry": [{"MjrRevIlglArry": []}]},
        "9000000049": lambda: {"QryRsltArry": [{"IdyCmrcAdmntvPnltyArry": [{
            "WritPrefNo": "粤市监罚字〔2024〕1号", "MainIlglFct": "示例处罚（mock）",
            "AdmntvPnltyCntnt": "罚款2万元", "PnltyInstNm": "示例监管机构",
            "PnltyDtrmBkIssuDt": "20240110",
        }]}]},
        "9000000050": lambda: {"QryRsltArry": [{"EnvPrtctAdmntvPnltyArry": []}]},
        "9000000051": lambda: {"QryRsltArry": [{"CrpnAbnArry": []}]},
        "9000000052": lambda: {"QryRsltArry": [{"JdctrHlpArry": []}]},
        "9000000053": lambda: {"QryRsltArry": [{"RfrFinBlklistArry": []}]},
        "9000000054": lambda: {"QryRsltArry": [{"PnltyEntpInfArry": []}]},
        "9000000055": lambda: {"QryRsltArry": [{"ClsCaseArray": []}]},
        "9000000056": lambda: {"RetDataStruct": {
            "CrtAncmListArry": [], "JdctrCaseListArry": [], "JdgntWritListArry": [],
            "BnkrPcdrListArry": [], "ClrgInfArry": [], "LandMrtgArry": [],
            "StkRghtPlgArry": [], "AdmntvPnltyArry": [], "LosCrdtBlklistArry": [],
        }},
    }

    builder = builders.get(txn_code)
    if builder:
        return builder()
    return {"MockNote": f"未配置专用 mock，交易码={txn_code}", "MockArry": []}
