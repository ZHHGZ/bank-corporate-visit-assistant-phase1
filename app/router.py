"""统一的API路由."""
import base64

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse

from app import models, services
from app.config import get_settings


router = APIRouter(prefix="/api/v1/enterprise", tags=["企业报告API"])


@router.post("/basic-info", response_model=models.ApiResponse, summary="API-01: 获取企业基础信息")
async def api_basic_info(req: models.BasicInfoRequest):
    data = await services.get_basic_info(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/key-person", response_model=models.ApiResponse, summary="API-02: 获取企业关键人信息")
async def api_key_person(req: models.KeyPersonRequest):
    data = await services.get_key_person(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/business", response_model=models.ApiResponse, summary="API-03: 获取企业经营能力")
async def api_business(req: models.BusinessRequest):
    data = await services.get_business(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/contact", response_model=models.ApiResponse, summary="API-04: 获取企业触达路径")
async def api_contact(req: models.ContactRequest):
    data = await services.get_contact(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/funding", response_model=models.ApiResponse, summary="API-05: 获取企业资金需求")
async def api_funding(req: models.FundingRequest):
    data = await services.get_funding(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/risk", response_model=models.ApiResponse, summary="API-06: 获取企业风险信息")
async def api_risk(req: models.RiskRequest):
    data = await services.get_risk(req.company_name)
    return models.ApiResponse(data=data)


@router.post("/analysis", response_model=models.ApiResponse, summary="API-07: 生成AI智能分析")
async def api_analysis(req: models.AnalysisRequest):
    all_data = {
        "basic_info": req.basic_info,
        "key_person": req.key_person,
        "business": req.business,
        "contact": req.contact,
        "funding": req.funding,
        "risk": req.risk,
    }
    data = await services.generate_analysis(req.company_name, all_data)
    return models.ApiResponse(data=data)


@router.post("/report", response_model=models.ApiResponse, summary="API-08: 生成完整报告")
async def api_report(req: models.ReportRequest, request: Request):
    all_data = {
        "basic_info": req.basic_info,
        "key_person": req.key_person,
        "business": req.business,
        "contact": req.contact,
        "funding": req.funding,
        "risk": req.risk,
        "analysis": req.analysis,
    }
    data = await services.generate_report(
        req.company_name,
        all_data,
        req.output_format,
        option_payload=req.option_payload,
        bsn_data=req.bsn_data,
        modules=req.modules,
    )
    report_id = data.get("reportId")
    if report_id:
        data["downloadUrl"] = str(
            request.url_for("download_report", report_id=report_id)
        )
    return models.ApiResponse(data=data)


@router.get(
    "/reports/{report_id}/download",
    name="download_report",
    summary="下载已生成的报告文件",
    response_class=FileResponse,
)
async def download_report(report_id: str):
    report_path = services.resolve_report_file(report_id)
    if report_path is None:
        raise HTTPException(status_code=404, detail="报告不存在或已被清理")

    media_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if report_path.suffix.lower() == ".docx"
        else "application/json"
    )
    return FileResponse(
        path=report_path,
        filename=report_path.name,
        media_type=media_type,
    )


@router.get("/report/download/{encoded_path}", summary="下载生成的报告文件")
async def api_download_report(encoded_path: str):
    cfg = get_settings()
    base_dir = cfg.resolve_path(cfg.report_output_dir)
    try:
        relative = base64.urlsafe_b64decode(encoded_path.encode()).decode("utf-8")
    except Exception:
        raise HTTPException(status_code=400, detail="路径参数无效")
    file_path = base_dir / relative
    resolved = file_path.resolve()
    if not str(resolved).startswith(str(base_dir.resolve())):
        raise HTTPException(status_code=403, detail="不允许访问该路径")
    if not resolved.is_file():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(str(resolved), filename=resolved.name)
