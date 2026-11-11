"""所有Pydantic模型定义."""
from typing import Any, Literal
from pydantic import BaseModel, Field


# ========== 通用模型 ==========

class ApiResponse(BaseModel):
    """统一响应格式."""
    code: str = "200"
    message: str = "success"
    data: Any = None


class BaseRequest(BaseModel):
    """基础请求参数（仅支持按企业名称查询）."""
    company_name: str = Field(..., alias="companyName")
    use_cache: bool = Field(True, alias="useCache")

    model_config = {"populate_by_name": True}


# ========== API-01 ~ API-06 请求模型 ==========

class BasicInfoRequest(BaseRequest):
    """API-01."""
    pass


class KeyPersonRequest(BaseRequest):
    """API-02."""
    relation_depth: int = Field(3, alias="relationDepth")


class BusinessRequest(BaseRequest):
    """API-03."""
    pass


class ContactRequest(BaseRequest):
    """API-04."""
    pass


class FundingRequest(BaseRequest):
    """API-05."""
    pass


class RiskRequest(BaseRequest):
    """API-06."""
    include_detail: bool = Field(True, alias="includeDetail")


# ========== API-07 / API-08：上游 data 必填，不给空对象默认值 ==========

class AnalysisRequest(BaseModel):
    """API-07: AI智能分析。上游字段填各接口返回的整个 data。"""
    company_name: str = Field(..., alias="companyName")
    basic_info: dict = Field(..., alias="basicInfo")
    key_person: dict = Field(..., alias="keyPerson")
    business: dict = Field(..., alias="business")
    contact: dict = Field(..., alias="contact")
    funding: dict = Field(..., alias="funding")
    risk: dict = Field(..., alias="risk")

    model_config = {"populate_by_name": True}


class ReportRequest(BaseModel):
    """API-08: 生成报告。上游字段填各接口返回的整个 data。

    DOCX 走正式渲染链路：按企业名从原始 ESB 响应上下文读取完整报文并补齐缺失交易码，
    与 llm-standalone 官方模板/渲染器生成结构、字段、样式一致的报告。上游 basicInfo…
    analysis 仍保留以兼容既有工作流，但正式 DOCX 不再依赖这些聚合结果。
    可选 optionPayload（中知数通科技力评分等字段覆盖）、bsnData（CRM 行内信息）缺失时，
    正式渲染器按规则显示“无数据”或省略可选章节。
    """
    company_name: str = Field(..., alias="companyName")
    basic_info: dict = Field(..., alias="basicInfo")
    key_person: dict = Field(..., alias="keyPerson")
    business: dict = Field(..., alias="business")
    contact: dict = Field(..., alias="contact")
    funding: dict = Field(..., alias="funding")
    risk: dict = Field(..., alias="risk")
    analysis: dict = Field(..., alias="analysis")
    output_format: Literal["docx", "json"] = Field(..., alias="outputFormat")
    # 可选：正式报告字段覆盖（如中知数通科技力评分）与 CRM 行内信息
    option_payload: dict | None = Field(None, alias="optionPayload")
    bsn_data: dict | None = Field(None, alias="bsnData")
    modules: list[str] | None = Field(None, alias="modules")

    model_config = {"populate_by_name": True}


# ========== 响应数据模型 ==========

class DataResponse(BaseModel):
    """通用数据响应."""
    source_interfaces: list[str] = Field(default_factory=list, alias="sourceInterfaces")
    failed_interfaces: list[str] = Field(default_factory=list, alias="failedInterfaces")
    timestamp: str | None = None

    model_config = {"populate_by_name": True}
