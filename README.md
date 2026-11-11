# 银行对公客户经营智能助手（一期代码展示版）

这是面向企业客户经理的访前信息整合与报告生成服务。仓库展示一期代码中的主要工程链路：FastAPI 接收任务、服务层组织企业信息查询、多源响应解析为统一结构、聚合企业画像与分析字段，并生成结构化分析和 DOCX 报告。

> 本仓库是经过脱敏的作品集副本。项目所属银行名称、调用标识和接入配置已匿名化；不包含原始报告文件、开发者本地配置或真实凭据。它不是银行生产部署包，也不应连接真实银行系统。公开前仍应确认原单位对源代码和业务规则的披露许可。

## 主要链路

```text
HTTP 请求
  → FastAPI 路由与请求校验
  → 服务层组织企业信息查询
  → ESB Client 适配边界
  → 分主题 Parser 规范化响应
  → Aggregator 汇总企业画像与分析字段
  → 报告构建、字段校验与 DOCX 渲染
  → API 响应 / 报告文件
```

外部接口配置已替换为不可路由占位值，调用标识为演示值。测试路径使用仓库内的模拟响应；本仓库不会自动连接真实 ESB。

## 技术栈

- Python 3.10+
- FastAPI / Pydantic
- asyncio / httpx
- pytest
- DOCX 报告生成与结构化字段校验

实际依赖版本以 `requirements.txt` 为准。

## 本地测试

建议在隔离的 Python 环境中安装依赖并运行测试：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:PYTEST_DISABLE_PLUGIN_AUTOLOAD = "1"
python -m pytest tests/ -q
```

测试使用模拟响应，不应访问真实银行接口。运行结果以实际测试为准；接口标识经过匿名化后，部分旧用例可能需要依据测试输出进一步调整。不要把单元测试结果解释为银行生产环境验证。

## 运行服务

```powershell
Copy-Item .env.example .env
python -m uvicorn app.main:app --host 127.0.0.1 --port 8001
```

服务绑定本机。ESB、LLM 和用户身份配置未提供；ESB 默认值为不可路由占位符。未配置获批依赖时，外部查询不可用是预期行为。不要将真实凭据补入公开仓库。

## 目录概览

- `app/router.py`：HTTP 接口与请求处理
- `app/services.py`：企业信息查询和报告流程编排
- `app/core/esb_client.py`：外部查询适配边界
- `app/parsers/`：分主题响应解析
- `app/aggregator.py`：跨接口结果整合
- `app/customer_report/`：报告构建、DOCX 渲染和字段校验
- `configs/`：脱敏后的配置样例
- `tests/`：解析、契约、聚合和报告相关测试

## 真实性说明

仓库用于展示代码组织和实现思路。具体机构的生产部署、运行效果、业务收益、性能指标和个人职责，需以可核实材料为准；本 README 不对这些事项作声明。二期 Agent 设计和代码不在本仓库中。