"""FastAPI应用主入口."""
import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.openapi.docs import (
    get_redoc_html,
    get_swagger_ui_html,
    get_swagger_ui_oauth2_redirect_html,
)
from fastapi.staticfiles import StaticFiles

from app.router import router
from app.utils.logger import get_logger


logger = get_logger(__name__)


def _ensure_startup_dirs() -> None:
    """启动阶段检查并创建必要目录，写入测试验证可写性.

    REPORT_OUTPUT_DIR 不可写时直接 raise 阻断启动；
    ESB 缓存目录不可写时输出 ERROR 日志但不阻断。
    """
    from app.config import get_settings

    cfg = get_settings()
    report_dir = cfg.resolve_path(cfg.report_output_dir)
    cache_parent = cfg.resolve_path(cfg.esb_cache_db).parent if cfg.esb_cache_db else None

    for label, path, block_on_failure in (
        ("REPORT_OUTPUT_DIR", report_dir, True),
        ("ESB_CACHE_DB 所在目录", cache_parent, False),
    ):
        if path is None:
            continue
        try:
            path.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            msg = f"启动目录创建失败 [{label}] {path}: {exc}"
            if block_on_failure:
                raise RuntimeError(msg) from exc
            logger.error(msg)
            continue

        try:
            probe = path / f".write_test_{os.getpid()}"
            probe.write_text("1", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            msg = f"目录无写权限 [{label}] {path}: {exc}"
            if block_on_failure:
                raise RuntimeError(msg) from exc
            logger.error(msg)
            continue

        logger.info("启动目录就绪 [%s]: %s", label, path)

# 本地化的 Swagger/ReDoc 静态资源目录（内网无公网 CDN 时用于渲染 /docs）
STATIC_DIR = Path(__file__).parent / "static"


def create_app() -> FastAPI:
    app = FastAPI(
        title="企业报告聚合API（独立版）",
        description="""
## 为HiAgent工作流设计的企业报告生成API

将78个ESB接口聚合为8个业务接口，完全独立实现，无外部依赖。

### 8个API

1. **API-01** POST /api/v1/enterprise/basic-info - 企业基础信息
2. **API-02** POST /api/v1/enterprise/key-person - 关键人信息
3. **API-03** POST /api/v1/enterprise/business - 经营能力
4. **API-04** POST /api/v1/enterprise/contact - 触达路径
5. **API-05** POST /api/v1/enterprise/funding - 资金需求
6. **API-06** POST /api/v1/enterprise/risk - 风险信息
7. **API-07** POST /api/v1/enterprise/analysis - AI智能分析
8. **API-08** POST /api/v1/enterprise/report - 生成完整报告

### HiAgent工作流

```
输入: 企业名称

┌─ 并行执行 ─────────────┐
│ API-01 ~ API-06        │
└────────────────────────┘
           ↓
       API-07 (AI分析)
           ↓
       API-08 (生成报告)
           ↓
输出: DOCX/JSON文件
```

### 快速开始

```bash
# 1. 配置环境
cp .env.example .env
# 编辑.env填入ESB配置

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动服务
uvicorn app.main:app --reload --port 8001

# 4. 访问文档
http://localhost:8001/docs
```
        """,
        version="1.0.0",
        # 关闭默认 docs（其静态资源来自公网 CDN，内网打不开），改用本地静态资源渲染
        docs_url=None,
        redoc_url=None,
    )
    
    # CORS
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # 挂载本地静态资源并自定义 /docs、/redoc，避免依赖公网 CDN（内网可用）
    if STATIC_DIR.is_dir():
        app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

        @app.get("/docs", include_in_schema=False)
        def custom_swagger_ui_html():
            return get_swagger_ui_html(
                openapi_url=app.openapi_url,
                title=f"{app.title} - Swagger UI",
                oauth2_redirect_url=app.swagger_ui_oauth2_redirect_url,
                swagger_js_url="/static/swagger-ui-bundle.js",
                swagger_css_url="/static/swagger-ui.css",
            )

        @app.get(app.swagger_ui_oauth2_redirect_url, include_in_schema=False)
        def swagger_ui_redirect():
            return get_swagger_ui_oauth2_redirect_html()

        @app.get("/redoc", include_in_schema=False)
        def custom_redoc_html():
            return get_redoc_html(
                openapi_url=app.openapi_url,
                title=f"{app.title} - ReDoc",
                redoc_js_url="/static/redoc.standalone.js",
            )
    else:
        logger.warning("未找到本地 docs 静态资源目录，/docs 将不可用: %s", STATIC_DIR)
    
    # 注册路由
    app.include_router(router)
    
    @app.get("/", tags=["系统"])
    def root():
        return {
            "service": "企业报告聚合API（独立版）",
            "version": "1.0.0",
            "docs": "/docs",
            "apis": {
                "API-01": "/api/v1/enterprise/basic-info",
                "API-02": "/api/v1/enterprise/key-person",
                "API-03": "/api/v1/enterprise/business",
                "API-04": "/api/v1/enterprise/contact",
                "API-05": "/api/v1/enterprise/funding",
                "API-06": "/api/v1/enterprise/risk",
                "API-07": "/api/v1/enterprise/analysis",
                "API-08": "/api/v1/enterprise/report",
            },
        }
    
    @app.get("/health", tags=["系统"])
    def health():
        return {"status": "healthy", "version": "1.0.0"}
    
    @app.on_event("startup")
    async def startup():
        _ensure_startup_dirs()
        logger.info("🚀 企业报告API服务启动")
    
    @app.on_event("shutdown")
    async def shutdown():
        logger.info("👋 企业报告API服务关闭")
    
    return app


app = create_app()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.main:app", host="0.0.0.0", port=8001, reload=True)
