"""应用配置管理."""
from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict

# enterprise-report-api/ 项目根目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    """环境变量配置."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
    )

    # ESB配置
    esb_url: str = "https://example.invalid/service"
    esb_big_message_url: str = "https://example.invalid/service"
    esb_consumer_sys_id: str = "DEMO"
    esb_original_consumer_sys_id: str = "DEMO"
    esb_timeout: int = 60
    esb_company: str = "DEMO"
    esb_channel_type: str = "DEMO"
    esb_inst_id: str = "DEMO"
    esb_user_no: str = "DEMO"
    esb_dept_no: str = "DEMO"
    esb_service_no: str = "DEMO"
    esb_system_no: str = "DEMO"
    esb_center_no: str = "DEMO"
    # 真实 ESB 失败时是否用 mock 数据兜底（本地联调建议 true，生产建议 false）
    esb_mock_fallback: bool = False
    # 原始 ESB 响应上下文缓存（供 API-08 正式报告复用完整报文）
    esb_cache_db: str = "temp/esb_context.db"
    esb_cache_ttl_seconds: int = 86400

    # LLM配置
    llm_enabled: bool = True
    llm_api_key: str = ""
    llm_base_url: str = "https://example.invalid/service"
    llm_model: str = "gpt-4"
    llm_max_tokens: int = 2000

    # 报告配置（相对路径相对于项目根目录）
    report_output_dir: str = "temp"
    report_template_dir: str = "templates"
    # API-08 生成 DOCX 时是否在 HiAgent 数据之外再调一次 ESB（默认 true 兼容现网）
    report_esb_fallback: bool = True

    # 服务配置
    log_level: str = "INFO"
    api_base_url: str = "https://example.invalid/service"

    def resolve_path(self, path_str: str) -> Path:
        """相对路径落到项目根下，绝对路径原样使用."""
        path = Path(path_str).expanduser()
        if not path.is_absolute():
            path = PROJECT_ROOT / path
        return path


_settings: Settings | None = None


def get_settings() -> Settings:
    """获取全局配置单例."""
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings
