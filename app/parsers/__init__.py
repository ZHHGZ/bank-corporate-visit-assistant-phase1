"""ESB 响应解析层：各业务模块统一取数口径（与 llm-standalone renderer 对齐）."""
from app.parsers.basic_profile import parse_basic_profile

__all__ = ["parse_basic_profile"]
