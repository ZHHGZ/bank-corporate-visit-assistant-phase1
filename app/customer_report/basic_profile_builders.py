"""客户全貌字段构建（兼容层）— 请改用 app.parsers.basic_profile."""
from app.parsers.basic_profile import (
    build_certification_report_text,
    build_certification_display,
    normalize_change_records,
    parse_basic_info_record,
    parse_basic_profile,
)

__all__ = [
    "build_certification_display",
    "build_certification_report_text",
    "normalize_change_records",
    "parse_basic_info_record",
    "parse_basic_profile",
]
