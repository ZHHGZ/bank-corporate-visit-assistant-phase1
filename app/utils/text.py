"""纯文本工具 — 无任何业务依赖，可安全被所有层 import."""
from __future__ import annotations

import json
from typing import Any


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list, tuple)):
        try:
            return json.dumps(value, ensure_ascii=False)
        except TypeError:
            return str(value)
    return str(value)
