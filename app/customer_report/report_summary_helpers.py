"""向后兼容：保留 normalize_text 导入路径，实际定义在 app.utils.text."""
from __future__ import annotations

from app.utils.text import normalize_text

__all__ = ["normalize_text"]
