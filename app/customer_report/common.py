"""Shared helpers for customer report scripts."""
from __future__ import annotations


def safe_filename(value: str, *, default: str = "企业") -> str:
    forbidden = '<>:"/\\|?*'
    cleaned = "".join("_" if char in forbidden else char for char in str(value).strip())
    return cleaned or default
