"""从 ESB BODY 递归提取数组/结构体（避免 parsers 与 services 循环依赖）."""
from __future__ import annotations

from typing import Any


def records(data: Any, *array_keys: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if isinstance(data, dict):
        for key, value in data.items():
            if key in array_keys and isinstance(value, list):
                rows.extend(item for item in value if isinstance(item, dict))
            else:
                rows.extend(records(value, *array_keys))
    elif isinstance(data, list):
        for item in data:
            rows.extend(records(item, *array_keys))
    return rows


def first_record(data: Any, *array_keys: str) -> dict[str, Any]:
    rows = records(data, *array_keys)
    return rows[0] if rows else {}


def first_struct(data: Any, struct_key: str) -> dict[str, Any]:
    if isinstance(data, dict):
        found = data.get(struct_key)
        if isinstance(found, dict):
            return found
        for child in data.values():
            result = first_struct(child, struct_key)
            if result:
                return result
    elif isinstance(data, list):
        for item in data:
            result = first_struct(item, struct_key)
            if result:
                return result
    return {}


def find_first_scalar(data: Any, field: str) -> str:
    if isinstance(data, dict):
        value = data.get(field)
        if value is not None and not isinstance(value, (dict, list)):
            return str(value)
        for child in data.values():
            result = find_first_scalar(child, field)
            if result:
                return result
    elif isinstance(data, list):
        for item in data:
            result = find_first_scalar(item, field)
            if result:
                return result
    return ""
