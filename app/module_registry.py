"""9大模块注册表 - 新需求模块定义、旧模板映射、过滤与排序统一入口.

后续所有模块改名、过滤、排序都走此注册表，避免在渲染器里到处写死映射。
"""
from __future__ import annotations

from collections import OrderedDict
from typing import Any, Sequence

from app.customer_report.opinion_excel_docx_renderer import OpinionReportField

_FALLBACK_FIELD_DEFAULT = OpinionReportField(
    module="",
    submodule="综合信息",
    name="数据暂不可用",
    interface_source="",
    remark="",
    interface_codes=[],
)

MODULE_REGISTRY: dict[str, dict[str, Any]] = OrderedDict(
    (
        (
            "ai_summary",
            {
                "id": "ai_summary",
                "name": "AI综合智能分析建议",
                "order": 1,
                "old_modules": [],
                "fallback": "AI综合分析数据暂不可用，建议结合其他模块信息综合判断。",
            },
        ),
        (
            "basic",
            {
                "id": "basic",
                "name": "客户全貌",
                "order": 2,
                "old_modules": ["客户全貌", "触达路径"],
                "fallback": "客户全貌数据暂不可用。",
            },
        ),
        (
            "equity",
            {
                "id": "equity",
                "name": "股权关联",
                "order": 3,
                "old_modules": ["关键人", "关键人物"],
                "fallback": "股权关联数据暂不可用。",
            },
        ),
        (
            "business",
            {
                "id": "business",
                "name": "经营能力",
                "order": 4,
                "old_modules": ["经营能力", "资金需求"],
                "fallback": "经营能力数据暂不可用。",
            },
        ),
        (
            "finance",
            {
                "id": "finance",
                "name": "财务分析",
                "order": 5,
                "old_modules": [],
                "fallback": "财务分析数据暂不可用，当前仅展示年报财务摘要。",
            },
        ),
        (
            "risk",
            {
                "id": "risk",
                "name": "风险提示",
                "order": 6,
                "old_modules": ["风险合规", "风险提示"],
                "fallback": "风险提示数据暂不可用。",
            },
        ),
        (
            "crm",
            {
                "id": "crm",
                "name": "行内合作情况",
                "order": 7,
                "old_modules": ["行内信息", "行内合作情况"],
                "fallback": "行内合作情况数据暂不可用。",
            },
        ),
        (
            "interview",
            {
                "id": "interview",
                "name": "访谈脚本",
                "order": 8,
                "old_modules": [],
                "fallback": "待补充：后续可结合AI智能分析、风险提示和营销方案生成客户经理现场访谈问题清单。",
            },
        ),
        (
            "marketing",
            {
                "id": "marketing",
                "name": "营销方案",
                "order": 9,
                "old_modules": [],
                "fallback": "营销方案暂不可用，建议结合企业经营状况选择对公产品。",
            },
        ),
    )
)

_OLD_TO_NEW: dict[str, str] = {}
for _mod_id, _mod_def in MODULE_REGISTRY.items():
    for _old_name in _mod_def["old_modules"]:
        _OLD_TO_NEW[_old_name] = _mod_def["name"]


def old_to_new_module_name(old_name: str) -> str:
    """旧模块名 → 新模块中文名."""
    return _OLD_TO_NEW.get(old_name, old_name)


def module_order_key(field: OpinionReportField) -> int:
    """字段排序键：按模块注册表 order 排序."""
    for mod_id, mod_def in MODULE_REGISTRY.items():
        if field.module == mod_def["name"]:
            return mod_def["order"]
    return 99


def is_new_module_rendered(fields: Sequence[OpinionReportField], module_id: str) -> bool:
    """检查 fields 中是否已包含指定新模块."""
    mod_def = MODULE_REGISTRY.get(module_id)
    if not mod_def:
        return False
    target_name = mod_def["name"]
    return any(f.module == target_name for f in fields)


def normalize_module_list(raw_modules: list[str] | None) -> set[str] | None:
    """将外部传入的 modules 列表归一化为新模块中文名集合.

    - "basic" → "客户全貌"
    - "关键人" / "关键人物" → "股权关联"
    - "风险合规" → "风险提示"
    - 未识别的值原样保留
    - None 返回 None 表示不过滤
    """
    if raw_modules is None:
        return None
    result: set[str] = set()
    for m in raw_modules:
        if not isinstance(m, str) or not m.strip():
            continue
        m = m.strip()
        mod_def = MODULE_REGISTRY.get(m)
        if mod_def:
            result.add(mod_def["name"])
            continue
        mapped = _OLD_TO_NEW.get(m)
        if mapped:
            result.add(mapped)
            continue
        result.add(m)
    return result if result else None


def remap_fields_to_new_modules(
    fields: Sequence[OpinionReportField],
) -> list[OpinionReportField]:
    """将旧模板字段的 module 属性改为新模块中文名.

    注意：直接修改 OpinionReportField.module（dataclass 非 frozen）.
    """
    result: list[OpinionReportField] = []
    for f in fields:
        new_name = old_to_new_module_name(f.module)
        if new_name != f.module:
            f.module = new_name
        result.append(f)
    return result


def inject_fallback_fields(
    fields: list[OpinionReportField],
) -> list[OpinionReportField]:
    """为注册表中缺失的模块注入占位字段."""
    existing_names = {f.module for f in fields}
    for mod_id, mod_def in MODULE_REGISTRY.items():
        name = mod_def["name"]
        if name not in existing_names:
            placeholder = OpinionReportField(
                module=name,
                submodule="综合信息",
                name=mod_def.get("fallback", "数据暂不可用"),
                interface_source="",
                remark="",
                interface_codes=[],
            )
            fields.append(placeholder)
    return fields


def sort_fields_by_module_order(
    fields: list[OpinionReportField],
) -> list[OpinionReportField]:
    """按 9 大模块顺序排序."""
    fields.sort(key=module_order_key)
    return fields


def get_new_module_names() -> list[str]:
    """按注册表顺序返回所有新模块中文名."""
    return [mod_def["name"] for mod_def in MODULE_REGISTRY.values()]
