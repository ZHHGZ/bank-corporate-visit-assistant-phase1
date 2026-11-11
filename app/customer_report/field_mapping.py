"""9大模块字段映射台账 - 加载、校验、查询.

现阶段仅做台账和测试校验，不接入 DOCX 渲染逻辑。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import yaml

from app.module_registry import MODULE_REGISTRY

VALID_SOURCE_TYPES = frozenset(
    {"行外-中数智汇", "行外-年报", "行内-CRM", "AI生成", "规则生成", "待定"},
)
VALID_STATUSES = frozenset(
    {"已接入", "部分接入", "待接入", "待确认", "待补充"},
)
VALID_COMPUTE_TYPES = frozenset(
    {
        "direct",
        "first_non_empty",
        "recursive_first",
        "recursive_count",
        "count_nonzero",
        "multi_array_count",
        "latest_by_year",
        "join_by_key",
        "derived",
        "manual_confirm",
    },
)
VALID_VERIFIED_BY = frozenset(
    {
        "current_code",
        "current_code_partial",
        "pending_bank_confirm",
        "guessed",
    },
)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_MAPPING_PATH = PROJECT_ROOT / "configs" / "report_field_mapping.yaml"


@dataclass
class FieldMappingEntry:
    """单字段映射条目."""

    key: str
    module: str
    submodule: str
    label: str
    source_type: str
    interface_code: Optional[str]
    response_path: Optional[str]
    table_key: Optional[str]
    status: str
    fallback: str
    compute: Optional[str] = None
    aliases: list[str] = field(default_factory=list)
    source_arrays: list[str] = field(default_factory=list)
    verified_by: Optional[str] = None
    notes: str = ""


@dataclass
class ModuleMappingConfig:
    """单个模块的字段映射配置."""

    id: str
    name: str
    order: int
    fields: list[FieldMappingEntry] = field(default_factory=list)


@dataclass
class ReportFieldMapping:
    """完整的报告字段映射台账."""

    version: str
    modules: list[ModuleMappingConfig] = field(default_factory=list)

    def all_fields(self) -> list[FieldMappingEntry]:
        """返回所有模块的所有字段（扁平列表）."""
        result: list[FieldMappingEntry] = []
        for mod in self.modules:
            result.extend(mod.fields)
        return result

    def get_module(self, module_id: str) -> Optional[ModuleMappingConfig]:
        for mod in self.modules:
            if mod.id == module_id:
                return mod
        return None

    def get_fields_by_module(self, module_id: str) -> list[FieldMappingEntry]:
        mod = self.get_module(module_id)
        return list(mod.fields) if mod else []

    def get_field_by_key(self, key: str) -> Optional[FieldMappingEntry]:
        for entry in self.all_fields():
            if entry.key == key:
                return entry
        return None

    def field_keys(self) -> set[str]:
        return {e.key for e in self.all_fields()}


def load_field_mapping(path: Optional[Path] = None) -> ReportFieldMapping:
    """加载并解析字段映射 YAML 配置."""
    filepath = path or DEFAULT_MAPPING_PATH
    with open(filepath, "r", encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)

    modules: list[ModuleMappingConfig] = []
    for m in raw.get("modules", []):
        entries: list[FieldMappingEntry] = []
        for f in m.get("fields", []):
            entries.append(FieldMappingEntry(
                key=f["key"],
                module=m["id"],
                submodule=f.get("submodule", "综合信息"),
                label=f["label"],
                source_type=f.get("source_type", "待定"),
                interface_code=f.get("interface_code"),
                response_path=f.get("response_path"),
                table_key=f.get("table_key"),
                status=f.get("status", "待确认"),
                fallback=f.get("fallback", ""),
                compute=f.get("compute"),
                aliases=list(f.get("aliases") or []),
                source_arrays=list(f.get("source_arrays") or []),
                verified_by=f.get("verified_by"),
                notes=str(f.get("notes") or ""),
            ))
        modules.append(ModuleMappingConfig(
            id=m["id"],
            name=m["name"],
            order=m.get("order", 99),
            fields=entries,
        ))

    return ReportFieldMapping(
        version=raw.get("version", "unknown"),
        modules=modules,
    )


def validate_field_mapping(mapping: ReportFieldMapping) -> list[str]:
    """校验字段映射台账，返回所有错误信息（空列表表示通过）."""
    errors: list[str] = []

    valid_module_ids = set(MODULE_REGISTRY.keys())

    registered_ids: set[str] = set()
    for mod in mapping.modules:
        registered_ids.add(mod.id)
        if mod.id not in valid_module_ids:
            errors.append(f"模块 '{mod.id}' 不在 9 大模块注册表中")

    missing = valid_module_ids - registered_ids
    if missing:
        errors.append(f"台账缺少模块: {sorted(missing)}")

    for mod in mapping.modules:
        if mod.id in MODULE_REGISTRY:
            expected_name = MODULE_REGISTRY[mod.id]["name"]
            if mod.name != expected_name:
                errors.append(
                    f"模块 '{mod.id}' 的 name='{mod.name}' 与注册表 '{expected_name}' 不一致"
                )

    for mod in mapping.modules:
        if mod.id in MODULE_REGISTRY:
            expected_order = MODULE_REGISTRY[mod.id]["order"]
            if mod.order != expected_order:
                errors.append(
                    f"模块 '{mod.id}' 的 order={mod.order} 与注册表 order={expected_order} 不一致"
                )

    seen_keys: dict[str, int] = {}
    for entry in mapping.all_fields():
        if not entry.key or not entry.key.strip():
            errors.append(f"模块 '{entry.module}' 中存在空 key")
            continue
        if entry.key in seen_keys:
            errors.append(
                f"字段 key '{entry.key}' 重复: 模块 '{entry.module}' 与 "
                f"第 {seen_keys[entry.key]} 次出现冲突"
            )
        else:
            seen_keys[entry.key] = 1

    for entry in mapping.all_fields():
        if entry.source_type not in VALID_SOURCE_TYPES:
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') 的 source_type="
                f"'{entry.source_type}' 不合法，允许值: {sorted(VALID_SOURCE_TYPES)}"
            )

    for entry in mapping.all_fields():
        if entry.status not in VALID_STATUSES:
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') 的 status="
                f"'{entry.status}' 不合法，允许值: {sorted(VALID_STATUSES)}"
            )

    for entry in mapping.all_fields():
        if entry.compute and entry.compute not in VALID_COMPUTE_TYPES:
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') 的 compute="
                f"'{entry.compute}' 不合法，允许值: {sorted(VALID_COMPUTE_TYPES)}"
            )

    for entry in mapping.all_fields():
        if entry.verified_by and entry.verified_by not in VALID_VERIFIED_BY:
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') 的 verified_by="
                f"'{entry.verified_by}' 不合法，允许值: {sorted(VALID_VERIFIED_BY)}"
            )

    for entry in mapping.all_fields():
        if not entry.fallback and entry.status in ("已接入", "部分接入", "待接入"):
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') status='{entry.status}' "
                f"但 fallback 为空"
            )

    complex_compute = {
        "recursive_count",
        "multi_array_count",
        "join_by_key",
        "derived",
        "count_nonzero",
        "manual_confirm",
    }
    for entry in mapping.all_fields():
        if entry.compute in complex_compute and not entry.source_arrays and not entry.notes:
            errors.append(
                f"字段 '{entry.key}' (模块 '{entry.module}') compute='{entry.compute}' "
                f"但缺少 source_arrays 或 notes"
            )

    table_keys: dict[str, str] = {}
    for entry in mapping.all_fields():
        if entry.table_key:
            if entry.table_key in table_keys:
                errors.append(
                    f"table_key '{entry.table_key}' 重复: 字段 '{entry.key}'"
                    f"(模块 '{entry.module}') 与字段 '{table_keys[entry.table_key]}' 冲突"
                )
            else:
                table_keys[entry.table_key] = f"{entry.key}({entry.module})"

    for mod in mapping.modules:
        seen_labels: dict[str, str] = {}
        for entry in mod.fields:
            if entry.label in seen_labels:
                errors.append(
                    f"模块 '{mod.id}' 内 label '{entry.label}' 重复: "
                    f"字段 '{seen_labels[entry.label]}' 与 '{entry.key}'"
                )
            else:
                seen_labels[entry.label] = entry.key

    return errors
