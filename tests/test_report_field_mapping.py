"""阶段四-小块一：字段映射台账测试."""
from __future__ import annotations

from copy import deepcopy

from app.customer_report.field_mapping import (
    VALID_COMPUTE_TYPES,
    VALID_SOURCE_TYPES,
    VALID_STATUSES,
    VALID_VERIFIED_BY,
    FieldMappingEntry,
    ModuleMappingConfig,
    ReportFieldMapping,
    load_field_mapping,
    validate_field_mapping,
)


def test_load_field_mapping_succeeds() -> None:
    """加载真实 YAML 台账不报错."""
    mapping = load_field_mapping()
    assert isinstance(mapping, ReportFieldMapping)
    assert mapping.version == "1.2"
    assert len(mapping.modules) == 9


def test_has_nine_modules() -> None:
    """台账必须包含全部 9 个模块 ID."""
    mapping = load_field_mapping()
    module_ids = {m.id for m in mapping.modules}
    assert module_ids == {
        "ai_summary", "basic", "equity", "business",
        "finance", "risk", "crm", "interview", "marketing",
    }


def test_module_names_match_registry() -> None:
    """模块中文名必须与 module_registry 一致."""
    from app.module_registry import MODULE_REGISTRY

    mapping = load_field_mapping()
    for m in mapping.modules:
        expected = MODULE_REGISTRY[m.id]["name"]
        assert m.name == expected, (
            f"模块 '{m.id}' name='{m.name}' 预期 '{expected}'"
        )


def test_module_orders_match_registry() -> None:
    """模块 order 必须与 module_registry 一致."""
    from app.module_registry import MODULE_REGISTRY

    mapping = load_field_mapping()
    for m in mapping.modules:
        expected = MODULE_REGISTRY[m.id]["order"]
        assert m.order == expected, (
            f"模块 '{m.id}' order={m.order} 预期 order={expected}"
        )


def test_field_keys_unique() -> None:
    """所有字段 key 必须全局唯一."""
    mapping = load_field_mapping()
    keys = mapping.field_keys()
    all_fields = mapping.all_fields()
    assert len(keys) == len(all_fields), (
        f"key 数量 ({len(keys)}) != 字段总数 ({len(all_fields)})，存在重复"
    )


def test_source_types_valid() -> None:
    """所有字段 source_type 必须在合法枚举中."""
    mapping = load_field_mapping()
    for entry in mapping.all_fields():
        assert entry.source_type in VALID_SOURCE_TYPES, (
            f"字段 '{entry.key}': source_type='{entry.source_type}' 不合法"
        )


def test_statuses_valid() -> None:
    """所有字段 status 必须在合法枚举中."""
    mapping = load_field_mapping()
    for entry in mapping.all_fields():
        assert entry.status in VALID_STATUSES, (
            f"字段 '{entry.key}': status='{entry.status}' 不合法"
        )


def test_labels_unique_per_module() -> None:
    """同一模块内 label 不能重复."""
    mapping = load_field_mapping()
    for m in mapping.modules:
        labels = [e.label for e in m.fields]
        assert len(labels) == len(set(labels)), (
            f"模块 '{m.id}' 内存在重复 label: {[l for l in labels if labels.count(l) > 1]}"
        )


def test_table_keys_unique() -> None:
    """有 table_key 的字段，table_key 不能跨模块重复."""
    mapping = load_field_mapping()
    seen: dict[str, str] = {}
    for entry in mapping.all_fields():
        if entry.table_key:
            assert entry.table_key not in seen, (
                f"table_key '{entry.table_key}' 重复: "
                f"字段 '{entry.key}'({entry.module}) 与 '{seen[entry.table_key]}' 冲突"
            )
            seen[entry.table_key] = f"{entry.key}({entry.module})"


def test_validation_passes() -> None:
    """正式台账 validate_field_mapping 必须返回空错误列表."""
    mapping = load_field_mapping()
    errors = validate_field_mapping(mapping)
    assert errors == [], f"台账校验失败:\n" + "\n".join(errors)


def test_get_field_by_key() -> None:
    """通过 key 查询字段."""
    mapping = load_field_mapping()
    entry = mapping.get_field_by_key("company_name")
    assert entry is not None
    assert entry.label == "企业名称"
    assert entry.module == "basic"
    assert entry.status == "已接入"

    assert mapping.get_field_by_key("nonexistent_key") is None


def test_get_fields_by_module() -> None:
    """按模块 ID 查询字段列表."""
    mapping = load_field_mapping()

    risk_fields = mapping.get_fields_by_module("risk")
    assert len(risk_fields) == 18
    risk_keys = {f.key for f in risk_fields}
    assert "risk_judgments" in risk_keys
    assert "risk_executed_person" in risk_keys

    assert mapping.get_fields_by_module("nonexistent") == []


def test_basic_module_has_27_fields() -> None:
    """客户全貌模块有 27 个已接入字段（来自旧模板）."""
    mapping = load_field_mapping()
    basic = mapping.get_fields_by_module("basic")
    assert len(basic) == 27, f"客户全貌应有 27 字段，实际 {len(basic)}"


def test_finance_module_has_table_fields() -> None:
    """财务分析模块字段必须有 table_key."""
    mapping = load_field_mapping()
    finance = mapping.get_fields_by_module("finance")
    for entry in finance:
        assert entry.table_key is not None, (
            f"财务分析字段 '{entry.key}' 缺少 table_key"
        )


def test_all_accessible_fields_have_fallback() -> None:
    """status 为已接入/部分接入/待接入的字段，fallback 不能为空."""
    mapping = load_field_mapping()
    for entry in mapping.all_fields():
        if entry.status in ("已接入", "部分接入", "待接入"):
            assert entry.fallback, (
                f"字段 '{entry.key}' status='{entry.status}' 但 fallback 为空"
            )


def test_compute_and_verified_by_valid_when_present() -> None:
    """已声明 compute / verified_by 的字段必须在合法枚举中."""
    mapping = load_field_mapping()
    for entry in mapping.all_fields():
        if entry.compute:
            assert entry.compute in VALID_COMPUTE_TYPES, (
                f"字段 '{entry.key}': compute='{entry.compute}' 不合法"
            )
        if entry.verified_by:
            assert entry.verified_by in VALID_VERIFIED_BY, (
                f"字段 '{entry.key}': verified_by='{entry.verified_by}' 不合法"
            )


def test_high_risk_fields_downgraded_or_annotated() -> None:
    """高风险字段应已降级为部分接入/待确认，并带有 compute 说明."""
    mapping = load_field_mapping()
    patent = mapping.get_field_by_key("patent_count")
    assert patent is not None
    assert patent.status == "部分接入"
    assert patent.compute == "recursive_count"

    financing = mapping.get_field_by_key("financing")
    assert financing is not None
    assert financing.status == "部分接入"
    assert financing.compute == "derived"
    assert "FincngEvArry" in financing.source_arrays

    risk_hit = mapping.get_field_by_key("risk_hit_list")
    assert risk_hit is not None
    assert risk_hit.status == "待确认"


# ===== 校验逻辑单元测试 =====


def _make_mapping(
    modules: list[ModuleMappingConfig],
) -> ReportFieldMapping:
    return ReportFieldMapping(version="test", modules=modules)


def _make_module(
    mod_id: str,
    name: str = "",
    order: int = 0,
    fields: list[FieldMappingEntry] | None = None,
) -> ModuleMappingConfig:
    return ModuleMappingConfig(id=mod_id, name=name, order=order, fields=fields or [])


def _make_field(
    key: str,
    module: str = "basic",
    label: str = "",
    source_type: str = "行外-中数智汇",
    status: str = "已接入",
    fallback: str = "兜底",
) -> FieldMappingEntry:
    return FieldMappingEntry(
        key=key,
        module=module,
        submodule="测试",
        label=label or key,
        source_type=source_type,
        interface_code=None,
        response_path=None,
        table_key=None,
        status=status,
        fallback=fallback,
    )


def test_validate_rejects_invalid_module() -> None:
    """不在 9 大模块注册表中的模块应报错."""
    mapping = _make_mapping([_make_module("fake_module")])
    errors = validate_field_mapping(mapping)
    assert any("fake_module" in e for e in errors)


def test_validate_rejects_missing_modules() -> None:
    """台账缺少模块应报错."""
    mapping = _make_mapping([_make_module("basic"), _make_module("risk")])
    errors = validate_field_mapping(mapping)
    assert any("缺少模块" in e for e in errors)


def test_validate_rejects_duplicate_keys() -> None:
    """重复的 key 应报错."""
    fields = [
        _make_field("dup_key", module="basic"),
        _make_field("dup_key", module="risk"),
    ]
    mapping = _make_mapping([
        _make_module("basic", fields=[fields[0]]),
        _make_module("risk", fields=[fields[1]]),
    ])
    errors = validate_field_mapping(mapping)
    assert len(errors) > 0  # duplicate key triggers error
    # 不期望整体通过，至少有一条关于重复 key
    assert any("dup_key" in e for e in errors)


def test_validate_rejects_invalid_source_type() -> None:
    """非法的 source_type 应报错."""
    mapping = _make_mapping([
        _make_module("basic", fields=[_make_field("test", source_type="非法类型")]),
    ])
    errors = validate_field_mapping(mapping)
    assert any("source_type" in e and "非法类型" in e for e in errors)


def test_validate_rejects_invalid_status() -> None:
    """非法的 status 应报错."""
    mapping = _make_mapping([
        _make_module("basic", fields=[_make_field("test", status="未定义")]),
    ])
    errors = validate_field_mapping(mapping)
    assert any("status" in e and "未定义" in e for e in errors)


def test_validate_empty_key() -> None:
    """空 key 应报错."""
    mapping = _make_mapping([
        _make_module("basic", fields=[_make_field("", module="basic")]),
    ])
    errors = validate_field_mapping(mapping)
    assert any("空 key" in e for e in errors)


def test_validate_access_status_without_fallback() -> None:
    """已接入字段缺少 fallback 应报错."""
    mapping = _make_mapping([
        _make_module(
            "basic",
            fields=[_make_field("no_fb", module="basic", status="已接入", fallback="")],
        ),
    ])
    errors = validate_field_mapping(mapping)
    assert any("fallback" in e and "no_fb" in e for e in errors)
