"""阶段三：9大模块注册表、字段重映射、过滤与排序单元测试."""
from __future__ import annotations

from app.customer_report.opinion_excel_docx_renderer import OpinionReportField
from app.module_registry import (
    MODULE_REGISTRY,
    inject_fallback_fields,
    normalize_module_list,
    old_to_new_module_name,
    remap_fields_to_new_modules,
    sort_fields_by_module_order,
)


def _make_field(module: str, name: str = "测试字段") -> OpinionReportField:
    return OpinionReportField(
        module=module,
        submodule="综合信息",
        name=name,
        interface_source="",
        remark="",
        interface_codes=[],
    )


def test_registry_has_nine_modules() -> None:
    """注册表必须包含 9 个模块."""
    assert len(MODULE_REGISTRY) == 9, f"期望 9 个模块，实际 {len(MODULE_REGISTRY)}"
    assert set(MODULE_REGISTRY.keys()) == {
        "ai_summary", "basic", "equity", "business",
        "finance", "risk", "crm", "interview", "marketing",
    }


def test_registry_order_monotonic() -> None:
    """模块 order 必须从 1 递增到 9."""
    orders = [defn["order"] for defn in MODULE_REGISTRY.values()]
    assert orders == list(range(1, 10)), f"模块 order 不连续: {orders}"


def test_registry_all_have_fallback() -> None:
    """每个模块必须有 fallback 文案."""
    for mod_id, defn in MODULE_REGISTRY.items():
        assert defn.get("fallback"), f"模块 {mod_id} 缺少 fallback"


def test_old_to_new_mapping() -> None:
    """旧模块名正确映射到新模块中文名."""
    assert old_to_new_module_name("风险合规") == "风险提示"
    assert old_to_new_module_name("关键人") == "股权关联"
    assert old_to_new_module_name("关键人物") == "股权关联"
    assert old_to_new_module_name("行内信息") == "行内合作情况"
    assert old_to_new_module_name("资金需求") == "经营能力"
    assert old_to_new_module_name("触达路径") == "客户全貌"


def test_old_to_new_passthrough() -> None:
    """未知模块名原样返回."""
    assert old_to_new_module_name("客户全貌") == "客户全貌"
    assert old_to_new_module_name("不存在的模块") == "不存在的模块"


def test_normalize_module_list_english_ids() -> None:
    """英文 ID 归一化为中文模块名."""
    result = normalize_module_list(["basic", "risk", "marketing"])
    assert result == {"客户全貌", "风险提示", "营销方案"}


def test_normalize_module_list_old_names() -> None:
    """旧中文名归一化."""
    result = normalize_module_list(["风险合规", "关键人", "行内信息"])
    assert result == {"风险提示", "股权关联", "行内合作情况"}


def test_normalize_module_list_mixed() -> None:
    """中英文混合归一化."""
    result = normalize_module_list(["basic", "风险合规", "equity"])
    assert result == {"客户全貌", "风险提示", "股权关联"}


def test_normalize_module_list_none() -> None:
    """None 返回 None 表示不过滤."""
    assert normalize_module_list(None) is None


def test_remap_fields_to_new_modules() -> None:
    """字段 module 正确改为新模块名."""
    fields = [
        _make_field("客户全貌", "企业名称"),
        _make_field("风险合规", "裁判文书"),
        _make_field("关键人", "股东信息"),
        _make_field("资金需求", "动产抵押"),
        _make_field("行内信息", "行内合作"),
    ]
    remapped = remap_fields_to_new_modules(fields)
    modules = {f.module for f in remapped}
    assert "客户全貌" in modules
    assert "风险提示" in modules
    assert "股权关联" in modules
    assert "经营能力" in modules
    assert "行内合作情况" in modules
    assert "风险合规" not in modules
    assert "关键人" not in modules


def test_inject_fallback_fields_fills_nine_modules() -> None:
    """注入兜底字段后必须包含 9 个模块名."""
    fields = [_make_field("客户全貌", "企业名称")]
    filled = inject_fallback_fields(fields)
    module_names = {f.module for f in filled}
    assert len(module_names) == 9, f"期望 9 个模块，实际 {len(module_names)}: {module_names}"
    for mod_id, defn in MODULE_REGISTRY.items():
        assert defn["name"] in module_names, f"缺少模块: {defn['name']}"


def test_sort_fields_by_module_order() -> None:
    """排序后模块按注册表 order 1-9 排列."""
    fields = [
        _make_field("营销方案", "占位"),
        _make_field("客户全貌", "企业名称"),
        _make_field("风险提示", "裁判文书"),
        _make_field("AI综合智能分析建议", "占位"),
    ]
    sorted_fields = sort_fields_by_module_order(fields)
    sorted_modules = [f.module for f in sorted_fields]
    expected_order = [
        "AI综合智能分析建议",
        "客户全貌",
        "股权关联",  # from inject step
    ]
    # 验证排前面的模块顺序
    assert sorted_modules[0] == "AI综合智能分析建议"
    assert sorted_modules[1] == "客户全貌"
    assert sorted_modules[-1] == "营销方案"
    assert sorted_modules[-2] == "风险提示"


def test_modules_filter_returns_only_selected() -> None:
    """modules=["basic"] 只返回客户全貌，不包含未选模块."""
    fields = remap_fields_to_new_modules([
        _make_field("客户全貌", "企业名称"),
        _make_field("客户全貌", "经营范围"),
        _make_field("风险合规", "裁判文书"),
    ])
    fields = inject_fallback_fields(fields)
    fields = sort_fields_by_module_order(fields)

    allowed = normalize_module_list(["basic"])
    filtered = [f for f in fields if f.module in allowed]
    filtered_modules = {f.module for f in filtered}

    assert "客户全貌" in filtered_modules
    assert "风险提示" not in filtered_modules
    assert "股权关联" not in filtered_modules
    assert "AI综合智能分析建议" not in filtered_modules
    assert len(filtered_modules) == 1
