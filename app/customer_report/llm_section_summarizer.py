"""LLM summaries for customer visit report sections.

The client uses an OpenAI-compatible chat completion API. Model connection
settings come from the root `.env` `LLM_MODELS` configuration.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PROMPT_CONFIG_PATH = PROJECT_ROOT / "configs" / "customer_report_llm_prompts.yaml"
ALL_SECTION_NAMES = ("客户全貌", "关键人", "经营能力", "触达路径", "资金需求", "风险合规")

from app.customer_report.common import safe_filename
from app.customer_report._env import get_env


def get_customer_report_model_config(model_alias: str) -> dict[str, Any]:
    models = json.loads(get_env("LLM_MODELS", "[]"))
    if not isinstance(models, list):
        raise ValueError("LLM_MODELS must be a JSON array")
    normalized_alias = model_alias.strip().lower()
    for item in models:
        if isinstance(item, Mapping) and str(item.get("name") or "").strip().lower() == normalized_alias:
            return dict(item)
    available = [str(item.get("name")) for item in models if isinstance(item, Mapping) and item.get("name")]
    raise ValueError(f"Unknown customer report LLM model alias: {model_alias}. Available: {available}")


def env_path_value(env_name: str, default: Path) -> str:
    value = get_env(env_name, str(default))
    path = Path(value)
    if path.is_absolute():
        return str(path)
    return str(PROJECT_ROOT / path)


SECTION_INTERFACE_CODES = {
    "客户全貌": [
        "9000000061",
        "9000000001",
        "9000000004",
        "9000000006",
        "9000000062",
        "9000000057",
        "9000000063",
        "9000000058",
        "9000000064",
        "9000000065",
        "9000000066",
    ],
    "关键人": ["9000000010", "9000000067", "9000000016"],
    "经营能力": [
        "9000000068",
        "9000000015",
        "9000000069",
        "9000000019",
        "9000000070",
        "9000000005",
        "9000000002",
        "9000000003",
        "9000000007",
        "9000000071",
        "9000000008",
        "9000000018",
        "9000000072",
        "9000000073",
        "9000000074",
        "9000000075",
        "9000000076",
    ],
    "触达路径": [
        "9000000011",
        "9000000014",
        "9000000030",
        "9000000031",
        "9000000012",
        "9000000013",
        "9000000032",
        "9000000033",
        "9000000017",
        "9000000034",
    ],
    "资金需求": [
        "9000000020",
        "9000000021",
        "9000000022",
        "9000000023",
        "9000000024",
        "9000000025",
        "9000000026",
        "9000000027",
        "9000000028",
        "9000000029",
    ],
    "风险合规": [
        "9000000035",
        "9000000036",
        "9000000037",
        "9000000060",
        "9000000038",
        "9000000039",
        "9000000040",
        "9000000041",
        "9000000042",
        "9000000043",
        "9000000044",
        "9000000045",
        "9000000046",
        "9000000047",
        "9000000048",
        "9000000049",
        "9000000050",
        "9000000051",
        "9000000052",
        "9000000053",
        "9000000054",
        "9000000055",
        "9000000077",
        "9000000056",
        "9000000059",
        "9000000078",
    ],
}
LLM_EXCLUDED_FIELD_SOURCES = {
    "JdgntWritDtlTxtCntnt",
}
MAX_LLM_FIELD_TEXT_LENGTH = 300


@dataclass
class CustomerReportLLMConfig:
    """Configuration for report section summaries."""

    provider: str = "qwen"
    base_url: str = ""
    model: str = ""
    api_key: str = ""
    timeout: int = 120
    max_tokens: int = 800
    temperature: float = 0.2
    debug_payload_dir: str = field(default_factory=lambda: get_env("CUSTOMER_REPORT_LLM_DEBUG_PAYLOAD_DIR", ""))
    trace_path: str = field(default_factory=lambda: get_env("CUSTOMER_REPORT_LLM_TRACE_PATH", ""))
    prompt_config_path: str = field(default_factory=lambda: env_path_value("CUSTOMER_REPORT_LLM_PROMPT_CONFIG_PATH", DEFAULT_PROMPT_CONFIG_PATH))

    def __post_init__(self) -> None:
        model_alias = get_env("LLM_DEFAULT_MODEL") or self.provider
        model_config = get_customer_report_model_config(model_alias)
        self.provider = str(model_config.get("name") or model_alias)
        self.base_url = str(model_config.get("base_url") or "")
        self.model = str(model_config.get("model_name") or model_config.get("name") or model_alias)
        self.api_key = str(model_config.get("api_key") or "")


class CustomerReportLLMSummarizer:
    """Generate concise section summaries from normalized ESB data."""

    def __init__(self, config: CustomerReportLLMConfig | None = None) -> None:
        from openai import OpenAI

        self.config = config or CustomerReportLLMConfig()
        self.client = OpenAI(
            base_url=self.config.base_url,
            api_key=self.config.api_key or "notempty",
            timeout=self.config.timeout,
        )

    def summarize_section(
        self,
        *,
        company_name: str,
        section_name: str,
        normalized_json: Mapping[str, Any],
        metrics_json: Mapping[str, Any],
    ) -> str:
        if not self.config.api_key and self.config.provider != "private":
            raise ValueError("LLM_MODELS selected model api_key is empty")
        section_payload = build_section_payload(
            company_name=company_name,
            section_name=section_name,
            normalized_json=normalized_json,
            metrics_json=metrics_json,
        )
        self.save_debug_payload(company_name, section_name, section_payload)
        messages = [
            {"role": "system", "content": build_system_prompt(self.config.prompt_config_path)},
            {"role": "user", "content": build_section_prompt(section_name, self.config.prompt_config_path)},
            {"role": "user", "content": json.dumps(section_payload, ensure_ascii=False, indent=2)},
        ]
        response = self.client.chat.completions.create(
            model=self.config.model,
            messages=messages,
            temperature=self.config.temperature,
            max_tokens=self.config.max_tokens,
        )
        summary = clamp_summary(response.choices[0].message.content or "")
        self.append_trace(company_name, section_name, section_payload, summary)
        return summary

    def save_debug_payload(self, company_name: str, section_name: str, payload: Mapping[str, Any]) -> None:
        if not self.config.debug_payload_dir:
            return
        from pathlib import Path

        output_dir = Path(self.config.debug_payload_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        path = output_dir / (
            f"{safe_filename(company_name, default='未命名')}_"
            f"{safe_filename(section_name, default='未命名')}_llm_payload.json"
        )
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def append_trace(
        self,
        company_name: str,
        section_name: str,
        payload: Mapping[str, Any],
        summary: str,
    ) -> None:
        if not self.config.trace_path:
            return
        path = Path(self.config.trace_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        if not path.exists():
            path.write_text(build_trace_header(company_name, self.config), encoding="utf-8")
        with path.open("a", encoding="utf-8") as file:
            file.write("\n")
            file.write("=" * 80 + "\n")
            file.write(f"模块：{section_name}\n")
            file.write(f"调用时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            file.write("\n【模块提示词】\n")
            file.write(build_section_prompt(section_name, self.config.prompt_config_path))
            file.write("\n【送给大模型的数据】\n")
            file.write(json.dumps(payload, ensure_ascii=False, indent=2))
            file.write("\n\n【大模型返回结果】\n")
            file.write(summary.strip())
            file.write("\n")


def build_section_payload(
    *,
    company_name: str,
    section_name: str,
    normalized_json: Mapping[str, Any],
    metrics_json: Mapping[str, Any],
) -> dict[str, Any]:
    interface_codes = SECTION_INTERFACE_CODES.get(section_name) or []
    interfaces = normalized_json.get("interfaces") or {}
    section_metrics = ((metrics_json.get("sections") or {}).get(section_name) or {}).get("metrics") or {}
    generated_counts = ((metrics_json.get("sections") or {}).get(section_name) or {}).get("generated_array_counts") or {}
    compact_interfaces = {}
    for code in interface_codes:
        interface = interfaces.get(code)
        if not interface:
            continue
        if not is_success_interface(interface):
            continue
        interface_name = str(interface.get("interface_name") or code)
        compact_interfaces[f"{interface_name}({code})"] = compact_interface_for_llm(interface)
    return {
        "company_name": company_name,
        "section_name": section_name,
        "模块统计": section_metrics,
        "接口数据量": generated_counts,
        "接口数据": compact_interfaces,
    }


def compact_interface_for_llm(interface: Mapping[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, item in (interface.get("fields") or {}).items():
        if not isinstance(item, Mapping) or should_skip_llm_field(key, item):
            continue
        value = item.get("value") if isinstance(item, Mapping) else None
        if value not in (None, "") and not is_placeholder_value(value):
            result[str(item.get("name") or key)] = compact_llm_value(value)
    array_counts = {}
    for path, item in (interface.get("arrays") or {}).items():
        count = item.get("count") if isinstance(item, Mapping) else 0
        if count:
            array_counts[str(item.get("name") or path)] = count
    if array_counts:
        result["数据条数"] = array_counts
    if interface.get("business_code"):
        result["业务返回码"] = interface.get("business_code")
    if interface.get("business_message"):
        result["业务返回说明"] = interface.get("business_message")
    return result


def is_success_interface(interface: Mapping[str, Any]) -> bool:
    business_code = str(interface.get("business_code") or "").strip()
    if business_code and business_code != "200":
        return False
    business_message = str(interface.get("business_message") or "").strip().lower()
    if any(marker in business_message for marker in ("not found", "参数错误", "参数不能为空", "请求参数错误", "查询量已耗尽")):
        return False
    if not business_code and business_message:
        return False
    return True


def should_skip_llm_field(field_key: Any, item: Mapping[str, Any]) -> bool:
    source = str(item.get("source") or "")
    name = str(item.get("name") or "")
    path = str(item.get("path") or "")
    if source in LLM_EXCLUDED_FIELD_SOURCES:
        return True
    text = f"{field_key} {name} {path}"
    return "JdgntWritDtlTxtCntnt" in text


def compact_llm_value(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    text = " ".join(value.split())
    if len(text) <= MAX_LLM_FIELD_TEXT_LENGTH:
        return text
    return text[:MAX_LLM_FIELD_TEXT_LENGTH].rstrip() + "...[已截断]"


def is_placeholder_value(value: Any) -> bool:
    text = str(value)
    return text.endswith("-模拟值") or text.endswith("测试值")


def clamp_summary(text: str, limit: int = 800) -> str:
    lines = [
        " ".join(line.strip().split())
        for line in str(text or "").strip().splitlines()
    ]
    cleaned = "\n".join(line for line in lines if line)
    if len(cleaned) <= limit:
        return cleaned
    truncated = cleaned[:limit]
    for mark in ("\n", "。", "；", "，"):
        pos = truncated.rfind(mark)
        if pos >= int(limit * 0.75):
            return truncated[:pos].rstrip() if mark == "\n" else truncated[: pos + 1]
    return (truncated.rstrip("，；、 ") + "。")[:limit]


def build_system_prompt(prompt_config_path: str | Path = DEFAULT_PROMPT_CONFIG_PATH) -> str:
    config = load_prompt_config(prompt_config_path)
    return "\n\n".join(
        part
        for part in (
            config.get("system_prompt"),
            config.get("common_output_requirements"),
        )
        if part
    ) or (
        "你是银行客户经理访前报告分析助手。请基于给定接口数据生成模块小结，"
        "输出内容控制在300-800字之间，不要编造未提供的数据。"
    )


def build_section_prompt(section_name: str, prompt_config_path: str | Path = DEFAULT_PROMPT_CONFIG_PATH) -> str:
    config = load_prompt_config(prompt_config_path)
    section_prompts = config.get("section_prompts") or {}
    prompt = section_prompts.get(section_name)
    return str(prompt or f"请为“{section_name}”模块生成银行客户经理访前报告小结。")


def load_prompt_config(path: str | Path = DEFAULT_PROMPT_CONFIG_PATH) -> dict[str, Any]:
    prompt_path = Path(path)
    if not prompt_path.exists():
        return {}
    try:
        import yaml  # type: ignore

        data = yaml.safe_load(prompt_path.read_text(encoding="utf-8")) or {}
        return data if isinstance(data, dict) else {}
    except Exception:
        return parse_prompt_config_limited(prompt_path.read_text(encoding="utf-8"))


def parse_prompt_config_limited(text: str) -> dict[str, Any]:
    data: dict[str, Any] = {"section_prompts": {}}
    current_key = ""
    current_section = ""
    current_lines: list[str] = []
    mode = ""

    def flush() -> None:
        nonlocal current_key, current_section, current_lines
        value = "\n".join(current_lines).rstrip()
        if mode == "top" and current_key:
            data[current_key] = value
        elif mode == "section" and current_section:
            data["section_prompts"][current_section] = value
        current_key = ""
        current_section = ""
        current_lines = []

    for raw in text.splitlines():
        if raw.startswith("system_prompt: |-"):
            flush()
            mode = "top"
            current_key = "system_prompt"
            continue
        if raw.startswith("common_output_requirements: |-"):
            flush()
            mode = "top"
            current_key = "common_output_requirements"
            continue
        if raw.startswith("section_prompts:"):
            flush()
            mode = ""
            continue
        section_match = raw.startswith("  ") and raw.strip().endswith(": |-")
        if section_match:
            flush()
            mode = "section"
            current_section = raw.strip()[:-3].strip()
            continue
        if mode == "top" and raw.startswith("  "):
            current_lines.append(raw[2:])
        elif mode == "section" and raw.startswith("    "):
            current_lines.append(raw[4:])
    flush()
    return data


def build_trace_header(company_name: str, config: CustomerReportLLMConfig) -> str:
    return (
        "客户报告大模型调用记录\n"
        f"公司名称：{company_name}\n"
        f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"模型提供方：{config.provider}\n"
        f"模型地址：{config.base_url}\n"
        f"模型名称：{config.model}\n"
        f"温度：{config.temperature}\n"
        f"最大输出Token：{config.max_tokens}\n"
    )


def apply_llm_section_summary(
    payload: dict[str, Any],
    *,
    company_name: str,
    section_name: str = "客户全貌",
    summarizer: CustomerReportLLMSummarizer | None = None,
) -> dict[str, Any]:
    """Replace one section summary in payload with an LLM-generated summary."""
    summarizer = summarizer or CustomerReportLLMSummarizer()
    summary = summarizer.summarize_section(
        company_name=company_name,
        section_name=section_name,
        normalized_json=payload["normalized_json"],
        metrics_json=payload["metrics_json"],
    )
    payload["llm_analysis_json"][section_name] = summary
    pdf_section = (payload["pdf_data_json"].get("sections") or {}).get(section_name)
    if isinstance(pdf_section, dict):
        pdf_section["summary"] = summary
    return payload


def apply_llm_section_summaries(
    payload: dict[str, Any],
    *,
    company_name: str,
    section_names: tuple[str, ...] = ALL_SECTION_NAMES,
    summarizer: CustomerReportLLMSummarizer | None = None,
) -> dict[str, Any]:
    summarizer = summarizer or CustomerReportLLMSummarizer()
    reset_trace_file(summarizer.config.trace_path)
    for section_name in section_names:
        payload = apply_llm_section_summary(
            payload,
            company_name=company_name,
            section_name=section_name,
            summarizer=summarizer,
        )
    return payload


def reset_trace_file(trace_path: str) -> None:
    if not trace_path:
        return
    path = Path(trace_path)
    if path.exists():
        path.unlink()
