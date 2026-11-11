"""环境变量读取兼容层.

正式渲染器与 LLM 摘要模块原本依赖 `src.utils.logger.get_env`，从 llm-standalone
项目根 `.env` 读取 `LLM_MODELS` / `LLM_DEFAULT_MODEL` 等配置。迁移到本项目后，
本项目沿用 `LLM_API_KEY` / `LLM_BASE_URL` / `LLM_MODEL` 风格的配置，这里提供一个
等价的 `get_env`：

- 优先返回真实环境变量（含 `.env`）；
- 当请求 `LLM_MODELS` / `LLM_DEFAULT_MODEL` 且未显式配置时，从本项目 LLM 配置
  合成一个单模型别名，使正式渲染器的 AI 分析可直接复用本项目 LLM 设置。
"""
from __future__ import annotations

import json
import os
from pathlib import Path

from dotenv import load_dotenv

# app/customer_report/_env.py -> parents[2] == enterprise-report-api/
PROJECT_ROOT = Path(__file__).resolve().parents[2]
ENV_PATH = PROJECT_ROOT / ".env"
_ENV_LOADED = False
_SYNTH_ALIAS = "enterprise-report-llm"


def _load_project_env() -> None:
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    load_dotenv(ENV_PATH)
    _ENV_LOADED = True


def _synth_llm_models() -> str:
    """用本项目 LLM_* 配置合成 llm-standalone 风格的 LLM_MODELS 数组."""
    api_key = os.getenv("LLM_API_KEY", "")
    base_url = os.getenv("LLM_BASE_URL", "")
    model = os.getenv("LLM_MODEL", "")
    if not (base_url and model):
        return "[]"
    return json.dumps(
        [
            {
                "name": _SYNTH_ALIAS,
                "base_url": base_url,
                "model_name": model,
                "api_key": api_key,
            }
        ],
        ensure_ascii=False,
    )


def get_env(key: str, default: str = "") -> str:
    """读取环境变量，必要时为正式渲染器合成 LLM 配置."""
    _load_project_env()
    value = os.getenv(key)
    if value not in (None, ""):
        return value  # type: ignore[return-value]
    if key == "LLM_MODELS":
        synthesized = _synth_llm_models()
        if synthesized != "[]":
            return synthesized
    if key == "LLM_DEFAULT_MODEL":
        if os.getenv("LLM_BASE_URL") and os.getenv("LLM_MODEL"):
            return _SYNTH_ALIAS
    return default
