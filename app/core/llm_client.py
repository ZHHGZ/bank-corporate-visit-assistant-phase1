"""LLM客户端 - 独立实现，支持OpenAI兼容接口."""
from __future__ import annotations

from app.config import get_settings
from app.utils.logger import get_logger


logger = get_logger(__name__)


async def chat_completion(prompt: str, system: str = "你是某商业银行的智能分析助手，专注于企业信用和风险分析。") -> str:
    """调用LLM生成文本，失败时返回空字符串."""
    cfg = get_settings()
    if not cfg.llm_enabled or not cfg.llm_api_key:
        logger.info("LLM未启用，跳过AI分析")
        return ""

    try:
        from openai import AsyncOpenAI
        client = AsyncOpenAI(api_key=cfg.llm_api_key, base_url=cfg.llm_base_url)
        resp = await client.chat.completions.create(
            model=cfg.llm_model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
            max_tokens=cfg.llm_max_tokens,
        )
        return resp.choices[0].message.content or ""
    except Exception as e:
        logger.warning(f"LLM调用失败: {e}")
        return ""
