"""日志工具."""
import logging
import sys


def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        h = logging.StreamHandler(sys.stdout)
        h.setFormatter(logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(h)
    level = _resolve_log_level()
    logger.setLevel(level)
    return logger


def _resolve_log_level() -> int:
    valid_levels = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")
    try:
        from app.config import get_settings

        cfg = get_settings()
        raw = (cfg.log_level or "").strip().upper()
        if not raw:
            return logging.INFO
        if raw in valid_levels:
            return getattr(logging, raw)
        logging.getLogger("app.utils.logger").warning(
            "无效的 LOG_LEVEL 配置值 '%s'，已降级为 INFO（有效值：%s）",
            raw,
            ", ".join(valid_levels),
        )
        return logging.INFO
    except Exception:
        return logging.INFO
