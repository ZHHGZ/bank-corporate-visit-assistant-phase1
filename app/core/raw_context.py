"""原始 ESB 响应上下文缓存.

正式客户画像报告（API-08）需要按交易码组织的**完整原始 ESB 响应**才能复现原项目
的报告结构与字段。但 API-01～07 面向 HiAgent 输出的是聚合结果，会丢弃大报文。

本模块用 SQLite（WAL 模式）按 `企业名称 + 交易码` 落地每次调用（含 mock 兜底）的
完整响应 JSON，并带 TTL，避免旧企业数据混入新报告，也避免把大报文透传给 HiAgent。
API-08 生成报告前从这里读取上下文，缺失的交易码再自动补调。
"""
from __future__ import annotations

import json
import sqlite3
import threading
import time
from pathlib import Path
from typing import Any

from app.config import get_settings
from app.utils.logger import get_logger

logger = get_logger(__name__)

_LOCK = threading.Lock()
_INITIALIZED_DBS: set[str] = set()


def _db_path() -> Path:
    cfg = get_settings()
    path = cfg.resolve_path(cfg.esb_cache_db)
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _connect() -> sqlite3.Connection:
    path = _db_path()
    conn = sqlite3.connect(str(path), timeout=30)
    key = str(path)
    if key not in _INITIALIZED_DBS:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS esb_raw (
                company_name TEXT NOT NULL,
                txn_code     TEXT NOT NULL,
                response     TEXT NOT NULL,
                updated_at   REAL NOT NULL,
                PRIMARY KEY (company_name, txn_code)
            )
            """
        )
        conn.commit()
        _INITIALIZED_DBS.add(key)
    return conn


def save_responses(company_name: str, responses: dict[str, Any]) -> None:
    """将一批 `{交易码: 完整响应}` 落地上下文缓存（覆盖同名同交易码旧值）."""
    company = (company_name or "").strip()
    if not company or not responses:
        return
    now = time.time()
    rows = [
        (company, str(code), json.dumps(resp, ensure_ascii=False), now)
        for code, resp in responses.items()
        if resp is not None
    ]
    if not rows:
        return
    with _LOCK:
        conn = _connect()
        try:
            conn.executemany(
                "INSERT OR REPLACE INTO esb_raw (company_name, txn_code, response, updated_at) "
                "VALUES (?, ?, ?, ?)",
                rows,
            )
            conn.commit()
        finally:
            conn.close()


def load_responses(company_name: str, ttl_seconds: int | None = None) -> dict[str, Any]:
    """读取某企业未过期的全部原始响应，返回 `{交易码: 完整响应}`."""
    company = (company_name or "").strip()
    if not company:
        return {}
    cfg = get_settings()
    ttl = cfg.esb_cache_ttl_seconds if ttl_seconds is None else ttl_seconds
    with _LOCK:
        conn = _connect()
        try:
            cursor = conn.execute(
                "SELECT txn_code, response, updated_at FROM esb_raw WHERE company_name = ?",
                (company,),
            )
            rows = cursor.fetchall()
        finally:
            conn.close()
    now = time.time()
    result: dict[str, Any] = {}
    for txn_code, response, updated_at in rows:
        if ttl and ttl > 0 and (now - float(updated_at)) > ttl:
            continue
        try:
            result[str(txn_code)] = json.loads(response)
        except (ValueError, TypeError):
            continue
    return result


def get_response(company_name: str, txn_code: str, ttl_seconds: int | None = None) -> Any | None:
    """读取单个交易码的原始响应（未过期时）."""
    return load_responses(company_name, ttl_seconds).get(str(txn_code))


def purge_expired(ttl_seconds: int | None = None) -> int:
    """清理过期记录，返回删除条数."""
    cfg = get_settings()
    ttl = cfg.esb_cache_ttl_seconds if ttl_seconds is None else ttl_seconds
    if not ttl or ttl <= 0:
        return 0
    threshold = time.time() - ttl
    with _LOCK:
        conn = _connect()
        try:
            cursor = conn.execute("DELETE FROM esb_raw WHERE updated_at < ?", (threshold,))
            conn.commit()
            return cursor.rowcount or 0
        finally:
            conn.close()
