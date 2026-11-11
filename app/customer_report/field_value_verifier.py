"""字段映射取值验证 - 对照 YAML 台账检查 ESB 响应能否取到值.

仅用于台账联调验收，不接入 DOCX 渲染或 API 主链路。
"""
from __future__ import annotations

import json
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from app.core.esb_client import get_body, is_success
from app.customer_report.field_mapping import (
    FieldMappingEntry,
    ReportFieldMapping,
    load_field_mapping,
)


@dataclass
class FieldVerifyResult:
  """单字段验证结果."""

  key: str
  module: str
  label: str
  status: str
  interface_code: Optional[str]
  compute: Optional[str]
  detail: str
  sample_value: Optional[str] = None


def load_esb_responses_from_json(path: str | Path) -> dict[str, Any]:
  """从 JSON 文件加载 ESB 响应.

  支持两种格式:
  - {"9000000001": {...}, "9000000018": {...}}
  - {"company_name": "xx", "responses": {...}}
  """
  payload = json.loads(Path(path).read_text(encoding="utf-8"))
  if not isinstance(payload, dict):
    raise ValueError(f"ESB 响应文件必须是 JSON 对象: {path}")
  if isinstance(payload.get("responses"), dict):
    responses = payload["responses"]
  else:
    responses = payload
  return {str(code): resp for code, resp in responses.items()}


def build_mock_esb_responses(
  company_name: str,
  interface_codes: Sequence[str] | None = None,
) -> dict[str, Any]:
  """基于 mock_data 构造可用于验证的 ESB 响应."""
  from app.core.mock_data import build_mock_response

  codes = list(interface_codes) if interface_codes else [
    "9000000001", "9000000002", "9000000003", "9000000004", "9000000005",
    "9000000006", "9000000007", "9000000008", "9000000018", "9000000019",
    "9000000023", "9000000020", "9000000029", "9000000036", "9000000056",
  ]
  return {code: build_mock_response(code, company_name) for code in codes}


def verify_field(
  entry: FieldMappingEntry,
  responses: Mapping[str, Any],
) -> FieldVerifyResult:
  """验证单个字段在当前 ESB 响应下是否可取到值."""
  compute = entry.compute or "direct"
  if compute in {"derived", "join_by_key", "manual_confirm"}:
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="skipped",
      interface_code=entry.interface_code,
      compute=compute,
      detail=entry.notes or "复杂字段需代码逻辑验证，当前仅校验台账定义",
    )
  if entry.status in {"待确认", "待补充"} and not entry.interface_code:
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="skipped",
      interface_code=entry.interface_code,
      compute=compute,
      detail="台账状态为待确认/待补充，跳过自动验证",
    )
  if not entry.interface_code:
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="skipped",
      interface_code=None,
      compute=compute,
      detail="无 interface_code，跳过自动验证",
    )

  response = responses.get(entry.interface_code)
  if response is None:
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="miss",
      interface_code=entry.interface_code,
      compute=compute,
      detail="接口响应缺失",
    )
  if not is_success(response):
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="miss",
      interface_code=entry.interface_code,
      compute=compute,
      detail="接口响应失败或非成功状态",
    )

  body = get_body(response) if isinstance(response, dict) else {}
  if compute in {"first_non_empty", "direct"}:
    value, used_path = _resolve_first_non_empty(response, entry.response_path, entry.aliases)
    if _is_displayable(value):
      return FieldVerifyResult(
        key=entry.key,
        module=entry.module,
        label=entry.label,
        status="hit",
        interface_code=entry.interface_code,
        compute=compute,
        detail=f"命中路径: {used_path}",
        sample_value=_preview_value(value),
      )
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="miss",
      interface_code=entry.interface_code,
      compute=compute,
      detail="路径未取到有效值",
    )

  if compute == "recursive_first":
    key_name = _leaf_key(entry.response_path)
    value = _walk_find_key(body, key_name) if key_name else None
    used_path = key_name
    if not _is_displayable(value):
      for alias in entry.aliases:
        alias_key = _leaf_key(alias)
        value = _walk_find_key(body, alias_key) if alias_key else None
        if _is_displayable(value):
          used_path = alias
          break
    if _is_displayable(value):
      return FieldVerifyResult(
        key=entry.key,
        module=entry.module,
        label=entry.label,
        status="hit",
        interface_code=entry.interface_code,
        compute=compute,
        detail=f"递归命中: {used_path}",
        sample_value=_preview_value(value),
      )
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="miss",
      interface_code=entry.interface_code,
      compute=compute,
      detail="递归未找到有效值",
    )

  if compute in {"recursive_count", "multi_array_count"}:
    keys = entry.source_arrays or _array_keys_from_path(entry.response_path)
    counts = {key: len(_records_by_key(body, key)) for key in keys}
    total = sum(counts.values())
    if total > 0:
      status = "hit" if len([c for c in counts.values() if c > 0]) == len(keys) or compute == "recursive_count" else "partial"
      detail = ", ".join(f"{key}={count}" for key, count in counts.items())
      return FieldVerifyResult(
        key=entry.key,
        module=entry.module,
        label=entry.label,
        status=status,
        interface_code=entry.interface_code,
        compute=compute,
        detail=detail,
        sample_value=str(total),
      )
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="miss",
      interface_code=entry.interface_code,
      compute=compute,
      detail="目标数组为空或不存在",
    )

  if compute == "count_nonzero":
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="skipped",
      interface_code=entry.interface_code,
      compute=compute,
      detail=entry.notes or "count_nonzero 需结合业务字段清单验证",
    )

  value, used_path = _resolve_first_non_empty(response, entry.response_path, entry.aliases)
  if _is_displayable(value):
    return FieldVerifyResult(
      key=entry.key,
      module=entry.module,
      label=entry.label,
      status="hit",
      interface_code=entry.interface_code,
      compute=compute,
      detail=f"命中路径: {used_path}",
      sample_value=_preview_value(value),
    )
  return FieldVerifyResult(
    key=entry.key,
    module=entry.module,
    label=entry.label,
    status="miss",
    interface_code=entry.interface_code,
    compute=compute,
    detail="未支持的 compute 或路径未命中",
  )


def verify_module(
  mapping: ReportFieldMapping,
  module_id: str,
  responses: Mapping[str, Any],
) -> list[FieldVerifyResult]:
  return [
    verify_field(entry, responses)
    for entry in mapping.get_fields_by_module(module_id)
  ]


def verify_all(
  mapping: ReportFieldMapping,
  responses: Mapping[str, Any],
  *,
  module_id: str | None = None,
) -> list[FieldVerifyResult]:
  if module_id:
    return verify_module(mapping, module_id, responses)
  return [verify_field(entry, responses) for entry in mapping.all_fields()]


def summarize_results(results: Sequence[FieldVerifyResult]) -> dict[str, int]:
  summary = {"hit": 0, "miss": 0, "partial": 0, "skipped": 0}
  for item in results:
    summary[item.status] = summary.get(item.status, 0) + 1
  return summary


def format_report(results: Sequence[FieldVerifyResult]) -> str:
  lines = ["字段取值验证报告", "=" * 40]
  summary = summarize_results(results)
  lines.append(
    f"命中 {summary.get('hit', 0)} | 部分命中 {summary.get('partial', 0)} | "
    f"未命中 {summary.get('miss', 0)} | 跳过 {summary.get('skipped', 0)}"
  )
  lines.append("")
  for item in results:
    sample = f" => {item.sample_value}" if item.sample_value else ""
    lines.append(
      f"[{item.status}] {item.module}.{item.key} ({item.label}) "
      f"{item.interface_code or '-'} {item.detail}{sample}"
    )
  return "\n".join(lines)


def _is_displayable(value: Any) -> bool:
  if value is None:
    return False
  if isinstance(value, str) and not value.strip():
    return False
  if isinstance(value, (list, dict, tuple, set)) and len(value) == 0:
    return False
  return True


def _preview_value(value: Any, *, limit: int = 80) -> str:
  text = str(value).replace("\n", " ").strip()
  if len(text) > limit:
    return text[: limit - 3] + "..."
  return text


def _leaf_key(path: Optional[str]) -> Optional[str]:
  if not path:
    return None
  part = path.split(".")[-1]
  if part.endswith("[]"):
    return part[:-2]
  if "(" in part:
    return part.split("(", 1)[0]
  return part


def _array_keys_from_path(path: Optional[str]) -> list[str]:
  if not path:
    return []
  keys: list[str] = []
  for part in path.split("."):
    if part.endswith("[]"):
      keys.append(part[:-2])
  return keys


def _resolve_first_non_empty(
  response: Any,
  response_path: Optional[str],
  aliases: Sequence[str],
) -> tuple[Any, Optional[str]]:
  candidates: list[tuple[Optional[str], Any]] = []
  if response_path:
    candidates.append((response_path, _resolve_path_value(response, response_path)))
  for alias in aliases:
    candidates.append((alias, _resolve_path_value(response, alias)))
  for path, value in candidates:
    if _is_displayable(value):
      return value, path
  return None, response_path


def _resolve_path_value(payload: Any, path: str) -> Any:
  path = path.strip()
  if not path:
    return None
  current: Any = payload
  parts = [part for part in path.split(".") if part]
  for index, raw_part in enumerate(parts):
    is_array = raw_part.endswith("[]")
    key = raw_part[:-2] if is_array else raw_part.split("(", 1)[0]
    if isinstance(current, list):
      current = current[0] if current else None
    if not isinstance(current, Mapping):
      return None
    current = current.get(key)
    if is_array:
      if not isinstance(current, list):
        return [] if index == len(parts) - 1 else None
      if index == len(parts) - 1:
        return current
      current = current[0] if current else None
  return current


def _records_by_key(payload: Any, target_key: str) -> list[Mapping[str, Any]]:
  records: list[Mapping[str, Any]] = []
  if isinstance(payload, Mapping):
    for key, value in payload.items():
      if key == target_key and isinstance(value, list):
        records.extend(item for item in value if isinstance(item, Mapping))
      else:
        records.extend(_records_by_key(value, target_key))
  elif isinstance(payload, list):
    for item in payload:
      records.extend(_records_by_key(item, target_key))
  return records


def _walk_find_key(payload: Any, key: str) -> Any:
  if isinstance(payload, Mapping):
    if key in payload and _is_displayable(payload[key]):
      return payload[key]
    for value in payload.values():
      found = _walk_find_key(value, key)
      if found is not None:
        return found
  elif isinstance(payload, list):
    for item in payload:
      found = _walk_find_key(item, key)
      if found is not None:
        return found
  return None
