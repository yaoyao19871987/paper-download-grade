from __future__ import annotations

import hashlib
import json
import os
import re
import time
from pathlib import Path
from typing import Any

import requests
import win32crypt


REPO_ROOT_MARKERS = (
    ".git",
    "config/pipeline/pipeline.config.json",
    "pipeline/pipeline.config.json",
)
KIMI_DEFAULT_BASE_URL = "https://api.kimi.com/coding/v1"
KIMI_DEFAULT_MODEL = "kimi-for-coding"
FEEDBACK_TIMEOUT_SECONDS = 180
FEEDBACK_RETRYABLE_STATUS_CODES = {408, 409, 425, 429, 500, 502, 503, 504}


def score_ratio(item: dict[str, Any]) -> float:
    raw_score = item.get("score")
    raw_max = item.get("max_score")
    try:
        score = float(raw_score or 0.0)
        max_score = float(raw_max or 0.0)
    except (TypeError, ValueError):
        return 0.0
    if max_score <= 0:
        return 0.0
    return max(0.0, min(score / max_score, 1.0))


def clean_feedback_text(text: Any) -> str:
    cleaned = re.sub(r"\s+", " ", str(text or "")).strip()
    cleaned = cleaned.strip(" ;；,.。，、")
    return cleaned


def build_student_feedback(entry: dict[str, Any], grade_data: dict[str, Any]) -> str:
    return generate_student_feedback_artifact(entry, grade_data)["markdown"]


def generate_student_feedback_artifact(entry: dict[str, Any], grade_data: dict[str, Any]) -> dict[str, Any]:
    valid_grade_data = _has_valid_grade_data(grade_data)
    prompt_payload = _build_feedback_input(entry, grade_data)
    signature = hashlib.sha256(
        json.dumps(prompt_payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()
    artifact_dir = _feedback_artifact_dir(entry)
    raw_response_path = artifact_dir / "student_feedback_kimi_raw.json"
    metadata_path = artifact_dir / "student_feedback_kimi_meta.json"

    force_retry = entry.get("force_feedback_retry", False)
    if not force_retry and metadata_path.exists():
        try:
            cached_meta = json.loads(metadata_path.read_text(encoding="utf-8"))
            if cached_meta.get("input_signature") == signature and cached_meta.get("status") == "ok":
                markdown_path = artifact_dir / "student_feedback_kimi.md"
                # If we don't have the markdown file cached directly, we might need to recreate it from raw?
                # Actually, the markdown is saved outside this function by the caller. But we return it.
                # If status is OK, we can load the parsed json and re-render.
                if raw_response_path.exists():
                    raw_data = json.loads(raw_response_path.read_text(encoding="utf-8"))
                    parsed = raw_data.get("parsed")
                    if parsed:
                        markdown = _render_feedback_markdown(entry, grade_data, parsed)
                        return {"markdown": markdown, **cached_meta}
        except Exception:
            pass

    if not valid_grade_data:
        markdown = _build_missing_grade_feedback(entry)
        metadata = {
            "status": "upstream_incomplete",
            "model": None,
            "attempts": 0,
            "error": entry.get("grade_error") or "invalid or empty grade data",
            "raw_response_path": None,
            "input_signature": signature,
        }
        _write_feedback_metadata(metadata_path, metadata)
        return {"markdown": markdown, **metadata}

    try:
        response_json, feedback_json, model_name, attempts = _generate_feedback_with_kimi(prompt_payload)
        raw_response_path.write_text(
            json.dumps(
                {
                    "input": prompt_payload,
                    "response": _sanitize_raw_response(response_json),
                    "parsed": feedback_json,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        markdown = _render_feedback_markdown(entry, grade_data, feedback_json)
        metadata = {
            "status": "ok",
            "model": model_name,
            "attempts": attempts,
            "error": None,
            "raw_response_path": str(raw_response_path.resolve()),
            "input_signature": signature,
        }
        _write_feedback_metadata(metadata_path, metadata)
        return {"markdown": markdown, **metadata}
    except Exception as exc:
        markdown = _build_generation_failed_feedback(entry, grade_data, str(exc))
        metadata = {
            "status": "generation_failed",
            "model": KIMI_DEFAULT_MODEL,
            "attempts": 0,
            "error": str(exc),
            "raw_response_path": str(raw_response_path.resolve()) if raw_response_path.exists() else None,
            "input_signature": signature,
        }
        _write_feedback_metadata(metadata_path, metadata)
        return {"markdown": markdown, **metadata}


def _has_valid_grade_data(grade_data: dict[str, Any]) -> bool:
    if not isinstance(grade_data, dict) or not grade_data:
        return False
    summary = grade_data.get("summary")
    if not isinstance(summary, dict):
        return False
    if summary.get("decision"):
        return True
    return any(
        grade_data.get(key)
        for key in ("format_items", "content_items", "text_review", "visual_review", "reference_audit")
    )


def _build_feedback_input(entry: dict[str, Any], grade_data: dict[str, Any]) -> dict[str, Any]:
    summary = grade_data.get("summary", {})
    visual_review = grade_data.get("visual_review", {})
    text_review = grade_data.get("text_review", {})
    reference_audit = grade_data.get("reference_audit", {})
    extracted = grade_data.get("extracted", {})
    gate = grade_data.get("gate", {})
    return {
        "student": {
            "sid": entry.get("sid"),
            "name": entry.get("name"),
            "paper_title": extracted.get("title") or entry.get("paper_title"),
            "stage": summary.get("stage") or entry.get("stage"),
            "teacher_name": entry.get("teacher_name"),
        },
        "upstream_summary": {
            "decision": summary.get("decision"),
            "total_score": summary.get("total_score"),
            "gate_reasons": _limit_list(summary.get("gate_reasons"), 8),
            "grade_error": entry.get("grade_error"),
        },
        "format_scores": _compact_items(grade_data.get("format_items"), limit=8),
        "content_scores": _compact_items(grade_data.get("content_items"), limit=8),
        "visual_review": {
            "mode": visual_review.get("mode"),
            "model": visual_review.get("model"),
            "overall_verdict": visual_review.get("overall_verdict"),
            "visual_order_score": visual_review.get("visual_order_score"),
            "major_issues": _limit_list(visual_review.get("major_issues"), 6),
            "minor_issues": _limit_list(visual_review.get("minor_issues"), 4),
            "evidence": _limit_list(visual_review.get("evidence"), 4),
            "notes": _limit_list(visual_review.get("notes"), 4),
        },
        "text_review": {
            "mode": text_review.get("mode"),
            "fused_score": text_review.get("fused_score"),
            "confidence": text_review.get("confidence"),
            "major_problems": _limit_list(text_review.get("major_problems"), 8),
            "revision_actions": _limit_list(text_review.get("revision_actions"), 8),
            "strengths": _limit_list(text_review.get("strengths"), 4),
            "notes": _limit_list(text_review.get("notes"), 4),
        },
        "reference_review": {
            "notes": _limit_list(reference_audit.get("notes"), 8),
            "summary": _compact_reference_summary(reference_audit),
        },
        "gate": {
            "reasons": _limit_list(gate.get("reasons"), 8),
        },
    }


def _compact_items(items: Any, limit: int) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for item in list(items or []):
        if not isinstance(item, dict):
            continue
        normalized.append(
            {
                "name": item.get("name"),
                "score": item.get("score"),
                "max_score": item.get("max_score"),
                "status": item.get("status"),
                "suggestions": _limit_list(item.get("suggestions"), 3),
            }
        )
    normalized.sort(key=lambda item: (score_ratio(item), str(item.get("name") or "")))
    return normalized[:limit]


def _compact_reference_summary(reference_audit: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(reference_audit, dict):
        return {}
    summary: dict[str, Any] = {}
    for key in (
        "citation_coverage",
        "verified_reference_count",
        "unverified_reference_count",
        "unmatched_citations",
        "orphan_references",
    ):
        if key in reference_audit:
            summary[key] = reference_audit.get(key)
    return summary


def _limit_list(values: Any, limit: int) -> list[str]:
    result: list[str] = []
    for value in list(values or []):
        text = clean_feedback_text(value)
        if text and text not in result:
            result.append(text)
        if len(result) >= limit:
            break
    return result


def _generate_feedback_with_kimi(prompt_payload: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], str, int]:
    config = _resolve_kimi_config()
    url = _build_chat_completions_url(config["api_base_url"])
    headers = {
        "Authorization": f"Bearer {config['api_key']}",
        "Content-Type": "application/json",
        "User-Agent": os.getenv("KIMI_CODING_USER_AGENT") or "KimiCLI/1.0",
    }
    payload = {
        "model": config["model"],
        "messages": [
            {"role": "system", "content": _feedback_system_prompt()},
            {"role": "user", "content": json.dumps(prompt_payload, ensure_ascii=False)},
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
        "stream": True,
    }
    max_retries = max(1, int(os.getenv("PIPELINE_FEEDBACK_MAX_RETRIES", "3")))
    backoff_seconds = max(1.0, float(os.getenv("PIPELINE_FEEDBACK_RETRY_BACKOFF_SECONDS", "8")))
    last_error: Exception | None = None

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(
                url,
                headers=headers,
                json=payload,
                timeout=FEEDBACK_TIMEOUT_SECONDS,
                stream=True,
            )
        except (requests.Timeout, requests.ConnectionError) as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(backoff_seconds * attempt)
                continue
            raise RuntimeError(f"Kimi student feedback request failed after {max_retries} attempts: {exc}") from exc

        if response.status_code >= 400:
            retryable = response.status_code in FEEDBACK_RETRYABLE_STATUS_CODES
            error = RuntimeError(_format_provider_error(response))
            last_error = error
            if retryable and attempt < max_retries:
                time.sleep(backoff_seconds * attempt)
                continue
            raise error

        response_json = _decode_streamed_chat_response(response)
        feedback_json = _parse_feedback_json(response_json)
        validation_error = _validate_feedback_json(feedback_json)
        if validation_error:
            last_error = RuntimeError(validation_error)
            if attempt < max_retries:
                time.sleep(backoff_seconds * attempt)
                continue
            raise last_error
        return response_json, feedback_json, config["model"], attempt

    raise RuntimeError(f"Kimi student feedback request failed after {max_retries} attempts: {last_error}")


def _feedback_system_prompt() -> str:
    return (
        "你是毕业论文指导老师，现在负责把上游评分系统的结构化结果，改写成给学生看的最终评语。"
        "上游已经完成了文档抽取、规则评分、视觉审查、文本审查、引用核验，你不能重新编造评分，只能基于提供的数据解释问题。"
        "输出必须是中文，必须简单、明确、可操作，像老师直接布置修改任务。"
        "不要写空话，不要重复系统字段名，不要泛泛鼓励，不要套模板。"
        "如果数据提示论文问题严重，就直说严重在哪里；如果上游数据不完整，也要直说本次批改没有完整完成。"
        "必须严格返回 JSON 对象，字段固定为：status、title、overall、priority_actions、format_fix、content_fix、writing_fix、citation_fix、strengths、next_steps、risk_notice。"
        "其中 status 只能是 ok 或 incomplete。overall 必须是一段完整中文。priority_actions 和 next_steps 至少各给 2 条。各 list 元素都必须是完整句子。"
    )


def _resolve_kimi_config() -> dict[str, str]:
    api_key = os.getenv("MOONSHOT_API_KEY")
    api_base_url = os.getenv("MOONSHOT_BASE_URL")
    model = os.getenv("PIPELINE_FEEDBACK_MODEL") or os.getenv("KIMI_FEEDBACK_MODEL") or KIMI_DEFAULT_MODEL
    entry_path = _credential_entry_path("moonshot_kimi")
    if entry_path.exists():
        entry = _load_credential_entry("moonshot_kimi")
        metadata = entry.get("metadata") or {}
        if not api_key:
            api_key = entry.get("fields", {}).get("api_key")
        if not api_base_url:
            api_base_url = str(metadata.get("api_base_url") or "")
        if model == KIMI_DEFAULT_MODEL and metadata.get("default_model"):
            model = str(metadata["default_model"])
    if not api_key:
        raise RuntimeError("Kimi credential is missing. Save moonshot_kimi first.")
    return {
        "api_key": api_key,
        "api_base_url": api_base_url or KIMI_DEFAULT_BASE_URL,
        "model": model,
    }


def _build_chat_completions_url(api_base_url: str) -> str:
    base = (api_base_url or "").rstrip("/")
    if not base:
        raise RuntimeError("Kimi API base URL is missing.")
    if base.endswith("/v1"):
        return base + "/chat/completions"
    return base + "/v1/chat/completions"


def _decode_streamed_chat_response(response: requests.Response) -> dict[str, Any]:
    content_type = str(response.headers.get("Content-Type") or "").lower()
    if "text/event-stream" not in content_type:
        return response.json()
    final_payload: dict[str, Any] | None = None
    fragments: list[str] = []
    for _, data in _iter_sse_events(response):
        if not data or data == "[DONE]":
            continue
        try:
            payload = json.loads(data)
        except ValueError:
            continue
        if not isinstance(payload, dict):
            continue
        if not isinstance(payload.get("choices"), list):
            continue
        final_payload = payload
        fragments.extend(_extract_choice_delta_text(payload.get("choices") or []))

    content = "".join(fragments)
    if not content and isinstance(final_payload, dict):
        content = _extract_choice_message_text(final_payload.get("choices") or [])
    if not content:
        raise RuntimeError("Kimi returned no parsable feedback content.")

    base_choice: dict[str, Any] = {}
    if isinstance(final_payload, dict):
        choices = final_payload.get("choices") or []
        if choices and isinstance(choices[0], dict):
            base_choice = dict(choices[0])
    message = dict(base_choice.get("message") or {})
    message["content"] = content
    base_choice["message"] = message
    base_choice.pop("delta", None)
    return {"choices": [base_choice]}


def _iter_sse_events(response: requests.Response) -> list[tuple[str, str]]:
    events: list[tuple[str, str]] = []
    event_name = ""
    data_lines: list[str] = []
    for raw_line in response.iter_lines(decode_unicode=True):
        if raw_line is None:
            continue
        line = str(raw_line).lstrip("\ufeff")
        if not line:
            if data_lines:
                events.append((event_name, "\n".join(data_lines)))
                event_name = ""
                data_lines = []
            continue
        if line.startswith(":"):
            continue
        field, _, value = line.partition(":")
        value = value.lstrip()
        if field == "event":
            event_name = value
        elif field == "data":
            data_lines.append(value)
    if data_lines:
        events.append((event_name, "\n".join(data_lines)))
    return events


def _extract_choice_delta_text(choices: list[Any]) -> list[str]:
    texts: list[str] = []
    for choice in choices:
        if not isinstance(choice, dict):
            continue
        delta = choice.get("delta") or {}
        texts.extend(_coerce_stream_text(delta.get("content")))
        texts.extend(_coerce_stream_text(choice.get("text")))
    return texts


def _extract_choice_message_text(choices: list[Any]) -> str:
    parts: list[str] = []
    for choice in choices:
        if not isinstance(choice, dict):
            continue
        message = choice.get("message") or {}
        parts.extend(_coerce_stream_text(message.get("content")))
    return "".join(parts)


def _coerce_stream_text(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value] if value.strip() else []
    if isinstance(value, list):
        items: list[str] = []
        for item in value:
            if isinstance(item, str):
                if item.strip():
                    items.append(item)
                continue
            if not isinstance(item, dict):
                continue
            text = item.get("text") or item.get("content")
            if isinstance(text, str) and text.strip():
                items.append(text)
        return items
    if isinstance(value, dict):
        text = value.get("text") or value.get("content")
        if isinstance(text, str) and text.strip():
            return [text]
    return []


def _parse_feedback_json(response_json: dict[str, Any]) -> dict[str, Any]:
    choices = response_json.get("choices") or []
    if not choices:
        raise RuntimeError("Kimi feedback response has no choices.")
    content = (choices[0].get("message") or {}).get("content")
    if isinstance(content, list):
        content = "".join(_coerce_stream_text(content))
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("Kimi feedback response has empty content.")
    return json.loads(_extract_json_object(content))


def _validate_feedback_json(payload: dict[str, Any]) -> str | None:
    if not isinstance(payload, dict):
        return "Feedback payload is not a JSON object."
    status = str(payload.get("status") or "").strip().lower()
    if status not in {"ok", "incomplete"}:
        return "Feedback status is invalid."
    overall = clean_feedback_text(payload.get("overall"))
    if len(overall) < 20:
        return "Feedback overall comment is too short."
    priority_actions = _limit_list(payload.get("priority_actions"), 6)
    next_steps = _limit_list(payload.get("next_steps"), 6)
    if len(priority_actions) < 2:
        return "Feedback priority_actions is too short."
    if len(next_steps) < 2:
        return "Feedback next_steps is too short."
    meaningful_lists = sum(
        1
        for key in ("format_fix", "content_fix", "writing_fix", "citation_fix")
        if _limit_list(payload.get(key), 6)
    )
    if meaningful_lists == 0:
        return "Feedback has no actionable sections."
    return None


def _render_feedback_markdown(entry: dict[str, Any], grade_data: dict[str, Any], payload: dict[str, Any]) -> str:
    summary = grade_data.get("summary", {})
    extracted = grade_data.get("extracted", {})
    lines = [
        f"# {clean_feedback_text(payload.get('title')) or ((entry.get('name') or '学生') + '论文修改建议')}",
        "",
        f"- 学号: {entry.get('sid') or '-'}",
        f"- 姓名: {entry.get('name') or '-'}",
        f"- 论文题目: {extracted.get('title') or entry.get('paper_title') or '-'}",
        f"- 当前结论: {summary.get('decision') or '-'}",
        f"- 当前总分: {summary.get('total_score') if summary.get('total_score') is not None else '-'} / 100",
        "",
        "## 总评",
        clean_feedback_text(payload.get("overall")) or "本次评语生成失败，请重新执行。",
        "",
        "## 先改这几件事",
    ]
    lines.extend(_render_numbered_lines(payload.get("priority_actions"), fallback="先把最影响通过的硬伤逐条改掉。"))
    lines.extend(["", "## 格式怎么改"])
    lines.extend(_render_numbered_lines(payload.get("format_fix"), fallback="对照学校模板逐项检查封面、摘要、目录、页眉页码和标题层级。"))
    lines.extend(["", "## 内容怎么改"])
    lines.extend(_render_numbered_lines(payload.get("content_fix"), fallback="按章节补充分析和论证，不要只列条目。"))
    lines.extend(["", "## 写作怎么改"])
    lines.extend(_render_numbered_lines(payload.get("writing_fix"), fallback="把空话和套话删掉，用具体事实、分析和结论替换。"))
    lines.extend(["", "## 引用怎么改"])
    lines.extend(_render_numbered_lines(payload.get("citation_fix"), fallback="逐条核对文内引用和参考文献是否一一对应、是否真实存在。"))
    strengths = _limit_list(payload.get("strengths"), 4)
    if strengths:
        lines.extend(["", "## 可以先保留的地方"])
        lines.extend(_render_numbered_lines(strengths))
    risk_notice = clean_feedback_text(payload.get("risk_notice"))
    if risk_notice:
        lines.extend(["", "## 当前风险提醒", risk_notice])
    lines.extend(["", "## 建议修改顺序"])
    lines.extend(_render_numbered_lines(payload.get("next_steps"), fallback="先改硬伤，再补正文，再统一检查引用和排版。"))
    lines.append("")
    return "\n".join(lines)


def _render_numbered_lines(values: Any, fallback: str | None = None) -> list[str]:
    items = _limit_list(values, 8)
    if not items and fallback:
        items = [fallback]
    return [f"{index}. {item}" for index, item in enumerate(items, 1)]


def _build_missing_grade_feedback(entry: dict[str, Any]) -> str:
    return "\n".join(
        [
            f"# {(entry.get('name') or '学生')}论文批改状态说明",
            "",
            f"- 学号: {entry.get('sid') or '-'}",
            f"- 姓名: {entry.get('name') or '-'}",
            f"- 论文题目: {entry.get('paper_title') or '-'}",
            "- 当前结论: 本次没有生成有效评分结果",
            "",
            "## 当前情况",
            "本次批改流程没有完整产出有效评分数据，所以这份评语不能当正式批改结果使用。",
            "",
            "## 需要立即处理的事",
            "1. 重新执行完整批改流程，不要只刷新 tracking 日志。",
            "2. 检查评分 JSON、评分报告和模型原始响应文件是否都已生成。",
            "3. 确认异常环节后再重新生成学生评语。",
            "",
        ]
    )


def _build_generation_failed_feedback(entry: dict[str, Any], grade_data: dict[str, Any], error: str) -> str:
    summary = grade_data.get("summary", {})
    extracted = grade_data.get("extracted", {})
    return "\n".join(
        [
            f"# {(entry.get('name') or '学生')}论文修改建议",
            "",
            f"- 学号: {entry.get('sid') or '-'}",
            f"- 姓名: {entry.get('name') or '-'}",
            f"- 论文题目: {extracted.get('title') or entry.get('paper_title') or '-'}",
            f"- 当前结论: {summary.get('decision') or '-'}",
            f"- 当前总分: {summary.get('total_score') if summary.get('total_score') is not None else '-'} / 100",
            "",
            "## 当前情况",
            "评分结果已经生成，但本次给学生的最终评语在大模型改写环节失败了，所以这里只能先给出系统失败提示，不能当正式评语使用。",
            "",
            "## 异常信息",
            clean_feedback_text(error) or "最终评语生成失败。",
            "",
            "## 需要立即处理的事",
            "1. 重新执行学生评语生成，不需要重跑前面的评分步骤。",
            "2. 检查 Kimi 调用是否成功返回完整 JSON。",
            "3. 检查重试后是否仍然失败，并记录失败环节。",
            "",
        ]
    )


def _extract_json_object(raw_text: str) -> str:
    raw_text = str(raw_text or "").strip()
    if raw_text.startswith("{") and raw_text.endswith("}"):
        return raw_text
    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise RuntimeError("Feedback response does not contain a JSON object.")
    return raw_text[start : end + 1]


def _feedback_artifact_dir(entry: dict[str, Any]) -> Path:
    run_root = entry.get("run_root")
    if run_root:
        path = Path(str(run_root)).resolve() / "feedback"
    else:
        path = _repo_root() / "runtime" / "tracking" / "feedback_artifacts"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _write_feedback_metadata(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _sanitize_raw_response(payload: Any) -> Any:
    if isinstance(payload, dict):
        sanitized: dict[str, Any] = {}
        for key, value in payload.items():
            if str(key).lower() == "reasoning_content":
                continue
            sanitized[str(key)] = _sanitize_raw_response(value)
        return sanitized
    if isinstance(payload, list):
        return [_sanitize_raw_response(item) for item in payload]
    return payload


def _format_provider_error(response: requests.Response) -> str:
    detail = f"HTTP {response.status_code}"
    try:
        payload = response.json()
    except ValueError:
        payload = None
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            message = str(error.get("message") or "").strip()
            if message:
                detail = f"{detail}: {message}"
    return detail


def _repo_root() -> Path:
    current = Path(__file__).resolve().parent
    for candidate in (current, *current.parents):
        if any((candidate / marker).exists() for marker in REPO_ROOT_MARKERS):
            return candidate
    raise RuntimeError(f"Unable to locate repository root from {__file__}")


def _credential_store_root() -> Path:
    override = os.getenv("PAPER_PIPELINE_CREDENTIAL_STORE_DIR")
    if override:
        return Path(override).expanduser().resolve()
    return _repo_root() / "runtime" / "secrets" / "credential_store"


def _credential_entry_path(service: str) -> Path:
    return _credential_store_root() / f"{service}.json"


def _load_credential_entry(service: str) -> dict[str, Any]:
    path = _credential_entry_path(service)
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    decrypted_fields: dict[str, str] = {}
    for key, value in (payload.get("fields") or {}).items():
        cipher_bytes = bytes.fromhex(value)
        plain_bytes = win32crypt.CryptUnprotectData(cipher_bytes, None, None, None, 0)[1]
        decrypted_fields[key] = plain_bytes.decode("utf-16-le").rstrip("\x00")
    return {
        "service": payload.get("service") or service,
        "saved_at": payload.get("saved_at"),
        "metadata": payload.get("metadata") or {},
        "fields": decrypted_fields,
        "path": str(path.resolve()),
    }
