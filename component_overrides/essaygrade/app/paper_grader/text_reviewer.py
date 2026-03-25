from __future__ import annotations

from dataclasses import asdict, dataclass
import json
import os
from pathlib import Path
import re
import time
from typing import Any

import requests

from .credential_store import credential_entry_exists, load_credential_entry
from .rubric import (
    KEYWORD_RANGE,
    MIN_BODY_CHARS,
    OFFICIAL_WRITING_REQUIREMENTS,
    STRICT_ABSTRACT_RANGE,
    TITLE_MAX_CHARS,
)


SILICONFLOW_DEFAULT_BASE_URL = "https://api.siliconflow.cn/v1"
MOONSHOT_DEFAULT_BASE_URL = "https://api.moonshot.cn/v1"
TEXT_REVIEW_TIMEOUT_SECONDS = 180
DEFAULT_TEXT_PRIMARY_MODEL = "deepseek-ai/DeepSeek-V3.2"
DEFAULT_TEXT_SECONDARY_MODEL = "kimi-for-coding"
DEFAULT_KIMI_CODING_USER_AGENT = "KimiCLI/1.0"
BODY_REQUIREMENT_CONTEXT_HINTS = ("正文", "全文", "主体", "论文", "字数", "篇幅", "补充", "扩写", "充实")
BODY_REQUIREMENT_EXCLUDE_HINTS = ("摘要", "标题", "题目", "关键词", "参考文献", "附录")
TEXT_REVIEW_RETRYABLE_STATUS_CODES = {408, 409, 425, 429, 500, 502, 503, 504}


@dataclass
class TextModelReview:
    provider: str
    model: str
    score: float | None
    confidence: float | None
    dimensions: dict[str, float]
    strengths: list[str]
    major_problems: list[str]
    revision_actions: list[str]
    notes: list[str]
    raw_response_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class TextReviewResult:
    mode: str
    fused_score: float | None
    confidence: float | None
    adjustment: float
    primary: TextModelReview | None
    secondary: TextModelReview | None
    strengths: list[str]
    major_problems: list[str]
    revision_actions: list[str]
    notes: list[str]

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        return payload


def has_siliconflow_text_credentials() -> bool:
    if os.getenv("SILICONFLOW_API_KEY"):
        return True
    return credential_entry_exists("siliconflow")


def has_moonshot_text_credentials() -> bool:
    if os.getenv("MOONSHOT_API_KEY"):
        return True
    return credential_entry_exists("moonshot_kimi")


def review_writing_quality(
    text_payload: dict[str, Any],
    stage: str,
    mode: str = "off",
    primary_model: str | None = None,
    secondary_model: str | None = None,
    output_dir: str | None = None,
) -> TextReviewResult:
    normalized_mode = (mode or "off").lower()
    if normalized_mode == "off":
        return TextReviewResult(
            mode="off",
            fused_score=None,
            confidence=None,
            adjustment=0.0,
            primary=None,
            secondary=None,
            strengths=[],
            major_problems=[],
            revision_actions=[],
            notes=["文本大模型写作质量审稿已关闭。"],
        )

    review_dir = _ensure_output_dir(output_dir)

    primary_result: TextModelReview | None = None
    secondary_result: TextModelReview | None = None
    notes: list[str] = []

    if normalized_mode in {"siliconflow", "auto", "expert"}:
        try:
            primary_result = _review_with_siliconflow(
                text_payload=text_payload,
                stage=stage,
                model=primary_model or DEFAULT_TEXT_PRIMARY_MODEL,
                output_dir=review_dir,
            )
        except Exception as exc:
            notes.append(f"SiliconFlow 文本专家未成功执行: {type(exc).__name__}: {exc}")

    if normalized_mode in {"moonshot", "auto", "expert"}:
        try:
            secondary_result = _review_with_moonshot(
                text_payload=text_payload,
                stage=stage,
                model=secondary_model or DEFAULT_TEXT_SECONDARY_MODEL,
                output_dir=review_dir,
            )
        except Exception as exc:
            notes.append(f"Moonshot/Kimi 文本专家未成功执行: {type(exc).__name__}: {exc}")

    if normalized_mode == "siliconflow":
        secondary_result = None
    if normalized_mode == "moonshot":
        primary_result = None

    if not primary_result and not secondary_result:
        return TextReviewResult(
            mode="unavailable",
            fused_score=None,
            confidence=None,
            adjustment=0.0,
            primary=None,
            secondary=None,
            strengths=[],
            major_problems=[],
            revision_actions=[],
            notes=notes or ["文本大模型写作质量审稿不可用。"],
        )

    fused_score, fused_confidence = _fuse_scores(primary_result, secondary_result)
    strengths = _merge_unique(
        (primary_result.strengths if primary_result else []),
        (secondary_result.strengths if secondary_result else []),
    )
    major_problems = _merge_unique(
        (primary_result.major_problems if primary_result else []),
        (secondary_result.major_problems if secondary_result else []),
    )
    revision_actions = _merge_unique(
        (primary_result.revision_actions if primary_result else []),
        (secondary_result.revision_actions if secondary_result else []),
    )
    fused_notes = _merge_unique(
        notes,
        (primary_result.notes if primary_result else []),
        (secondary_result.notes if secondary_result else []),
    )

    return TextReviewResult(
        mode="expert" if primary_result and secondary_result else "single",
        fused_score=fused_score,
        confidence=fused_confidence,
        adjustment=_score_to_adjustment(fused_score),
        primary=primary_result,
        secondary=secondary_result,
        strengths=strengths[:6],
        major_problems=major_problems[:8],
        revision_actions=revision_actions[:8],
        notes=fused_notes[:8],
    )


def _review_with_siliconflow(
    text_payload: dict[str, Any],
    stage: str,
    model: str,
    output_dir: Path,
) -> TextModelReview:
    config = _resolve_siliconflow_config(model=model)
    payload = _build_text_payload(
        model=config["model"],
        stage=stage,
        text_payload=text_payload,
    )
    url = _build_chat_completions_url(config["api_base_url"])
    response_json = _post_json_request(url, payload, config["api_key"])

    raw_path = output_dir / "text_primary_siliconflow.json"
    raw_path.write_text(json.dumps(_sanitize_raw_response(response_json), ensure_ascii=False, indent=2), encoding="utf-8")
    parsed = _parse_chat_json_response(response_json)
    parsed = _normalize_with_official_requirements(parsed)
    parsed["notes"].insert(0, f"主文本专家: SiliconFlow({config['model']})。")
    return TextModelReview(
        provider="siliconflow",
        model=config["model"],
        score=parsed["overall_score"],
        confidence=parsed["confidence"],
        dimensions=parsed["dimensions"],
        strengths=parsed["strengths"],
        major_problems=parsed["major_problems"],
        revision_actions=parsed["revision_actions"],
        notes=parsed["notes"],
        raw_response_path=str(raw_path),
    )


def _review_with_moonshot(
    text_payload: dict[str, Any],
    stage: str,
    model: str,
    output_dir: Path,
) -> TextModelReview:
    config = _resolve_moonshot_config(model=model)
    payload = _build_text_payload(
        model=config["model"],
        stage=stage,
        text_payload=text_payload,
    )
    url = _build_chat_completions_url(config["api_base_url"])
    response_json = _post_json_request(url, payload, config["api_key"])

    raw_path = output_dir / "text_secondary_moonshot.json"
    raw_path.write_text(json.dumps(_sanitize_raw_response(response_json), ensure_ascii=False, indent=2), encoding="utf-8")
    parsed = _parse_chat_json_response(response_json)
    parsed = _normalize_with_official_requirements(parsed)
    parsed["notes"].insert(0, f"次文本专家: Moonshot/Kimi({config['model']})。")
    return TextModelReview(
        provider="moonshot",
        model=config["model"],
        score=parsed["overall_score"],
        confidence=parsed["confidence"],
        dimensions=parsed["dimensions"],
        strengths=parsed["strengths"],
        major_problems=parsed["major_problems"],
        revision_actions=parsed["revision_actions"],
        notes=parsed["notes"],
        raw_response_path=str(raw_path),
    )


def _resolve_siliconflow_config(model: str) -> dict[str, str]:
    api_key = os.getenv("SILICONFLOW_API_KEY")
    api_base_url = os.getenv("SILICONFLOW_BASE_URL") or SILICONFLOW_DEFAULT_BASE_URL
    if credential_entry_exists("siliconflow"):
        entry = load_credential_entry("siliconflow")
        metadata = entry.get("metadata") or {}
        if not api_key:
            api_key = entry.get("fields", {}).get("api_key")
        if metadata.get("api_base_url") and api_base_url == SILICONFLOW_DEFAULT_BASE_URL:
            api_base_url = str(metadata.get("api_base_url"))
        if metadata.get("default_model") and not model:
            model = str(metadata.get("default_model"))
    if not api_key:
        raise RuntimeError("SiliconFlow API Key 未配置。")
    return {
        "api_key": api_key,
        "api_base_url": api_base_url,
        "model": model or DEFAULT_TEXT_PRIMARY_MODEL,
    }


def _resolve_moonshot_config(model: str) -> dict[str, str]:
    api_key = os.getenv("MOONSHOT_API_KEY")
    api_base_url = os.getenv("MOONSHOT_BASE_URL") or MOONSHOT_DEFAULT_BASE_URL
    if credential_entry_exists("moonshot_kimi"):
        entry = load_credential_entry("moonshot_kimi")
        metadata = entry.get("metadata") or {}
        if not api_key:
            api_key = entry.get("fields", {}).get("api_key")
        if metadata.get("api_base_url") and api_base_url == MOONSHOT_DEFAULT_BASE_URL:
            api_base_url = str(metadata.get("api_base_url"))
        if metadata.get("default_model") and not model:
            model = str(metadata.get("default_model"))
    if not api_key:
        raise RuntimeError("Moonshot/Kimi API Key 未配置。")
    return {
        "api_key": api_key,
        "api_base_url": api_base_url,
        "model": model or DEFAULT_TEXT_SECONDARY_MODEL,
    }


def _build_text_payload(model: str, stage: str, text_payload: dict[str, Any]) -> dict[str, Any]:
    keyword_min, keyword_max = KEYWORD_RANGE
    abstract_min, abstract_max = STRICT_ABSTRACT_RANGE
    strategy = (
        "你是严苛的毕业论文文本质量审稿员，只评写作质量，不评版式。"
        "评分必须从严，不得宽松给分。"
        "评分维度与权重：选题贴合10、结构逻辑20、论证深度25、证据与引用15、语言表达15、学术规范15。"
        f"学校硬性要求：正文字数不少于{MIN_BODY_CHARS}字；标题不超过{TITLE_MAX_CHARS}字；关键词{keyword_min}-{keyword_max}个；摘要约{abstract_min}-{abstract_max}字。"
        "整改建议必须服从学校硬性要求，不得自创更高最低字数（例如不得要求8000字）。"
        "你必须遵守以下硬规则："
        "1) 选题贴合=标题与正文严格对应，不评题目是否新潮。若标题关键词未在正文持续展开，topic_alignment不得高于40；若正文缺失，topic_alignment不得高于20。"
        "2) 结构逻辑看“问题-方法-分析-结论”链条。仅有目录式罗列、套模板、章节空壳时，structure_logic不得高于45。"
        "3) 语言表达必须检查车轱辘话、空泛口号、重复句和无信息句。出现明显套话堆砌时，language_clarity不得高于45。"
        "4) 若正文字符<1200或章节<2，overall_score不得高于45；若正文字符=0，overall_score不得高于20。"
        "请严格返回 JSON，不要输出 markdown。"
    )
    user_prompt = {
        "stage": stage,
        "task": "请基于给定论文文本片段做严格评分（从严打分，不做鼓励式宽松）。",
        "rubric": {
            "topic_alignment": 10,
            "structure_logic": 20,
            "argument_depth": 25,
            "evidence_citation": 15,
            "language_clarity": 15,
            "academic_integrity": 15,
        },
        "dimension_definitions": {
            "topic_alignment": "只看标题与正文是否严格对应；不看题目新旧。标题关键词需在正文反复、实质展开。",
            "structure_logic": "是否形成问题提出-方法路径-分析论证-结论回扣的完整链条，拒绝目录式空壳。",
            "argument_depth": "是否有实质分析、因果解释、对比论证，而不是结论口号。",
            "evidence_citation": "关键结论是否有数据/案例/文献支撑，且文内文献映射清楚。",
            "language_clarity": "是否存在套话、车轱辘话、重复句、冗长空话，是否简洁准确。",
            "academic_integrity": "引用规范、文献真实性、学术表达是否合规。",
        },
        "strict_penalties": {
            "topic_mismatch": "标题与正文不对应时，topic_alignment <= 40",
            "body_missing_topic": "正文缺失时，topic_alignment <= 20",
            "outline_only": "结构空壳或目录堆砌时，structure_logic <= 45",
            "boilerplate_language": "明显套话/重复话时，language_clarity <= 45",
            "short_body_global_cap": "body_chars < 1200 或 chapter_count < 2 时，overall_score <= 45",
            "empty_body_global_cap": "body_chars = 0 时，overall_score <= 20",
        },
        "official_requirements": {
            "min_body_chars": MIN_BODY_CHARS,
            "title_max_chars": TITLE_MAX_CHARS,
            "keyword_range": [keyword_min, keyword_max],
            "strict_abstract_range": [abstract_min, abstract_max],
            "use_school_rules_as_final_authority": True,
        },
        "official_writing_requirements": OFFICIAL_WRITING_REQUIREMENTS,
        "scoring_principles": [
            "先校验学校硬约束，再判断写作质量",
            "题文不一致时，选题贴合度必须严格降分",
            "结构空壳、套话堆砌、论证薄弱必须重扣",
            "整改建议必须引用学校要求，不得自创更高硬指标",
        ],
        "paper": text_payload,
        "output_requirements": {
            "overall_score": "0-100",
            "confidence": "0-1",
            "dimensions": "每个维度0-100",
            "strengths": "2-6条",
            "major_problems": "3-10条",
            "revision_actions": "3-10条，可执行",
            "notes": "可选，说明触发了哪些硬规则与扣分理由",
        },
    }
    return {
        "model": model,
        "messages": [
            {"role": "system", "content": strategy},
            {"role": "user", "content": json.dumps(user_prompt, ensure_ascii=False)},
        ],
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }


def _post_json_request(url: str, payload: dict[str, Any], api_key: str) -> dict[str, Any]:
    headers = _build_request_headers(url, api_key)
    max_retries = max(1, int(os.getenv("TEXT_REVIEW_MAX_RETRIES", "4")))
    backoff_seconds = max(1.0, float(os.getenv("TEXT_REVIEW_RETRY_BACKOFF_SECONDS", "8")))
    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=TEXT_REVIEW_TIMEOUT_SECONDS)
        except (requests.Timeout, requests.ConnectionError) as exc:
            last_error = exc
            if attempt < max_retries - 1:
                time.sleep(backoff_seconds * (attempt + 1))
                continue
            raise RuntimeError(f"文本评分请求超时/连接失败（重试{max_retries}次后仍失败）: {exc}") from exc

        if response.status_code >= 400:
            retryable = response.status_code in TEXT_REVIEW_RETRYABLE_STATUS_CODES
            if retryable and attempt < max_retries - 1:
                time.sleep(backoff_seconds * (attempt + 1))
                continue
            raise RuntimeError(_format_provider_error(response, url))
        return response.json()

    raise RuntimeError(f"文本评分请求失败（重试{max_retries}次后仍失败）: {last_error}")


def _build_request_headers(url: str, api_key: str) -> dict[str, str]:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    normalized_url = (url or "").lower()
    if "api.kimi.com/coding" in normalized_url:
        headers["User-Agent"] = os.getenv("KIMI_CODING_USER_AGENT") or DEFAULT_KIMI_CODING_USER_AGENT
    return headers


def _build_chat_completions_url(api_base_url: str) -> str:
    base = (api_base_url or "").rstrip("/")
    if not base:
        raise RuntimeError("API base URL 未配置。")
    if base.endswith("/v1"):
        return base + "/chat/completions"
    return base + "/v1/chat/completions"


def _format_provider_error(response: requests.Response, url: str) -> str:
    detail = f"HTTP {response.status_code}"
    error_type = ""
    try:
        payload = response.json()
    except ValueError:
        payload = None

    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            error_type = str(error.get("type") or "").strip()
            message = str(error.get("message") or "").strip()
            if message:
                detail = f"{detail}: {message}"

    normalized_url = url.lower()
    if "api.kimi.com/coding" in normalized_url and error_type == "access_terminated_error":
        return "Kimi Code 拒绝了当前请求；请确认使用 OpenAI 兼容 chat/completions，并携带 KimiCLI 兼容 User-Agent。"

    return detail


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


def _parse_chat_json_response(response_json: dict[str, Any]) -> dict[str, Any]:
    choices = response_json.get("choices") or []
    if not choices:
        raise RuntimeError("文本评审接口未返回 choices。")
    content = choices[0].get("message", {}).get("content")
    if isinstance(content, list):
        content = "\n".join(
            str(item.get("text") or "")
            for item in content
            if isinstance(item, dict)
        ).strip()
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("文本评审接口未返回可解析文本。")

    parsed = json.loads(_extract_json_object(content))
    dimensions = parsed.get("dimensions") if isinstance(parsed.get("dimensions"), dict) else {}
    dimension_values = {
        "topic_alignment": _clamp_score(dimensions.get("topic_alignment")),
        "structure_logic": _clamp_score(dimensions.get("structure_logic")),
        "argument_depth": _clamp_score(dimensions.get("argument_depth")),
        "evidence_citation": _clamp_score(dimensions.get("evidence_citation")),
        "language_clarity": _clamp_score(dimensions.get("language_clarity")),
        "academic_integrity": _clamp_score(dimensions.get("academic_integrity")),
    }

    weighted = (
        0.10 * dimension_values["topic_alignment"]
        + 0.20 * dimension_values["structure_logic"]
        + 0.25 * dimension_values["argument_depth"]
        + 0.15 * dimension_values["evidence_citation"]
        + 0.15 * dimension_values["language_clarity"]
        + 0.15 * dimension_values["academic_integrity"]
    )
    overall_score = _clamp_score(parsed.get("overall_score"), fallback=weighted)
    confidence = _clamp_confidence(parsed.get("confidence"))

    return {
        "overall_score": overall_score,
        "confidence": confidence,
        "dimensions": dimension_values,
        "strengths": _coerce_string_list(parsed.get("strengths")),
        "major_problems": _coerce_string_list(parsed.get("major_problems")),
        "revision_actions": _coerce_string_list(parsed.get("revision_actions")),
        "notes": _coerce_string_list(parsed.get("notes")),
    }


def _normalize_with_official_requirements(parsed: dict[str, Any]) -> dict[str, Any]:
    normalized = dict(parsed or {})
    for field in ("major_problems", "revision_actions", "notes"):
        values = _coerce_string_list(normalized.get(field))
        cleaned: list[str] = []
        for value in values:
            fixed = _normalize_requirement_sentence(value)
            if fixed and fixed not in cleaned:
                cleaned.append(fixed)
        normalized[field] = cleaned
    return normalized


def _normalize_requirement_sentence(text: str) -> str:
    sentence = re.sub(r"\s+", " ", str(text or "")).strip()
    if not sentence:
        return ""

    sentence = _replace_inflated_body_wording(sentence)
    sentence = re.sub(r"（按学校要求）{2,}", "（按学校要求）", sentence)
    return sentence


def _replace_inflated_body_wording(sentence: str) -> str:
    def _replace(match: re.Match[str]) -> str:
        number_text = match.group("num")
        try:
            number_value = int(number_text)
        except (TypeError, ValueError):
            return match.group(0)

        if number_value <= MIN_BODY_CHARS:
            return match.group(0)

        start = max(0, match.start() - 18)
        end = min(len(sentence), match.end() + 18)
        context = sentence[start:end]

        if any(token in context for token in BODY_REQUIREMENT_EXCLUDE_HINTS):
            return match.group(0)
        if not any(token in context for token in BODY_REQUIREMENT_CONTEXT_HINTS):
            return match.group(0)

        return f"{MIN_BODY_CHARS}字（按学校要求）"

    return re.sub(r"(?P<num>\d{4,5})\s*字", _replace, sentence)


def _clamp_score(value: Any, fallback: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return round(max(0.0, min(100.0, float(fallback))), 2)
        return round(max(0.0, min(100.0, float(value))), 2)
    except (TypeError, ValueError):
        return round(max(0.0, min(100.0, float(fallback))), 2)


def _clamp_confidence(value: Any) -> float:
    try:
        if value is None or value == "":
            return 0.65
        return round(max(0.0, min(1.0, float(value))), 3)
    except (TypeError, ValueError):
        return 0.65


def _score_to_adjustment(score: float | None) -> float:
    if score is None:
        return 0.0
    if score >= 88:
        return 3.0
    if score >= 78:
        return 1.5
    if score >= 68:
        return 0.0
    if score >= 58:
        return -1.5
    if score >= 48:
        return -3.0
    return -5.0


def _fuse_scores(primary: TextModelReview | None, secondary: TextModelReview | None) -> tuple[float | None, float | None]:
    if primary and secondary and primary.score is not None and secondary.score is not None:
        p_conf = max(0.3, min(0.95, primary.confidence or 0.7))
        s_conf = max(0.3, min(0.95, secondary.confidence or 0.6))
        numerator = primary.score * p_conf * 0.75 + secondary.score * s_conf * 0.25
        denominator = p_conf * 0.75 + s_conf * 0.25
        fused = round(numerator / max(denominator, 1e-6), 2)
        confidence = round(min(0.95, p_conf * 0.8 + s_conf * 0.2), 3)
        return fused, confidence
    if primary and primary.score is not None:
        return round(primary.score, 2), round(primary.confidence or 0.65, 3)
    if secondary and secondary.score is not None:
        return round(secondary.score, 2), round(secondary.confidence or 0.6, 3)
    return None, None


def _extract_json_object(raw_text: str) -> str:
    raw_text = raw_text.strip()
    if raw_text.startswith("{") and raw_text.endswith("}"):
        return raw_text
    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise RuntimeError("文本评审响应中没有找到 JSON 对象。")
    return raw_text[start : end + 1]


def _coerce_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        text = value.strip()
        return [text] if text else []
    if isinstance(value, list):
        result: list[str] = []
        for item in value:
            text = str(item).strip()
            if text:
                result.append(text)
        return result
    if isinstance(value, dict):
        result: list[str] = []
        for key, item in value.items():
            text = f"{key}: {item}".strip()
            if text:
                result.append(text)
        return result
    text = str(value).strip()
    return [text] if text else []


def _merge_unique(*collections: list[str]) -> list[str]:
    merged: list[str] = []
    for collection in collections:
        for item in collection:
            text = str(item).strip()
            if text and text not in merged:
                merged.append(text)
    return merged


def _ensure_output_dir(output_dir: str | None) -> Path:
    if output_dir:
        path = Path(output_dir).expanduser().resolve()
    else:
        path = Path.cwd() / ".text_review_artifacts"
    path.mkdir(parents=True, exist_ok=True)
    return path
