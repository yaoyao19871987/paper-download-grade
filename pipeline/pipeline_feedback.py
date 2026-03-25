from __future__ import annotations

import re
from typing import Any

from pipeline_utils import dedupe_keep_order, format_score


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
    cleaned = re.sub(r"\s+", " ", str(text or "")).strip(" ;；。")
    if not cleaned:
        return ""
    return cleaned


def number_lines(values: list[str]) -> list[str]:
    return [f"{index}. {value}" for index, value in enumerate(values, 1)]


def feedback_severity(item: dict[str, Any], category: str) -> str:
    ratio = score_ratio(item)
    labels = {
        "format": [
            (0.12, "格式问题非常严重"),
            (0.25, "格式问题严重"),
            (0.45, "格式问题较大"),
            (0.65, "格式问题明显"),
            (0.80, "格式仍需修整"),
            (1.01, "格式基本合格"),
        ],
        "content": [
            (0.12, "内容基本未达到要求"),
            (0.25, "内容问题很大"),
            (0.45, "内容存在较大问题"),
            (0.65, "内容存在明显问题"),
            (0.80, "内容基本过关但还需补强"),
            (1.01, "内容整体尚可"),
        ],
    }
    for threshold, label in labels[category]:
        if ratio <= threshold:
            return label
    return labels[category][-1][1]


def build_teacher_action_text(suggestions: list[str], fallback: str) -> str:
    cleaned = [clean_feedback_text(item) for item in suggestions]
    cleaned = [item for item in cleaned if item]
    if not cleaned:
        return fallback
    return "你需要修改的地方: " + "; ".join(cleaned) + "。"


def sorted_feedback_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ranked = [item for item in items if item.get("suggestions")]
    ranked.sort(key=lambda item: (score_ratio(item), str(item.get("name") or "")))
    return ranked


def overall_teacher_comment(decision: str) -> str:
    mapping = {
        "打回重写": "这篇稿件当前问题较多，还不能按毕业论文成稿来提交，先把硬伤修好。",
        "引用退修": "这篇稿件当前最大的问题在引用与学术规范，需要先把文献和正文引用关系做实。",
        "引用待核": "结构已经有了，但引用证据还不够扎实，先补引文和来源。",
        "通过": "这篇稿件可以继续往下走，但仍需按下面意见继续修改。",
    }
    return mapping.get(decision or "", "这篇稿件还需要继续修改，按下面的顺序逐项处理。")


def build_student_feedback(entry: dict[str, Any], grade_data: dict[str, Any]) -> str:
    summary = grade_data.get("summary", {})
    extracted = grade_data.get("extracted", {})
    format_items = grade_data.get("format_items", [])
    content_items = grade_data.get("content_items", [])
    reference_audit = grade_data.get("reference_audit", {})
    visual_review = grade_data.get("visual_review", {})
    text_review = grade_data.get("text_review", {})

    text_major = [
        clean_feedback_text(item)
        for item in list(text_review.get("major_problems", []) or [])
        if clean_feedback_text(item)
    ]
    text_actions_raw = [
        clean_feedback_text(item)
        for item in list(text_review.get("revision_actions", []) or [])
        if clean_feedback_text(item)
    ]

    urgent_items = dedupe_keep_order(
        list(summary.get("gate_reasons", []) or [])
        + list(grade_data.get("gate", {}).get("reasons", []) or [])
        + list(grade_data.get("reference_gate", {}).get("reasons", []) or [])
        + list(visual_review.get("major_issues", []) or [])
        + text_major[:3]
    )

    format_actions: list[str] = []
    for item in sorted_feedback_items(format_items):
        format_actions.append(
            f"{item.get('name', '格式项')}: {feedback_severity(item, 'format')}。"
            f"{build_teacher_action_text(item.get('suggestions') or [], '按学校格式要求重新核对这一项。')}"
        )

    content_actions: list[str] = []
    for item in sorted_feedback_items(content_items):
        content_actions.append(
            f"{item.get('name', '内容项')}: {feedback_severity(item, 'content')}。"
            f"{build_teacher_action_text(item.get('suggestions') or [], '把这一部分重新补充完整。')}"
        )

    strengths: list[str] = []
    for item in format_items + content_items:
        if item.get("status") in {"优秀", "良好"} or score_ratio(item) >= 0.75:
            strengths.append(f"{item.get('name', '项目')}这一项基础还可以，先保持住。")

    reference_actions = dedupe_keep_order(
        [clean_feedback_text(item) for item in list(reference_audit.get("notes", []) or [])]
        + [
            build_teacher_action_text(
                item.get("suggestions") or [],
                f"重点检查 {item.get('name', '这一项')}。",
            )
            for item in content_items
            if item.get("key") in {"references_support", "academic_integrity"} and item.get("suggestions")
        ]
    )

    text_actions = dedupe_keep_order(
        [f"写作主要问题: {item}" for item in text_major]
        + [f"具体修改动作: {item}" for item in text_actions_raw]
    )

    lines = [
        f"# {(entry.get('name') or '学生')}论文修改建议",
        "",
        f"- 学号: {entry.get('sid') or '-'}",
        f"- 姓名: {entry.get('name') or '-'}",
        f"- 论文题目: {extracted.get('title') or entry.get('paper_title') or '-'}",
        f"- 当前判定: {summary.get('decision') or '-'}",
        f"- 当前总分: {format_score(summary.get('total_score'))} / 100",
        f"- 视觉模式: {visual_review.get('mode') or 'off'}",
        f"- 文本模式: {text_review.get('mode') or 'off'}",
        f"- 总评: {overall_teacher_comment(summary.get('decision') or '')}",
        "",
        "## 优先修改",
    ]

    lines.extend(number_lines([clean_feedback_text(item) for item in urgent_items if clean_feedback_text(item)]) or ["1. 当前没有额外门槛问题，按下面顺序继续修改。"])
    lines.extend(["", "## 格式问题"])
    lines.extend(number_lines(format_actions) or ["1. 当前没有明显的格式硬伤。"])
    lines.extend(["", "## 内容问题"])
    lines.extend(number_lines(content_actions) or ["1. 当前没有特别突出的内容短板。"])
    lines.extend(["", "## 文本与表达"])
    lines.extend(number_lines(text_actions) or ["1. 本次文本专家没有补充额外写作动作。"])
    lines.extend(["", "## 引用与学术规范"])
    lines.extend(number_lines(reference_actions) or ["1. 当前没有额外引用风险提示，但仍需逐条核对文献真实性。"])
    lines.extend(["", "## 做得比较好的地方"])
    lines.extend(number_lines(dedupe_keep_order(strengths)) or ["1. 先把前面的硬伤处理掉，再看亮点打磨。"])
    lines.extend(
        [
            "",
            "## 建议修改顺序",
            "1. 先改门槛问题，再改格式硬伤。",
            "2. 格式改完后，再处理摘要、目录、正文和参考文献。",
            "3. 全部改完以后重新导出 Word，再走同一套复核流程。",
            "",
        ]
    )
    return "\n".join(lines)
