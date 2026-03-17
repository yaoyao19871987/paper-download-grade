from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re
from typing import Iterable

from .rubric import (
    CONTENT_WEIGHTS,
    COVER_REQUIRED_LABELS,
    EXPECTED_PAGE_SETUP_CM,
    FORMAT_GATE_RULES,
    FORMAT_WEIGHTS,
    GRADE_BANDS,
    HEADER_VARIANTS,
    KEYWORD_RANGE,
    MANDATORY_SECTIONS,
    MIN_BODY_CHARS,
    MIN_REFERENCE_COUNT,
    PAGE_SETUP_TOLERANCE_CM,
    RECOMMENDED_SECTIONS,
    RUBRIC_NAME,
    SOFT_ABSTRACT_RANGE,
    STRICT_ABSTRACT_RANGE,
    TITLE_MAX_CHARS,
)
from .reference_verifier import ReferenceAuditResult, audit_references
from .visual_reviewer import (
    DEFAULT_MOONSHOT_MODEL,
    DEFAULT_SILICONFLOW_SECONDARY_VISUAL_MODEL,
    DEFAULT_SILICONFLOW_VISUAL_MODEL,
    DEFAULT_VISUAL_MODEL,
    VisualReviewResult,
    export_document_to_pdf,
    has_moonshot_visual_credentials,
    has_siliconflow_visual_credentials,
    review_document_with_moonshot,
    review_document_with_openai,
    review_document_with_siliconflow,
    review_document_with_siliconflow_ensemble,
    resolve_moonshot_visual_model_name,
    resolve_siliconflow_secondary_visual_model_name,
    resolve_siliconflow_visual_model_name,
)
from .word_inspector import DocumentSnapshot, ParagraphSnapshot, inspect_document


ACADEMIC_CUES = ["研究", "分析", "方法", "模型", "数据", "结果", "表明", "提出", "问题", "对策", "案例", "验证"]
INTRO_CUES = ["绪论", "引言", "导论", "前言", "研究背景", "研究意义"]
METHOD_CUES = ["研究方法", "方法", "采用", "通过", "文献研究法", "案例分析法", "问卷", "实验", "实证"]
RESULT_CUES = ["结果", "表明", "发现", "得出", "说明", "验证"]
PROBLEM_CUES = ["问题", "不足", "风险", "挑战", "瓶颈", "困境"]
SOLUTION_CUES = ["对策", "建议", "优化", "改进", "措施", "路径"]
CONCLUSION_CUES = ["结论", "总结", "展望", "结束语"]
OVERVIEW_CUES = ["概述", "定义", "特征", "现状", "需求", "基础"]
APPLICATION_CUES = ["应用", "实践", "案例", "实证", "实现", "设计"]
PLACEHOLDER_RE = re.compile(r"×××|XXX|____|Lorem", re.IGNORECASE)
GENERIC_TERMS = {
    "研究",
    "分析",
    "应用",
    "设计",
    "实现",
    "系统",
    "论文",
    "问题",
    "对策",
    "探讨",
    "浅析",
    "思考",
    "策略",
}
A4_PAGE_SIZE_CM = (21.0, 29.7)
BODY_SERIF_HINTS = ("宋", "仿宋", "fangsong", "song", "明", "ming", "times")
BODY_HEAVY_HINTS = ("黑", "hei", "雅黑", "yahei", "gothic", "arial", "calibri")
BODY_KAI_HINTS = ("楷", "kai")
HEADING_HEAVY_HINTS = ("黑", "hei", "雅黑", "yahei", "小标宋", "biaosong")


@dataclass
class ScoreItem:
    key: str
    name: str
    score: float
    max_score: float
    status: str
    evidence: list[str]
    suggestions: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class GateResult:
    decision: str
    rewrite_required: bool
    score_cap: float | None
    reasons: list[str]

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class ChapterBlock:
    heading: ParagraphSnapshot
    paragraphs: list[ParagraphSnapshot]
    text: str


def grade_document(
    path: str,
    reference_docs: list[str] | None = None,
    stage: str = "initial_draft",
    visual_mode: str = "auto",
    visual_model: str = DEFAULT_VISUAL_MODEL,
    visual_output_dir: str | None = None,
) -> dict:
    snapshot = inspect_document(path)
    reference_snapshots = _load_reference_snapshots(snapshot.path, reference_docs or [])
    chapter_blocks = _chapter_blocks(snapshot)
    reference_entries = _extract_reference_entries(snapshot)
    reference_audit = audit_references(reference_entries, _extract_body_text(snapshot), online=True)
    visual_review = _build_visual_review(
        snapshot,
        stage=stage,
        mode=visual_mode,
        model=visual_model,
        output_dir=visual_output_dir,
    )

    format_items = [
        _score_page_setup(snapshot),
        _score_cover(snapshot),
        _score_header_pagination(snapshot),
        _score_abstract_keywords_format(snapshot),
        _score_toc_and_sections(snapshot),
        _score_body_typography(snapshot),
    ]
    content_items = [
        _score_topic_relevance(snapshot, chapter_blocks),
        _score_structure_logic(snapshot, chapter_blocks),
        _score_abstract_quality(snapshot),
        _score_chapter_development(snapshot, chapter_blocks),
        _score_coherence_alignment(snapshot, chapter_blocks),
        _score_references_support(snapshot, reference_audit),
        _score_language_quality(snapshot),
        _score_academic_integrity(snapshot, reference_snapshots, reference_audit),
    ]

    base_format_score = round(sum(item.score for item in format_items), 2)
    format_visual_adjustment = _visual_score_adjustment(visual_review)
    format_score = round(max(0.0, base_format_score + format_visual_adjustment), 2)
    content_score = round(sum(item.score for item in content_items), 2)
    raw_total_score = round(format_score + content_score, 2)
    format_gate = _evaluate_format_gate(snapshot, format_items, stage)
    reference_gate = _evaluate_reference_gate(reference_audit)
    visual_gate = _evaluate_visual_gate(visual_review, stage)
    gate = _combine_gates(format_gate, reference_gate, visual_gate)
    total_score = round(min(raw_total_score, gate.score_cap), 2) if gate.score_cap is not None else raw_total_score

    if gate.decision == "通过":
        band = _grade_band(total_score)
    else:
        band = {
            "label": gate.decision,
            "description": "当前稿件存在必须先处理的门槛问题，暂不应按正常论文进入完整评阅。",
        }

    return {
        "document": str(Path(path).expanduser().resolve()),
        "rubric_name": RUBRIC_NAME,
        "summary": {
            "stage": stage,
            "decision": gate.decision,
            "rewrite_required": gate.rewrite_required,
            "gate_reasons": gate.reasons,
            "raw_total_score": raw_total_score,
            "total_score": total_score,
            "base_format_score": base_format_score,
            "format_visual_adjustment": format_visual_adjustment,
            "format_score": format_score,
            "content_score": content_score,
            "grade_band": band["label"],
            "band_description": band["description"],
        },
        "gate": gate.to_dict(),
        "format_gate": format_gate.to_dict(),
        "reference_gate": reference_gate.to_dict(),
        "visual_gate": visual_gate.to_dict(),
        "visual_review": visual_review.to_dict(),
        "extracted": _build_extracted_data(snapshot, chapter_blocks, reference_snapshots, reference_audit),
        "reference_audit": reference_audit.to_dict(),
        "format_items": [item.to_dict() for item in format_items],
        "content_items": [item.to_dict() for item in content_items],
    }


def render_text_report(result: dict) -> str:
    summary = result["summary"]
    extracted = result["extracted"]
    format_gate = result.get("format_gate") or {}
    reference_gate = result.get("reference_gate") or {}
    visual_gate = result.get("visual_gate") or {}
    visual_review = result.get("visual_review") or {}
    lines = [
        f"评分对象: {result['document']}",
        f"评分模型: {result['rubric_name']}",
        f"评阅阶段: {summary['stage']}",
        f"初稿裁定: {summary['decision']}",
        f"总分: {summary['total_score']} / 100",
        f"原始总分: {summary['raw_total_score']} / 100",
        f"格式分: {summary['format_score']} / {sum(FORMAT_WEIGHTS.values())}",
        f"格式基础分: {summary['base_format_score']}，视觉修正: {summary['format_visual_adjustment']}",
        f"内容分: {summary['content_score']} / {sum(CONTENT_WEIGHTS.values())}",
        f"档位: {summary['grade_band']} ({summary['band_description']})",
    ]

    if format_gate.get("reasons"):
        lines.append("格式门槛说明:")
        for reason in format_gate["reasons"]:
            lines.append(f"- {reason}")

    if reference_gate.get("reasons"):
        lines.append("引用门槛说明:")
        for reason in reference_gate["reasons"]:
            lines.append(f"- {reason}")

    if visual_gate.get("reasons"):
        lines.append("视觉门槛说明:")
        for reason in visual_gate["reasons"]:
            lines.append(f"- {reason}")

    if not format_gate.get("reasons") and not reference_gate.get("reasons") and not visual_gate.get("reasons") and summary["gate_reasons"]:
        lines.append("门槛说明:")
        for reason in summary["gate_reasons"]:
            lines.append(f"- {reason}")

    lines.extend(
        [
            "",
            "提取结果:",
            f"- 标题: {extracted['title'] or '未识别'}",
            f"- 专业: {extracted['major'] or '未识别'}",
            f"- 摘要字数: {extracted['abstract_chars']}",
            f"- 正文字数: {extracted['body_chars']}",
            f"- 关键词数: {extracted['keyword_count']}",
            f"- 章节数: {extracted['chapter_count']}",
            f"- 结论字数: {extracted['conclusion_chars']}",
            f"- 参考文献数: {extracted['reference_count']}",
            f"- 参考比对相似度: {extracted['reference_similarity']}",
            "",
            "格式评分:",
        ]
    )
    for item in result["format_items"]:
        lines.extend(_render_item(item))
    lines.append("")
    lines.append("内容评分:")
    for item in result["content_items"]:
        lines.extend(_render_item(item))

    if visual_review:
        lines.extend(
            [
                "",
                "视觉评判:",
                f"- 来源: {_visual_mode_label(visual_review.get('mode'))}",
                f"- 模型: {visual_review.get('model') or '未启用'}",
                f"- 视觉裁定: {_visual_verdict_label(visual_review.get('overall_verdict'))}",
                f"- 版面秩序分: {visual_review.get('visual_order_score')}",
                f"- 置信度: {visual_review.get('confidence')}",
                f"- 视觉工件: {visual_review.get('pdf_path') or '未导出'}",
            ]
        )
        if visual_review.get("notes"):
            for note in visual_review["notes"]:
                lines.append(f"- 说明: {note}")
        if visual_review.get("major_issues"):
            lines.append("视觉大问题:")
            for issue in visual_review["major_issues"]:
                lines.append(f"- {issue}")
        if visual_review.get("minor_issues"):
            lines.append("视觉小问题:")
            for issue in visual_review["minor_issues"]:
                lines.append(f"- {issue}")
        if visual_review.get("evidence"):
            lines.append("视觉依据:")
            for item in visual_review["evidence"]:
                lines.append(f"- {item}")
        if visual_review.get("page_observations"):
            lines.append("页面观察:")
            for item in visual_review["page_observations"]:
                lines.append(f"- {item}")

    audit = result.get("reference_audit")
    if audit:
        lines.extend(
            [
                "",
                "引用核验:",
                f"- 正文编号引用: {audit['citation_numbers'] or '未检测到'}",
                f"- 已验证: {audit['verified_count']}，疑似匹配: {audit['possible_count']}，未找到: {audit['not_found_count']}，检索异常: {audit['search_error_count']}",
            ]
        )
        for note in audit["notes"]:
            lines.append(f"- {note}")
        for check in audit["checks"]:
            title = check["title"] or "标题未解析"
            cited_text = "正文已引" if check["cited_in_body"] else "正文未引"
            lines.append(f"- [{check['index']}] {cited_text} / {check['status']} / {title}")
            details: list[str] = []
            if check["matched_title"]:
                details.append(f"命中标题: {check['matched_title']}")
            if check["matched_url"]:
                details.append(f"命中链接: {check['matched_url']}")
            if check["notes"]:
                details.append("说明: " + "；".join(check["notes"]))
            if details:
                lines.append("  " + " | ".join(details))

    return "\n".join(lines)


def dump_json(result: dict, path: str) -> None:
    Path(path).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


def _render_item(item: dict) -> list[str]:
    lines = [f"- {item['name']}: {item['score']} / {item['max_score']} ({item['status']})"]
    if item["evidence"]:
        lines.append(f"  依据: {'; '.join(item['evidence'])}")
    if item["suggestions"]:
        lines.append(f"  建议: {'; '.join(item['suggestions'])}")
    return lines


def _load_reference_snapshots(document_path: str, reference_docs: Iterable[str]) -> list[DocumentSnapshot]:
    source = str(Path(document_path).resolve())
    snapshots: list[DocumentSnapshot] = []
    for reference_doc in reference_docs:
        resolved = str(Path(reference_doc).expanduser().resolve())
        if resolved == source:
            continue
        snapshots.append(inspect_document(resolved))
    return snapshots


def _build_extracted_data(
    snapshot: DocumentSnapshot,
    chapter_blocks: list[ChapterBlock],
    references: list[DocumentSnapshot],
    reference_audit: ReferenceAuditResult,
) -> dict:
    similarity = _max_similarity(snapshot, references) if references else None
    return {
        "title": _extract_cover_value(snapshot, "题目"),
        "major": _extract_cover_value(snapshot, "专业"),
        "abstract_chars": _chinese_char_count(_extract_abstract_text(snapshot)),
        "body_chars": _chinese_char_count(_extract_body_text(snapshot)),
        "keyword_count": len(_extract_keywords(snapshot)),
        "chapter_count": len(chapter_blocks),
        "conclusion_chars": _chinese_char_count(_extract_conclusion_text(snapshot, chapter_blocks)),
        "reference_count": len(_extract_reference_entries(snapshot)),
        "reference_verified_count": reference_audit.verified_count,
        "reference_not_found_count": reference_audit.not_found_count,
        "reference_similarity": round(similarity, 3) if similarity is not None else None,
    }


def _score_page_setup(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["page_setup"]
    if not snapshot.sections:
        return ScoreItem("page_setup", "页面设置", 0.0, max_score, "未通过", ["无法读取分节信息"], ["确认文档可被 Word 正常打开。"])

    weighted_hits = 0.0
    checks = 0
    evidence: list[str] = []
    major_deviations: list[str] = []
    for section in snapshot.sections:
        values = {
            "top": section.top_margin_cm,
            "bottom": section.bottom_margin_cm,
            "left": section.left_margin_cm,
            "right": section.right_margin_cm,
            "header_distance": section.header_distance_cm,
            "footer_distance": section.footer_distance_cm,
        }
        for name, expected in EXPECTED_PAGE_SETUP_CM.items():
            checks += 1
            diff = abs(values[name] - expected)
            weighted_hits += _soft_visual_match(diff, PAGE_SETUP_TOLERANCE_CM, 0.45, 0.85)
            if diff > 0.45:
                major_deviations.append(f"第{section.index}节{name}偏差约 {diff:.2f} cm")

        page_size_diffs = (
            abs(section.page_width_cm - A4_PAGE_SIZE_CM[0]),
            abs(section.page_height_cm - A4_PAGE_SIZE_CM[1]),
        )
        checks += 2
        weighted_hits += _soft_visual_match(page_size_diffs[0], 0.15, 0.35, 0.80)
        weighted_hits += _soft_visual_match(page_size_diffs[1], 0.15, 0.35, 0.80)
        if max(page_size_diffs) > 0.35:
            major_deviations.append(f"第{section.index}节纸张尺寸约 {section.page_width_cm} x {section.page_height_cm} cm")

        evidence.append(
            f"第{section.index}节页边距 {section.top_margin_cm}/{section.bottom_margin_cm}/{section.left_margin_cm}/{section.right_margin_cm} cm，页眉/页脚 {section.header_distance_cm}/{section.footer_distance_cm} cm，纸张 {section.page_width_cm} x {section.page_height_cm} cm"
        )

    score = round(max_score * weighted_hits / max(checks, 1), 2)
    suggestions = []
    if major_deviations:
        suggestions.append("页面设置与标准差异较大，需统一为 A4，上下 2.54cm，左右 3.17cm，页眉 1.5cm，页脚 1.75cm。")
    elif score < max_score:
        suggestions.append("页面设置存在轻微偏差，建议统一到学校模板参数。")
    return ScoreItem("page_setup", "页面设置", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_cover(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["cover"]
    cover_text = "\n".join(paragraph.text for paragraph in _cover_paragraphs(snapshot))
    evidence: list[str] = []
    suggestions: list[str] = []
    hits = 0

    if "黑龙江省经济管理干部学院" in cover_text:
        hits += 1
        evidence.append("封面识别到学校名称。")
    else:
        suggestions.append("封面缺少学校名称。")

    for label in COVER_REQUIRED_LABELS:
        if _extract_cover_value(snapshot, label):
            hits += 1
            evidence.append(f"封面已填写{label}。")
        else:
            suggestions.append(f"补全封面字段：{label}。")

    if _placeholder_count(cover_text) == 0:
        hits += 1
        evidence.append("封面未发现模板占位符。")
    else:
        suggestions.append("删除封面的模板占位符。")

    score = round(max_score * hits / 7, 2)
    return ScoreItem("cover", "封面完整度", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_header_pagination(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["header_pagination"]
    if not snapshot.sections:
        return ScoreItem("header_pagination", "页眉与页码", 0.0, max_score, "未通过", ["无法读取页眉信息"], ["检查 Word 页眉页脚设置。"])

    evidence: list[str] = []
    suggestions: list[str] = []
    header_hits = 0
    roman_like = False
    arabic_like = False
    field_hits = 0

    for section in snapshot.sections:
        text = section.header_text
        evidence.append(f"第{section.index}节页眉: {text or '空'}")
        if any(variant in text for variant in HEADER_VARIANTS):
            header_hits += 1
        if re.search(r"^[IVXLC]+\s", text):
            roman_like = True
        if re.search(r"第\s*\d+\s*页", text):
            arabic_like = True
        if section.header_field_count > 0 or section.footer_field_count > 0:
            field_hits += 1

    score = 0.0
    score += max_score * 0.45 * (header_hits / len(snapshot.sections))
    score += max_score * 0.20 * float(roman_like)
    score += max_score * 0.25 * float(arabic_like)
    score += max_score * 0.10 * (field_hits / len(snapshot.sections))
    score = round(score, 2)

    if header_hits < len(snapshot.sections):
        suggestions.append("统一页眉为学校规定文本。")
    if not roman_like:
        suggestions.append("摘要、目录等前置部分建议使用罗马数字页码。")
    if not arabic_like:
        suggestions.append("正文部分建议使用“第M页”的页码格式。")
    return ScoreItem("header_pagination", "页眉与页码", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_abstract_keywords_format(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["abstract_keywords_format"]
    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0

    heading = _find_paragraph(snapshot, lambda paragraph: paragraph.normalized == "摘要")
    abstract_chars = _chinese_char_count(_extract_abstract_text(snapshot))
    keyword_paragraph = _find_keyword_paragraph(snapshot)
    keywords = _extract_keywords(snapshot)

    if heading:
        evidence.append(f"摘要标题字号约 {heading.font_size} 磅。")
        if _is_close(heading.font_size, 14.0, 0.6):
            score += 1.5
        if heading.alignment == 1:
            score += 0.8
    else:
        suggestions.append("补充独立的“摘要”标题。")

    evidence.append(f"摘要约 {abstract_chars} 字。")
    if STRICT_ABSTRACT_RANGE[0] <= abstract_chars <= STRICT_ABSTRACT_RANGE[1]:
        score += 2.0
    elif SOFT_ABSTRACT_RANGE[0] <= abstract_chars <= SOFT_ABSTRACT_RANGE[1]:
        score += 1.0
        suggestions.append("摘要长度偏离学校建议值，建议控制在 150 字左右。")
    else:
        suggestions.append("摘要长度明显不合适，需压缩或补写。")

    evidence.append(f"关键词数量 {len(keywords)}。")
    if keyword_paragraph and _is_close(keyword_paragraph.font_size, 10.5, 0.6):
        score += 0.8
    if KEYWORD_RANGE[0] <= len(keywords) <= KEYWORD_RANGE[1]:
        score += 0.9
    elif keywords:
        suggestions.append("关键词建议保留 3-5 个。")
    else:
        suggestions.append("补充“关键词”并使用分号分隔。")

    score = round(min(score, max_score), 2)
    return ScoreItem("abstract_keywords_format", "摘要与关键词格式", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_toc_and_sections(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["toc_and_sections"]
    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0

    toc_heading = _find_paragraph(snapshot, lambda paragraph: paragraph.normalized == "目录")
    toc_entries = _toc_entries(snapshot)
    if toc_heading:
        score += 1.5
        evidence.append("检测到目录标题。")
    else:
        suggestions.append("补充目录页。")

    if len(toc_entries) >= 4:
        score += 1.5
        evidence.append(f"目录项约 {len(toc_entries)} 条。")
    elif toc_entries:
        score += 0.8
        suggestions.append("目录项偏少，建议覆盖主要二级标题。")
    else:
        suggestions.append("未识别到带页码的目录项。")

    required_hits = 0
    for section_name in MANDATORY_SECTIONS:
        if _has_section(snapshot, section_name):
            required_hits += 1
            evidence.append(f"存在必备部分: {section_name}")
        else:
            suggestions.append(f"补充必备部分: {section_name}")
    score += 1.5 * required_hits / len(MANDATORY_SECTIONS)

    recommended_hits = 0
    for section_name in RECOMMENDED_SECTIONS:
        if _has_section(snapshot, section_name):
            recommended_hits += 1
            evidence.append(f"存在推荐部分: {section_name}")
    if recommended_hits == 0:
        suggestions.append("可补充致谢或附录，提升结构完整度。")
    else:
        score += 0.5

    score = round(min(score, max_score), 2)
    return ScoreItem("toc_and_sections", "目录与结构完整性", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_body_typography(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = FORMAT_WEIGHTS["body_typography"]
    metrics = _body_visual_metrics(snapshot)
    if metrics["sample_count"] == 0:
        return ScoreItem("body_typography", "正文字体与标题", 0.0, max_score, "未通过", ["未识别到正文内容"], ["确认正文位于目录和参考文献之间。"])

    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0
    score += 1.8 * metrics["body_font_quality"]
    score += 1.0 * metrics["size_ratio"]
    score += 0.9 * metrics["indent_ratio"]
    score += 1.8 * (
        0.45 * metrics["rhythm_cleanliness"]
        + 0.30 * metrics["alignment_ratio"]
        + 0.25 * metrics["spacing_consistency"]
    )
    score += 1.5 * (0.6 * metrics["dominant_group_ratio"] + 0.4 * metrics["heading_visual_ratio"])

    if metrics["hostile_body_ratio"] >= 0.25:
        score = min(score, 4.4)
    if metrics["blank_ratio"] >= 0.16 or metrics["max_blank_run"] >= 2:
        score = min(score, 4.8)
    if metrics["spacing_outlier_ratio"] >= 0.30:
        score = min(score, 4.6)
    if metrics["dominant_group_ratio"] < 0.55:
        score = min(score, 4.3)
    if metrics["hostile_body_ratio"] >= 0.25 and metrics["blank_ratio"] >= 0.16:
        score = min(score, 3.8)

    score = round(min(score, max_score), 2)

    evidence.append(
        f"抽样正文 {metrics['sample_count']} 段，主导字族 {metrics['dominant_group_label']}，覆盖率 {metrics['dominant_group_ratio']:.2f}，正文视觉兼容率 {metrics['body_font_quality']:.2f}。"
    )
    evidence.append(
        f"五号附近字号命中 {metrics['size_hits']} 段，首行缩进命中 {metrics['indent_hits']} 段，正文对齐稳定率 {metrics['alignment_ratio']:.2f}。"
    )
    evidence.append(
        f"正文区域空段 {metrics['blank_paragraphs']} 个，空段占比 {metrics['blank_ratio']:.2f}，最大连续空段 {metrics['max_blank_run']} 个，段前段后异常占比 {metrics['spacing_outlier_ratio']:.2f}。"
    )
    evidence.append(
        f"章标题 {metrics['chapter_count']} 个，视觉规范命中率 {metrics['heading_visual_ratio']:.2f}。"
    )

    if metrics["hostile_body_ratio"] >= 0.18:
        suggestions.append("正文存在较多黑体或过重字族，第一眼观感明显偏离规范论文。")
    elif metrics["body_font_quality"] < 0.75:
        suggestions.append("正文主字体不够稳定，建议统一到宋体或仿宋一类的正文风格。")

    if metrics["blank_ratio"] >= 0.10 or metrics["max_blank_run"] >= 2:
        suggestions.append("正文里有较多空段或多余回车，版面节奏发散，翻页时会很乱。")
    if metrics["spacing_outlier_ratio"] >= 0.20:
        suggestions.append("正文段前段后留白偏大或不统一，视觉上会出现明显跳行感。")
    if metrics["size_ratio"] < 0.75:
        suggestions.append("正文字号不够统一，建议稳定在五号附近。")
    if metrics["indent_ratio"] < 0.75:
        suggestions.append("正文首行缩进不统一，建议统一为两个汉字。")
    if metrics["dominant_group_ratio"] < 0.68:
        suggestions.append("正文混用多种字体风格，整体看上去不像一套版式。")
    if metrics["heading_visual_ratio"] < 0.75:
        suggestions.append("章标题建议统一为居中、醒目且稳定的标题样式。")
    return ScoreItem("body_typography", "正文字体与标题", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_topic_relevance(snapshot: DocumentSnapshot, chapter_blocks: list[ChapterBlock]) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["topic_relevance"]
    title = _extract_cover_value(snapshot, "题目")
    major = _extract_cover_value(snapshot, "专业")
    body_text = _extract_body_text(snapshot)
    heading_text = " ".join(block.heading.text for block in chapter_blocks)
    title_terms = _concept_terms(title)
    major_terms = _concept_terms(major)

    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0

    if title:
        title_chars = _chinese_char_count(title)
        evidence.append(f"标题: {title}")
        if 6 <= title_chars <= TITLE_MAX_CHARS:
            score += 1.8
        elif 4 <= title_chars <= TITLE_MAX_CHARS + 4:
            score += 0.8
            suggestions.append(f"标题长度不够理想，建议控制在 {TITLE_MAX_CHARS} 字以内并保持聚焦。")
        else:
            suggestions.append("标题长度明显不合适。")
    else:
        suggestions.append("封面题目未识别。")

    if title_terms:
        body_ratio = _term_coverage_ratio(title_terms, body_text)
        heading_ratio = _term_coverage_ratio(title_terms, heading_text)
        evidence.append(f"标题核心词正文覆盖率 {body_ratio:.2f}，章节标题覆盖率 {heading_ratio:.2f}。")
        score += 2.0 * min(body_ratio, 1.0)
        score += 0.8 * min(heading_ratio, 1.0)
        if body_ratio < 0.6:
            suggestions.append("正文对标题核心概念的回应不足，存在跑题风险。")
    else:
        suggestions.append("标题核心概念不够清晰，建议更明确研究对象。")

    if major_terms:
        major_ratio = _term_coverage_ratio(major_terms, title + " " + body_text)
        evidence.append(f"专业相关词覆盖率 {major_ratio:.2f}。")
        score += 1.4 * min(major_ratio, 1.0)
        if major_ratio < 0.5:
            suggestions.append("题目与正文和专业的贴合度偏弱。")
    else:
        suggestions.append("封面专业未识别。")

    score = round(min(score, max_score), 2)
    return ScoreItem("topic_relevance", "选题与专业相关性", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_structure_logic(snapshot: DocumentSnapshot, chapter_blocks: list[ChapterBlock]) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["structure_logic"]
    level2 = _find_heading_matches(snapshot, r"^[一二三四五六七八九十]+、")
    level3 = _find_heading_matches(snapshot, r"^（[一二三四五六七八九十]+）|^\([一二三四五六七八九十]+\)")
    intro_ok = bool(chapter_blocks) and _contains_any(chapter_blocks[0].heading.text, INTRO_CUES + OVERVIEW_CUES)
    conclusion_ok = bool(chapter_blocks) and _contains_any(chapter_blocks[-1].heading.text, CONCLUSION_CUES)

    evidence = [
        f"主章节 {len(chapter_blocks)} 个。",
        f"二级标题 {len(level2)} 个。",
        f"三级标题 {len(level3)} 个。",
    ]
    suggestions: list[str] = []
    score = 0.0

    if len(chapter_blocks) >= 5:
        score += 3.0
    elif len(chapter_blocks) >= 4:
        score += 2.0
    elif len(chapter_blocks) >= 3:
        score += 1.0
        suggestions.append("主章节偏少，结构深度不够。")
    else:
        suggestions.append("未形成像样的章节结构。")

    if intro_ok:
        score += 1.8
        evidence.append("首章具备引入或背景性质。")
    else:
        suggestions.append("首章没有明显承担引言/绪论功能。")

    if conclusion_ok:
        score += 1.8
        evidence.append("末章具备结论或展望性质。")
    else:
        suggestions.append("缺少明确的结论或展望章节。")

    if level2:
        score += 1.4
    else:
        suggestions.append("缺少二级标题，层次不够清楚。")

    if level3:
        score += 1.0
    else:
        suggestions.append("缺少三级标题，展开不够细。")

    required_hits = sum(1 for name in MANDATORY_SECTIONS if _has_section(snapshot, name))
    score += 1.0 * required_hits / len(MANDATORY_SECTIONS)
    if required_hits < len(MANDATORY_SECTIONS):
        suggestions.append("摘要、目录、参考文献没有全部落齐。")

    middle_substantial = sum(1 for block in chapter_blocks[1:-1] if _chinese_char_count(block.text) >= 350)
    evidence.append(f"中间章节中，实质展开章节 {middle_substantial} 个。")
    if middle_substantial >= max(2, len(chapter_blocks) - 2):
        score += 1.0
    elif middle_substantial >= 2:
        score += 0.5
    else:
        suggestions.append("中间章节偏空，前后像骨架但中间论证撑不起来。")

    score = round(min(score, max_score), 2)
    return ScoreItem("structure_logic", "结构完整性与逻辑", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_abstract_quality(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["abstract_quality"]
    abstract_text = _extract_abstract_text(snapshot)
    abstract_chars = _chinese_char_count(abstract_text)
    evidence = [f"摘要长度约 {abstract_chars} 字。"]
    suggestions: list[str] = []
    score = 0.0

    if STRICT_ABSTRACT_RANGE[0] <= abstract_chars <= STRICT_ABSTRACT_RANGE[1]:
        score += 2.5
    elif SOFT_ABSTRACT_RANGE[0] <= abstract_chars <= SOFT_ABSTRACT_RANGE[1]:
        score += 1.2
        suggestions.append("摘要长度偏离推荐值，需压缩或补足。")
    else:
        suggestions.append("摘要长度明显不规范。")

    cue_groups = {
        "研究对象": ["本文", "研究", "围绕", "以"],
        "研究方法": METHOD_CUES,
        "研究发现": RESULT_CUES,
        "研究意义": ["意义", "价值", "启示", "作用"],
    }

    matched_groups = 0
    for label, terms in cue_groups.items():
        if _contains_any(abstract_text, terms):
            matched_groups += 1
            evidence.append(f"摘要覆盖 {label}。")
        else:
            suggestions.append(f"摘要缺少 {label}。")

    score += 3.5 * matched_groups / len(cue_groups)
    if len(_extract_keywords(snapshot)) == 0:
        suggestions.append("摘要后缺少关键词。")
    elif len(_extract_keywords(snapshot)) > KEYWORD_RANGE[1]:
        suggestions.append("关键词过多，摘要也显得偏散。")

    score = round(min(score, max_score), 2)
    return ScoreItem("abstract_quality", "摘要内容质量", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_chapter_development(snapshot: DocumentSnapshot, chapter_blocks: list[ChapterBlock]) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["chapter_development"]
    body_text = _extract_body_text(snapshot)
    body_chars = _chinese_char_count(body_text)
    body_paragraphs = _body_paragraphs(snapshot)
    evidence = [f"正文约 {body_chars} 字。", f"主章节 {len(chapter_blocks)} 个。"]
    suggestions: list[str] = []
    score = 0.0

    if body_chars >= MIN_BODY_CHARS + 800:
        score += 3.0
    elif body_chars >= MIN_BODY_CHARS:
        score += 2.2
    elif body_chars >= 2500:
        score += 1.0
        suggestions.append("正文字数接近下限，但仍偏薄。")
    else:
        suggestions.append("正文字数明显不足。")

    avg_per_chapter = body_chars / max(len(chapter_blocks), 1)
    evidence.append(f"平均每章约 {int(avg_per_chapter)} 字。")
    if avg_per_chapter >= 500:
        score += 2.0
    elif avg_per_chapter >= 320:
        score += 1.0
        suggestions.append("章节展开偏薄。")
    else:
        suggestions.append("章节过短，论证深度不足。")

    substantial_blocks = sum(1 for block in chapter_blocks if _chinese_char_count(block.text) >= 450)
    evidence.append(f"达到 450 字以上的章节 {substantial_blocks} 个。")
    if substantial_blocks >= max(3, len(chapter_blocks) - 1):
        score += 2.0
    elif substantial_blocks >= 3:
        score += 1.0
    else:
        suggestions.append("真正写开的章节太少。")

    cue_hits = sum(body_text.count(term) for term in ACADEMIC_CUES)
    evidence.append(f"分析信号词累计 {cue_hits} 次。")
    if cue_hits >= 35:
        score += 2.0
    elif cue_hits >= 18:
        score += 1.2
    else:
        suggestions.append("分析性表达不足，正文更像资料堆叠。")

    long_paragraphs = sum(1 for paragraph in body_paragraphs if _chinese_char_count(paragraph.text) >= 80)
    evidence.append(f"80 字以上段落 {long_paragraphs} 个。")
    if long_paragraphs >= 10:
        score += 1.0
    elif long_paragraphs < 6:
        suggestions.append("有实质内容的段落偏少。")

    score = round(min(score, max_score), 2)
    return ScoreItem("chapter_development", "正文展开与论证深度", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_coherence_alignment(snapshot: DocumentSnapshot, chapter_blocks: list[ChapterBlock]) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["coherence_alignment"]
    title_terms = _core_terms(snapshot)
    abstract_text = _extract_abstract_text(snapshot)
    body_text = _extract_body_text(snapshot)
    conclusion_text = _extract_conclusion_text(snapshot, chapter_blocks)
    heading_text = " ".join(block.heading.text for block in chapter_blocks)

    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0

    body_ratio = _term_coverage_ratio(title_terms, body_text)
    heading_ratio = _term_coverage_ratio(title_terms, heading_text)
    conclusion_ratio = _term_coverage_ratio(title_terms, conclusion_text)
    evidence.append(
        f"标题/关键词核心词在正文、章标题、结论中的覆盖率分别为 {body_ratio:.2f}/{heading_ratio:.2f}/{conclusion_ratio:.2f}。"
    )
    score += 5.0 * min((body_ratio * 0.5 + heading_ratio * 0.2 + conclusion_ratio * 0.3), 1.0)
    if body_ratio < 0.7:
        suggestions.append("正文没有持续围绕标题核心概念展开，存在明显跑题风险。")
    if conclusion_ratio < 0.5:
        suggestions.append("结论没有回扣标题核心概念，前后照应不足。")

    abstract_terms = _concept_terms(abstract_text)[:8] or title_terms
    abstract_body_ratio = _term_coverage_ratio(abstract_terms, body_text)
    abstract_conclusion_ratio = _term_coverage_ratio(abstract_terms, conclusion_text)
    evidence.append(f"摘要核心词在正文和结论中的覆盖率分别为 {abstract_body_ratio:.2f}/{abstract_conclusion_ratio:.2f}。")
    score += 4.0 * min((abstract_body_ratio * 0.7 + abstract_conclusion_ratio * 0.3), 1.0)
    if abstract_body_ratio < 0.65:
        suggestions.append("摘要写了不少内容，但正文没有一一落地。")
    if abstract_conclusion_ratio < 0.4:
        suggestions.append("摘要中的核心结论没有在结尾得到回收。")

    method_chain = 0
    if _contains_any(abstract_text, METHOD_CUES):
        method_chain += 1
    if _contains_any(body_text, METHOD_CUES):
        method_chain += 1
    if _contains_any(conclusion_text, RESULT_CUES + ["总结", "结论"]):
        method_chain += 1
    problem_solution_pair = int(_contains_any(body_text, PROBLEM_CUES) and _contains_any(body_text, SOLUTION_CUES))
    evidence.append(f"方法链命中 {method_chain}/3，问题-对策配对命中 {problem_solution_pair}/1。")
    score += 4.0 * (method_chain + problem_solution_pair) / 4
    if method_chain < 2:
        suggestions.append("研究方法、过程、结论之间没有形成闭环。")
    if not problem_solution_pair and _contains_any(body_text, PROBLEM_CUES):
        suggestions.append("文中提出了问题，但没有给出足够的对策或改进路径。")

    progression_score, progression_evidence, progression_suggestions = _chapter_progression_score(chapter_blocks)
    evidence.extend(progression_evidence)
    suggestions.extend(progression_suggestions)
    score += progression_score

    conclusion_chars = _chinese_char_count(conclusion_text)
    evidence.append(f"结论部分约 {conclusion_chars} 字。")
    if conclusion_chars >= 220:
        score += 2.0
    elif conclusion_chars >= 120:
        score += 1.0
        suggestions.append("结论存在，但总结力度偏弱。")
    else:
        suggestions.append("结论太短，无法承担收束全文的作用。")

    score = round(min(score, max_score), 2)
    return ScoreItem("coherence_alignment", "前后逻辑自洽与闭环", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_references_support(snapshot: DocumentSnapshot, reference_audit: ReferenceAuditResult) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["references_support"]
    references = reference_audit.entries
    evidence = [f"参考文献 {len(references)} 条。"]
    suggestions: list[str] = []
    score = 0.0

    if not references:
        suggestions.append("参考文献缺失，正文论证没有外部文献支撑。")
        return ScoreItem("references_support", "参考文献与支撑材料", 0.0, max_score, "未通过", evidence, suggestions)

    parseable_count = sum(1 for entry in references if entry.title)
    verified_equivalent = reference_audit.verified_count + 0.5 * reference_audit.possible_count
    cited_not_found = sum(1 for check in reference_audit.checks if check.cited_in_body and check.status == "not_found")

    evidence.append(f"可解析标题 {parseable_count}/{len(references)} 条。")
    evidence.append(
        f"联网核验通过 {reference_audit.verified_count} 条，疑似匹配 {reference_audit.possible_count} 条，未找到 {reference_audit.not_found_count} 条。"
    )
    evidence.append(f"正文编号引用 {len(reference_audit.citation_numbers)} 处。")
    if reference_audit.dangling_citations:
        evidence.append(f"存在无文后来源的引用编号 {reference_audit.dangling_citations}。")
    if reference_audit.uncited_reference_indexes:
        evidence.append(f"列出但未在正文出现的文献编号 {reference_audit.uncited_reference_indexes}。")

    if len(references) >= MIN_REFERENCE_COUNT + 2:
        score += 1.0
    elif len(references) >= MIN_REFERENCE_COUNT:
        score += 0.7
    elif len(references) >= 3:
        score += 0.3
        suggestions.append("参考文献数量偏少，对初稿论证支撑仍显单薄。")
    else:
        suggestions.append("参考文献严重不足。")

    score += 0.8 * parseable_count / len(references)
    if parseable_count < len(references):
        suggestions.append("部分参考文献条目格式混乱，连标题都无法稳定解析。")

    if not reference_audit.citation_numbers:
        suggestions.append("正文未使用规范的 [1][2] 编号引用，无法建立文内外映射。")
    elif reference_audit.citation_mapping_ok:
        score += 1.6
    elif not reference_audit.dangling_citations and len(reference_audit.uncited_reference_indexes) <= max(1, len(references) // 4):
        score += 0.8
        suggestions.append("文内引用与文后文献并未完全一一对应。")
    else:
        suggestions.append("文内引用与文后参考文献对应关系明显失衡。")

    if cited_not_found:
        suggestions.append(f"正文实际引用的文献里有 {cited_not_found} 条未核验到真实来源，存在杜撰风险。")
    else:
        authenticity_ratio = verified_equivalent / len(references)
        if authenticity_ratio >= 0.75 and reference_audit.verified_count >= max(1, len(references) // 2):
            score += 1.6
        elif authenticity_ratio >= 0.5:
            score += 0.9
        elif authenticity_ratio >= 0.25:
            score += 0.4
            suggestions.append("已有少量文献可核验，但整体真实性支撑仍偏弱。")
        else:
            suggestions.append("联网核验未形成足够的真实文献支撑。")

    if reference_audit.search_error_count:
        evidence.append(f"有 {reference_audit.search_error_count} 条文献因联网异常未完成核验。")

    if cited_not_found or reference_audit.dangling_citations:
        score = min(score, 0.8)
    if not reference_audit.citation_numbers:
        score = min(score, 1.0)
    if reference_audit.not_found_count >= max(2, (len(references) + 2) // 3):
        score = min(score, 0.5)

    score = round(min(score, max_score), 2)
    return ScoreItem("references_support", "参考文献与支撑材料", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_language_quality(snapshot: DocumentSnapshot) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["language_quality"]
    body_text = _extract_body_text(snapshot)
    sentences = _split_sentences(body_text)
    avg_sentence_chars = round(sum(_chinese_char_count(sentence) for sentence in sentences) / max(len(sentences), 1), 1)
    placeholder_hits = _placeholder_count(snapshot.raw_text)
    repeated_punct = len(re.findall(r"[。！？]{2,}|[.,]{3,}", body_text))
    tokens = _tokenize_keywords(body_text)
    diversity = len(set(tokens)) / max(len(tokens), 1)
    repetition_ratio, repeated_sentences, analyzable_sentences = _sentence_repetition_stats(body_text)

    evidence = [
        f"句子数约 {len(sentences)}，平均句长 {avg_sentence_chars} 字。",
        f"词汇离散度约 {diversity:.3f}。",
        f"可分析句子 {analyzable_sentences} 句，重复句 {repeated_sentences} 句，重复率约 {repetition_ratio:.3f}。",
    ]
    suggestions: list[str] = []
    score = 0.0

    if 18 <= avg_sentence_chars <= 65:
        score += 1.6
    elif 12 <= avg_sentence_chars <= 85:
        score += 0.8
        suggestions.append("句式节奏一般，建议再润色。")
    else:
        suggestions.append("句长控制较差，读起来不稳定。")

    if placeholder_hits == 0 and repeated_punct == 0:
        score += 0.8
        evidence.append("未发现明显占位符或重复标点。")
    else:
        suggestions.append("存在模板残留或标点问题。")

    if diversity >= 0.45:
        score += 1.0
    elif diversity >= 0.35:
        score += 0.5
    else:
        suggestions.append("表达重复度偏高。")

    if repetition_ratio <= 0.05:
        score += 0.6
    elif repetition_ratio <= 0.12:
        score += 0.3
        suggestions.append("局部已经出现重复表述，建议压缩套话。")
    else:
        suggestions.append("重复句比例偏高，存在车轱辘话或 AI 套话风险。")

    score = round(min(score, max_score), 2)
    return ScoreItem("language_quality", "语言表达质量", score, max_score, _status(score, max_score), evidence, suggestions)


def _score_academic_integrity(
    snapshot: DocumentSnapshot,
    references: list[DocumentSnapshot],
    reference_audit: ReferenceAuditResult,
) -> ScoreItem:
    max_score = CONTENT_WEIGHTS["academic_integrity"]
    template_markers = _placeholder_count(snapshot.raw_text)
    evidence: list[str] = []
    suggestions: list[str] = []
    score = max_score

    if template_markers:
        evidence.append(f"检测到模板占位符 {template_markers} 处。")
        suggestions.append("先清除模板占位符，再提交论文。")

    similarity = _max_similarity(snapshot, references) if references else None
    if similarity is None:
        evidence.append("未提供参考文档，仅做内部风险提示。")
        score = min(score, 3.5 if template_markers == 0 else 1.0)
    else:
        evidence.append(f"与参考文档最高相似度 {similarity:.3f}。")
        if similarity < 0.12:
            pass
        elif similarity < 0.20:
            score -= 1.0
            suggestions.append("与参考文档有一定相似度，仍建议继续改写。")
        elif similarity < 0.30:
            score -= 3.0
            suggestions.append("与参考文档相似度偏高，需要重点改写。")
        else:
            score = 0.0
            suggestions.append("与参考文档高度相似，存在明显套用风险。")

    cited_not_found = sum(1 for check in reference_audit.checks if check.cited_in_body and check.status == "not_found")
    evidence.append(
        f"引用核验结果: 已验证 {reference_audit.verified_count} 条，疑似匹配 {reference_audit.possible_count} 条，未找到 {reference_audit.not_found_count} 条。"
    )

    if not reference_audit.entries:
        score = 0.0
        suggestions.append("未列出参考文献，学术规范不成立。")
    else:
        if cited_not_found:
            score = 0.0
            suggestions.append(f"正文引用的 {cited_not_found} 条文献未核验到真实来源，疑似存在杜撰。")
        elif reference_audit.not_found_count >= max(2, (len(reference_audit.entries) + 2) // 3):
            score = min(score, 1.0)
            suggestions.append("参考文献中未核验到的条目过多，真实性风险很高。")
        elif reference_audit.not_found_count > 0:
            score = min(score, 2.5)
            suggestions.append("参考文献里存在未核验到的条目，需要逐条复核。")

        if not reference_audit.citation_numbers:
            score = min(score, 1.5)
            suggestions.append("正文没有规范编号引用，文内外无法建立稳定映射。")
        if reference_audit.dangling_citations:
            score = min(score, 1.0)
            suggestions.append("正文存在找不到文后来源的引用编号。")
        if reference_audit.citation_numbers and len(reference_audit.uncited_reference_indexes) > max(1, len(reference_audit.entries) // 4):
            score = min(score, 2.0)
            suggestions.append("参考文献列表与正文脱节，存在列而不用的条目。")
        if reference_audit.possible_count and reference_audit.verified_count == 0 and reference_audit.not_found_count:
            score = min(score, 1.0)
            suggestions.append("当前文献多为疑似匹配或未找到状态，真实性不足以支撑正常评阅。")

    if template_markers:
        score = min(score, 1.0)
    score = round(min(score, max_score), 2)
    return ScoreItem("academic_integrity", "学术规范与相似度风险", score, max_score, _status(score, max_score), evidence, suggestions)


def _evaluate_reference_gate(reference_audit: ReferenceAuditResult) -> GateResult:
    reasons: list[str] = []
    total = len(reference_audit.entries)
    cited_not_found = sum(1 for check in reference_audit.checks if check.cited_in_body and check.status == "not_found")

    severe = False
    moderate = False

    if total == 0:
        severe = True
        reasons.append("参考文献缺失，无法对论文来源与支撑材料进行核验。")

    if total > 0 and not reference_audit.citation_numbers:
        severe = True
        reasons.append("正文未检测到规范的 [1][2] 编号引用，文内外无法建立一一对应关系。")

    if reference_audit.dangling_citations:
        severe = True
        reasons.append(f"正文存在找不到文后条目的引用编号: {reference_audit.dangling_citations}。")

    if cited_not_found:
        severe = True
        reasons.append(f"正文实际引用的文献中有 {cited_not_found} 条未核验到真实来源，疑似杜撰。")
    elif reference_audit.not_found_count >= max(2, (total + 2) // 3):
        severe = True
        reasons.append(f"参考文献中未核验到的条目达到 {reference_audit.not_found_count} 条，真实性风险过高。")
    elif reference_audit.not_found_count > 0:
        moderate = True
        reasons.append(f"参考文献中仍有 {reference_audit.not_found_count} 条未核验到，需要人工复核。")

    if reference_audit.citation_numbers and len(reference_audit.uncited_reference_indexes) > max(1, total // 3):
        moderate = True
        reasons.append(f"参考文献列表中有较多条目未在正文出现: {reference_audit.uncited_reference_indexes}。")

    if severe:
        return GateResult("引用退修", True, 59.0, reasons)

    if moderate:
        return GateResult("引用待核", True, 69.0, reasons)

    return GateResult("通过", False, None, [])


def _build_visual_review(
    snapshot: DocumentSnapshot,
    stage: str,
    mode: str,
    model: str,
    output_dir: str | None,
) -> VisualReviewResult:
    normalized_mode = (mode or "auto").lower()
    if normalized_mode == "off":
        return VisualReviewResult(
            mode="skipped",
            model=None,
            pdf_path=None,
            overall_verdict="pass",
            visual_order_score=None,
            confidence=None,
            major_issues=[],
            minor_issues=[],
            evidence=[],
            page_observations=[],
            notes=["视觉审稿已关闭，本次仅使用规则引擎。"],
        )

    if normalized_mode == "moonshot":
        try:
            moonshot_model = model if model and not model.startswith("gpt-") else None
            resolved_moonshot_model = resolve_moonshot_visual_model_name(moonshot_model)
            return review_document_with_moonshot(
                snapshot.path,
                output_dir=output_dir,
                model=resolved_moonshot_model,
                stage=stage,
            )
        except Exception as exc:
            return VisualReviewResult(
                mode="unavailable",
                model=resolved_moonshot_model,
                pdf_path=None,
                overall_verdict="minor_issues",
                visual_order_score=None,
                confidence=None,
                major_issues=[],
                minor_issues=[],
                evidence=[],
                page_observations=[],
                notes=[f"Moonshot/Kimi 视觉审稿未成功执行: {type(exc).__name__}: {exc}"],
            )

    if normalized_mode == "siliconflow":
        try:
            siliconflow_model = model if model and not model.startswith("gpt-") else None
            resolved_siliconflow_model = resolve_siliconflow_visual_model_name(siliconflow_model)
            return review_document_with_siliconflow(
                snapshot.path,
                output_dir=output_dir,
                model=resolved_siliconflow_model,
                stage=stage,
            )
        except Exception as exc:
            return VisualReviewResult(
                mode="unavailable",
                model=resolved_siliconflow_model,
                pdf_path=None,
                overall_verdict="minor_issues",
                visual_order_score=None,
                confidence=None,
                major_issues=[],
                minor_issues=[],
                evidence=[],
                page_observations=[],
                notes=[f"SiliconFlow 视觉审稿未成功执行: {type(exc).__name__}: {exc}"],
            )

    if normalized_mode in {"expert", "ensemble"}:
        try:
            primary_model = model if model and not model.startswith("gpt-") else None
            resolved_primary_model = resolve_siliconflow_visual_model_name(primary_model)
            resolved_secondary_model = resolve_siliconflow_secondary_visual_model_name(
                None if primary_model else DEFAULT_SILICONFLOW_SECONDARY_VISUAL_MODEL
            )
            return review_document_with_siliconflow_ensemble(
                snapshot.path,
                output_dir=output_dir,
                primary_model=resolved_primary_model,
                secondary_model=resolved_secondary_model,
                stage=stage,
            )
        except Exception as exc:
            return VisualReviewResult(
                mode="unavailable",
                model=resolved_primary_model,
                pdf_path=None,
                overall_verdict="minor_issues",
                visual_order_score=None,
                confidence=None,
                major_issues=[],
                minor_issues=[],
                evidence=[],
                page_observations=[],
                notes=[f"双专家视觉审稿未成功执行: {type(exc).__name__}: {exc}"],
            )

    if normalized_mode == "auto":
        try:
            if has_siliconflow_visual_credentials():
                resolved_primary_model = resolve_siliconflow_visual_model_name(
                    model if model and not model.startswith("gpt-") else None
                )
                return review_document_with_siliconflow(
                    snapshot.path,
                    output_dir=output_dir,
                    model=resolved_primary_model,
                    stage=stage,
                )
            if has_moonshot_visual_credentials():
                moonshot_model = model if model and not model.startswith("gpt-") else None
                resolved_moonshot_model = resolve_moonshot_visual_model_name(moonshot_model)
                return review_document_with_moonshot(
                    snapshot.path,
                    output_dir=output_dir,
                    model=resolved_moonshot_model,
                    stage=stage,
                )
            return review_document_with_openai(
                snapshot.path,
                output_dir=output_dir,
                model=model,
                stage=stage,
            )
        except Exception as exc:
            fallback = _heuristic_visual_review(snapshot)
            if output_dir:
                try:
                    fallback.pdf_path = str(export_document_to_pdf(snapshot.path, output_dir))
                except Exception as pdf_exc:
                    fallback.notes.append(f"回退视觉审稿时导出 PDF 失败: {type(pdf_exc).__name__}: {pdf_exc}")
            fallback.notes.insert(0, f"未启用真实大模型视觉审稿，已降级为本地版面感知: {type(exc).__name__}: {exc}")
            fallback.model = model
            return fallback

    if normalized_mode in {"openai"}:
        try:
            return review_document_with_openai(
                snapshot.path,
                output_dir=output_dir,
                model=model,
                stage=stage,
            )
        except Exception as exc:
            return VisualReviewResult(
                mode="unavailable",
                model=model,
                pdf_path=None,
                overall_verdict="minor_issues",
                visual_order_score=None,
                confidence=None,
                major_issues=[],
                minor_issues=[],
                evidence=[],
                page_observations=[],
                notes=[f"大模型视觉审稿未成功执行: {type(exc).__name__}: {exc}"],
            )

    fallback = _heuristic_visual_review(snapshot, model=model)
    if output_dir:
        try:
            fallback.pdf_path = str(export_document_to_pdf(snapshot.path, output_dir))
        except Exception as pdf_exc:
            fallback.notes.append(f"导出视觉审稿 PDF 失败: {type(pdf_exc).__name__}: {pdf_exc}")
    return fallback


def _heuristic_visual_review(snapshot: DocumentSnapshot, model: str | None = None) -> VisualReviewResult:
    metrics = _body_visual_metrics(snapshot)
    major_issues = _format_visual_red_flags(snapshot)
    minor_issues: list[str] = []
    evidence: list[str] = []

    if metrics["sample_count"] == 0:
        return VisualReviewResult(
            mode="heuristic",
            model=model,
            pdf_path=None,
            overall_verdict="rewrite",
            visual_order_score=1.0,
            confidence=0.45,
            major_issues=["正文主体未形成可供视觉评阅的版面。"],
            minor_issues=[],
            evidence=["未识别到稳定的正文页面结构。"],
            page_observations=[],
            notes=["当前为本地启发式视觉判断，不等同于大模型视觉结论。"],
        )

    evidence.append(
        f"正文主导字族为 {metrics['dominant_group_label']}，覆盖率 {metrics['dominant_group_ratio']:.2f}，正文视觉兼容率 {metrics['body_font_quality']:.2f}。"
    )
    evidence.append(
        f"正文空段占比 {metrics['blank_ratio']:.2f}，最大连续空段 {metrics['max_blank_run']}，段前段后异常占比 {metrics['spacing_outlier_ratio']:.2f}。"
    )
    evidence.append(
        f"正文对齐稳定率 {metrics['alignment_ratio']:.2f}，标题视觉规范率 {metrics['heading_visual_ratio']:.2f}。"
    )

    if metrics["dominant_group_ratio"] < 0.70:
        minor_issues.append("正文字体风格混用较多，整体不够像一套统一版式。")
    if metrics["indent_ratio"] < 0.75:
        minor_issues.append("正文首行缩进不够稳定。")
    if metrics["size_ratio"] < 0.80:
        minor_issues.append("正文字号稳定性一般。")
    if 0.08 <= metrics["blank_ratio"] < 0.16:
        minor_issues.append("正文留白略多，翻页观感偏松。")
    if 0.12 <= metrics["spacing_outlier_ratio"] < 0.30:
        minor_issues.append("部分段落留白不一致。")

    if len(major_issues) >= 2:
        verdict = "rewrite"
        visual_order_score = 3.0
    elif major_issues:
        verdict = "major_revision"
        visual_order_score = 5.0
    elif minor_issues:
        verdict = "minor_issues"
        visual_order_score = 7.0
    else:
        verdict = "pass"
        visual_order_score = 8.5

    return VisualReviewResult(
        mode="heuristic",
        model=model,
        pdf_path=None,
        overall_verdict=verdict,
        visual_order_score=visual_order_score,
        confidence=0.45,
        major_issues=major_issues,
        minor_issues=minor_issues,
        evidence=evidence,
        page_observations=[],
        notes=["当前为本地启发式视觉判断，不等同于大模型视觉结论。"],
    )


def _visual_score_adjustment(visual_review: VisualReviewResult) -> float:
    if visual_review.mode in {"skipped", "unavailable"}:
        return 0.0
    if visual_review.overall_verdict == "pass":
        return 0.0
    if visual_review.overall_verdict == "minor_issues":
        return -round(min(2.0, 0.35 * len(visual_review.minor_issues) + 0.3), 2)
    if visual_review.overall_verdict == "major_revision":
        return -round(min(4.0, 1.2 + 0.5 * len(visual_review.major_issues) + 0.2 * len(visual_review.minor_issues)), 2)
    if visual_review.overall_verdict == "rewrite":
        return -round(min(6.0, 2.5 + 0.6 * len(visual_review.major_issues)), 2)
    return 0.0


def _evaluate_visual_gate(visual_review: VisualReviewResult, stage: str) -> GateResult:
    if stage != "initial_draft":
        return GateResult("通过", False, None, [])

    if visual_review.mode in {"skipped", "unavailable", "heuristic"}:
        return GateResult("通过", False, None, [])

    if visual_review.overall_verdict == "rewrite":
        reasons = visual_review.major_issues or ["视觉审稿判断该稿件整体版面失序，暂不适合作为初稿进入正常评阅。"]
        return GateResult("视觉打回", True, 54.0, reasons)

    if visual_review.overall_verdict == "major_revision":
        reasons = visual_review.major_issues or ["视觉审稿发现明显排版失衡，需要先整稿。"]
        return GateResult("视觉退修", True, 59.0, reasons)

    return GateResult("通过", False, None, [])


def _combine_gates(*gates: GateResult) -> GateResult:
    priority = {
        "打回重写": 0,
        "视觉打回": 1,
        "格式退修": 2,
        "视觉退修": 3,
        "引用退修": 4,
        "引用待核": 5,
        "通过": 6,
    }
    selected = min(gates, key=lambda gate: priority.get(gate.decision, 99))
    reasons: list[str] = []
    for gate in gates:
        for reason in gate.reasons:
            if reason not in reasons:
                reasons.append(reason)
    score_caps = [gate.score_cap for gate in gates if gate.score_cap is not None]
    return GateResult(
        selected.decision,
        any(gate.rewrite_required for gate in gates),
        min(score_caps) if score_caps else None,
        reasons,
    )


def _evaluate_format_gate(snapshot: DocumentSnapshot, format_items: list[ScoreItem], stage: str) -> GateResult:
    if stage != "initial_draft":
        return GateResult("通过", False, None, [])

    item_map = {item.key: item for item in format_items}
    format_score = sum(item.score for item in format_items)
    reasons: list[str] = []
    critical_failures = 0
    severe_failures = 0

    for key, min_score in FORMAT_GATE_RULES["critical_item_mins"].items():
        item = item_map[key]
        if item.score < min_score:
            critical_failures += 1
            reasons.append(f"{item.name} 仅 {item.score}/{item.max_score}，未达到初稿最低门槛。")

    for key, min_score in FORMAT_GATE_RULES["severe_item_mins"].items():
        item = item_map[key]
        if item.score < min_score:
            severe_failures += 1
            reasons.append(f"{item.name} 明显失范，当前状态不适合继续细评内容。")

    if not _body_paragraphs(snapshot):
        severe_failures += 1
        reasons.append("正文主体没有被正常识别，文档仍停留在模板或半成品状态。")

    if _placeholder_count(snapshot.raw_text) >= 5:
        severe_failures += 1
        reasons.append("文档存在较多模板占位符，说明基础清稿尚未完成。")

    missing_sections = [name for name in MANDATORY_SECTIONS if not _has_section(snapshot, name)]
    if missing_sections:
        critical_failures += 1
        reasons.append(f"必备部分缺失：{', '.join(missing_sections)}。")

    if severe_failures > 0 or format_score < FORMAT_GATE_RULES["rewrite_threshold"]:
        if not reasons:
            reasons.append("初稿格式整体失范，先打回处理格式。")
        return GateResult("打回重写", True, FORMAT_GATE_RULES["cap_for_rewrite"], reasons)

    if critical_failures > 0 or format_score < FORMAT_GATE_RULES["revision_threshold"]:
        if not reasons:
            reasons.append("初稿格式尚未过线，必须先完成格式整改。")
        return GateResult("格式退修", True, FORMAT_GATE_RULES["cap_for_revision"], reasons)

    return GateResult("通过", False, None, [])


def _grade_band(total_score: float) -> dict:
    for band in GRADE_BANDS:
        if total_score >= band.min_score:
            return {"label": band.label, "description": band.description}
    last = GRADE_BANDS[-1]
    return {"label": last.label, "description": last.description}


def _status(score: float, max_score: float) -> str:
    ratio = score / max(max_score, 1)
    if ratio >= 0.90:
        return "优秀"
    if ratio >= 0.75:
        return "良好"
    if ratio >= 0.55:
        return "一般"
    return "未通过"


def _visual_mode_label(mode: str | None) -> str:
    return {
        "openai": "大模型视觉审稿",
        "moonshot": "Moonshot/Kimi 视觉审稿",
        "siliconflow": "SiliconFlow 视觉审稿",
        "expert": "双专家视觉审稿",
        "heuristic": "本地启发式视觉回退",
        "skipped": "未启用视觉审稿",
        "unavailable": "视觉审稿不可用",
    }.get(mode or "", "未启用视觉审稿")


def _visual_verdict_label(verdict: str | None) -> str:
    return {
        "pass": "视觉通过",
        "minor_issues": "整体可接受，但有小毛病",
        "major_revision": "存在明显视觉失衡，需要退修",
        "rewrite": "整体观感失范，建议打回重整",
    }.get(verdict or "", "未判定")


def _is_close(value: float | None, target: float, tolerance: float) -> bool:
    return value is not None and abs(value - target) <= tolerance


def _soft_visual_match(diff: float, soft_limit: float, medium_limit: float, hard_limit: float) -> float:
    if diff <= soft_limit:
        return 1.0
    if diff <= medium_limit:
        return 0.8
    if diff <= hard_limit:
        return 0.45
    return 0.0


def _normalize_font_name(font_name: str) -> str:
    return re.sub(r"[\s_]+", "", (font_name or "").strip().lower())


def _font_visual_group(font_name: str) -> str:
    normalized = _normalize_font_name(font_name)
    if not normalized:
        return "unknown"
    if any(token in normalized for token in ("仿宋", "fangsong")):
        return "fangsong"
    if any(token in normalized for token in BODY_KAI_HINTS):
        return "kai"
    if any(token in normalized for token in BODY_HEAVY_HINTS):
        return "sans"
    if any(token in normalized for token in BODY_SERIF_HINTS):
        return "serif"
    return "other"


def _font_group_label(group: str) -> str:
    return {
        "fangsong": "仿宋/正文宋体类",
        "serif": "宋体类",
        "kai": "楷体类",
        "sans": "黑体/无衬线类",
        "other": "其他字体",
        "unknown": "未识别",
    }.get(group, "未识别")


def _body_font_visual_score(font_name: str) -> float:
    group = _font_visual_group(font_name)
    if group == "fangsong":
        return 1.0
    if group == "serif":
        return 0.92
    if group == "kai":
        return 0.72
    if group == "other":
        return 0.55
    if group == "sans":
        return 0.18
    return 0.50


def _heading_font_visual_score(paragraph: ParagraphSnapshot) -> float:
    normalized = _normalize_font_name(paragraph.font_name)
    if any(token in normalized for token in HEADING_HEAVY_HINTS):
        score = 1.0
    else:
        group = _font_visual_group(paragraph.font_name)
        if group == "sans":
            score = 0.9
        elif group in {"serif", "fangsong", "kai"}:
            score = 0.45
        else:
            score = 0.55
    if paragraph.bold and paragraph.bold > 0:
        score = min(score + 0.1, 1.0)
    return score


def _bucket_spacing(value: float | None) -> float | None:
    if value is None:
        return None
    return round(value / 2.0) * 2.0


def _dominant_ratio(values: list) -> tuple[object | None, float]:
    if not values:
        return None, 0.0
    counter = Counter(values)
    dominant, count = counter.most_common(1)[0]
    return dominant, round(count / len(values), 3)


def _find_paragraph(snapshot: DocumentSnapshot, predicate) -> ParagraphSnapshot | None:
    for paragraph in snapshot.non_empty_paragraphs:
        if predicate(paragraph):
            return paragraph
    return None


def _cover_paragraphs(snapshot: DocumentSnapshot) -> list[ParagraphSnapshot]:
    paragraphs: list[ParagraphSnapshot] = []
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "摘要":
            break
        paragraphs.append(paragraph)
    return paragraphs[:25]


def _extract_cover_value(snapshot: DocumentSnapshot, label: str) -> str:
    pattern = _cover_label_pattern(label)
    cover_paragraphs = _cover_paragraphs(snapshot)
    for index, paragraph in enumerate(cover_paragraphs):
        match = pattern.search(paragraph.text)
        if not match:
            continue
        value = match.group(1).strip(" ：:")
        value = re.sub(r"^[·.。]+", "", value)
        if value and value != label:
            return value
        if index + 1 < len(cover_paragraphs):
            next_text = cover_paragraphs[index + 1].text.strip()
            if next_text and not any(next_text.startswith(field) for field in COVER_REQUIRED_LABELS):
                return next_text
    return ""


def _find_keyword_paragraph(snapshot: DocumentSnapshot) -> ParagraphSnapshot | None:
    return _find_paragraph(snapshot, lambda paragraph: paragraph.text.startswith("关键词"))


def _extract_keywords(snapshot: DocumentSnapshot) -> list[str]:
    paragraph = _find_keyword_paragraph(snapshot)
    if not paragraph:
        return []
    content = paragraph.text.split("：", 1)[-1]
    content = content.split(":", 1)[-1]
    parts = re.split(r"[；;、，,]\s*", content)
    return [part.strip() for part in parts if part.strip()]


def _extract_abstract_text(snapshot: DocumentSnapshot) -> str:
    start = None
    end = None
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "摘要":
            start = paragraph.index
            continue
        if start and (paragraph.text.startswith("关键词") or paragraph.normalized == "目录"):
            end = paragraph.index
            break
    if start is None:
        return ""
    pieces = [
        paragraph.text
        for paragraph in snapshot.paragraphs
        if paragraph.index > start and (end is None or paragraph.index < end) and paragraph.text
    ]
    return "\n".join(pieces)


def _toc_entries(snapshot: DocumentSnapshot) -> list[str]:
    entries: list[str] = []
    in_toc = False
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "目录":
            in_toc = True
            continue
        if not in_toc:
            continue
        if re.search(r"\t\d+$|…+\d+$|\.{3,}\d+$", paragraph.text) or re.match(r"^第[一二三四五六七八九十0-9]+章", paragraph.text):
            entries.append(paragraph.text)
            continue
        if entries and re.match(r"^第[一二三四五六七八九十0-9]+章", paragraph.text):
            break
    return entries


def _has_section(snapshot: DocumentSnapshot, section_name: str) -> bool:
    target = re.sub(r"\s+", "", section_name)
    return any(paragraph.normalized == target for paragraph in snapshot.non_empty_paragraphs)


def _chapter_paragraphs(snapshot: DocumentSnapshot) -> list[ParagraphSnapshot]:
    candidates = _find_heading_matches(snapshot, r"^第[一二三四五六七八九十0-9]+章")
    return [paragraph for paragraph in candidates if not _looks_like_toc_entry(paragraph.text)]


def _find_heading_matches(snapshot: DocumentSnapshot, pattern: str) -> list[ParagraphSnapshot]:
    regex = re.compile(pattern)
    return [paragraph for paragraph in snapshot.non_empty_paragraphs if regex.search(paragraph.text)]


def _chapter_blocks(snapshot: DocumentSnapshot) -> list[ChapterBlock]:
    chapters = _chapter_paragraphs(snapshot)
    if not chapters:
        return []

    reference_start = None
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "参考文献":
            reference_start = paragraph.index
            break

    blocks: list[ChapterBlock] = []
    terminal_index = snapshot.paragraphs[-1].index + 1 if snapshot.paragraphs else 1
    for index, chapter in enumerate(chapters):
        next_index = chapters[index + 1].index if index + 1 < len(chapters) else reference_start or terminal_index
        paragraphs = [
            paragraph
            for paragraph in snapshot.paragraphs
            if chapter.index < paragraph.index < next_index and paragraph.text
        ]
        blocks.append(ChapterBlock(chapter, paragraphs, "\n".join(paragraph.text for paragraph in paragraphs)))
    return blocks


def _body_paragraphs(snapshot: DocumentSnapshot) -> list[ParagraphSnapshot]:
    paragraphs: list[ParagraphSnapshot] = []
    for block in _chapter_blocks(snapshot):
        paragraphs.extend(block.paragraphs)
    return paragraphs


def _body_region_paragraphs(snapshot: DocumentSnapshot) -> list[ParagraphSnapshot]:
    chapters = _chapter_paragraphs(snapshot)
    if not chapters:
        return []

    reference_start = None
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "参考文献":
            reference_start = paragraph.index
            break

    start_index = chapters[0].index
    end_index = reference_start or (snapshot.paragraphs[-1].index + 1 if snapshot.paragraphs else 1)
    return [paragraph for paragraph in snapshot.paragraphs if start_index < paragraph.index < end_index]


def _body_visual_metrics(snapshot: DocumentSnapshot) -> dict:
    body_paragraphs = _body_paragraphs(snapshot)
    chapters = _chapter_paragraphs(snapshot)
    body_region = _body_region_paragraphs(snapshot)
    samples = [paragraph for paragraph in body_paragraphs if len(paragraph.normalized) >= 20][:30]
    if not samples:
        return {
            "sample_count": 0,
            "dominant_group_label": "未识别",
            "dominant_group_ratio": 0.0,
            "body_font_quality": 0.0,
            "hostile_body_ratio": 0.0,
            "size_hits": 0,
            "size_ratio": 0.0,
            "indent_hits": 0,
            "indent_ratio": 0.0,
            "alignment_ratio": 0.0,
            "spacing_consistency": 0.0,
            "blank_paragraphs": 0,
            "blank_ratio": 0.0,
            "max_blank_run": 0,
            "spacing_outlier_ratio": 0.0,
            "chapter_count": len(chapters),
            "heading_visual_ratio": 0.0,
            "rhythm_cleanliness": 0.0,
        }

    body_groups = [_font_visual_group(paragraph.font_name) for paragraph in samples]
    dominant_group, dominant_group_ratio = _dominant_ratio(body_groups)
    body_font_scores = [_body_font_visual_score(paragraph.font_name) for paragraph in samples]
    body_font_quality = sum(body_font_scores) / len(body_font_scores)
    hostile_body_ratio = sum(1 for score in body_font_scores if score <= 0.3) / len(body_font_scores)

    size_hits = sum(1 for paragraph in samples if _is_close(paragraph.font_size, 10.5, 0.8))
    indent_hits = sum(
        1
        for paragraph in samples
        if paragraph.first_line_indent_pt is not None and abs(paragraph.first_line_indent_pt - 21.0) <= 4.0
    )

    alignment_values = [paragraph.alignment for paragraph in samples if paragraph.alignment is not None]
    alignment_ratio = (
        sum(1 for alignment in alignment_values if alignment in {0, 3}) / len(alignment_values)
        if alignment_values
        else 0.8
    )

    spacing_values = [_bucket_spacing(paragraph.line_spacing_pt) for paragraph in samples if paragraph.line_spacing_pt is not None]
    _, line_spacing_ratio = _dominant_ratio(spacing_values)
    if not spacing_values:
        line_spacing_ratio = 0.85

    spacing_outliers = sum(
        1
        for paragraph in samples
        if (paragraph.space_before_pt or 0.0) > 6.0 or (paragraph.space_after_pt or 0.0) > 6.0
    )
    spacing_outlier_ratio = spacing_outliers / len(samples)
    spacing_consistency = 0.55 * line_spacing_ratio + 0.45 * (1.0 - min(spacing_outlier_ratio / 0.35, 1.0))

    blank_paragraphs = sum(1 for paragraph in body_region if not paragraph.normalized)
    visible_body_lines = sum(1 for paragraph in body_region if paragraph.normalized)
    blank_ratio = blank_paragraphs / max(visible_body_lines, 1)
    max_blank_run = 0
    current_blank_run = 0
    for paragraph in body_region:
        if paragraph.normalized:
            max_blank_run = max(max_blank_run, current_blank_run)
            current_blank_run = 0
        else:
            current_blank_run += 1
    max_blank_run = max(max_blank_run, current_blank_run)
    rhythm_cleanliness = 0.65 * (1.0 - min(blank_ratio / 0.20, 1.0)) + 0.35 * (1.0 - min(spacing_outlier_ratio / 0.35, 1.0))

    heading_scores: list[float] = []
    for paragraph in chapters:
        score = 0.0
        score += 0.40 if paragraph.alignment == 1 else 0.0
        score += 0.25 if _is_close(paragraph.font_size, 12.0, 0.8) else 0.0
        score += 0.35 * _heading_font_visual_score(paragraph)
        heading_scores.append(score)
    heading_visual_ratio = sum(heading_scores) / len(heading_scores) if heading_scores else 0.0

    return {
        "sample_count": len(samples),
        "dominant_group_label": _font_group_label(dominant_group),
        "dominant_group_ratio": dominant_group_ratio,
        "body_font_quality": round(body_font_quality, 3),
        "hostile_body_ratio": round(hostile_body_ratio, 3),
        "size_hits": size_hits,
        "size_ratio": round(size_hits / len(samples), 3),
        "indent_hits": indent_hits,
        "indent_ratio": round(indent_hits / len(samples), 3),
        "alignment_ratio": round(alignment_ratio, 3),
        "spacing_consistency": round(spacing_consistency, 3),
        "blank_paragraphs": blank_paragraphs,
        "blank_ratio": round(blank_ratio, 3),
        "max_blank_run": max_blank_run,
        "spacing_outlier_ratio": round(spacing_outlier_ratio, 3),
        "chapter_count": len(chapters),
        "heading_visual_ratio": round(heading_visual_ratio, 3),
        "rhythm_cleanliness": round(rhythm_cleanliness, 3),
    }


def _format_visual_red_flags(snapshot: DocumentSnapshot) -> list[str]:
    metrics = _body_visual_metrics(snapshot)
    if metrics["sample_count"] == 0:
        return []

    reasons: list[str] = []
    if metrics["blank_ratio"] >= 0.16 or metrics["max_blank_run"] >= 2:
        reasons.append("正文空段或多余回车过多，版面节奏明显紊乱。")
    if metrics["spacing_outlier_ratio"] >= 0.30:
        reasons.append("正文段前段后留白明显失控，翻阅时会产生强烈跳行感。")
    if metrics["hostile_body_ratio"] >= 0.25:
        reasons.append("正文大量使用黑体等重字族，整体观感明显偏离规范论文。")
    if metrics["dominant_group_ratio"] < 0.55:
        reasons.append("正文混用多种字体风格，整体视觉不统一。")
    return reasons


def _extract_body_text(snapshot: DocumentSnapshot) -> str:
    return "\n".join(paragraph.text for paragraph in _body_paragraphs(snapshot))


def _extract_conclusion_text(snapshot: DocumentSnapshot, chapter_blocks: list[ChapterBlock] | None = None) -> str:
    blocks = chapter_blocks if chapter_blocks is not None else _chapter_blocks(snapshot)
    if not blocks:
        return ""
    for block in reversed(blocks):
        if _contains_any(block.heading.text, CONCLUSION_CUES):
            return block.text
    return blocks[-1].text


def _extract_reference_entries(snapshot: DocumentSnapshot) -> list[str]:
    start = None
    for paragraph in snapshot.non_empty_paragraphs:
        if paragraph.normalized == "参考文献":
            start = paragraph.index
            break
    if start is None:
        return []

    entries: list[str] = []
    for paragraph in snapshot.paragraphs:
        if paragraph.index <= start or not paragraph.text:
            continue
        if paragraph.normalized in {"附录", "致谢"}:
            break
        entries.append(paragraph.text)
    return entries


def _split_sentences(text: str) -> list[str]:
    return [segment.strip() for segment in re.split(r"[。！？!?]", text) if segment.strip()]


def _sentence_repetition_stats(text: str) -> tuple[float, int, int]:
    normalized_sentences = [
        _normalize_for_similarity(sentence)
        for sentence in _split_sentences(text)
        if len(_normalize_for_similarity(sentence)) >= 10
    ]
    if not normalized_sentences:
        return 0.0, 0, 0

    counts = Counter(normalized_sentences)
    repeated_sentences = sum(count - 1 for count in counts.values() if count > 1)
    repetition_ratio = round(repeated_sentences / len(normalized_sentences), 3)
    return repetition_ratio, repeated_sentences, len(normalized_sentences)


def _concept_terms(text: str) -> list[str]:
    if not text:
        return []

    working = text
    for connector in ["基于", "关于", "针对", "面向", "围绕", "以及", "与", "和", "及", "对", "的", "在"]:
        working = working.replace(connector, " ")
    for suffix in ["研究", "分析", "应用", "设计", "实现", "探讨", "浅析", "对策", "问题"]:
        working = working.replace(suffix, " ")

    raw_terms = re.split(r"[^\u4e00-\u9fffA-Za-z0-9]+", working)
    terms: list[str] = []
    for term in raw_terms:
        cleaned = term.strip()
        if len(cleaned) < 2 or cleaned in GENERIC_TERMS:
            continue
        if cleaned not in terms:
            terms.append(cleaned)
    return terms


def _core_terms(snapshot: DocumentSnapshot) -> list[str]:
    title = _extract_cover_value(snapshot, "题目")
    keywords = _extract_keywords(snapshot)
    terms: list[str] = []
    for source in [title, *keywords]:
        candidates = _concept_terms(source) or ([source.strip()] if source.strip() else [])
        for term in candidates:
            normalized = term.strip()
            if len(normalized) < 2 or normalized in GENERIC_TERMS:
                continue
            if normalized not in terms:
                terms.append(normalized)
    return terms[:8]


def _term_coverage_ratio(terms: list[str], text: str) -> float:
    if not terms:
        return 0.0
    hits = sum(1 for term in terms if term and term in text)
    return round(hits / len(terms), 3)


def _chapter_progression_score(chapter_blocks: list[ChapterBlock]) -> tuple[float, list[str], list[str]]:
    if not chapter_blocks:
        return 0.0, ["未识别到可分析的章节路径。"], ["正文还没有形成可评阅的章节推进链。"]

    evidence: list[str] = []
    suggestions: list[str] = []
    score = 0.0

    first_heading = chapter_blocks[0].heading.text
    last_heading = chapter_blocks[-1].heading.text
    if _contains_any(first_heading, INTRO_CUES + OVERVIEW_CUES):
        score += 1.0
        evidence.append("首章具有引入功能。")
    else:
        suggestions.append("首章没有明显承担开题、交代背景的作用。")

    if _contains_any(last_heading, CONCLUSION_CUES):
        score += 1.0
        evidence.append("末章具有收束功能。")
    else:
        suggestions.append("末章没有明确承担结论或展望功能。")

    middle_tags: set[str] = set()
    for block in chapter_blocks[1:-1]:
        heading = block.heading.text
        if _contains_any(heading, OVERVIEW_CUES):
            middle_tags.add("overview")
        if _contains_any(heading, METHOD_CUES):
            middle_tags.add("method")
        if _contains_any(heading, APPLICATION_CUES):
            middle_tags.add("application")
        if _contains_any(heading, PROBLEM_CUES):
            middle_tags.add("problem")
        if _contains_any(heading, SOLUTION_CUES):
            middle_tags.add("solution")

    evidence.append(f"中间章节功能标签: {sorted(middle_tags) if middle_tags else '未识别'}。")
    if len(middle_tags) >= 3:
        score += 2.0
    elif len(middle_tags) >= 2:
        score += 1.0
        suggestions.append("中间章节功能有一些区分，但论证路线还不够完整。")
    else:
        suggestions.append("中间章节几乎没有形成“分析-问题-对策”链条。")

    return score, evidence, suggestions


def _contains_any(text: str, cues: list[str]) -> bool:
    return any(cue in text for cue in cues)


def _chinese_char_count(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def _tokenize_keywords(text: str) -> list[str]:
    cleaned = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]+", " ", text or "")
    parts = [part.strip() for part in cleaned.split() if part.strip()]
    return [part for part in parts if len(part) >= 2]


def _normalize_for_similarity(text: str) -> str:
    return re.sub(r"[\W_]+", "", text.lower())


def _shingles(text: str, size: int = 8) -> set[str]:
    normalized = _normalize_for_similarity(text)
    if not normalized:
        return set()
    if len(normalized) < size:
        return {normalized}
    return {normalized[index : index + size] for index in range(len(normalized) - size + 1)}


def _max_similarity(snapshot: DocumentSnapshot, references: Iterable[DocumentSnapshot]) -> float | None:
    source = _shingles(_extract_body_text(snapshot) or snapshot.raw_text)
    if not source:
        return None
    scores: list[float] = []
    for reference in references:
        target = _shingles(_extract_body_text(reference) or reference.raw_text)
        if not target:
            continue
        union = len(source | target)
        if union == 0:
            continue
        scores.append(len(source & target) / union)
    return max(scores) if scores else None


def _cover_label_pattern(label: str) -> re.Pattern[str]:
    pieces = [re.escape(char) for char in label]
    joined = r"\s*".join(pieces)
    return re.compile(rf"{joined}\s*[:：]?\s*(.+)")


def _looks_like_toc_entry(text: str) -> bool:
    if "\t" in text:
        return True
    if re.search(r"[.…]{2,}\s*\d+\s*$", text):
        return True
    if re.search(r"\s+\d+\s*$", text) and re.match(r"^第[一二三四五六七八九十0-9]+章", text):
        return True
    return False


def _placeholder_count(text: str) -> int:
    return len(PLACEHOLDER_RE.findall(text or ""))
