from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ScoreBand:
    min_score: float
    label: str
    description: str


RUBRIC_NAME = "黑龙江省经济管理干部学院专科毕业论文评分系统"

EXPECTED_PAGE_SETUP_CM = {
    "top": 2.54,
    "bottom": 2.54,
    "left": 3.17,
    "right": 3.17,
    "header_distance": 1.50,
    "footer_distance": 1.75,
}

PAGE_SETUP_TOLERANCE_CM = 0.20

HEADER_VARIANTS = [
    "黑龙江省经济管理干部学院毕业论文（设计）",
    "黑龙江省经济管理干部学院毕业设计",
]

COVER_REQUIRED_LABELS = [
    "题目",
    "专业",
    "姓名",
    "指导教师",
    "完成日期",
]

MANDATORY_SECTIONS = [
    "摘要",
    "目录",
    "参考文献",
]

RECOMMENDED_SECTIONS = [
    "致谢",
    "附录",
]

FORMAT_WEIGHTS = {
    "page_setup": 8,
    "cover": 4,
    "header_pagination": 5,
    "abstract_keywords_format": 6,
    "toc_and_sections": 5,
    "body_typography": 7,
}

CONTENT_WEIGHTS = {
    "topic_relevance": 6,
    "structure_logic": 10,
    "abstract_quality": 7,
    "chapter_development": 10,
    "coherence_alignment": 18,
    "references_support": 5,
    "language_quality": 4,
    "academic_integrity": 5,
}

GRADE_BANDS = [
    ScoreBand(90, "优秀", "选题成熟，论证严密，结构完整，格式规范。"),
    ScoreBand(80, "良好", "论证较完整，结构较清晰，格式基本规范。"),
    ScoreBand(70, "中等", "有基本结构与内容，但论证和规范性仍有明显欠缺。"),
    ScoreBand(60, "及格", "勉强达到基本要求，仍需较多修改。"),
    ScoreBand(0, "不及格", "内容或格式未达到学校论文的基本要求。"),
]

STRICT_ABSTRACT_RANGE = (120, 220)
SOFT_ABSTRACT_RANGE = (100, 260)
KEYWORD_RANGE = (3, 5)
TITLE_MAX_CHARS = 20
MIN_BODY_CHARS = 3000
MIN_REFERENCE_COUNT = 5

OFFICIAL_FORMAT_REQUIREMENTS = {
    "paper_and_layout": {
        "paper_size": "A4",
        "margins_cm": EXPECTED_PAGE_SETUP_CM,
        "header_text": HEADER_VARIANTS[0],
        "header_font": "小四黑体",
        "header_paging": "页眉右端，正文使用“第M页”阿拉伯数字，摘要/目录前置部分使用罗马数字",
    },
    "title_abstract_keywords": {
        "title_font": "宋体三号",
        "title_max_chars": TITLE_MAX_CHARS,
        "abstract_font": "五号楷体",
        "abstract_target_chars": 150,
        "abstract_strict_range": STRICT_ABSTRACT_RANGE,
        "keyword_font": "五号楷体",
        "keyword_range": KEYWORD_RANGE,
    },
    "body_and_structure": {
        "heading_font": "小四黑体",
        "body_font": "五号宋体",
        "body_heading_levels": "建议三级标题：一、（二）1.",
        "line_break_rule": "章标题之间、节标题与正文之间空一个标准行",
        "mandatory_sections": MANDATORY_SECTIONS,
        "recommended_sections": RECOMMENDED_SECTIONS,
        "binding_order": [
            "封面",
            "摘要",
            "目录",
            "正文",
            "致谢",
            "参考文献",
            "附录",
        ],
    },
    "citation_and_figures": {
        "citation_style": "正文引用右上角标注，阿拉伯数字方括号 [1][2]...",
        "figure_table_rule": "图题置图下、表题置表上，按章编号（图2.1、表2.1），小五号统一字体",
        "equation_rule": "公式可按章编号（式2.1），公式居中，式号右侧",
    },
}

OFFICIAL_WRITING_REQUIREMENTS = {
    "hard_constraints": {
        "min_body_chars": MIN_BODY_CHARS,
        "title_max_chars": TITLE_MAX_CHARS,
        "keyword_range": KEYWORD_RANGE,
        "abstract_target_chars": 150,
        "abstract_strict_range": STRICT_ABSTRACT_RANGE,
        "min_reference_count": MIN_REFERENCE_COUNT,
    },
    "core_expectations": {
        "topic_alignment": "题目必须与专业相关，且标题与正文主线严格对应",
        "argumentation": "正文应围绕研究问题展开分析，避免口号化和车轱辘话",
        "structure": "应形成问题-分析-结论链条，章节层次清楚",
        "language": "文字简练通顺，概念使用准确，避免堆砌空泛语句",
        "academic_integrity": "引用可追溯、参考文献真实，禁止抄袭拼接",
    },
    "scoring_band_reference": {
        "excellent": "90-100，观点明确、论证严密、规范完整",
        "good": "80-89，论证较完整，表达较准确",
        "average": "70-79，结构与论证基本成立但存在明显欠缺",
        "pass": "60-69，达到最低要求但问题较多",
        "fail": "<60，内容空泛/结构混乱/规范不足或存在学术不端",
    },
}

FUSION_POLICY = {
    "final_score": {
        "format_score_max": float(sum(FORMAT_WEIGHTS.values())),
        "writing_score_max": float(sum(CONTENT_WEIGHTS.values())),
        "final_total_max": 100.0,
        "formula": "final_score = format_score + writing_score",
    },
    "format_fusion": {
        "rule_weight_range": (0.70, 0.88),
        "visual_weight_range": (0.12, 0.30),
        "heuristic_visual_weight": 0.08,
        "unavailable_visual_weight": 0.0,
        "major_revision_cap_ratio": 0.82,
        "rewrite_cap_ratio": 0.68,
    },
    "writing_fusion": {
        "rule_weight_range": (0.65, 0.90),
        "expert_weight_range": (0.10, 0.35),
        "single_expert_weight_range": (0.10, 0.22),
        "dual_expert_weight_range": (0.14, 0.35),
        "positive_adjust_cap": 4.0,
        "low_base_positive_adjust_cap": 2.5,
        "very_low_base_positive_adjust_cap": 1.0,
        "structure_lock": "body_chars < 1200 or chapter_count < 2 => positive_adjust_cap = 0.0",
    },
}

FORMAT_GATE_RULES = {
    "revision_threshold": 26.0,
    "rewrite_threshold": 22.0,
    "cap_for_revision": 59.0,
    "cap_for_rewrite": 54.0,
    "critical_item_mins": {
        "page_setup": 6.0,
        "cover": 2.0,
        "header_pagination": 2.0,
        "toc_and_sections": 3.0,
        "body_typography": 4.0,
    },
    "severe_item_mins": {
        "page_setup": 5.0,
        "toc_and_sections": 2.0,
        "body_typography": 2.0,
    },
}
