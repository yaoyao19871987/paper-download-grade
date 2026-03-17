from __future__ import annotations

from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from paper_grader.grader import (
    ScoreItem,
    _chapter_paragraphs,
    _evaluate_format_gate,
    _evaluate_reference_gate,
    _extract_cover_value,
    _format_visual_red_flags,
    _score_body_typography,
    _sentence_repetition_stats,
)
from paper_grader.reference_verifier import (
    ReferenceAuditResult,
    ReferenceCheck,
    ReferenceEntry,
    _is_plausible_reference_title,
    _is_possible_title_match,
    _is_strict_title_match,
    _title_similarity,
    audit_references,
    verify_reference_entry,
)
from paper_grader.word_inspector import DocumentSnapshot, ParagraphSnapshot


def make_paragraph(index: int, text: str, **overrides) -> ParagraphSnapshot:
    values = {
        "index": index,
        "text": text,
        "normalized": "".join(text.split()),
        "style_name": "正文",
        "font_name": "宋体",
        "font_size": 10.5,
        "bold": 0,
        "alignment": 3,
        "first_line_indent_pt": 21.0,
        "line_spacing_rule": 0,
        "line_spacing_pt": 20.0,
        "space_before_pt": 0.0,
        "space_after_pt": 0.0,
    }
    values.update(overrides)
    return ParagraphSnapshot(
        index=values["index"],
        text=values["text"],
        normalized=values["normalized"],
        style_name=values["style_name"],
        font_name=values["font_name"],
        font_size=values["font_size"],
        bold=values["bold"],
        alignment=values["alignment"],
        first_line_indent_pt=values["first_line_indent_pt"],
        line_spacing_rule=values["line_spacing_rule"],
        line_spacing_pt=values["line_spacing_pt"],
        space_before_pt=values["space_before_pt"],
        space_after_pt=values["space_after_pt"],
    )


class GraderHelpersTest(unittest.TestCase):
    def test_extract_cover_value_accepts_spaced_label(self) -> None:
        snapshot = DocumentSnapshot(
            path="demo.docx",
            paragraphs=[
                make_paragraph(1, "黑龙江省经济管理干部学院"),
                make_paragraph(2, "题    目 大数据技术在电商领域的应用研究"),
                make_paragraph(3, "专    业      大数据技术"),
                make_paragraph(4, "摘  要"),
            ],
        )

        self.assertEqual(_extract_cover_value(snapshot, "题目"), "大数据技术在电商领域的应用研究")
        self.assertEqual(_extract_cover_value(snapshot, "专业"), "大数据技术")

    def test_chapter_paragraphs_skip_toc_entries(self) -> None:
        snapshot = DocumentSnapshot(
            path="demo.docx",
            paragraphs=[
                make_paragraph(1, "目  录"),
                make_paragraph(2, "第一章 绪论\t1"),
                make_paragraph(3, "第二章 方法\t2"),
                make_paragraph(4, "第一章 绪论"),
                make_paragraph(5, "第二章 方法"),
                make_paragraph(6, "参考文献"),
            ],
        )

        chapters = _chapter_paragraphs(snapshot)
        self.assertEqual([paragraph.text for paragraph in chapters], ["第一章 绪论", "第二章 方法"])

    def test_format_gate_rejects_template_like_snapshot(self) -> None:
        snapshot = DocumentSnapshot(
            path="demo.docx",
            paragraphs=[
                make_paragraph(1, "黑龙江省经济管理干部学院"),
                make_paragraph(2, "×××"),
                make_paragraph(3, "摘要"),
            ],
        )
        format_items = [
            ScoreItem("page_setup", "页面设置", 4.0, 8.0, "未通过", [], []),
            ScoreItem("cover", "封面完整度", 1.0, 4.0, "未通过", [], []),
            ScoreItem("header_pagination", "页眉与页码", 1.0, 5.0, "未通过", [], []),
            ScoreItem("abstract_keywords_format", "摘要与关键词格式", 1.0, 6.0, "未通过", [], []),
            ScoreItem("toc_and_sections", "目录与结构完整性", 1.0, 5.0, "未通过", [], []),
            ScoreItem("body_typography", "正文字体与标题", 0.0, 7.0, "未通过", [], []),
        ]

        gate = _evaluate_format_gate(snapshot, format_items, "initial_draft")
        self.assertEqual(gate.decision, "打回重写")
        self.assertTrue(gate.rewrite_required)

    def test_reference_audit_flags_dangling_and_uncited_entries(self) -> None:
        audit = audit_references(
            [
                "[1] 张三. 电子商务发展研究[J]. 2024.",
                "[2] 李四. 数字经济分析[J]. 2023.",
            ],
            "正文先引用[1]，后面又出现了[3]。",
            online=False,
        )

        self.assertEqual(audit.dangling_citations, [3])
        self.assertEqual(audit.uncited_reference_indexes, [2])
        self.assertFalse(audit.citation_mapping_ok)
        self.assertTrue(audit.checks[0].cited_in_body)
        self.assertFalse(audit.checks[1].cited_in_body)

    def test_reference_gate_rejects_fabricated_cited_reference(self) -> None:
        audit = ReferenceAuditResult(
            entries=[
                ReferenceEntry(1, "[1] 张三. 电子商务发展研究[J]. 2024.", "电子商务发展研究", "张三", "2024"),
                ReferenceEntry(2, "[2] 李四. 数字经济分析[J]. 2023.", "数字经济分析", "李四", "2023"),
            ],
            checks=[
                ReferenceCheck(1, "[1] 张三. 电子商务发展研究[J]. 2024.", "电子商务发展研究", True, "not_found", 0.0, None, None, None, ["未找到"]),
                ReferenceCheck(2, "[2] 李四. 数字经济分析[J]. 2023.", "数字经济分析", False, "verified", 0.95, "crossref", "数字经济分析", "https://example.com", []),
            ],
            citation_numbers=[1],
            dangling_citations=[],
            uncited_reference_indexes=[2],
            verified_count=1,
            possible_count=0,
            not_found_count=1,
            search_error_count=0,
            citation_mapping_ok=False,
            notes=[],
        )

        gate = _evaluate_reference_gate(audit)
        self.assertEqual(gate.decision, "引用退修")
        self.assertTrue(gate.rewrite_required)

    def test_reference_title_match_rejects_partial_keyword_overlap(self) -> None:
        expected = "大数据在企业数字化转型中的应用研究"
        candidate = "人工智能在企业数字化转型的应用研究"

        self.assertLess(_title_similarity(expected, candidate), 0.82)
        self.assertFalse(_is_strict_title_match(expected, candidate))
        self.assertFalse(_is_possible_title_match(expected, candidate))

    def test_reference_title_match_accepts_normalized_exact_title(self) -> None:
        expected = "基于Python的校园图书管理系统设计与实现"
        candidate = "基于 python 的校园图书管理系统设计与实现"

        self.assertGreaterEqual(_title_similarity(expected, candidate), 0.98)
        self.assertTrue(_is_strict_title_match(expected, candidate))

    def test_reference_title_plausibility_rejects_garbage_fragments(self) -> None:
        self.assertFalse(_is_plausible_reference_title("49."))
        self.assertFalse(_is_plausible_reference_title("100-118."))
        self.assertFalse(_is_plausible_reference_title("2021(20):67-70"))
        self.assertFalse(_is_plausible_reference_title("张虹;李笑"))
        self.assertTrue(_is_plausible_reference_title("大数据技术在电商企业财务管理中的应用探讨"))

    def test_reference_verifier_rejects_implausible_title_before_search(self) -> None:
        entry = ReferenceEntry(1, "[1] 49.", "49.", "", "2024")
        with patch("paper_grader.reference_verifier._search_crossref") as mock_crossref, patch(
            "paper_grader.reference_verifier._search_bing_site"
        ) as mock_bing:
            check = verify_reference_entry(entry, cited_in_body=False, session=requests.Session())

        self.assertEqual(check.status, "not_found")
        self.assertIn("解析结果异常", " ".join(check.notes))
        mock_crossref.assert_not_called()
        mock_bing.assert_not_called()

    def test_sentence_repetition_stats_detects_duplicate_sentences(self) -> None:
        ratio, repeated, total = _sentence_repetition_stats(
            "数字经济正在深刻改变商业生态。数字经济正在深刻改变商业生态。"
            "平台治理能力直接影响企业竞争表现。"
        )

        self.assertGreater(total, 0)
        self.assertGreater(repeated, 0)
        self.assertGreater(ratio, 0.0)

    def test_body_typography_accepts_visually_close_body_font(self) -> None:
        snapshot = DocumentSnapshot(
            path="demo.docx",
            paragraphs=[
                make_paragraph(1, "第一章 绪论", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
                make_paragraph(2, "电子商务平台的演化速度持续加快，企业必须重新理解数据驱动决策的价值。", font_name="仿宋_GB2312"),
                make_paragraph(3, "大数据技术不仅改变了营销环节，也改变了库存管理、用户画像和风险控制的运行方式。", font_name="仿宋_GB2312"),
                make_paragraph(4, "第二章 平台数据分析", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
                make_paragraph(5, "从平台交易、评价、点击和物流等多源数据中提炼规律，是电商论文正文中常见的分析展开方式。", font_name="仿宋_GB2312"),
                make_paragraph(6, "参考文献", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
            ],
        )

        item = _score_body_typography(snapshot)
        self.assertGreaterEqual(item.score, 5.5)

    def test_visual_red_flags_catch_blank_lines_and_heavy_body_font(self) -> None:
        snapshot = DocumentSnapshot(
            path="demo.docx",
            paragraphs=[
                make_paragraph(1, "第一章 绪论", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
                make_paragraph(2, "数据驱动转型已经成为行业共识，但很多初稿在正文里仍然存在排版紊乱的问题。", font_name="黑体"),
                make_paragraph(3, "", normalized=""),
                make_paragraph(4, "如果每一段之间都额外敲一个回车，翻页时就会出现明显的跳行和散乱观感。", font_name="黑体"),
                make_paragraph(5, "", normalized=""),
                make_paragraph(6, "第二章 分析", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
                make_paragraph(7, "同时如果正文主体也大量使用黑体，整篇论文会显得发黑发重，不像正常装订稿。", font_name="黑体"),
                make_paragraph(8, "参考文献", style_name="标题 1", font_name="黑体", font_size=12.0, bold=1, alignment=1),
            ],
        )

        item = _score_body_typography(snapshot)
        flags = _format_visual_red_flags(snapshot)
        self.assertLess(item.score, 4.0)
        self.assertTrue(any("空段" in flag or "回车" in flag for flag in flags))
        self.assertTrue(any("黑体" in flag or "重字族" in flag for flag in flags))


if __name__ == "__main__":
    unittest.main()
