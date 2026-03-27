from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


PIPELINE_DIR = Path(__file__).resolve().parents[1] / "pipeline"
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from pipeline_feedback import _build_feedback_input, _render_feedback_markdown, clean_feedback_text, score_ratio
from pipeline_utils import normalize_repo_path, parse_filename, parse_ingested_filename


class PipelineHelperTests(unittest.TestCase):
    def test_parse_filename(self) -> None:
        self.assertEqual(parse_filename("24315018_张三"), ("24315018", "张三", None))
        self.assertEqual(parse_filename("24315018_张三_2"), ("24315018", "张三", "2"))
        self.assertEqual(parse_filename("bad-name"), (None, None, None))

    def test_parse_ingested_filename(self) -> None:
        self.assertEqual(parse_ingested_filename("20260317_24315060_张三_0c41d04c"), ("24315060", "张三"))
        self.assertEqual(parse_ingested_filename("24315060_张三"), ("24315060", "张三"))
        self.assertEqual(parse_ingested_filename("oops"), (None, None))

    def test_score_ratio(self) -> None:
        self.assertEqual(score_ratio({"score": 5, "max_score": 10}), 0.5)
        self.assertEqual(score_ratio({"score": "x", "max_score": 10}), 0.0)
        self.assertEqual(score_ratio({"score": 5, "max_score": 0}), 0.0)

    def test_clean_feedback_text(self) -> None:
        self.assertEqual(clean_feedback_text("  foo \n bar  "), "foo bar")
        self.assertEqual(clean_feedback_text(None), "")

    def test_normalize_repo_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo_root = Path(tmp) / "paper-download-grade"
            repo_root.mkdir()
            normalized = normalize_repo_path(
                "C:/Users/test/paper-download-grade/components/essaygrade/file.txt",
                repo_root,
            )
            self.assertEqual(
                normalized,
                str((repo_root / "components" / "essaygrade" / "file.txt").resolve()),
            )

    def test_render_feedback_markdown(self) -> None:
        entry = {"sid": "1", "name": "Test", "paper_title": "Demo"}
        grade_data = {
            "summary": {"decision": "通过", "total_score": 88},
            "extracted": {"title": "Demo"},
            "format_items": [],
            "content_items": [],
            "reference_audit": {},
            "visual_review": {"mode": "auto"},
            "text_review": {"mode": "expert", "major_problems": [], "revision_actions": []},
        }
        payload = {
            "title": "Test论文修改建议",
            "overall": "请先完成摘要、格式和引用的集中修订。",
            "priority_actions": ["先补摘要里的研究方法和结论。"],
            "format_fix": ["统一目录、页码和标题层级。"],
            "content_fix": ["补齐正文分析，不要只列概念。"],
            "writing_fix": ["删除空话，改成明确判断。"],
            "citation_fix": ["核对文内引用和参考文献一一对应。"],
            "next_steps": ["先改硬伤，再整体通读一遍。"],
        }
        text = _render_feedback_markdown(entry, grade_data, payload)
        self.assertIn("- 学号: 1", text)
        self.assertIn("## 总评", text)
        self.assertIn("## 主要问题与修改建议", text)
        self.assertIn("## 下一步", text)

    def test_build_feedback_input_with_audit_review_aliases(self) -> None:
        entry = {
            "sid": "1",
            "name": "Test",
            "paper_title": "Demo",
            "audit_review": {
                "audit_status": "ok",
                "audit_model": "gemini-2.5-flash",
                "review_verdict": "feedback_only",
                "reason": "摘要和结论的修改建议需要说得更明确。",
            },
        }
        grade_data = {
            "summary": {"decision": "通过", "total_score": 88},
            "extracted": {"title": "Demo"},
            "format_items": [],
            "content_items": [],
            "reference_audit": {},
            "visual_review": {"mode": "auto"},
            "text_review": {"mode": "expert", "major_problems": [], "revision_actions": []},
        }
        payload = _build_feedback_input(entry, grade_data)
        self.assertEqual(payload["audit_review"]["status"], "ok")
        self.assertEqual(payload["audit_review"]["model"], "gemini-2.5-flash")
        self.assertEqual(payload["audit_review"]["review_verdict"], "feedback_only")


if __name__ == "__main__":
    unittest.main()
