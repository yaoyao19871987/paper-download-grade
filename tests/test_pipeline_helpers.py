from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


PIPELINE_DIR = Path(__file__).resolve().parents[1] / "pipeline"
if str(PIPELINE_DIR) not in sys.path:
    sys.path.insert(0, str(PIPELINE_DIR))

from pipeline_feedback import build_student_feedback, clean_feedback_text, score_ratio
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

    def test_build_student_feedback(self) -> None:
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
        text = build_student_feedback(entry, grade_data)
        self.assertIn("# Test论文修改建议", text)
        self.assertIn("## 优先修改", text)
        self.assertIn("## 建议修改顺序", text)


if __name__ == "__main__":
    unittest.main()
