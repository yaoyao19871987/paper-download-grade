from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from paper_grader import credential_store
from paper_grader.visual_reviewer import (
    VisualReviewResult,
    _fuse_visual_reviews,
    _normalize_visual_review_json,
    _post_json_request,
    export_document_to_pdf,
)


class VisualAndCredentialHelpersTest(unittest.TestCase):
    def test_normalize_visual_review_json_from_wrapped_text(self) -> None:
        raw = """
prefix text
{
  "overall_verdict": "minor_issues",
  "visual_order_score": 7.3,
  "confidence": 0.82,
  "major_issues": [],
  "minor_issues": ["Spacing around heading is inconsistent"],
  "evidence": ["Page 2 heading spacing"],
  "page_observations": ["Page 2 has extra blank line before section title"],
  "notes": ["Use unified paragraph spacing settings"]
}
suffix text
"""
        parsed = _normalize_visual_review_json(raw)
        self.assertEqual(parsed["overall_verdict"], "minor_issues")
        self.assertAlmostEqual(parsed["visual_order_score"], 7.3, places=2)
        self.assertAlmostEqual(parsed["confidence"], 0.82, places=2)
        self.assertEqual(len(parsed["minor_issues"]), 1)

    def test_fuse_visual_reviews_prefers_primary_hard_fail(self) -> None:
        primary = VisualReviewResult(
            mode="siliconflow",
            model="primary-model",
            pdf_path="demo.pdf",
            overall_verdict="rewrite",
            visual_order_score=3.6,
            confidence=0.85,
            major_issues=["Title page and TOC are visually broken"],
            minor_issues=["Small punctuation spacing issues"],
            evidence=["Page 1 and 2 layout drift"],
            page_observations=["Page 2 has obvious blank block"],
            notes=["Primary expert flagged structural issues"],
            raw_response_path=None,
        )
        secondary = VisualReviewResult(
            mode="siliconflow",
            model="secondary-model",
            pdf_path="demo.pdf",
            overall_verdict="pass",
            visual_order_score=8.7,
            confidence=0.70,
            major_issues=[],
            minor_issues=["Footer alignment slightly off"],
            evidence=["Page 5 footer"],
            page_observations=["Most pages acceptable"],
            notes=["Secondary expert is more lenient"],
            raw_response_path=None,
        )

        fused = _fuse_visual_reviews(primary, secondary, pdf_path="demo.pdf")
        self.assertEqual(fused.overall_verdict, "rewrite")
        self.assertTrue(any("Title page" in issue for issue in fused.major_issues))
        self.assertIsNotNone(fused.visual_order_score)
        self.assertLess(fused.visual_order_score or 10.0, 6.0)

    def test_repo_root_found_by_marker_walk(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            repo_root = Path(tmp) / "paper-download-grade"
            (repo_root / ".git").mkdir(parents=True)
            nested_dir = repo_root / "component_overrides" / "essaygrade" / "app" / "paper_grader"
            nested_dir.mkdir(parents=True)
            fake_file = nested_dir / "credential_store.py"
            fake_file.write_text("# fake", encoding="utf-8")

            with patch.object(credential_store, "__file__", str(fake_file)):
                resolved = credential_store.repo_root()

            self.assertEqual(resolved, repo_root.resolve())

    def test_repo_root_raises_without_markers(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            nested_dir = Path(tmp) / "x" / "y" / "z"
            nested_dir.mkdir(parents=True)
            fake_file = nested_dir / "credential_store.py"
            fake_file.write_text("# fake", encoding="utf-8")

            with patch.object(credential_store, "__file__", str(fake_file)):
                with self.assertRaises(RuntimeError):
                    credential_store.repo_root()

    def test_visual_post_json_request_retries_then_succeeds(self) -> None:
        throttled = Mock()
        throttled.status_code = 429
        throttled.json.return_value = {"error": {"message": "rate limit"}}

        success = Mock()
        success.status_code = 200
        success.json.return_value = {"ok": True}

        with patch("paper_grader.visual_reviewer.requests.post", side_effect=[throttled, success]) as mock_post:
            with patch("paper_grader.visual_reviewer.time.sleep") as mock_sleep:
                with patch.dict(
                    "os.environ",
                    {
                        "VISUAL_REVIEW_MAX_RETRIES": "3",
                        "VISUAL_REVIEW_RETRY_BACKOFF_SECONDS": "1",
                    },
                    clear=False,
                ):
                    result = _post_json_request("https://api.example.com/v1/chat/completions", {"x": 1}, "token", None)

        self.assertEqual(result, {"ok": True})
        self.assertEqual(mock_post.call_count, 2)
        mock_sleep.assert_called_once()

    def test_export_document_to_pdf_uses_dispatchex(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "demo.docx"
            source.write_text("x", encoding="utf-8")
            out_dir = Path(tmp) / "out"
            out_dir.mkdir(parents=True, exist_ok=True)

            fake_document = Mock()
            fake_word = Mock()
            fake_word.Documents.Open.return_value = fake_document

            lock_cm = Mock()
            lock_cm.__enter__ = Mock(return_value=None)
            lock_cm.__exit__ = Mock(return_value=False)

            with patch("paper_grader.visual_reviewer._acquire_word_export_lock", return_value=lock_cm) as mock_lock:
                with patch("paper_grader.visual_reviewer.win32.DispatchEx", return_value=fake_word) as mock_dispatchex:
                    export_path = export_document_to_pdf(str(source), str(out_dir))

            mock_lock.assert_called_once()
            mock_dispatchex.assert_called_once_with("Word.Application")
            fake_word.Documents.Open.assert_called_once()
            fake_document.SaveAs.assert_called_once()
            fake_document.Close.assert_called_once_with(False)
            fake_word.Quit.assert_called_once()
            self.assertEqual(export_path.name, "demo.visual_review.pdf")


if __name__ == "__main__":
    unittest.main()
