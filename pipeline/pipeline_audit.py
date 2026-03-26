from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from pipeline_utils import read_json

logger = logging.getLogger(__name__)

AUDIT_REQUIRED_FIELDS = {
    "review_verdict",
    "reason",
    "score_adjustment_needed",
    "new_decision_if_any",
    "must_fix_fields",
    "student_feedback_rewrite_needed",
}

AUDIT_TIMEOUT_SECONDS = int(os.getenv("PIPELINE_AUDIT_TIMEOUT_SECONDS", "240"))
AUDIT_MODEL = os.getenv("PIPELINE_AUDIT_MODEL", "gemini-3-flash-preview")


def _classify_error_text(text: str) -> tuple[str, bool, str | None]:
    message = str(text or "")
    lowered = message.lower()
    if "exhausted your capacity" in lowered or "quota_exhausted" in lowered or "terminalquotaerror" in lowered:
        reset_match = re.search(r"reset after ([0-9hms ]+)", message, re.IGNORECASE)
        reset_after = reset_match.group(1).strip() if reset_match else None
        return "quota_exhausted", False, reset_after
    if "timed out" in lowered:
        return "timeout", True, None
    if "429" in message or "503" in message or "500" in message:
        return "provider_error", True, None
    return "cli_error", True, None


def _load_grade_data(entry: dict[str, Any]) -> dict[str, Any]:
    grade_data = entry.get("grade_data", {})
    if grade_data:
        return grade_data

    candidates: list[Path] = []
    for key in ("grading_json_path", "grade_json_path"):
        raw = str(entry.get(key) or "").strip()
        if raw:
            candidates.append(Path(raw))

    run_root = str(entry.get("run_root") or "").strip()
    if run_root:
        candidates.append(Path(run_root) / "json" / "grade_result.json")

    seen: set[str] = set()
    for candidate in candidates:
        try:
            resolved = str(candidate.resolve())
        except Exception:
            resolved = str(candidate)
        if resolved in seen:
            continue
        seen.add(resolved)
        if candidate.exists():
            return read_json(candidate, {})

    return {}


def build_audit_input(entry: dict[str, Any]) -> dict[str, Any]:
    grade_data = _load_grade_data(entry)
    summary = grade_data.get("summary", {})
    visual_review = grade_data.get("visual_review", {})
    text_review = grade_data.get("text_review", {})
    reference_audit = grade_data.get("reference_audit", {})
    
    # Compact format/content
    format_items = grade_data.get("format_items", [])
    content_items = grade_data.get("content_items", [])
    
    weak_format = [
        {"name": i.get("name"), "score": i.get("score"), "max_score": i.get("max_score")} 
        for i in format_items if i.get("score", 0) < i.get("max_score", 0) * 0.6
    ]
    weak_content = [
        {"name": i.get("name"), "score": i.get("score"), "max_score": i.get("max_score")} 
        for i in content_items if i.get("score", 0) < i.get("max_score", 0) * 0.6
    ]

    return {
        "student_info": {
            "sid": entry.get("sid"),
            "name": entry.get("name"),
            "teacher": entry.get("teacher_name"),
            "stage": entry.get("stage"),
        },
        "grading_status": {
            "score": summary.get("total_score"),
            "decision": summary.get("decision"),
            "grade_data_valid": entry.get("grade_data_valid"),
            "grade_error": entry.get("grade_error"),
        },
        "weak_items": {
            "format": weak_format[:5],
            "content": weak_content[:5],
        },
        "issues": {
            "visual": visual_review.get("major_issues", [])[:3],
            "text_major": text_review.get("major_problems", [])[:3],
            "reference_notes": reference_audit.get("notes", [])[:3],
        },
        "feedback_status": entry.get("feedback_status"),
    }


def _strip_code_fences(text: str) -> str:
    cleaned = str(text or "").strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    return cleaned.strip()


def _normalize_audit_payload(raw_output: str) -> dict[str, Any]:
    parsed = json.loads(_strip_code_fences(raw_output))
    meta: dict[str, Any] = {}
    payload = parsed

    if isinstance(parsed, dict) and "response" in parsed and "review_verdict" not in parsed:
        meta = {k: v for k, v in parsed.items() if k != "response"}
        payload = json.loads(_strip_code_fences(parsed.get("response", "")))

    if not isinstance(payload, dict):
        raise ValueError("Gemini audit output is not a JSON object")

    missing = sorted(AUDIT_REQUIRED_FIELDS - set(payload.keys()))
    if missing:
        raise ValueError(f"Gemini audit output missing fields: {', '.join(missing)}")

    result = {
        "review_verdict": payload.get("review_verdict"),
        "reason": payload.get("reason"),
        "score_adjustment_needed": payload.get("score_adjustment_needed"),
        "new_decision_if_any": payload.get("new_decision_if_any"),
        "must_fix_fields": payload.get("must_fix_fields") or [],
        "student_feedback_rewrite_needed": bool(payload.get("student_feedback_rewrite_needed")),
    }
    if meta:
        result["_meta"] = meta
    return result


def run_gemini_audit(student_record: dict[str, Any]) -> dict[str, Any]:
    input_data = build_audit_input(student_record)
    prompt = f"""You are a senior academic thesis auditor. Review the following student grading data.
Analyze if the grading decision and score match the severity of the issues.
Identify if the student feedback needs a rewrite or if the grading should be rerun.

Input Data:
{json.dumps(input_data, ensure_ascii=False, indent=2)}

Output strictly valid JSON with the following schema:
{{
  "review_verdict": "keep" | "feedback_only" | "rescore_and_feedback" | "rerun_grading",
  "reason": "string explaining the verdict",
  "score_adjustment_needed": number or null,
  "new_decision_if_any": "string or null",
  "must_fix_fields": ["list of strings"],
  "student_feedback_rewrite_needed": boolean
}}
"""
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", delete=False) as f:
        f.write(prompt)
        temp_path = f.name
    
    try:
        cmd = [
            "cmd", "/c", "gemini.cmd", 
            "--model", AUDIT_MODEL, 
            "-o", "json",
            "-p", "Output strictly valid JSON only without markdown codeblocks.",
        ]
        
        with open(temp_path, "r", encoding="utf-8") as f:
            result = subprocess.run(
                cmd,
                stdin=f,
                capture_output=True,
                text=True,
                check=True,
                encoding="utf-8",
                errors="replace",
                timeout=AUDIT_TIMEOUT_SECONDS,
            )
            
        output_text = result.stdout.strip()
        try:
            parsed = _normalize_audit_payload(output_text)
            parsed["model"] = AUDIT_MODEL
            return parsed
        except (json.JSONDecodeError, ValueError) as exc:
            return {"error": f"Failed to parse audit JSON: {exc}", "raw_output": output_text, "model": AUDIT_MODEL}
    except subprocess.CalledProcessError as e:
        stderr_text = str(e.stderr or "")
        error_type, retryable, reset_after = _classify_error_text(stderr_text)
        payload: dict[str, Any] = {
            "error": f"Gemini CLI execution failed: {stderr_text}",
            "error_type": error_type,
            "retryable": retryable,
            "model": AUDIT_MODEL,
        }
        if reset_after:
            payload["quota_reset_after"] = reset_after
        return payload
    except subprocess.TimeoutExpired as e:
        stderr = ""
        try:
            stderr = (e.stderr or b"").decode("utf-8", errors="replace") if isinstance(e.stderr, (bytes, bytearray)) else str(e.stderr or "")
        except Exception:
            stderr = str(e)
        return {
            "error": f"Gemini CLI timed out after {AUDIT_TIMEOUT_SECONDS} seconds",
            "error_type": "timeout",
            "retryable": True,
            "stderr": stderr,
            "model": AUDIT_MODEL,
        }
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception:
            pass
