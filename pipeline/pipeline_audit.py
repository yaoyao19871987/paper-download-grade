from __future__ import annotations

import json
import logging
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from pipeline_utils import read_json, write_json

logger = logging.getLogger(__name__)

def build_audit_input(entry: dict[str, Any]) -> dict[str, Any]:
    grade_data = entry.get("grade_data", {})
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
            "--model", "gemini-3.1-pro-preview", 
            "-o", "json",
            "-p", "Output strictly valid JSON only without markdown codeblocks.",
        ]
        
        with open(temp_path, "r", encoding="utf-8") as f:
            result = subprocess.run(cmd, stdin=f, capture_output=True, text=True, check=True)
            
        output_text = result.stdout.strip()
        # Clean potential markdown wrapping if gemini CLI didn't respect raw json
        if output_text.startswith("```json"):
            output_text = output_text[7:]
        if output_text.startswith("```"):
            output_text = output_text[3:]
        if output_text.endswith("```"):
            output_text = output_text[:-3]
        output_text = output_text.strip()
        
        try:
            return json.loads(output_text)
        except json.JSONDecodeError as e:
            return {"error": f"Failed to parse JSON: {e}", "raw_output": output_text}
    except subprocess.CalledProcessError as e:
        return {"error": f"Gemini CLI execution failed: {e.stderr}"}
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception:
            pass
