# 2026-03-25 Feedback Pipeline Resume Point

## Current Branch
- `codex/feedback-pipeline-resume`

## Saved Commits
- `829c929` Refactor pipeline feedback generation and add audit commands
- `b6caffc` Make Gemini audit resumable and parse CLI output
- `de19a98` Add retry and failure summary for audit runs
- pending: skip full tracking refresh for `audit-students` / `rebuild-anomalies`

## What Is Working
- Tracking now records `feedback_status`, `feedback_model`, `feedback_attempts`, `feedback_error`, `feedback_raw_response_path`.
- Invalid or missing `grade_result.json` no longer silently turns into a fake normal student feedback.
- Student feedback generation uses Kimi and stores run-level artifacts under each grading run:
  - `feedback/student_feedback_kimi_raw.json`
  - `feedback/student_feedback_kimi_meta.json`
- `audit-students` is now resumable:
  - loads existing `runtime/pipeline/state/audit_results.json`
  - skips only students with a valid parsed `review_verdict`
  - writes results after each student
- Gemini CLI audit output is normalized from the `-o json` wrapper into the target audit schema.
- Audit retries are now in place:
  - immediate retries: `2` via `PIPELINE_AUDIT_IMMEDIATE_RETRIES`
  - deferred retry passes: `1` via `PIPELINE_AUDIT_DEFERRED_RETRY_PASSES`
  - final failure summary file: `runtime/pipeline/state/audit_failures.json`
- `audit-students` and `rebuild-anomalies` no longer force a full `refresh_tracking_outputs()` on exit, so they can finish promptly and remain interruptible.

## Sample Verification Done Today
- Two audit samples were rerun successfully and written to:
  - `runtime/pipeline/state/audit_results.json`
- Current sample verdicts:
  - `24110001`: `keep`
  - `24110002`: `rescore_and_feedback`

## Important Runtime Files
- Student progress log:
  - `runtime/tracking/student_progress_log.json`
- Student feedback directory:
  - `runtime/tracking/student_feedback`
- Audit result file:
  - `runtime/pipeline/state/audit_results.json`

## Known Gaps Before Full 221 Audit
- Some Chinese text in existing logs and some Gemini audit outputs still shows mojibake. This needs a dedicated normalization pass before final export.
- Formal anomaly rebuild is still partial. `rebuild-anomalies` identifies anomalies and refreshes feedback, but the full regrade/rescore queue is not finished.

## Resume Commands
```powershell
git checkout codex/feedback-pipeline-resume
python -m py_compile pipeline\pipeline.py pipeline\pipeline_audit.py pipeline\pipeline_feedback.py pipeline\pipeline_tracking.py
powershell -ExecutionPolicy Bypass -File pipeline\run_pipeline.ps1 audit-students --limit 10
```

## Suggested Next Steps
1. Fix mojibake on student names and Chinese decision text before exporting review results.
2. Run `audit-students` in small batches and let `runtime/pipeline/state/audit_results.json` accumulate.
3. Inspect `runtime/pipeline/state/audit_failures.json` after each batch and rerun only the remaining failures if needed.
4. After audit output is stable, continue `rebuild-anomalies` on the abnormal batches only.
