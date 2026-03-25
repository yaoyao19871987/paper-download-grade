from __future__ import annotations

from collections import defaultdict
import logging
from datetime import datetime
from pathlib import Path
from typing import Any

from pipeline_feedback import generate_student_feedback_artifact
from pipeline_utils import (
    format_score,
    normalize_repo_path,
    now_iso,
    parse_filename,
    parse_ingested_filename,
    parse_run_info,
    read_json,
    relative_display,
    safe_name,
    write_json,
)


BASE_STUDENT_RECORD: dict[str, Any] = {
    "student_key": None,
    "sid": None,
    "name": None,
    "downloaded": False,
    "downloaded_count": 0,
    "downloaded_at": None,
    "downloaded_files": [],
    "latest_download_file": None,
    "ingested": False,
    "graded": False,
    "score": None,
    "decision": None,
    "paper_title": None,
    "paper_path": None,
    "grading_report_path": None,
    "grading_json_path": None,
    "feedback_path": None,
    "feedback_status": None,
    "feedback_model": None,
    "feedback_attempts": 0,
    "feedback_error": None,
    "feedback_raw_response_path": None,
    "run_root": None,
    "stage": None,
    "visual_mode": None,
    "visual_model": None,
    "text_mode": None,
    "text_primary_model": None,
    "text_secondary_model": None,
    "grade_time": None,
    "grade_data_valid": False,
    "grade_error": None,
    "source_key": None,
    "teacher_name": None,
    "stage_label": None,
    "source_target_page_url": None,
    "delivery_case_name": None,
}


logger = logging.getLogger("paper_pipeline")


def new_student_record(student_key: str, sid: str | None, name: str | None, **overrides: Any) -> dict[str, Any]:
    record = dict(BASE_STUDENT_RECORD)
    record["student_key"] = student_key
    record["sid"] = sid
    record["name"] = name
    record.update(overrides)
    if "downloaded_files" not in overrides:
        record["downloaded_files"] = []
    return record


def _resolve_run_artifact(
    run_dir: Path,
    recorded_path: str | None,
    repo_root: Path,
    default_parts: tuple[str, ...],
) -> tuple[Path, str | None]:
    candidates: list[Path] = []
    if recorded_path:
        candidates.append(Path(recorded_path))
        normalized = normalize_repo_path(recorded_path, repo_root)
        if normalized:
            normalized_path = Path(normalized)
            if normalized_path not in candidates:
                candidates.append(normalized_path)
    candidates.append(run_dir.joinpath(*default_parts))

    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve(), None

    missing_path = candidates[0] if candidates else run_dir.joinpath(*default_parts)
    fallback_path = run_dir.joinpath(*default_parts)
    return fallback_path.resolve(), f"missing artifact: recorded={missing_path} fallback={fallback_path}"


def _has_valid_grade_data(grade_data: dict[str, Any]) -> bool:
    if not isinstance(grade_data, dict) or not grade_data:
        return False
    summary = grade_data.get("summary")
    if not isinstance(summary, dict):
        return False
    if summary.get("decision"):
        return True
    return any(
        grade_data.get(key)
        for key in ("format_items", "content_items", "text_review", "visual_review", "reference_audit")
    )


def refresh_tracking_outputs(pipeline: Any) -> dict[str, Any]:
    download_index_path = pipeline.cfg.download_output_root / "state" / "downloaded_index.json"
    ingest_state = pipeline._load_state()
    file_source_map = pipeline._load_file_source_map()
    download_index = read_json(download_index_path, {})
    student_records: dict[str, dict[str, Any]] = {}
    incomplete_runs: list[dict[str, Any]] = []
    feedback_failures: list[dict[str, Any]] = []

    for base, item in (download_index.get("students", {}) or {}).items():
        student_records[base] = new_student_record(
            base,
            item.get("sid"),
            item.get("name"),
            downloaded=True,
            downloaded_count=len(item.get("files", []) or []),
            downloaded_at=item.get("lastDownloadedAt") or item.get("firstDownloadedAt"),
            downloaded_files=[
                str((pipeline.cfg.download_output_root / "downloads" / file_name).resolve())
                for file_name in (item.get("files", []) or [])
            ],
            latest_download_file=(
                str((pipeline.cfg.download_output_root / "downloads" / item.get("files", [])[-1]).resolve())
                if item.get("files")
                else None
            ),
        )
        pipeline._apply_source_metadata(student_records[base], file_source_map)

    for src, item in (ingest_state.get("source_to_dest", {}) or {}).items():
        source_path = normalize_repo_path(src, pipeline.repo_root)
        source_name = Path(source_path).name
        sid, name, _ = parse_filename(Path(source_name).stem)
        if not sid or not name:
            continue
        key = f"{sid}_{name}"
        record = student_records.setdefault(
            key,
            new_student_record(
                key,
                sid,
                name,
                latest_download_file=source_path,
            ),
        )
        record["ingested"] = True
        record["paper_path"] = normalize_repo_path(str(item.get("dest") or ""), pipeline.repo_root)
        if not record.get("latest_download_file"):
            record["latest_download_file"] = source_path
        if source_path and source_path not in record["downloaded_files"]:
            record["downloaded_files"].append(source_path)
        pipeline._apply_source_metadata(record, file_source_map)

    grade_by_student: dict[str, list[dict[str, Any]]] = defaultdict(list)
    grading_runs_dir = pipeline.cfg.grading_runs_dir
    if grading_runs_dir.exists():
        for run_dir in sorted(grading_runs_dir.iterdir(), key=lambda path: path.stat().st_mtime):
            if not run_dir.is_dir():
                continue
            run_info_path = run_dir / "notes" / "run_info.txt"
            if not run_info_path.exists():
                continue
            run_info = parse_run_info(run_info_path, pipeline.repo_root)
            paper_path = run_info.get("paper")
            if not paper_path:
                continue
            sid, name = parse_ingested_filename(Path(paper_path).stem)
            if not sid or not name:
                continue
            key = f"{sid}_{name}"
            grade_json_path, grade_json_warning = _resolve_run_artifact(
                run_dir=run_dir,
                recorded_path=run_info.get("json"),
                repo_root=pipeline.repo_root,
                default_parts=("json", "grade_result.json"),
            )
            grade_report_path, grade_report_warning = _resolve_run_artifact(
                run_dir=run_dir,
                recorded_path=run_info.get("report"),
                repo_root=pipeline.repo_root,
                default_parts=("reports", "grade_report.txt"),
            )
            grade_data = read_json(grade_json_path, {})
            grade_data_valid = _has_valid_grade_data(grade_data)
            grade_error_parts = [item for item in (grade_json_warning, grade_report_warning) if item]
            if not grade_data_valid:
                grade_error_parts.append(f"invalid or empty grade data: {grade_json_path}")
            grade_error = " | ".join(grade_error_parts) if grade_error_parts else None
            if grade_error:
                logger.warning("Run %s has grade artifact issues: %s", run_dir.name, grade_error)
            summary = grade_data.get("summary", {})
            extracted = grade_data.get("extracted", {})
            grade_by_student[key].append(
                {
                    "sid": sid,
                    "name": name,
                    "paper_path": normalize_repo_path(paper_path, pipeline.repo_root),
                    "run_root": str(run_dir.resolve()),
                    "grade_json_path": normalize_repo_path(str(grade_json_path), pipeline.repo_root),
                    "grade_report_path": normalize_repo_path(str(grade_report_path), pipeline.repo_root),
                    "grade_time": datetime.fromtimestamp(run_dir.stat().st_mtime).astimezone().isoformat(timespec="seconds"),
                    "score": summary.get("total_score"),
                    "decision": summary.get("decision"),
                    "stage": summary.get("stage") or run_info.get("stage"),
                    "visual_mode": grade_data.get("visual_review", {}).get("mode") or run_info.get("visual_mode"),
                    "visual_model": grade_data.get("visual_review", {}).get("model") or run_info.get("visual_model"),
                    "text_mode": grade_data.get("text_review", {}).get("mode") or run_info.get("text_mode"),
                    "text_primary_model": (grade_data.get("text_review", {}).get("primary") or {}).get("model") or run_info.get("text_primary_model"),
                    "text_secondary_model": (grade_data.get("text_review", {}).get("secondary") or {}).get("model") or run_info.get("text_secondary_model"),
                    "paper_title": extracted.get("title"),
                    "grade_data": grade_data,
                    "grade_data_valid": grade_data_valid,
                    "grade_error": grade_error,
                }
            )

    pipeline.feedback_dir.mkdir(parents=True, exist_ok=True)
    for key, runs in grade_by_student.items():
        latest = runs[-1]
        selected = next((item for item in reversed(runs) if item.get("grade_data_valid")), latest)
        if not latest.get("grade_data_valid"):
            incomplete_runs.append(
                {
                    "sid": latest["sid"],
                    "name": latest["name"],
                    "run_root": latest["run_root"],
                    "paper_path": latest["paper_path"],
                    "grade_json_path": latest["grade_json_path"],
                    "error": latest.get("grade_error") or "invalid or empty grade data",
                }
            )
        record = student_records.setdefault(
            key,
            new_student_record(
                key,
                selected["sid"],
                selected["name"],
                ingested=True,
                paper_path=selected["paper_path"],
            ),
        )
        feedback_file = pipeline.feedback_dir / f"{safe_name(record['sid'] or '')}_{safe_name(record['name'] or 'student')}.md"
        feedback_result = generate_student_feedback_artifact(
            {
                **record,
                "paper_title": selected["paper_title"],
                "run_root": selected["run_root"],
                "stage": selected["stage"],
                "teacher_name": record.get("teacher_name"),
                "grade_error": latest.get("grade_error"),
            },
            selected["grade_data"],
        )
        feedback_file.write_text(feedback_result["markdown"], encoding="utf-8")
        if feedback_result.get("status") != "ok":
            feedback_failures.append(
                {
                    "sid": selected["sid"],
                    "name": selected["name"],
                    "run_root": selected["run_root"],
                    "status": feedback_result.get("status"),
                    "error": feedback_result.get("error"),
                }
            )

        record.update(
            {
                "graded": bool(selected.get("grade_data_valid")),
                "ingested": True,
                "score": selected["score"],
                "decision": selected["decision"] or ("评分结果缺失" if not selected.get("grade_data_valid") else None),
                "paper_title": selected["paper_title"],
                "paper_path": selected["paper_path"],
                "grading_report_path": selected["grade_report_path"],
                "grading_json_path": selected["grade_json_path"],
                "feedback_path": str(feedback_file.resolve()),
                "feedback_status": feedback_result.get("status"),
                "feedback_model": feedback_result.get("model"),
                "feedback_attempts": feedback_result.get("attempts"),
                "feedback_error": feedback_result.get("error"),
                "feedback_raw_response_path": feedback_result.get("raw_response_path"),
                "run_root": selected["run_root"],
                "stage": selected["stage"],
                "visual_mode": selected["visual_mode"],
                "visual_model": selected["visual_model"],
                "text_mode": selected["text_mode"],
                "text_primary_model": selected["text_primary_model"],
                "text_secondary_model": selected["text_secondary_model"],
                "grade_time": selected["grade_time"],
                "grade_data_valid": bool(selected.get("grade_data_valid")),
                "grade_error": latest.get("grade_error"),
            }
        )
        pipeline._apply_source_metadata(record, file_source_map)

    entries = sorted(
        student_records.values(),
        key=lambda item: (
            item.get("sid") or "",
            item.get("downloaded_at") or "",
            item.get("grade_time") or "",
        ),
    )

    log_payload = {
        "updated_at": now_iso(),
        "repo_root": str(pipeline.repo_root.resolve()),
        "download_root": str(pipeline.cfg.download_output_root.resolve()),
        "incoming_dir": str(pipeline.cfg.incoming_dir.resolve()),
        "feedback_dir": str(pipeline.feedback_dir.resolve()),
        "active_source": pipeline._get_active_source(),
        "summary": {
            "downloaded_students": sum(1 for item in entries if item.get("downloaded")),
            "ingested_students": sum(1 for item in entries if item.get("ingested")),
            "graded_students": sum(1 for item in entries if item.get("graded")),
            "feedback_failed_students": sum(1 for item in entries if item.get("feedback_status") not in {None, "ok"}),
        },
        "incomplete_runs": incomplete_runs,
        "feedback_failures": feedback_failures,
        "students": entries,
    }
    write_json(pipeline.student_log_json_path, log_payload)

    lines = [
        "# 学生论文下载与评分总日志",
        "",
        f"- 更新时间: {log_payload['updated_at']}",
        f"- 已下载学生数: {log_payload['summary']['downloaded_students']}",
        f"- 已入队学生数: {log_payload['summary']['ingested_students']}",
        f"- 已评分学生数: {log_payload['summary']['graded_students']}",
        f"- 当前活动来源: {((log_payload.get('active_source') or {}).get('folder_name') or '-')}",
        "",
        "| 学号 | 姓名 | 来源老师 | 阶段 | 下载 | 入队 | 评分 | 分数 | 裁定 | 下载文件 | 论文路径 | 评分报告 | 学生评语 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for item in entries:
        lines.append(
            "| {sid} | {name} | {teacher_name} | {stage_label} | {downloaded} | {ingested} | {graded} | {score} | {decision} | {download_file} | {paper} | {report} | {feedback} |".format(
                sid=item.get("sid") or "-",
                name=item.get("name") or "-",
                teacher_name=item.get("teacher_name") or "-",
                stage_label=item.get("stage_label") or "-",
                downloaded="是" if item.get("downloaded") else "否",
                ingested="是" if item.get("ingested") else "否",
                graded="是" if item.get("graded") else "否",
                score=format_score(item.get("score")),
                decision=item.get("decision") or "-",
                download_file=relative_display(item.get("latest_download_file"), pipeline.repo_root),
                paper=relative_display(item.get("paper_path"), pipeline.repo_root),
                report=relative_display(item.get("grading_report_path"), pipeline.repo_root),
                feedback=relative_display(item.get("feedback_path"), pipeline.repo_root),
            )
        )

    # Add anomaly and error summary
    lines.extend([
        "",
        "## 异常与错误汇总 (Anomaly & Error Summary)",
        "| 学号 | 姓名 | 教师/批次 | 失败环节 | 最后错误信息 | 重试次数 | 需要人工介入 |",
        "| --- | --- | --- | --- | --- | --- | --- |"
    ])
    
    anomalies = []
    for item in entries:
        is_anomaly = False
        reasons = []
        failure_stage = "-"
        last_error = "-"
        retries = 0
        still_needs_manual = "否"
        
        if item.get("feedback_status") not in {None, "ok"} and item.get("graded"):
            is_anomaly = True
            reasons.append("反馈生成失败或为空")
            failure_stage = "Feedback Generation"
            last_error = item.get("feedback_error") or "Unknown Kimi Error"
            retries = item.get("feedback_attempts") or 0
            still_needs_manual = "是"
        
        if item.get("ingested") and not item.get("grade_data_valid"):
            is_anomaly = True
            reasons.append("评分数据无效")
            failure_stage = "Grading"
            last_error = item.get("grade_error") or "Invalid JSON/missing artifacts"
            still_needs_manual = "是"
            
        teacher = str(item.get("teacher_name") or "")
        if "3C" in teacher:
            is_anomaly = True
            reasons.append("教师_3C_特殊批次")
            if failure_stage == "-":
                failure_stage = "None"
                last_error = "3C teacher requires formal rebuild"
            
        if is_anomaly:
            lines.append(
                "| {sid} | {name} | {teacher} | {stage} | {error} | {retries} | {manual} |".format(
                    sid=item.get("sid") or "-",
                    name=item.get("name") or "-",
                    teacher=item.get("teacher_name") or "-",
                    stage=failure_stage,
                    error=str(last_error).replace("\n", " ")[:100],
                    retries=retries,
                    manual=still_needs_manual
                )
            )

    pipeline.student_log_md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    return {
        "student_log_json": str(pipeline.student_log_json_path.resolve()),
        "student_log_md": str(pipeline.student_log_md_path.resolve()),
        "feedback_dir": str(pipeline.feedback_dir.resolve()),
        "student_count": len(entries),
        "graded_count": log_payload["summary"]["graded_students"],
        "incomplete_runs": incomplete_runs,
        "incomplete_run_count": len(incomplete_runs),
        "feedback_failures": feedback_failures,
        "feedback_failure_count": len(feedback_failures),
    }
