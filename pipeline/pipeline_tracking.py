from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from pipeline_feedback import build_student_feedback
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
    "run_root": None,
    "stage": None,
    "visual_mode": None,
    "visual_model": None,
    "text_mode": None,
    "text_primary_model": None,
    "text_secondary_model": None,
    "grade_time": None,
    "source_key": None,
    "teacher_name": None,
    "stage_label": None,
    "source_target_page_url": None,
    "delivery_case_name": None,
}


def new_student_record(student_key: str, sid: str | None, name: str | None, **overrides: Any) -> dict[str, Any]:
    record = dict(BASE_STUDENT_RECORD)
    record["student_key"] = student_key
    record["sid"] = sid
    record["name"] = name
    record.update(overrides)
    if "downloaded_files" not in overrides:
        record["downloaded_files"] = []
    return record


def refresh_tracking_outputs(pipeline: Any) -> dict[str, Any]:
    download_index_path = pipeline.cfg.download_output_root / "state" / "downloaded_index.json"
    ingest_state = pipeline._load_state()
    file_source_map = pipeline._load_file_source_map()
    download_index = read_json(download_index_path, {})
    student_records: dict[str, dict[str, Any]] = {}

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
            grade_json_path = Path(run_info.get("json", "")) if run_info.get("json") else run_dir / "json" / "grade_result.json"
            grade_report_path = Path(run_info.get("report", "")) if run_info.get("report") else run_dir / "reports" / "grade_report.txt"
            grade_data = read_json(grade_json_path, {})
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
                }
            )

    pipeline.feedback_dir.mkdir(parents=True, exist_ok=True)
    for key, runs in grade_by_student.items():
        latest = runs[-1]
        record = student_records.setdefault(
            key,
            new_student_record(
                key,
                latest["sid"],
                latest["name"],
                ingested=True,
                paper_path=latest["paper_path"],
            ),
        )
        feedback_file = pipeline.feedback_dir / f"{safe_name(record['sid'] or '')}_{safe_name(record['name'] or 'student')}.md"
        feedback_file.write_text(build_student_feedback(record, latest["grade_data"]), encoding="utf-8")

        record.update(
            {
                "graded": True,
                "ingested": True,
                "score": latest["score"],
                "decision": latest["decision"],
                "paper_title": latest["paper_title"],
                "paper_path": latest["paper_path"],
                "grading_report_path": latest["grade_report_path"],
                "grading_json_path": latest["grade_json_path"],
                "feedback_path": str(feedback_file.resolve()),
                "run_root": latest["run_root"],
                "stage": latest["stage"],
                "visual_mode": latest["visual_mode"],
                "visual_model": latest["visual_model"],
                "text_mode": latest["text_mode"],
                "text_primary_model": latest["text_primary_model"],
                "text_secondary_model": latest["text_secondary_model"],
                "grade_time": latest["grade_time"],
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
        },
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
    pipeline.student_log_md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    return {
        "student_log_json": str(pipeline.student_log_json_path.resolve()),
        "student_log_md": str(pipeline.student_log_md_path.resolve()),
        "feedback_dir": str(pipeline.feedback_dir.resolve()),
        "student_count": len(entries),
        "graded_count": log_payload["summary"]["graded_students"],
    }
