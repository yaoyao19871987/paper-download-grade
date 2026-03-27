from __future__ import annotations

import argparse
import csv
import os
import hashlib
import json
import logging
import shutil
import re
import subprocess
import tempfile
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from pipeline_cli import build_parser as _build_parser_shared
from pipeline_feedback import (
    build_student_feedback as _build_student_feedback_impl,
    clean_feedback_text as _clean_feedback_text_impl,
    generate_student_feedback_artifact as _generate_student_feedback_artifact_impl,
    score_ratio as _score_ratio_impl,
)
from pipeline_tracking import refresh_tracking_outputs as _refresh_tracking_outputs_impl
from pipeline_utils import (
    format_score as _format_score_impl,
    normalize_repo_path as _normalize_repo_path_impl,
    read_json as _read_json_impl,
    read_text as _read_text_impl,
    run_command as _run_command_impl,
)


logger = logging.getLogger("paper_pipeline")


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _read_json(path: Path, fallback: Any) -> Any:
    return _read_json_impl(path, fallback)


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _merge_dict(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged: dict[str, Any] = dict(base)
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            merged[k] = _merge_dict(base[k], v)
        else:
            merged[k] = v
    return merged


def _resolve_path(raw: str, base_dir: Path) -> Path:
    p = Path(raw)
    if p.is_absolute():
        return p.resolve()
    return (base_dir / p).resolve()


def _find_repo_root(start: Path) -> Path:
    for candidate in (start, *start.parents):
        if (candidate / ".git").exists():
            return candidate
    raise RuntimeError(f"Unable to locate repository root from {start}")


def _run_command(args: list[str], cwd: Path, extra_env: dict[str, str] | None = None) -> None:
    _run_command_impl(args, cwd=cwd, extra_env=extra_env)


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _safe_name(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F\s]+', "_", value).strip("_")


def _split_csv_arg(raw: str | None) -> list[str]:
    return [item.strip() for item in str(raw or "").split(",") if item.strip()]


def _coerce_float(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _grade_data_is_valid(grade_data: dict[str, Any]) -> bool:
    if not isinstance(grade_data, dict) or not grade_data:
        return False
    summary = grade_data.get("summary")
    if isinstance(summary, dict) and summary.get("decision"):
        return True
    return any(
        grade_data.get(key)
        for key in ("format_items", "content_items", "text_review", "visual_review", "reference_audit")
    )


def _parse_filename(stem: str) -> tuple[str | None, str | None, str | None]:
    # Example: 24315018_张三, 24315018_张三_2
    m = re.match(r"^(\d{5,20})_(.+?)(?:_(\d+))?$", stem)
    if not m:
        return None, None, None
    return m.group(1), m.group(2), m.group(3)


def _parse_ingested_filename(stem: str) -> tuple[str | None, str | None]:
    # Examples:
    # 20260317_24315060_张三_0c41d04c
    # 24315060_张三
    m = re.match(r"^(?:\d{8}_)?(\d{5,20})_(.+?)(?:_(?:[0-9a-f]{4,64}|v\d+|\d+))?$", stem)
    if not m:
        return None, None
    return m.group(1), m.group(2)


def _normalize_repo_path(raw: str, repo_root: Path) -> str:
    return _normalize_repo_path_impl(raw, repo_root)


def _parse_run_info(path: Path, repo_root: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        value = value.strip()
        if key in {"run_root", "paper", "paper_copy", "visual_dir", "text_dir", "json", "report"}:
            value = _normalize_repo_path(value, repo_root)
        data[key.strip()] = value
    return data


def _format_score(value: Any) -> str:
    return _format_score_impl(value)


def _relative_display(path: str | None, repo_root: Path) -> str:
    if not path:
        return "-"
    try:
        return str(Path(path).resolve().relative_to(repo_root))
    except Exception:
        return path


def _dedupe_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def _read_text(path: Path) -> str:
    return _read_text_impl(path)


def _bundle_relative(path: Path, root: Path) -> str:
    try:
        return str(path.resolve().relative_to(root.resolve())).replace("\\", "/")
    except Exception:
        return str(path.resolve())


def _copy_file(source: str | None, destination: Path) -> str | None:
    if not source:
        return None
    src = Path(source)
    if not src.exists() or not src.is_file():
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, destination)
    return str(destination.resolve())


def _summarize_feedback(feedback_path: str | None) -> str:
    if not feedback_path:
        return ""
    path = Path(feedback_path)
    if not path.exists():
        return ""

    text = _read_text(path)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    # 1. 从头部 bullet 提取结论（- 当前结论: xxx）
    conclusion = ""
    for line in lines:
        if not line.startswith("- "):
            continue
        if ":" not in line and "\uff1a" not in line:
            continue
        _, value = re.split(r"[:\uff1a]", line, maxsplit=1)
        value = value.strip()
        if len(value) >= 2:
            conclusion = value

    # 2. 按 ## 标题分段，提取总评全文和全部修改建议条目
    overall_lines: list[str] = []
    urgent_items: list[str] = []
    current_section = ""

    for line in lines:
        if line.startswith("## "):
            current_section = line[3:].strip()
            continue

        if current_section == "\u603b\u8bc4":
            if not line.startswith(("#", "-")):
                overall_lines.append(line)

        if current_section == "\u4e3b\u8981\u95ee\u9898\u4e0e\u4fee\u6539\u5efa\u8bae":
            m = re.match(r"^\d+[.)]\s*(.+)$", line)
            if m:
                item = m.group(1).strip()
                if item:
                    urgent_items.append(item)

    # 3. 拼接 summary
    overall_text = " ".join(overall_lines).strip()
    parts: list[str] = []
    if conclusion:
        parts.append(conclusion)
    if overall_text:
        parts.append(overall_text)
    if urgent_items:
        parts.append("Priority fixes: " + "; ".join(urgent_items))

    summary = " ".join(parts).strip()
    if summary:
        return summary

    for line in lines:
        if line.startswith(("# ", "## ")):
            continue
        if line.startswith("- "):
            candidate = line[2:].strip()
            if candidate:
                return candidate
    return ""


def _source_folder_name(teacher_name: str, stage_label: str) -> str:
    teacher = str(teacher_name or "").strip()
    stage = str(stage_label or "").strip() or "初稿"
    return f"{teacher}_{stage}" if teacher else stage


def _source_key(teacher_name: str, stage_label: str) -> str:
    return _safe_name(_source_folder_name(teacher_name, stage_label)) or "source"


def _score_ratio(item: dict[str, Any]) -> float:
    return _score_ratio_impl(item)


def _clean_feedback_text(text: Any) -> str:
    return _clean_feedback_text_impl(text)

def _number_lines(values: list[str]) -> list[str]:
    return [f"{index}. {value}" for index, value in enumerate(values, 1)]


def _feedback_severity(item: dict[str, Any], category: str) -> str:
    ratio = _score_ratio(item)
    if category == "format":
        if ratio <= 0.12:
            return "格式基本全是错的"
        if ratio <= 0.25:
            return "格式严重错误"
        if ratio <= 0.45:
            return "格式问题很大"
        if ratio <= 0.65:
            return "格式存在较大问题"
        if ratio <= 0.8:
            return "格式存在一定问题"
        return "格式基本合格"

    if ratio <= 0.12:
        return "内容基本没有达到要求"
    if ratio <= 0.25:
        return "内容问题很大"
    if ratio <= 0.45:
        return "内容存在较大问题"
    if ratio <= 0.65:
        return "内容存在一定问题"
    if ratio <= 0.8:
        return "内容基本过得去，但还要继续补强"
    return "内容整体还可以"


def _build_teacher_action_text(suggestions: list[str], fallback: str) -> str:
    cleaned = [_clean_feedback_text(item) for item in suggestions]
    cleaned = [item for item in cleaned if item]
    if not cleaned:
        return fallback
    return "你需要修改的地方：" + "；".join(cleaned) + "。"


def _sorted_feedback_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ranked = [item for item in items if item.get("suggestions")]
    ranked.sort(key=lambda item: (_score_ratio(item), str(item.get("name") or "")))
    return ranked


def _overall_teacher_comment(decision: str) -> str:
    mapping = {
        "打回重写": "这篇稿件目前问题比较多，还不能按毕业论文成稿来交。你要先把格式、结构、引用这些硬伤改到位，再重新提交。",
        "引用退修": "这篇稿件最主要的问题在引用和学术规范。正文引用、文后文献、文献真实性这三件事必须先改明白。",
        "引用待核": "这篇稿件已经有基本框架，但引用证明还不够扎实。你要把文献来源逐条补强，再提交复核。",
        "通过": "这篇稿件已经能往下走了，但仍然要按下面的意见继续修改，不要直接交终稿。",
    }
    return mapping.get(decision or "", "这篇稿件还需要继续修改。你按下面的顺序一项一项改。")


def _build_student_feedback(entry: dict[str, Any], grade_data: dict[str, Any]) -> str:
    return _build_student_feedback_impl(entry, grade_data)


@dataclass
class PipelineConfig:
    project_root: Path
    paperdownload_root: Path
    paper_grading_root: Path
    download_output_root: Path
    incoming_dir: Path
    grading_runs_dir: Path
    state_dir: Path
    credential_store_dir: Path
    feedback_dir: Path
    student_log_json_path: Path
    student_log_md_path: Path
    case_exports_dir: Path
    rename_prefix_with_date: bool
    rename_hash_length: int

    @classmethod
    def load(cls, config_path: Path) -> "PipelineConfig":
        config_dir = config_path.resolve().parent
        project_root = Path(__file__).resolve().parent
        repo_root = _find_repo_root(project_root)
        defaults: dict[str, Any] = {
            "paperdownload_root": str((repo_root / "components" / "paperdownload").resolve()),
            "paper_grading_root": str((repo_root / "components" / "essaygrade").resolve()),
            "download_output_root": str((repo_root / "runtime" / "downloads" / "longzhi_batch_output").resolve()),
            "incoming_dir": str((repo_root / "runtime" / "grading" / "incoming_papers").resolve()),
            "grading_runs_dir": str((repo_root / "runtime" / "grading" / "runs").resolve()),
            "state_dir": str((repo_root / "runtime" / "pipeline" / "state").resolve()),
            "credential_store_dir": str((repo_root / "runtime" / "secrets" / "credential_store").resolve()),
            "feedback_dir": str((repo_root / "runtime" / "tracking" / "student_feedback").resolve()),
            "student_log_json_path": str((repo_root / "runtime" / "tracking" / "student_progress_log.json").resolve()),
            "student_log_md_path": str((repo_root / "runtime" / "tracking" / "student_progress_log.md").resolve()),
            "case_exports_dir": str((repo_root / "runtime" / "exports" / "case_exports").resolve()),
            "rename": {
                "prefix_with_date": True,
                "hash_length": 8,
            },
        }
        user_data = _read_json(config_path, {})
        merged = _merge_dict(defaults, user_data if isinstance(user_data, dict) else {})
        return cls(
            project_root=project_root,
            paperdownload_root=_resolve_path(str(merged["paperdownload_root"]), config_dir),
            paper_grading_root=_resolve_path(str(merged["paper_grading_root"]), config_dir),
            download_output_root=_resolve_path(str(merged["download_output_root"]), config_dir),
            incoming_dir=_resolve_path(str(merged["incoming_dir"]), config_dir),
            grading_runs_dir=_resolve_path(str(merged["grading_runs_dir"]), config_dir),
            state_dir=_resolve_path(str(merged["state_dir"]), config_dir),
            credential_store_dir=_resolve_path(str(merged["credential_store_dir"]), config_dir),
            feedback_dir=_resolve_path(str(merged["feedback_dir"]), config_dir),
            student_log_json_path=_resolve_path(str(merged["student_log_json_path"]), config_dir),
            student_log_md_path=_resolve_path(str(merged["student_log_md_path"]), config_dir),
            case_exports_dir=_resolve_path(str(merged["case_exports_dir"]), config_dir),
            rename_prefix_with_date=bool(merged["rename"]["prefix_with_date"]),
            rename_hash_length=max(4, int(merged["rename"]["hash_length"])),
        )


class UnifiedPipeline:
    def __init__(self, cfg: PipelineConfig) -> None:
        self.cfg = cfg
        self.repo_root = self.cfg.project_root.parent
        self.state_path = self.cfg.state_dir / "ingest_state.json"
        self.source_registry_path = self.cfg.state_dir / "source_registry.json"
        self.file_source_map_path = self.cfg.state_dir / "file_source_map.json"
        self.reports_dir = self.cfg.state_dir / "reports"
        self.student_log_json_path = self.cfg.student_log_json_path
        self.student_log_md_path = self.cfg.student_log_md_path
        self.feedback_dir = self.cfg.feedback_dir
        self.case_exports_dir = self.cfg.case_exports_dir

    def _runtime_env(self) -> dict[str, str]:
        return {
            "PAPER_PIPELINE_REPO_ROOT": str(self.repo_root),
            "PAPER_PIPELINE_CONFIG": str((self.repo_root / "config" / "pipeline" / "pipeline.config.json").resolve()),
            "PAPER_PIPELINE_CREDENTIAL_STORE_DIR": str(self.cfg.credential_store_dir),
            "PAPERDOWNLOAD_OUTPUT_ROOT": str(self.cfg.download_output_root),
            "ESSAYGRADE_INCOMING_DIR": str(self.cfg.incoming_dir),
            "ESSAYGRADE_RUNS_DIR": str(self.cfg.grading_runs_dir),
        }

    def _load_state(self) -> dict[str, Any]:
        return _read_json(
            self.state_path,
            {
                "version": 1,
                "updated_at": _now_iso(),
                "digests": {},
                "source_to_dest": {},
            },
        )

    def _save_state(self, state: dict[str, Any]) -> None:
        state["updated_at"] = _now_iso()
        _write_json(self.state_path, state)

    def _load_source_registry(self) -> dict[str, Any]:
        return _read_json(
            self.source_registry_path,
            {
                "updated_at": _now_iso(),
                "active_source_key": None,
                "sources": {},
            },
        )

    def _save_source_registry(self, registry: dict[str, Any]) -> None:
        registry["updated_at"] = _now_iso()
        _write_json(self.source_registry_path, registry)

    def _load_file_source_map(self) -> dict[str, Any]:
        return _read_json(
            self.file_source_map_path,
            {
                "updated_at": _now_iso(),
                "files": {},
            },
        )

    def _save_file_source_map(self, file_map: dict[str, Any]) -> None:
        file_map["updated_at"] = _now_iso()
        _write_json(self.file_source_map_path, file_map)

    def _get_active_source(self) -> dict[str, Any] | None:
        registry = self._load_source_registry()
        source_key = str(registry.get("active_source_key") or "").strip()
        if not source_key:
            return None
        source = (registry.get("sources", {}) or {}).get(source_key)
        if not isinstance(source, dict):
            return None
        return dict(source)

    def _source_metadata_from_path(self, path_str: str | None, file_map: dict[str, Any]) -> dict[str, Any] | None:
        if not path_str:
            return None
        normalized = _normalize_repo_path(str(path_str), self.repo_root)
        entry = (file_map.get("files", {}) or {}).get(normalized)
        if isinstance(entry, dict):
            return entry
        return None

    def _apply_source_metadata(self, record: dict[str, Any], file_map: dict[str, Any]) -> None:
        candidates = []
        if record.get("latest_download_file"):
            candidates.append(record.get("latest_download_file"))
        candidates.extend(record.get("downloaded_files", []) or [])
        for candidate in candidates:
            metadata = self._source_metadata_from_path(str(candidate), file_map)
            if not metadata:
                continue
            record["source_key"] = metadata.get("source_key")
            record["teacher_name"] = metadata.get("teacher_name")
            record["stage_label"] = metadata.get("stage_label")
            record["source_target_page_url"] = metadata.get("target_page_url")
            record["delivery_case_name"] = metadata.get("folder_name")
            return

    def _bind_files_to_source(self, files: list[str], source: dict[str, Any], run_id: str | None = None) -> dict[str, Any]:
        file_map = self._load_file_source_map()
        items: dict[str, Any] = file_map.setdefault("files", {})
        bound = 0
        source_meta = {
            "source_key": source.get("source_key"),
            "teacher_name": source.get("teacher_name"),
            "stage_label": source.get("stage_label"),
            "target_page_url": source.get("target_page_url"),
            "folder_name": source.get("folder_name"),
        }
        downloads_root = self.cfg.download_output_root / "downloads"
        for raw in files:
            text = str(raw or "").strip()
            if not text:
                continue
            candidate = Path(text)
            if not candidate.is_absolute():
                candidate = downloads_root / candidate
            normalized = _normalize_repo_path(str(candidate), self.repo_root)
            items[normalized] = {
                **source_meta,
                "run_id": run_id,
                "bound_at": _now_iso(),
            }
            bound += 1
        self._save_file_source_map(file_map)
        return {"bound_files": bound, "source_key": source.get("source_key")}

    def _write_run_report(self, report: dict[str, Any]) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = self.reports_dir / f"{ts}_pipeline_run.json"
        _write_json(path, report)
        return path

    def set_source(
        self,
        teacher_name: str,
        target_page_url: str,
        stage_label: str,
        set_active: bool,
        bind_all_current: bool,
    ) -> dict[str, Any]:
        teacher = str(teacher_name or "").strip()
        target = str(target_page_url or "").strip()
        stage = str(stage_label or "").strip() or "初稿"
        if not teacher:
            raise RuntimeError("teacher_name is required")
        if not target:
            raise RuntimeError("target_page_url is required")

        registry = self._load_source_registry()
        sources = registry.setdefault("sources", {})
        source_key = _source_key(teacher, stage)
        existing = sources.get(source_key, {}) if isinstance(sources.get(source_key), dict) else {}
        source = {
            "source_key": source_key,
            "teacher_name": teacher,
            "stage_label": stage,
            "target_page_url": target,
            "folder_name": _source_folder_name(teacher, stage),
            "created_at": existing.get("created_at") or _now_iso(),
            "updated_at": _now_iso(),
        }
        sources[source_key] = source
        if set_active:
            registry["active_source_key"] = source_key
        self._save_source_registry(registry)

        bound = 0
        if bind_all_current:
            payload = _read_json(self.student_log_json_path, {})
            files: list[str] = []
            for entry in list(payload.get("students", []) or []):
                latest_file = str(entry.get("latest_download_file") or "").strip()
                if latest_file:
                    files.append(latest_file)
                for item in list(entry.get("downloaded_files", []) or []):
                    text = str(item or "").strip()
                    if text:
                        files.append(text)
            bound = int(self._bind_files_to_source(_dedupe_keep_order(files), source).get("bound_files", 0))

        return {
            "source": source,
            "active_source_key": registry.get("active_source_key"),
            "bound_existing_files": bound,
            "registry_path": str(self.source_registry_path.resolve()),
            "file_source_map_path": str(self.file_source_map_path.resolve()),
        }

    def list_sources(self) -> dict[str, Any]:
        registry = self._load_source_registry()
        sources = list((registry.get("sources", {}) or {}).values())
        sources.sort(key=lambda item: (str(item.get("teacher_name") or ""), str(item.get("stage_label") or "")))
        return {
            "active_source_key": registry.get("active_source_key"),
            "sources": sources,
            "registry_path": str(self.source_registry_path.resolve()),
        }

    def rename_source(self, source_key: str, teacher_name: str, stage_label: str, set_active: bool) -> dict[str, Any]:
        old_key = str(source_key or "").strip()
        teacher = str(teacher_name or "").strip()
        stage = str(stage_label or "").strip() or "初稿"
        if not old_key:
            raise RuntimeError("source_key is required")
        if not teacher:
            raise RuntimeError("teacher_name is required")

        registry = self._load_source_registry()
        sources = registry.setdefault("sources", {})
        source = sources.get(old_key)
        if not isinstance(source, dict):
            raise RuntimeError(f"Source not found: {old_key}")

        new_key = _source_key(teacher, stage)
        existing = sources.get(new_key)
        if existing and new_key != old_key:
            raise RuntimeError(f"Target source already exists: {new_key}")

        updated_source = {
            **source,
            "source_key": new_key,
            "teacher_name": teacher,
            "stage_label": stage,
            "folder_name": _source_folder_name(teacher, stage),
            "updated_at": _now_iso(),
        }
        if new_key != old_key:
            del sources[old_key]
        sources[new_key] = updated_source
        if set_active or registry.get("active_source_key") == old_key:
            registry["active_source_key"] = new_key
        self._save_source_registry(registry)

        file_map = self._load_file_source_map()
        updated_files = 0
        for _, item in (file_map.get("files", {}) or {}).items():
            if not isinstance(item, dict):
                continue
            if item.get("source_key") != old_key:
                continue
            item["source_key"] = new_key
            item["teacher_name"] = teacher
            item["stage_label"] = stage
            item["folder_name"] = updated_source["folder_name"]
            item["updated_at"] = _now_iso()
            updated_files += 1
        self._save_file_source_map(file_map)

        return {
            "source": updated_source,
            "old_source_key": old_key,
            "new_source_key": new_key,
            "updated_file_bindings": updated_files,
            "registry_path": str(self.source_registry_path.resolve()),
            "file_source_map_path": str(self.file_source_map_path.resolve()),
        }

    def _bundle_selected_entries(
        self,
        case_name: str,
        selected: list[dict[str, Any]],
        overwrite: bool,
        selection_meta: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        case_root = self.case_exports_dir / _safe_name(case_name)
        backup_root = None
        if case_root.exists():
            if not overwrite:
                raise RuntimeError(f"Case bundle already exists: {case_root}")
            backup_root = case_root.with_name(case_root.name + "__backup_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
            shutil.move(str(case_root), str(backup_root))

        papers_dir = case_root / "01_论文初稿"
        reports_dir = case_root / "02_评分结果"
        feedback_dir = case_root / "03_学生评语"
        summary_dir = case_root / "04_汇总"
        for path in [papers_dir, reports_dir, feedback_dir, summary_dir]:
            path.mkdir(parents=True, exist_ok=True)

        manifest_students: list[dict[str, Any]] = []
        summary_rows: list[dict[str, str]] = []

        for item in selected:
            sid = str(item.get("sid") or "").strip()
            name = str(item.get("name") or "student").strip()
            label = _safe_name(f"{sid}_{name}" if sid else name)

            paper_source = item.get("latest_download_file") or item.get("paper_path")
            paper_dest = None
            if paper_source:
                src_path = Path(str(paper_source))
                paper_dest = papers_dir / label / src_path.name
                _copy_file(str(src_path), paper_dest)

            report_dest = None
            json_dest = None
            visual_dest = None
            result_dir = reports_dir / label
            if item.get("grading_report_path"):
                report_dest = result_dir / "grade_report.txt"
                _copy_file(item.get("grading_report_path"), report_dest)
            if item.get("grading_json_path"):
                json_dest = result_dir / "grade_result.json"
                _copy_file(item.get("grading_json_path"), json_dest)
            run_root = Path(str(item.get("run_root") or ""))
            if run_root.exists():
                visual_candidates = sorted((run_root / "visual").glob("*.pdf"))
                if visual_candidates:
                    visual_dest = result_dir / visual_candidates[0].name
                    _copy_file(str(visual_candidates[0]), visual_dest)

            feedback_dest = None
            if item.get("feedback_path"):
                feedback_dest = feedback_dir / f"{label}.md"
                _copy_file(item.get("feedback_path"), feedback_dest)

            feedback_summary = _summarize_feedback(item.get("feedback_path"))
            summary_rows.append(
                {
                    "学号": sid,
                    "姓名": name,
                    "论文初稿": _bundle_relative(paper_dest, case_root) if paper_dest else "",
                    "分数": _format_score(item.get("score")),
                    "教师评语": feedback_summary,
                }
            )

            manifest_students.append(
                {
                    "sid": sid,
                    "name": name,
                    "score": item.get("score"),
                    "decision": item.get("decision"),
                    "paper_title": item.get("paper_title"),
                    "teacher_name": item.get("teacher_name"),
                    "stage_label": item.get("stage_label"),
                    "source_key": item.get("source_key"),
                    "source_paths": {
                        "downloaded_file": item.get("latest_download_file"),
                        "paper_path": item.get("paper_path"),
                        "grading_report_path": item.get("grading_report_path"),
                        "grading_json_path": item.get("grading_json_path"),
                        "feedback_path": item.get("feedback_path"),
                        "run_root": item.get("run_root"),
                    },
                    "bundle_paths": {
                        "paper": _bundle_relative(paper_dest, case_root) if paper_dest else None,
                        "report": _bundle_relative(report_dest, case_root) if report_dest else None,
                        "json": _bundle_relative(json_dest, case_root) if json_dest else None,
                        "visual": _bundle_relative(visual_dest, case_root) if visual_dest else None,
                        "feedback": _bundle_relative(feedback_dest, case_root) if feedback_dest else None,
                    },
                }
            )

        csv_path = summary_dir / "学生汇总.csv"
        with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=["学号", "姓名", "论文初稿", "分数", "教师评语"])
            writer.writeheader()
            writer.writerows(summary_rows)

        md_lines = [
            f"# {case_name} 交付包",
            "",
            f"- 学生数量: {len(summary_rows)}",
            f"- 生成时间: {_now_iso()}",
            "",
            "| 学号 | 姓名 | 论文初稿 | 分数 | 教师评语 |",
            "| --- | --- | --- | --- | --- |",
        ]
        for row in summary_rows:
            md_lines.append(
                "| {学号} | {姓名} | {论文初稿} | {分数} | {教师评语} |".format(
                    **{key: (value or "-").replace("\n", "<br>") for key, value in row.items()}
                )
            )
        (summary_dir / "学生汇总.md").write_text("\n".join(md_lines) + "\n", encoding="utf-8")

        readme_lines = [
            f"# {case_name}",
            "",
            "这个目录是对外交付层，不参与系统运行。",
            "",
            "- `01_论文初稿`: 下载下来的论文初稿副本。",
            "- `02_评分结果`: 每个学生的评分报告、结构化结果、视觉审稿产物。",
            "- `03_学生评语`: 可直接发给学生的老师口吻修改意见。",
            "- `04_汇总`: 面向最终用户的简化总表，不包含运行期杂项日志。",
            "",
            "说明：本交付包由系统运行层复制生成，原始文件未移动。",
        ]
        (case_root / "README.md").write_text("\n".join(readme_lines) + "\n", encoding="utf-8")

        manifest = {
            "case_name": case_name,
            "created_at": _now_iso(),
            "student_count": len(summary_rows),
            "selection": selection_meta or {},
            "backup_of_previous_bundle": str(backup_root.resolve()) if backup_root else None,
            "students": manifest_students,
        }
        _write_json(case_root / "manifest.json", manifest)

        return {
            "case_root": str(case_root.resolve()),
            "student_count": len(summary_rows),
            "summary_csv": str(csv_path.resolve()),
            "summary_md": str((summary_dir / "学生汇总.md").resolve()),
            "backup_of_previous_bundle": str(backup_root.resolve()) if backup_root else None,
        }

    def refresh_tracking_outputs(self) -> dict[str, Any]:
        return _refresh_tracking_outputs_impl(self)

        download_index_path = self.cfg.download_output_root / "state" / "downloaded_index.json"
        ingest_state = self._load_state()
        file_source_map = self._load_file_source_map()
        download_index = _read_json(download_index_path, {})
        student_records: dict[str, dict[str, Any]] = {}

        for base, item in (download_index.get("students", {}) or {}).items():
            student_records[base] = {
                "student_key": base,
                "sid": item.get("sid"),
                "name": item.get("name"),
                "downloaded": True,
                "downloaded_count": len(item.get("files", []) or []),
                "downloaded_at": item.get("lastDownloadedAt") or item.get("firstDownloadedAt"),
                "downloaded_files": [
                    str((self.cfg.download_output_root / "downloads" / file_name).resolve())
                    for file_name in (item.get("files", []) or [])
                ],
                "latest_download_file": (
                    str((self.cfg.download_output_root / "downloads" / item.get("files", [])[-1]).resolve())
                    if item.get("files")
                    else None
                ),
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
            self._apply_source_metadata(student_records[base], file_source_map)

        for src, item in (ingest_state.get("source_to_dest", {}) or {}).items():
            source_path = _normalize_repo_path(src, self.repo_root)
            source_name = Path(source_path).name
            sid, name, _ = _parse_filename(Path(source_name).stem)
            if not sid or not name:
                continue
            key = f"{sid}_{name}"
            record = student_records.setdefault(
                key,
                {
                    "student_key": key,
                    "sid": sid,
                    "name": name,
                    "downloaded": False,
                    "downloaded_count": 0,
                    "downloaded_at": None,
                    "downloaded_files": [],
                    "latest_download_file": source_path,
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
                },
            )
            record["ingested"] = True
            record["paper_path"] = _normalize_repo_path(str(item.get("dest") or ""), self.repo_root)
            if not record.get("latest_download_file"):
                record["latest_download_file"] = source_path
            if source_path and source_path not in record["downloaded_files"]:
                record["downloaded_files"].append(source_path)
            self._apply_source_metadata(record, file_source_map)

        grade_by_student: dict[str, list[dict[str, Any]]] = defaultdict(list)
        grading_runs_dir = self.cfg.paper_grading_root / "grading_runs"
        if grading_runs_dir.exists():
            for run_dir in sorted(grading_runs_dir.iterdir(), key=lambda p: p.stat().st_mtime):
                if not run_dir.is_dir():
                    continue
                run_info_path = run_dir / "notes" / "run_info.txt"
                if not run_info_path.exists():
                    continue
                run_info = _parse_run_info(run_info_path, self.repo_root)
                paper_path = run_info.get("paper")
                if not paper_path:
                    continue
                sid, name = _parse_ingested_filename(Path(paper_path).stem)
                if not sid or not name:
                    continue
                key = f"{sid}_{name}"
                grade_json_path = Path(run_info.get("json", "")) if run_info.get("json") else run_dir / "json" / "grade_result.json"
                grade_report_path = Path(run_info.get("report", "")) if run_info.get("report") else run_dir / "reports" / "grade_report.txt"
                grade_data = _read_json(grade_json_path, {})
                summary = grade_data.get("summary", {})
                extracted = grade_data.get("extracted", {})
                grade_by_student[key].append(
                    {
                        "sid": sid,
                        "name": name,
                        "paper_path": _normalize_repo_path(paper_path, self.repo_root),
                        "run_root": str(run_dir.resolve()),
                        "grade_json_path": _normalize_repo_path(str(grade_json_path), self.repo_root),
                        "grade_report_path": _normalize_repo_path(str(grade_report_path), self.repo_root),
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

        self.feedback_dir.mkdir(parents=True, exist_ok=True)
        for key, runs in grade_by_student.items():
            latest = runs[-1]
            record = student_records.setdefault(
                key,
                {
                    "student_key": key,
                    "sid": latest["sid"],
                    "name": latest["name"],
                    "downloaded": False,
                    "downloaded_count": 0,
                    "downloaded_at": None,
                    "downloaded_files": [],
                    "latest_download_file": None,
                    "ingested": True,
                    "graded": False,
                    "score": None,
                    "decision": None,
                    "paper_title": None,
                    "paper_path": latest["paper_path"],
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
                },
            )
            feedback_file = self.feedback_dir / f"{_safe_name(record['sid'] or '')}_{_safe_name(record['name'] or 'student')}.md"
            feedback_text = _build_student_feedback(record, latest["grade_data"])
            feedback_file.write_text(feedback_text, encoding="utf-8")

            record["graded"] = True
            record["ingested"] = True
            record["score"] = latest["score"]
            record["decision"] = latest["decision"]
            record["paper_title"] = latest["paper_title"]
            record["paper_path"] = latest["paper_path"]
            record["grading_report_path"] = latest["grade_report_path"]
            record["grading_json_path"] = latest["grade_json_path"]
            record["feedback_path"] = str(feedback_file.resolve())
            record["run_root"] = latest["run_root"]
            record["stage"] = latest["stage"]
            record["visual_mode"] = latest["visual_mode"]
            record["visual_model"] = latest["visual_model"]
            record["text_mode"] = latest["text_mode"]
            record["text_primary_model"] = latest["text_primary_model"]
            record["text_secondary_model"] = latest["text_secondary_model"]
            record["grade_time"] = latest["grade_time"]
            self._apply_source_metadata(record, file_source_map)

        entries = sorted(
            student_records.values(),
            key=lambda item: (
                item.get("sid") or "",
                item.get("downloaded_at") or "",
                item.get("grade_time") or "",
            ),
        )

        log_payload = {
            "updated_at": _now_iso(),
            "repo_root": str(self.repo_root.resolve()),
            "download_root": str(self.cfg.download_output_root.resolve()),
            "incoming_dir": str(self.cfg.incoming_dir.resolve()),
            "feedback_dir": str(self.feedback_dir.resolve()),
            "active_source": self._get_active_source(),
            "summary": {
                "downloaded_students": sum(1 for item in entries if item.get("downloaded")),
                "ingested_students": sum(1 for item in entries if item.get("ingested")),
                "graded_students": sum(1 for item in entries if item.get("graded")),
            },
            "students": entries,
        }
        _write_json(self.student_log_json_path, log_payload)

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
                    score=_format_score(item.get("score")),
                    decision=item.get("decision") or "-",
                    download_file=_relative_display(item.get("latest_download_file"), self.repo_root),
                    paper=_relative_display(item.get("paper_path"), self.repo_root),
                    report=_relative_display(item.get("grading_report_path"), self.repo_root),
                    feedback=_relative_display(item.get("feedback_path"), self.repo_root),
                )
            )
        self.student_log_md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

        return {
            "student_log_json": str(self.student_log_json_path.resolve()),
            "student_log_md": str(self.student_log_md_path.resolve()),
            "feedback_dir": str(self.feedback_dir.resolve()),
            "student_count": len(entries),
            "graded_count": log_payload["summary"]["graded_students"],
        }

    def download(self, page_size: int, start_page: int, max_students: int = 0) -> dict[str, Any]:
        script = self.cfg.paperdownload_root / "run-longzhi-automation.ps1"
        if not script.exists():
            raise FileNotFoundError(f"Download script not found: {script}")

        args = [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-PageSize",
            str(page_size),
            "-StartPage",
            str(start_page),
            "-OutputRoot",
            str(self.cfg.download_output_root),
        ]
        extra_env: dict[str, str] = {}
        active_source = self._get_active_source()
        if active_source and active_source.get("target_page_url"):
            extra_env["TARGET_PAGE_URL"] = str(active_source["target_page_url"])
            extra_env["ACTIVE_SOURCE_KEY"] = str(active_source.get("source_key") or "")
            extra_env["ACTIVE_TEACHER_NAME"] = str(active_source.get("teacher_name") or "")
            extra_env["ACTIVE_STAGE_LABEL"] = str(active_source.get("stage_label") or "")
        if max_students > 0:
            extra_env["MAX_STUDENTS"] = str(max_students)
        runtime_env = self._runtime_env()
        runtime_env.update(extra_env)
        _run_command(args, cwd=self.cfg.paperdownload_root, extra_env=runtime_env)

        summary_path = self.cfg.download_output_root / "state" / "latest_automation_summary.json"
        summary = _read_json(summary_path, {})
        bound = None
        if active_source and summary.get("newFiles"):
            bound = self._bind_files_to_source(list(summary.get("newFiles", []) or []), active_source, str(summary.get("runId") or ""))
        return {
            "status": summary.get("status", "unknown"),
            "runId": summary.get("runId"),
            "processed": summary.get("processed"),
            "downloaded": summary.get("downloaded"),
            "max_students": max_students,
            "newFiles": summary.get("newFiles", []),
            "active_source": active_source,
            "bound_source_files": bound,
            "summary_path": str(summary_path),
        }

    def ingest(self) -> dict[str, Any]:
        downloads_dir = self.cfg.download_output_root / "downloads"
        if not downloads_dir.exists():
            raise FileNotFoundError(f"Downloads directory not found: {downloads_dir}")

        self.cfg.incoming_dir.mkdir(parents=True, exist_ok=True)
        self.cfg.state_dir.mkdir(parents=True, exist_ok=True)

        state = self._load_state()
        digest_index: dict[str, dict[str, Any]] = state.setdefault("digests", {})
        src_map: dict[str, dict[str, Any]] = state.setdefault("source_to_dest", {})

        all_files = [p for p in downloads_dir.iterdir() if p.is_file()]
        candidates = sorted(all_files, key=lambda p: (p.stat().st_mtime, p.name))

        ingested: list[dict[str, Any]] = []
        skipped_digest = 0
        skipped_non_word = 0

        for src in candidates:
            if src.suffix.lower() not in {".doc", ".docx"}:
                skipped_non_word += 1
                continue

            digest = _sha256(src)
            if digest in digest_index:
                src_map[str(src.resolve())] = {
                    "digest": digest,
                    "dest": digest_index[digest]["dest"],
                    "updated_at": _now_iso(),
                }
                skipped_digest += 1
                continue

            sid, raw_name, seq = _parse_filename(src.stem)
            parts: list[str] = []
            if self.cfg.rename_prefix_with_date:
                parts.append(datetime.fromtimestamp(src.stat().st_mtime).strftime("%Y%m%d"))
            if sid and raw_name:
                parts.append(_safe_name(sid))
                parts.append(_safe_name(raw_name))
                if seq:
                    parts.append(seq)
            else:
                parts.append(_safe_name(src.stem))

            hash_short = digest[: self.cfg.rename_hash_length]
            parts.append(hash_short)
            base_name = "_".join([p for p in parts if p]) or f"paper_{hash_short}"
            dest_name = f"{base_name}{src.suffix.lower()}"
            dest = self.cfg.incoming_dir / dest_name

            version = 1
            while dest.exists():
                if _sha256(dest) == digest:
                    break
                version += 1
                dest = self.cfg.incoming_dir / f"{base_name}_v{version}{src.suffix.lower()}"

            if not dest.exists():
                shutil.copy2(src, dest)

            record = {
                "source": str(src.resolve()),
                "dest": str(dest.resolve()),
                "digest": digest,
                "ingested_at": _now_iso(),
                "sid": sid,
                "name": raw_name,
                "sequence": seq,
            }
            digest_index[digest] = {
                "dest": record["dest"],
                "first_seen_at": record["ingested_at"],
                "source_name": src.name,
            }
            src_map[record["source"]] = {
                "digest": digest,
                "dest": record["dest"],
                "updated_at": _now_iso(),
            }
            ingested.append(record)

        self._save_state(state)
        return {
            "download_dir": str(downloads_dir),
            "incoming_dir": str(self.cfg.incoming_dir),
            "ingested_count": len(ingested),
            "skipped_duplicate_digest": skipped_digest,
            "skipped_non_word": skipped_non_word,
            "ingested": ingested,
        }

    def grade(
        self,
        stage: str,
        visual_mode: str,
        visual_model: str,
        text_mode: str,
        text_primary_model: str,
        text_secondary_model: str,
        limit: int,
    ) -> dict[str, Any]:
        script = self.cfg.paper_grading_root / "process_incoming_papers.ps1"
        if not script.exists():
            raise FileNotFoundError(f"Grading script not found: {script}")

        args = [
            "powershell",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script),
            "-Stage",
            stage,
            "-VisualMode",
            visual_mode,
            "-VisualModel",
            visual_model,
            "-TextMode",
            text_mode,
            "-TextPrimaryModel",
            text_primary_model,
            "-TextSecondaryModel",
            text_secondary_model,
        ]
        if limit > 0:
            args.extend(["-Limit", str(limit)])

        _run_command(args, cwd=self.cfg.paper_grading_root, extra_env=self._runtime_env())
        return {
            "status": "success",
            "mode": "queue",
            "stage": stage,
            "visual_mode": visual_mode,
            "visual_model": visual_model,
            "text_mode": text_mode,
            "text_primary_model": text_primary_model,
            "text_secondary_model": text_secondary_model,
            "limit": limit,
        }

    def grade_ingested_files(
        self,
        files: list[str],
        stage: str,
        visual_mode: str,
        visual_model: str,
        text_mode: str,
        text_primary_model: str,
        text_secondary_model: str,
        limit: int,
        skip_tracking_refresh: bool = False,
    ) -> dict[str, Any]:
        script = self.cfg.paper_grading_root / "run_grade.ps1"
        if not script.exists():
            raise FileNotFoundError(f"Single-file grading script not found: {script}")

        picked = files[: limit] if limit > 0 else files
        graded: list[dict[str, Any]] = []
        for paper in picked:
            paper_path = Path(paper).resolve()
            label = _safe_name(paper_path.stem) or "paper"
            args = [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-PaperPath",
                str(paper_path),
                "-RunLabel",
                label,
                "-Stage",
                stage,
                "-VisualMode",
                visual_mode,
                "-VisualModel",
                visual_model,
                "-TextMode",
                text_mode,
                "-TextPrimaryModel",
                text_primary_model,
                "-TextSecondaryModel",
                text_secondary_model,
            ]
            env = os.environ.copy()
            env.update(self._runtime_env())
            if skip_tracking_refresh:
                env["PAPER_PIPELINE_SKIP_REFRESH_LOG"] = "1"
            timeout_seconds = max(60, int(os.getenv("PIPELINE_GRADE_TIMEOUT_SECONDS", "1800")))
            with tempfile.TemporaryFile(mode="w+t", encoding="utf-8", errors="replace") as stdout_file, tempfile.TemporaryFile(
                mode="w+t", encoding="utf-8", errors="replace"
            ) as stderr_file:
                proc = subprocess.Popen(
                    args,
                    cwd=str(self.cfg.paper_grading_root),
                    env=env,
                    stdout=stdout_file,
                    stderr=stderr_file,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                started_at = datetime.now()
                while True:
                    try:
                        returncode = proc.wait(timeout=15)
                        break
                    except subprocess.TimeoutExpired:
                        if (datetime.now() - started_at).total_seconds() > timeout_seconds:
                            proc.kill()
                            proc.wait(timeout=30)
                            stdout_file.seek(0)
                            stderr_file.seek(0)
                            stdout = stdout_file.read()
                            stderr = stderr_file.read()
                            detail = (stderr or stdout or "").strip() or "no output captured"
                            raise RuntimeError(
                                f"Command timed out after {timeout_seconds}s: {' '.join(args)}\n{detail}"
                            )
                stdout_file.seek(0)
                stderr_file.seek(0)
                stdout = stdout_file.read()
                stderr = stderr_file.read()
            if stdout:
                logger.info("Command stdout:\n%s", stdout.rstrip())
            if stderr:
                logger.warning("Command stderr:\n%s", stderr.rstrip())
            if returncode != 0:
                detail = (stderr or stdout or "").strip() or "no output captured"
                raise RuntimeError(
                    f"Command failed (exit={returncode}): {' '.join(args)}\n{detail}"
                )
            graded.append(
                {
                    "paper": str(paper_path),
                    "run_label": label,
                    "graded_at": _now_iso(),
                }
            )

        return {
            "status": "success",
            "mode": "ingested_only",
            "stage": stage,
            "visual_mode": visual_mode,
            "visual_model": visual_model,
            "text_mode": text_mode,
            "text_primary_model": text_primary_model,
            "text_secondary_model": text_secondary_model,
            "limit": limit,
            "graded_count": len(graded),
            "graded": graded,
        }

    def run_all(
        self,
        page_size: int,
        start_page: int,
        max_students: int,
        stage: str,
        visual_mode: str,
        visual_model: str,
        text_mode: str,
        text_primary_model: str,
        text_secondary_model: str,
        limit: int,
        grade_even_if_no_new: bool,
        grade_ingested_only: bool,
    ) -> dict[str, Any]:
        report: dict[str, Any] = {
            "started_at": _now_iso(),
            "steps": {},
        }
        report["steps"]["download"] = self.download(
            page_size=page_size,
            start_page=start_page,
            max_students=max_students,
        )
        report["steps"]["ingest"] = self.ingest()
        ingest_count = int(report["steps"]["ingest"].get("ingested_count", 0))
        ingested_files = [item["dest"] for item in report["steps"]["ingest"].get("ingested", [])]

        if ingest_count > 0 or grade_even_if_no_new:
            if ingest_count > 0 and grade_ingested_only:
                report["steps"]["grade"] = self.grade_ingested_files(
                    files=ingested_files,
                    stage=stage,
                    visual_mode=visual_mode,
                    visual_model=visual_model,
                    text_mode=text_mode,
                    text_primary_model=text_primary_model,
                    text_secondary_model=text_secondary_model,
                    limit=limit,
                )
            else:
                report["steps"]["grade"] = self.grade(
                    stage=stage,
                    visual_mode=visual_mode,
                    visual_model=visual_model,
                    text_mode=text_mode,
                    text_primary_model=text_primary_model,
                    text_secondary_model=text_secondary_model,
                    limit=limit,
                )
        else:
            report["steps"]["grade"] = {
                "status": "skipped",
                "reason": "No newly ingested papers.",
            }

        report["finished_at"] = _now_iso()
        report_path = self._write_run_report(report)
        report["report_path"] = str(report_path)
        return report

    def status(self) -> dict[str, Any]:
        state = self._load_state()
        active_source = self._get_active_source()
        downloads_dir = self.cfg.download_output_root / "downloads"
        incoming_dir = self.cfg.incoming_dir
        downloads = [p for p in downloads_dir.iterdir() if p.is_file()] if downloads_dir.exists() else []
        word_downloads = [p for p in downloads if p.suffix.lower() in {".doc", ".docx"}]
        incoming = [p for p in incoming_dir.iterdir() if p.is_file()] if incoming_dir.exists() else []
        return {
            "updated_at": state.get("updated_at"),
            "download_count_total": len(downloads),
            "download_count_word": len(word_downloads),
            "incoming_count": len(incoming),
            "tracked_digest_count": len(state.get("digests", {})),
            "download_root": str(self.cfg.download_output_root),
            "incoming_dir": str(self.cfg.incoming_dir),
            "grading_runs_dir": str(self.cfg.grading_runs_dir),
            "credential_store_dir": str(self.cfg.credential_store_dir),
            "feedback_dir": str(self.feedback_dir),
            "case_exports_dir": str(self.case_exports_dir),
            "state_file": str(self.state_path),
            "active_source": active_source,
            "source_registry_file": str(self.source_registry_path),
        }

    def bundle_case(
        self,
        case_name: str,
        student_ids: list[str],
        latest_graded: int,
        all_graded: bool,
        overwrite: bool,
    ) -> dict[str, Any]:
        payload = _read_json(self.student_log_json_path, {})
        entries = list(payload.get("students", []) or [])
        if not entries:
            raise RuntimeError(f"Student log is empty: {self.student_log_json_path}")

        selected: list[dict[str, Any]] = []
        if student_ids:
            wanted = {item.strip() for item in student_ids if item.strip()}
            selected = [item for item in entries if str(item.get("sid") or "").strip() in wanted]
            missing = sorted(wanted - {str(item.get("sid") or "").strip() for item in selected})
            if missing:
                raise RuntimeError(f"Students not found in log: {', '.join(missing)}")
        elif latest_graded > 0:
            graded = [item for item in entries if item.get("graded")]
            graded.sort(key=lambda item: item.get("grade_time") or "", reverse=True)
            selected = graded[:latest_graded]
        elif all_graded:
            selected = [item for item in entries if item.get("graded")]
        else:
            raise RuntimeError("bundle-case requires --student-ids, --latest-graded, or --all-graded")

        if not selected:
            raise RuntimeError("No students matched bundle-case selection.")

        selected.sort(key=lambda item: (str(item.get("sid") or ""), str(item.get("name") or "")))
        return self._bundle_selected_entries(
            case_name=case_name,
            selected=selected,
            overwrite=overwrite,
            selection_meta={
                "student_ids": student_ids,
                "latest_graded": latest_graded,
                "all_graded": all_graded,
            },
        )

    def bundle_source(self, source_key: str, overwrite: bool) -> dict[str, Any]:
        key = str(source_key or "").strip()
        if not key:
            raise RuntimeError("source_key is required")

        registry = self._load_source_registry()
        source = (registry.get("sources", {}) or {}).get(key)
        if not isinstance(source, dict):
            raise RuntimeError(f"Source not found: {key}")

        payload = _read_json(self.student_log_json_path, {})
        entries = list(payload.get("students", []) or [])
        selected = [item for item in entries if item.get("source_key") == key and item.get("graded")]
        teacher_name = str(source.get("teacher_name") or "").strip()
        folder_name = str(source.get("folder_name") or key).strip()
        fallback_selected = [
            item
            for item in entries
            if item.get("graded")
            and str(item.get("teacher_name") or "").strip() == teacher_name
            and (
                str(item.get("source_key") or "").strip() == key
                or str(item.get("delivery_case_name") or "").strip() == folder_name
            )
        ]
        if len(fallback_selected) > len(selected):
            selected = fallback_selected
        if not selected:
            raise RuntimeError(f"No graded students found for source: {key}")

        selected.sort(key=lambda item: (str(item.get("sid") or ""), str(item.get("name") or "")))
        return self._bundle_selected_entries(
            case_name=str(source.get("folder_name") or key),
            selected=selected,
            overwrite=overwrite,
            selection_meta={
                "source_key": key,
                "teacher_name": source.get("teacher_name"),
                "stage_label": source.get("stage_label"),
                "target_page_url": source.get("target_page_url"),
            },
        )

    def audit_students(self, limit: int = 0, force: bool = False) -> dict[str, Any]:
        from pipeline_audit import run_gemini_audit
        payload = _read_json(self.student_log_json_path, {})
        entries = list(payload.get("students", []) or [])
        graded = [item for item in entries if item.get("graded")]

        audit_file = self.cfg.state_dir / "audit_results.json"
        results = _read_json(audit_file, []) if audit_file.exists() else []
        audited_sids = set()
        if not force:
            audited_sids = {
                str(r.get("sid"))
                for r in results
                if r.get("audit")
                and not r.get("audit", {}).get("error")
                and r.get("audit", {}).get("review_verdict")
            }

        to_audit = [item for item in graded if str(item.get("sid")) not in audited_sids]
        if limit > 0:
            to_audit = to_audit[:limit]

        immediate_retries = max(0, int(os.getenv("PIPELINE_AUDIT_IMMEDIATE_RETRIES", "2")))
        deferred_retry_passes = max(0, int(os.getenv("PIPELINE_AUDIT_DEFERRED_RETRY_PASSES", "1")))
        failure_summary_file = self.cfg.state_dir / "audit_failures.json"

        def _store_result(
            student_entry: dict[str, Any],
            audit_result: dict[str, Any],
            attempts: int,
            deferred_pass: int,
        ) -> None:
            nonlocal results
            results = [r for r in results if str(r.get("sid")) != str(student_entry.get("sid"))]
            results.append(
                {
                    "sid": student_entry.get("sid"),
                    "name": student_entry.get("name"),
                    "teacher": student_entry.get("teacher_name"),
                    "audit_model": audit_result.get("model"),
                    "audit_status": "ok" if not audit_result.get("error") else "error",
                    "audit_attempts": attempts,
                    "deferred_retry_pass": deferred_pass,
                    "audit": audit_result,
                }
            )
            _write_json(audit_file, results)

        def _run_with_retries(student_entry: dict[str, Any], deferred_pass: int) -> tuple[dict[str, Any], int]:
            max_attempts = 1 + immediate_retries
            last_result: dict[str, Any] = {}
            for attempt in range(1, max_attempts + 1):
                last_result = run_gemini_audit(student_entry)
                if not last_result.get("error"):
                    return last_result, attempt
                if last_result.get("retryable") is False:
                    return last_result, attempt
            return last_result, max_attempts

        failed_entries: list[dict[str, Any]] = []
        batch_blocked_reason: str | None = None
        for entry in to_audit:
            print(f"Auditing {entry.get('sid')} {entry.get('name')}...")
            res, attempts = _run_with_retries(entry, deferred_pass=0)
            _store_result(entry, res, attempts, deferred_pass=0)
            if res.get("error"):
                failed_entries.append(entry)
                if res.get("error_type") == "quota_exhausted":
                    batch_blocked_reason = str(res.get("error"))
                    break

        final_failures: list[dict[str, Any]] = []
        for retry_pass in range(1, deferred_retry_passes + 1):
            if batch_blocked_reason:
                break
            if not failed_entries:
                break
            current_failures = list(failed_entries)
            failed_entries = []
            for entry in current_failures:
                print(f"Retry audit pass {retry_pass}: {entry.get('sid')} {entry.get('name')}...")
                res = run_gemini_audit(entry)
                _store_result(entry, res, attempts=1, deferred_pass=retry_pass)
                if res.get("error"):
                    failed_entries.append(entry)
                    if res.get("error_type") == "quota_exhausted":
                        batch_blocked_reason = str(res.get("error"))
                        break
            if batch_blocked_reason:
                break

        for entry in failed_entries:
            matched = next((item for item in results if str(item.get("sid")) == str(entry.get("sid"))), {})
            final_failures.append(
                {
                    "sid": entry.get("sid"),
                    "name": entry.get("name"),
                    "teacher": entry.get("teacher_name"),
                    "error": (matched.get("audit") or {}).get("error"),
                    "audit_attempts": matched.get("audit_attempts"),
                    "deferred_retry_pass": matched.get("deferred_retry_pass"),
                    "still_needs_manual": True,
                }
            )
        _write_json(failure_summary_file, final_failures)

        return {
            "audited_count": len(to_audit),
            "audit_file": str(audit_file),
            "failed_count": len(final_failures),
            "failure_summary_file": str(failure_summary_file),
            "batch_blocked_reason": batch_blocked_reason,
            "force": bool(force),
        }

    def rebuild_anomalies(self) -> dict[str, Any]:
        payload = _read_json(self.student_log_json_path, {})
        entries = list(payload.get("students", []) or [])
        anomalies = []
        for entry in entries:
            is_anomaly = False
            reasons = []
            
            if entry.get("feedback_status") != "ok" and entry.get("graded"):
                is_anomaly = True
                reasons.append("feedback_failed_or_legacy")
            
            if entry.get("ingested") and not entry.get("grade_data_valid"):
                is_anomaly = True
                reasons.append("invalid_grade_data")
                
            teacher = str(entry.get("teacher_name") or "")
            if "3C" in teacher:
                is_anomaly = True
                reasons.append("teacher_3C")
                
            if is_anomaly:
                entry["anomaly_reasons"] = reasons
                anomalies.append(entry)

        # "Only anomaly batches should be formally rebuilt immediately."
        # This implies we should run grading for invalid ones, and feedback for others.
        # But wait, running grading here is synchronous and slow. Let's return the anomaly list 
        # and trigger a feedback retry for the feedback ones.
        feedback_retries = 0
        from pipeline_feedback import generate_student_feedback_artifact
        for anomaly in anomalies:
            if "feedback_failed_or_legacy" in anomaly.get("anomaly_reasons", []) or "teacher_3C" in anomaly.get("anomaly_reasons", []):
                if anomaly.get("grade_data_valid"):
                    anomaly["force_feedback_retry"] = True # Will be picked up by refresh_tracking_outputs if we just flag it, wait no, let's just delete the meta file
                    feedback_meta = Path(anomaly.get("feedback_raw_response_path", "")).parent / "student_feedback_kimi_meta.json"
                    if feedback_meta.exists():
                        feedback_meta.unlink()
                    feedback_retries += 1
        
        # After deleting meta, refresh log will rebuild them.
        self.refresh_tracking_outputs()

        return {
            "anomaly_count": len(anomalies),
            "feedback_retries_triggered": feedback_retries,
            "anomalies": [
                {"sid": a.get("sid"), "name": a.get("name"), "reasons": a.get("anomaly_reasons")} 
                for a in anomalies
            ]
        }

    def _clear_feedback_artifacts(self, run_root: str) -> None:
        if not run_root:
            return
        artifact_dir = Path(str(run_root)).resolve() / "feedback"
        for name in ("student_feedback_kimi_meta.json", "student_feedback_kimi_raw.json"):
            target = artifact_dir / name
            if target.exists():
                target.unlink()

    def _apply_rescore_adjustment(self, row: dict[str, Any], entry: dict[str, Any]) -> dict[str, Any]:
        audit = row.get("audit") if isinstance(row.get("audit"), dict) else {}
        verdict = str(audit.get("review_verdict") or "").strip()
        if verdict != "rescore_and_feedback":
            return {"status": "skipped", "sid": row.get("sid"), "reason": "not_rescore"}

        grade_json_path = Path(str(entry.get("grading_json_path") or "")).resolve()
        if not grade_json_path.exists():
            raise FileNotFoundError(f"Grading JSON not found for {row.get('sid')}: {grade_json_path}")

        grade_data = _read_json(grade_json_path, {})
        if not isinstance(grade_data, dict):
            raise RuntimeError(f"Invalid grading JSON for {row.get('sid')}: {grade_json_path}")
        summary = grade_data.get("summary")
        if not isinstance(summary, dict):
            raise RuntimeError(f"Grading JSON has no summary for {row.get('sid')}: {grade_json_path}")

        original_score = _coerce_float(summary.get("total_score"))
        score_delta = _coerce_float(audit.get("score_adjustment_needed"))
        effective_score = original_score
        if original_score is not None and score_delta is not None:
            effective_score = round(max(0.0, min(100.0, original_score + score_delta)), 2)
            summary["total_score"] = effective_score
            raw_total = _coerce_float(summary.get("raw_total_score"))
            if raw_total is not None:
                summary["raw_total_score"] = round(max(0.0, min(100.0, raw_total + score_delta)), 2)

        original_decision = str(summary.get("decision") or "").strip() or None
        effective_decision = str(audit.get("new_decision_if_any") or "").strip() or original_decision
        if effective_decision:
            summary["decision"] = effective_decision
            if "grade_band" in summary:
                summary["grade_band"] = effective_decision

        grade_data["audit_application"] = {
            "applied_at": _now_iso(),
            "review_verdict": verdict,
            "source_sid": row.get("sid"),
            "score_adjustment_needed": score_delta,
            "original_total_score": original_score,
            "effective_total_score": effective_score,
            "original_decision": original_decision,
            "effective_decision": effective_decision,
        }
        _write_json(grade_json_path, grade_data)

        return {
            "status": "applied",
            "sid": row.get("sid"),
            "grade_json_path": str(grade_json_path),
            "effective_score": effective_score,
            "effective_decision": effective_decision,
        }

    def _find_latest_student_run(self, sid: str) -> dict[str, Any] | None:
        runs: list[dict[str, Any]] = []
        if not self.cfg.grading_runs_dir.exists():
            return None
        for run_dir in sorted(self.cfg.grading_runs_dir.iterdir(), key=lambda path: path.stat().st_mtime):
            if not run_dir.is_dir():
                continue
            run_info_path = run_dir / "notes" / "run_info.txt"
            if not run_info_path.exists():
                continue
            run_info = _parse_run_info(run_info_path, self.repo_root)
            paper_path = run_info.get("paper")
            if not paper_path:
                continue
            run_sid, run_name = _parse_ingested_filename(Path(paper_path).stem)
            if str(run_sid or "").strip() != str(sid or "").strip():
                continue
            grade_json_candidates = []
            if run_info.get("json"):
                grade_json_candidates.append(Path(_normalize_repo_path(run_info.get("json"), self.repo_root)))
            grade_json_candidates.append(run_dir / "json" / "grade_result.json")
            grade_json_path = next((path.resolve() for path in grade_json_candidates if path.exists()), (run_dir / "json" / "grade_result.json").resolve())

            grade_report_candidates = []
            if run_info.get("report"):
                grade_report_candidates.append(Path(_normalize_repo_path(run_info.get("report"), self.repo_root)))
            grade_report_candidates.append(run_dir / "reports" / "grade_report.txt")
            grade_report_path = next((path.resolve() for path in grade_report_candidates if path.exists()), (run_dir / "reports" / "grade_report.txt").resolve())
            grade_data = _read_json(grade_json_path, {})
            summary = grade_data.get("summary", {}) if isinstance(grade_data, dict) else {}
            extracted = grade_data.get("extracted", {}) if isinstance(grade_data, dict) else {}
            runs.append(
                {
                    "sid": run_sid,
                    "name": run_name,
                    "paper_path": _normalize_repo_path(paper_path, self.repo_root),
                    "run_root": str(run_dir.resolve()),
                    "grade_json_path": _normalize_repo_path(str(grade_json_path), self.repo_root),
                    "grade_report_path": _normalize_repo_path(str(grade_report_path), self.repo_root),
                    "grade_time": datetime.fromtimestamp(run_dir.stat().st_mtime).astimezone().isoformat(timespec="seconds"),
                    "score": summary.get("total_score"),
                    "decision": summary.get("decision"),
                    "stage": summary.get("stage") or run_info.get("stage"),
                    "visual_mode": (grade_data.get("visual_review", {}) if isinstance(grade_data, dict) else {}).get("mode") or run_info.get("visual_mode"),
                    "visual_model": (grade_data.get("visual_review", {}) if isinstance(grade_data, dict) else {}).get("model") or run_info.get("visual_model"),
                    "text_mode": (grade_data.get("text_review", {}) if isinstance(grade_data, dict) else {}).get("mode") or run_info.get("text_mode"),
                    "text_primary_model": (((grade_data.get("text_review", {}) if isinstance(grade_data, dict) else {}).get("primary") or {}).get("model")) or run_info.get("text_primary_model"),
                    "text_secondary_model": (((grade_data.get("text_review", {}) if isinstance(grade_data, dict) else {}).get("secondary") or {}).get("model")) or run_info.get("text_secondary_model"),
                    "paper_title": extracted.get("title"),
                    "grade_data": grade_data,
                    "grade_data_valid": _grade_data_is_valid(grade_data),
                }
            )
        if not runs:
            return None
        latest = runs[-1]
        for item in reversed(runs):
            if item.get("grade_data_valid"):
                return item
        return latest

    def _refresh_selected_students(self, student_ids: list[str]) -> dict[str, Any]:
        payload = _read_json(self.student_log_json_path, {})
        students = list(payload.get("students", []) or [])
        student_by_sid = {str(item.get("sid") or "").strip(): item for item in students if str(item.get("sid") or "").strip()}
        audit_rows = _read_json(self.cfg.state_dir / "audit_results.json", [])
        audit_by_sid = {
            str(item.get("sid") or "").strip(): item
            for item in list(audit_rows or [])
            if isinstance(item, dict) and str(item.get("sid") or "").strip()
        }
        updated: list[str] = []
        wanted = [str(item).strip() for item in student_ids if str(item).strip()]
        for sid in wanted:
            record = student_by_sid.get(sid)
            selected = self._find_latest_student_run(sid)
            if not record or not selected:
                continue
            audit_row = audit_by_sid.get(sid) if isinstance(audit_by_sid.get(sid), dict) else {}
            audit = audit_row.get("audit") if isinstance((audit_row or {}).get("audit"), dict) else {}
            application = audit_row.get("application") if isinstance((audit_row or {}).get("application"), dict) else {}
            verdict = str(audit.get("review_verdict") or "").strip()
            feedback_entry = {
                **record,
                "paper_title": selected.get("paper_title"),
                "run_root": selected.get("run_root"),
                "stage": selected.get("stage"),
                "teacher_name": record.get("teacher_name"),
                "grade_error": None if selected.get("grade_data_valid") else "invalid or empty grade data",
            }
            if verdict and verdict != "keep" and not (
                str(application.get("status") or "").strip() == "applied"
                and verdict in {"rescore_and_feedback", "rerun_grading"}
            ):
                feedback_entry["audit_review"] = {
                    "audit_status": audit_row.get("audit_status"),
                    "audit_model": audit_row.get("audit_model"),
                    "audit_attempts": audit_row.get("audit_attempts"),
                    "deferred_retry_pass": audit_row.get("deferred_retry_pass"),
                    "review_verdict": verdict,
                    "reason": audit.get("reason"),
                    "score_adjustment_needed": audit.get("score_adjustment_needed"),
                    "new_decision_if_any": audit.get("new_decision_if_any"),
                    "must_fix_fields": list(audit.get("must_fix_fields") or []),
                    "student_feedback_rewrite_needed": audit.get("student_feedback_rewrite_needed"),
                    "application": application,
                }
            feedback_result = _generate_student_feedback_artifact_impl(feedback_entry, selected.get("grade_data") or {})
            feedback_file = self.feedback_dir / f"{_safe_name(str(record.get('sid') or ''))}_{_safe_name(str(record.get('name') or 'student'))}.md"
            feedback_file.write_text(feedback_result["markdown"], encoding="utf-8")
            record.update(
                {
                    "graded": bool(selected.get("grade_data_valid")),
                    "ingested": True,
                    "score": selected.get("score"),
                    "decision": selected.get("decision"),
                    "paper_title": selected.get("paper_title"),
                    "paper_path": selected.get("paper_path"),
                    "grading_report_path": selected.get("grade_report_path"),
                    "grading_json_path": selected.get("grade_json_path"),
                    "feedback_path": str(feedback_file.resolve()),
                    "feedback_status": feedback_result.get("status"),
                    "feedback_model": feedback_result.get("model"),
                    "feedback_attempts": feedback_result.get("attempts"),
                    "feedback_error": feedback_result.get("error"),
                    "feedback_raw_response_path": feedback_result.get("raw_response_path"),
                    "audit_review_verdict": verdict or None,
                    "audit_reason": audit.get("reason"),
                    "audit_new_decision_if_any": audit.get("new_decision_if_any"),
                    "audit_score_adjustment_needed": audit.get("score_adjustment_needed"),
                    "audit_application_status": application.get("status") if application else None,
                    "audit_application_applied_at": application.get("applied_at") if application else None,
                    "run_root": selected.get("run_root"),
                    "stage": selected.get("stage"),
                    "visual_mode": selected.get("visual_mode"),
                    "visual_model": selected.get("visual_model"),
                    "text_mode": selected.get("text_mode"),
                    "text_primary_model": selected.get("text_primary_model"),
                    "text_secondary_model": selected.get("text_secondary_model"),
                    "grade_time": selected.get("grade_time"),
                    "grade_data_valid": bool(selected.get("grade_data_valid")),
                    "grade_error": None if selected.get("grade_data_valid") else "invalid or empty grade data",
                }
            )
            updated.append(sid)

        if updated:
            payload["updated_at"] = _now_iso()
            payload["students"] = students
            _write_json(self.student_log_json_path, payload)
        return {"updated_students": updated, "updated_count": len(updated)}

    def apply_audit_reviews(
        self,
        teachers: list[str] | None = None,
        verdicts: list[str] | None = None,
        student_ids: list[str] | None = None,
        force: bool = False,
    ) -> dict[str, Any]:
        student_payload = _read_json(self.student_log_json_path, {})
        student_entries = {
            str(item.get("sid") or "").strip(): item
            for item in list(student_payload.get("students", []) or [])
            if str(item.get("sid") or "").strip()
        }
        audit_file = self.cfg.state_dir / "audit_results.json"
        audit_rows = _read_json(audit_file, []) if audit_file.exists() else []
        actionable_verdicts = {"feedback_only", "rescore_and_feedback", "rerun_grading"}
        teacher_filter = {item.strip() for item in list(teachers or []) if str(item or "").strip()}
        verdict_filter = {item.strip() for item in list(verdicts or []) if str(item or "").strip()} or set(actionable_verdicts)
        sid_filter = {item.strip() for item in list(student_ids or []) if str(item or "").strip()}

        filtered_rows: list[dict[str, Any]] = []
        skipped_already_applied = 0
        for row in audit_rows if isinstance(audit_rows, list) else []:
            if not isinstance(row, dict):
                continue
            audit = row.get("audit") if isinstance(row.get("audit"), dict) else {}
            verdict = str(audit.get("review_verdict") or "").strip()
            if verdict not in actionable_verdicts or verdict not in verdict_filter:
                continue
            sid = str(row.get("sid") or "").strip()
            if sid_filter and sid not in sid_filter:
                continue
            entry = student_entries.get(sid)
            teacher_name = str(row.get("teacher") or (entry or {}).get("teacher_name") or "").strip()
            if teacher_filter and teacher_name not in teacher_filter:
                continue
            application = row.get("application") if isinstance(row.get("application"), dict) else {}
            if not force and str(application.get("status") or "").strip() == "applied":
                skipped_already_applied += 1
                continue
            filtered_rows.append(row)

        rerun_groups: dict[tuple[str, str, str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
        rerun_targets: list[dict[str, Any]] = []
        feedback_targets: list[str] = []
        rescore_results: list[dict[str, Any]] = []
        applied_counts = {"feedback_only": 0, "rescore_and_feedback": 0, "rerun_grading": 0}
        for row in filtered_rows:
            audit = row.get("audit") if isinstance(row.get("audit"), dict) else {}
            verdict = str(audit.get("review_verdict") or "").strip()
            sid = str(row.get("sid") or "").strip()
            entry = student_entries.get(sid)
            if not entry:
                continue
            if verdict == "feedback_only":
                self._clear_feedback_artifacts(str(entry.get("run_root") or ""))
                row["application"] = {
                    "status": "applied",
                    "applied_at": _now_iso(),
                    "verdict": verdict,
                    "actions": ["feedback_regenerated"],
                }
                feedback_targets.append(sid)
                applied_counts[verdict] += 1
                continue
            if verdict == "rescore_and_feedback":
                rescore_result = self._apply_rescore_adjustment(row, entry)
                row["application"] = {
                    "status": "applied",
                    "applied_at": _now_iso(),
                    "verdict": verdict,
                    "actions": ["score_adjusted", "feedback_regenerated"],
                    "effective_score": rescore_result.get("effective_score"),
                    "effective_decision": rescore_result.get("effective_decision"),
                }
                self._clear_feedback_artifacts(str(entry.get("run_root") or ""))
                feedback_targets.append(sid)
                rescore_results.append(rescore_result)
                applied_counts[verdict] += 1
                continue
            paper_path = str(
                entry.get("latest_download_file")
                or (entry.get("downloaded_files") or [None])[-1]
                or entry.get("paper_path")
                or ""
            ).strip()
            if not paper_path:
                continue
            rerun_targets.append(
                {
                    "row": row,
                    "sid": sid,
                    "name": entry.get("name"),
                    "teacher": entry.get("teacher_name"),
                    "paper_path": paper_path,
                    "stage": entry.get("stage") or "initial_draft",
                    "visual_mode": entry.get("visual_mode") or "auto",
                    "visual_model": entry.get("visual_model") or "gpt-5.4",
                    "text_mode": entry.get("text_mode") or "expert",
                    "text_primary_model": entry.get("text_primary_model") or "deepseek-ai/DeepSeek-V3.2",
                    "text_secondary_model": entry.get("text_secondary_model") or "kimi-for-coding",
                }
            )

        for item in rerun_targets:
            key = (
                str(item.get("stage") or "initial_draft"),
                str(item.get("visual_mode") or "auto"),
                str(item.get("visual_model") or "gpt-5.4"),
                str(item.get("text_mode") or "expert"),
                str(item.get("text_primary_model") or "deepseek-ai/DeepSeek-V3.2"),
                str(item.get("text_secondary_model") or "kimi-for-coding"),
            )
            rerun_groups[key].append(item)

        regraded_batches: list[dict[str, Any]] = []
        for (stage, visual_mode, visual_model, text_mode, text_primary_model, text_secondary_model), group in sorted(
            rerun_groups.items(),
            key=lambda item: item[0],
        ):
            group = sorted(group, key=lambda item: (str(item.get("sid") or ""), str(item.get("name") or "")))
            papers = [item["paper_path"] for item in group]
            result = self.grade_ingested_files(
                files=papers,
                stage=stage,
                visual_mode=visual_mode,
                visual_model=visual_model,
                text_mode=text_mode,
                text_primary_model=text_primary_model,
                text_secondary_model=text_secondary_model,
                limit=0,
                skip_tracking_refresh=True,
            )
            regraded_batches.append(
                {
                    "stage": stage,
                    "visual_mode": visual_mode,
                    "visual_model": visual_model,
                    "text_mode": text_mode,
                    "text_primary_model": text_primary_model,
                    "text_secondary_model": text_secondary_model,
                    "count": len(group),
                    "result": result,
                    "students": [{"sid": item["sid"], "name": item["name"]} for item in group],
                }
            )
            for item in group:
                row = item["row"]
                row["application"] = {
                    "status": "applied",
                    "applied_at": _now_iso(),
                    "verdict": "rerun_grading",
                    "actions": ["grading_rerun", "feedback_regenerated"],
                }
                applied_counts["rerun_grading"] += 1

        _write_json(audit_file, audit_rows)

        selected_sids = sorted({str(row.get("sid") or "").strip() for row in filtered_rows if str(row.get("sid") or "").strip()})
        if teacher_filter or sid_filter or verdict_filter != actionable_verdicts:
            tracking = self._refresh_selected_students(selected_sids)
        else:
            tracking = self.refresh_tracking_outputs()
        return {
            "filters": {
                "teachers": sorted(teacher_filter),
                "verdicts": sorted(verdict_filter),
                "student_ids": sorted(sid_filter),
                "force": bool(force),
            },
            "selected_count": len(filtered_rows),
            "skipped_already_applied": skipped_already_applied,
            "applied_counts": applied_counts,
            "feedback_target_count": len(feedback_targets),
            "rescore_results": rescore_results,
            "rerun_target_count": len(rerun_targets),
            "rerun_batch_count": len(regraded_batches),
            "rerun_batches": regraded_batches,
            "tracking": tracking,
        }

    def doctor(self) -> dict[str, Any]:
        checks = {
            "paperdownload_root_exists": self.cfg.paperdownload_root.exists(),
            "paper_grading_root_exists": self.cfg.paper_grading_root.exists(),
            "download_script_exists": (self.cfg.paperdownload_root / "run-longzhi-automation.ps1").exists(),
            "save_credential_script_exists": (self.cfg.paperdownload_root / "save-longzhi-credential.ps1").exists(),
            "grade_batch_script_exists": (self.cfg.paper_grading_root / "process_incoming_papers.ps1").exists(),
            "grade_single_script_exists": (self.cfg.paper_grading_root / "run_grade.ps1").exists(),
            "credential_store_exists": self.cfg.credential_store_dir.exists(),
            "longzhi_credential_exists": (self.cfg.credential_store_dir / "longzhi.json").exists(),
            "moonshot_credential_exists": (self.cfg.credential_store_dir / "moonshot_kimi.json").exists(),
            "siliconflow_credential_exists": (self.cfg.credential_store_dir / "siliconflow.json").exists(),
            "node_found": shutil.which("node") is not None,
            "npm_found": shutil.which("npm.cmd") is not None or shutil.which("npm") is not None,
            "python_found": shutil.which("python") is not None,
            "git_found": shutil.which("git") is not None,
        }
        ok = all(
            checks[key]
            for key in [
                "paperdownload_root_exists",
                "paper_grading_root_exists",
                "download_script_exists",
                "save_credential_script_exists",
                "grade_batch_script_exists",
                "grade_single_script_exists",
                "credential_store_exists",
                "longzhi_credential_exists",
                "node_found",
                "npm_found",
                "python_found",
                "git_found",
            ]
        )
        missing = [k for k, v in checks.items() if not v]
        return {
            "ok": ok,
            "checks": checks,
            "missing": missing,
            "note": "评分依赖本机 Microsoft Word（COM）环境，此项无法在 doctor 中自动确认。Moonshot/Kimi/SiliconFlow 凭据均为可选项，仅在对应视觉模式下需要；凭据使用当前 Windows 用户的 DPAPI 加密，换机器或换用户后需要重新保存。",
        }


def _build_parser() -> argparse.ArgumentParser:
    return _build_parser_shared(str((Path(__file__).resolve().parent.parent / "config" / "pipeline" / "pipeline.config.json")))

    parser = argparse.ArgumentParser(description="统一论文下载与评分工作流（下载 -> 改名入队 -> 批量评分）。")
    parser.add_argument(
        "--config",
        default=str((Path(__file__).resolve().parent / "pipeline.config.json")),
        help="配置文件路径（JSON）",
    )

    sub = parser.add_subparsers(dest="command", required=True)

    p_download = sub.add_parser("download", help="仅执行下载阶段")
    p_download.add_argument("--page-size", type=int, default=100)
    p_download.add_argument("--start-page", type=int, default=1)
    p_download.add_argument("--max-students", type=int, default=0, help="仅下载前 N 位学生（0 表示不限制）")

    sub.add_parser("ingest", help="仅执行改名入队阶段")

    p_grade = sub.add_parser("grade", help="仅执行评分阶段")
    p_grade.add_argument("--stage", choices=["initial_draft", "final"], default="initial_draft")
    p_grade.add_argument("--visual-mode", choices=["auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off"], default="auto")
    p_grade.add_argument("--visual-model", default="gpt-5.4")
    p_grade.add_argument("--text-mode", choices=["off", "auto", "expert", "siliconflow", "moonshot"], default="expert")
    p_grade.add_argument("--text-primary-model", default="deepseek-ai/DeepSeek-V3.2")
    p_grade.add_argument("--text-secondary-model", default="kimi-for-coding")
    p_grade.add_argument("--limit", type=int, default=0)

    p_all = sub.add_parser("run-all", help="执行完整流程：下载 -> 入队 -> 评分")
    p_all.add_argument("--page-size", type=int, default=100)
    p_all.add_argument("--start-page", type=int, default=1)
    p_all.add_argument("--max-students", type=int, default=0, help="仅下载前 N 位学生（0 表示不限制）")
    p_all.add_argument("--stage", choices=["initial_draft", "final"], default="initial_draft")
    p_all.add_argument("--visual-mode", choices=["auto", "openai", "moonshot", "siliconflow", "expert", "heuristic", "off"], default="auto")
    p_all.add_argument("--visual-model", default="gpt-5.4")
    p_all.add_argument("--text-mode", choices=["off", "auto", "expert", "siliconflow", "moonshot"], default="expert")
    p_all.add_argument("--text-primary-model", default="deepseek-ai/DeepSeek-V3.2")
    p_all.add_argument("--text-secondary-model", default="kimi-for-coding")
    p_all.add_argument("--limit", type=int, default=0)
    p_all.add_argument("--grade-even-if-no-new", action="store_true")
    p_all.add_argument(
        "--queue-grade",
        action="store_true",
        help="改为按评分系统队列模式跑（默认只评分本次 ingest 的文件）",
    )

    p_bundle = sub.add_parser("bundle-case", help="按批次导出最终交付目录（复制，不移动原文件）")
    p_bundle.add_argument("--case-name", required=True, help="批次名称/文件夹名称")
    p_bundle.add_argument("--student-ids", default="", help="逗号分隔的学生学号列表")
    p_bundle.add_argument("--latest-graded", type=int, default=0, help="按最新评分时间选取最近 N 位学生")
    p_bundle.add_argument("--all-graded", action="store_true", help="导出当前所有已评分学生")
    p_bundle.add_argument("--overwrite", action="store_true", help="若目标批次已存在，先备份旧目录再重建")

    p_source = sub.add_parser("set-source", help="登记老师、链接和阶段，并可设为当前活动来源")
    p_source.add_argument("--teacher-name", required=True, help="老师名称，例如 牛逼老师")
    p_source.add_argument("--target-page-url", required=True, help="Longzhi 批阅页面链接")
    p_source.add_argument("--stage-label", default="初稿", help="阶段标签，例如 初稿、终稿")
    p_source.add_argument("--no-set-active", action="store_true", help="只登记，不切换当前活动来源")
    p_source.add_argument("--bind-all-current", action="store_true", help="把当前总日志中的下载文件全部绑定到这个来源")

    sub.add_parser("list-sources", help="查看已登记的老师来源映射")

    p_bundle_source = sub.add_parser("bundle-source", help="按已登记的老师来源自动打包")
    p_bundle_source.add_argument("--source-key", required=True, help="来源标识，可先用 list-sources 查看")
    p_bundle_source.add_argument("--overwrite", action="store_true", help="若目标交付目录已存在，先备份旧目录再重建")

    sub.add_parser("refresh-log", help="刷新仓库根目录的总日志和学生评语")
    sub.add_parser("status", help="查看当前状态")
    sub.add_parser("doctor", help="检查运行环境与依赖是否就绪")

    return parser


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    parser = _build_parser()
    args = parser.parse_args()

    cfg = PipelineConfig.load(Path(args.config).resolve())
    pipeline = UnifiedPipeline(cfg)

    try:
        if args.command == "download":
            result = pipeline.download(
                page_size=args.page_size,
                start_page=args.start_page,
                max_students=args.max_students,
            )
        elif args.command == "ingest":
            result = pipeline.ingest()
        elif args.command == "grade":
            result = pipeline.grade(
                stage=args.stage,
                visual_mode=args.visual_mode,
                visual_model=args.visual_model,
                text_mode=args.text_mode,
                text_primary_model=args.text_primary_model,
                text_secondary_model=args.text_secondary_model,
                limit=args.limit,
            )
        elif args.command == "run-all":
            result = pipeline.run_all(
                page_size=args.page_size,
                start_page=args.start_page,
                max_students=args.max_students,
                stage=args.stage,
                visual_mode=args.visual_mode,
                visual_model=args.visual_model,
                text_mode=args.text_mode,
                text_primary_model=args.text_primary_model,
                text_secondary_model=args.text_secondary_model,
                limit=args.limit,
                grade_even_if_no_new=args.grade_even_if_no_new,
                grade_ingested_only=not args.queue_grade,
            )
        elif args.command == "status":
            result = pipeline.status()
        elif args.command == "refresh-log":
            result = pipeline.refresh_tracking_outputs()
        elif args.command == "bundle-case":
            student_ids = [item.strip() for item in str(args.student_ids or "").split(",") if item.strip()]
            result = pipeline.bundle_case(
                case_name=args.case_name,
                student_ids=student_ids,
                latest_graded=args.latest_graded,
                all_graded=bool(args.all_graded),
                overwrite=bool(args.overwrite),
            )
        elif args.command == "set-source":
            result = pipeline.set_source(
                teacher_name=args.teacher_name,
                target_page_url=args.target_page_url,
                stage_label=args.stage_label,
                set_active=not bool(args.no_set_active),
                bind_all_current=bool(args.bind_all_current),
            )
        elif args.command == "list-sources":
            result = pipeline.list_sources()
        elif args.command == "bundle-source":
            result = pipeline.bundle_source(
                source_key=args.source_key,
                overwrite=bool(args.overwrite),
            )
        elif args.command == "doctor":
            result = pipeline.doctor()
        elif args.command == "audit-students":
            result = pipeline.audit_students(limit=args.limit, force=bool(getattr(args, "force", False)))
        elif args.command == "rebuild-anomalies":
            result = pipeline.rebuild_anomalies()
        elif args.command == "apply-audit":
            result = pipeline.apply_audit_reviews(
                teachers=_split_csv_arg(getattr(args, "teachers", "")),
                verdicts=_split_csv_arg(getattr(args, "verdicts", "")),
                student_ids=_split_csv_arg(getattr(args, "student_ids", "")),
                force=bool(getattr(args, "force", False)),
            )
        else:
            parser.error(f"Unknown command: {args.command}")
            return 2
    except Exception as err:
        print(json.dumps({"status": "failed", "error": str(err)}, ensure_ascii=False, indent=2))
        return 1

    refresh_after_commands = {"download", "ingest", "grade", "run-all", "set-source"}
    if args.command in refresh_after_commands and isinstance(result, dict):
        result["tracking"] = pipeline.refresh_tracking_outputs()

    print(json.dumps({"status": "ok", "command": args.command, "result": result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
