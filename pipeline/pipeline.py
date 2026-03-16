from __future__ import annotations

import argparse
import os
import hashlib
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _read_json(path: Path, fallback: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        try:
            return json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return fallback


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


def _run_command(args: list[str], cwd: Path, extra_env: dict[str, str] | None = None) -> None:
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    proc = subprocess.run(args, cwd=str(cwd), check=False, env=env)
    if proc.returncode != 0:
        raise RuntimeError(f"Command failed (exit={proc.returncode}): {' '.join(args)}")


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _safe_name(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F\s]+', "_", value).strip("_")


def _parse_filename(stem: str) -> tuple[str | None, str | None, str | None]:
    # Example: 24315018_张三, 24315018_张三_2
    m = re.match(r"^(\d{5,20})_(.+?)(?:_(\d+))?$", stem)
    if not m:
        return None, None, None
    return m.group(1), m.group(2), m.group(3)


@dataclass
class PipelineConfig:
    project_root: Path
    paperdownload_root: Path
    paper_grading_root: Path
    download_output_root: Path
    incoming_dir: Path
    state_dir: Path
    rename_prefix_with_date: bool
    rename_hash_length: int

    @classmethod
    def load(cls, config_path: Path) -> "PipelineConfig":
        project_root = config_path.resolve().parent
        defaults: dict[str, Any] = {
            "paperdownload_root": str((project_root.parent / "paperdownload").resolve()),
            "paper_grading_root": str((project_root.parent / "paper-grading-system").resolve()),
            "download_output_root": str((project_root.parent / "paperdownload" / "longzhi_batch_output").resolve()),
            "incoming_dir": str((project_root.parent / "paper-grading-system" / "assets" / "incoming_papers").resolve()),
            "state_dir": str((project_root / "state").resolve()),
            "rename": {
                "prefix_with_date": True,
                "hash_length": 8,
            },
        }
        user_data = _read_json(config_path, {})
        merged = _merge_dict(defaults, user_data if isinstance(user_data, dict) else {})
        return cls(
            project_root=project_root,
            paperdownload_root=Path(merged["paperdownload_root"]).resolve(),
            paper_grading_root=Path(merged["paper_grading_root"]).resolve(),
            download_output_root=Path(merged["download_output_root"]).resolve(),
            incoming_dir=Path(merged["incoming_dir"]).resolve(),
            state_dir=Path(merged["state_dir"]).resolve(),
            rename_prefix_with_date=bool(merged["rename"]["prefix_with_date"]),
            rename_hash_length=max(4, int(merged["rename"]["hash_length"])),
        )


class UnifiedPipeline:
    def __init__(self, cfg: PipelineConfig) -> None:
        self.cfg = cfg
        self.state_path = self.cfg.state_dir / "ingest_state.json"
        self.reports_dir = self.cfg.state_dir / "reports"

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

    def _write_run_report(self, report: dict[str, Any]) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = self.reports_dir / f"{ts}_pipeline_run.json"
        _write_json(path, report)
        return path

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
        if max_students > 0:
            extra_env["MAX_STUDENTS"] = str(max_students)
        _run_command(args, cwd=self.cfg.paperdownload_root, extra_env=extra_env or None)

        summary_path = self.cfg.download_output_root / "state" / "latest_automation_summary.json"
        summary = _read_json(summary_path, {})
        return {
            "status": summary.get("status", "unknown"),
            "runId": summary.get("runId"),
            "processed": summary.get("processed"),
            "downloaded": summary.get("downloaded"),
            "max_students": max_students,
            "newFiles": summary.get("newFiles", []),
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

    def grade(self, stage: str, visual_mode: str, visual_model: str, limit: int) -> dict[str, Any]:
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
        ]
        if limit > 0:
            args.extend(["-Limit", str(limit)])

        _run_command(args, cwd=self.cfg.paper_grading_root)
        return {
            "status": "success",
            "mode": "queue",
            "stage": stage,
            "visual_mode": visual_mode,
            "visual_model": visual_model,
            "limit": limit,
        }

    def grade_ingested_files(self, files: list[str], stage: str, visual_mode: str, visual_model: str, limit: int) -> dict[str, Any]:
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
            ]
            _run_command(args, cwd=self.cfg.paper_grading_root)
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
                    limit=limit,
                )
            else:
                report["steps"]["grade"] = self.grade(
                    stage=stage,
                    visual_mode=visual_mode,
                    visual_model=visual_model,
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
            "state_file": str(self.state_path),
        }


def _build_parser() -> argparse.ArgumentParser:
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
    p_grade.add_argument("--visual-mode", choices=["auto", "openai", "heuristic", "off"], default="auto")
    p_grade.add_argument("--visual-model", default="gpt-5.4")
    p_grade.add_argument("--limit", type=int, default=0)

    p_all = sub.add_parser("run-all", help="执行完整流程：下载 -> 入队 -> 评分")
    p_all.add_argument("--page-size", type=int, default=100)
    p_all.add_argument("--start-page", type=int, default=1)
    p_all.add_argument("--max-students", type=int, default=0, help="仅下载前 N 位学生（0 表示不限制）")
    p_all.add_argument("--stage", choices=["initial_draft", "final"], default="initial_draft")
    p_all.add_argument("--visual-mode", choices=["auto", "openai", "heuristic", "off"], default="auto")
    p_all.add_argument("--visual-model", default="gpt-5.4")
    p_all.add_argument("--limit", type=int, default=0)
    p_all.add_argument("--grade-even-if-no-new", action="store_true")
    p_all.add_argument(
        "--queue-grade",
        action="store_true",
        help="改为按评分系统队列模式跑（默认只评分本次 ingest 的文件）",
    )

    sub.add_parser("status", help="查看当前状态")

    return parser


def main() -> int:
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
                limit=args.limit,
                grade_even_if_no_new=args.grade_even_if_no_new,
                grade_ingested_only=not args.queue_grade,
            )
        elif args.command == "status":
            result = pipeline.status()
        else:
            parser.error(f"Unknown command: {args.command}")
            return 2
    except Exception as err:
        print(json.dumps({"status": "failed", "error": str(err)}, ensure_ascii=False, indent=2))
        return 1

    print(json.dumps({"status": "ok", "command": args.command, "result": result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
