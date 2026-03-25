from __future__ import annotations

import json
import logging
import os
import hashlib
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any


logger = logging.getLogger("paper_pipeline")

DEFAULT_COMMAND_TIMEOUT_SECONDS = 30 * 60


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def read_json(path: Path, fallback: Any) -> Any:
    for encoding in ("utf-8", "utf-8-sig"):
        try:
            return json.loads(path.read_text(encoding=encoding))
        except FileNotFoundError:
            return fallback
        except (json.JSONDecodeError, UnicodeDecodeError):
            continue

    logger.warning("Failed to decode JSON from %s, using fallback.", path)
    return fallback


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def merge_dict(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged: dict[str, Any] = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            merged[key] = merge_dict(base[key], value)
        else:
            merged[key] = value
    return merged


def resolve_path(raw: str, base_dir: Path) -> Path:
    candidate = Path(raw)
    if candidate.is_absolute():
        return candidate.resolve()
    return (base_dir / candidate).resolve()


def run_command(
    args: list[str],
    cwd: Path,
    extra_env: dict[str, str] | None = None,
    timeout_seconds: int = DEFAULT_COMMAND_TIMEOUT_SECONDS,
) -> None:
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)

    logger.info("Running command: %s", " ".join(args))
    try:
        proc = subprocess.run(
            args,
            cwd=str(cwd),
            check=False,
            env=env,
            timeout=timeout_seconds,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.TimeoutExpired as err:
        raise RuntimeError(
            f"Command timed out after {timeout_seconds}s: {' '.join(args)}"
        ) from err

    if proc.stdout:
        logger.info("Command stdout:\n%s", proc.stdout.rstrip())
    if proc.stderr:
        logger.warning("Command stderr:\n%s", proc.stderr.rstrip())

    if proc.returncode != 0:
        detail = proc.stderr.strip() or proc.stdout.strip() or "no output captured"
        raise RuntimeError(
            f"Command failed (exit={proc.returncode}): {' '.join(args)}\n{detail}"
        )


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_name(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1F\s]+', "_", value).strip("_")


def parse_filename(stem: str) -> tuple[str | None, str | None, str | None]:
    match = re.match(r"^(\d{5,20})_(.+?)(?:_(\d+))?$", stem)
    if not match:
        return None, None, None
    return match.group(1), match.group(2), match.group(3)


def parse_ingested_filename(stem: str) -> tuple[str | None, str | None]:
    match = re.match(r"^(?:\d{8}_)?(\d{5,20})_(.+?)(?:_(?:[0-9a-f]{4,64}|v\d+|\d+))?$", stem)
    if not match:
        return None, None
    return match.group(1), match.group(2)


def normalize_repo_path(raw: str, repo_root: Path) -> str:
    if not raw:
        return raw

    candidate = Path(raw)
    if candidate.exists():
        return str(candidate.resolve())

    normalized = raw.replace("\\", "/")
    marker = f"{repo_root.name}/"
    idx = normalized.lower().find(marker.lower())
    if idx == -1:
        return raw

    relative_part = normalized[idx + len(marker) :]
    return str((repo_root / Path(relative_part)).resolve())


def parse_run_info(path: Path, repo_root: Path) -> dict[str, str]:
    data: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        cleaned = value.strip()
        if key in {"run_root", "paper", "paper_copy", "visual_dir", "text_dir", "json", "report"}:
            cleaned = normalize_repo_path(cleaned, repo_root)
        data[key.strip()] = cleaned
    return data


def format_score(value: Any) -> str:
    if value is None:
        return "-"
    if isinstance(value, (int, float)):
        return f"{float(value):.2f}"
    try:
        return f"{float(str(value)): .2f}".strip()
    except ValueError:
        return str(value)


def relative_display(path: str | None, repo_root: Path) -> str:
    if not path:
        return "-"
    try:
        return str(Path(path).resolve().relative_to(repo_root))
    except ValueError:
        return path


def dedupe_keep_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        cleaned = str(value or "").strip()
        if not cleaned or cleaned in seen:
            continue
        seen.add(cleaned)
        result.append(cleaned)
    return result


def read_text(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "gbk"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue

    logger.warning("Falling back to lossy UTF-8 read for %s", path)
    return path.read_text(encoding="utf-8", errors="ignore")


def bundle_relative(path: Path, root: Path) -> str:
    try:
        return str(path.resolve().relative_to(root.resolve())).replace("\\", "/")
    except ValueError:
        return str(path.resolve())


def copy_file(source: str | None, destination: Path) -> str | None:
    if not source:
        return None
    src = Path(source)
    if not src.exists() or not src.is_file():
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, destination)
    return str(destination.resolve())


def summarize_feedback(feedback_path: str | None) -> str:
    if not feedback_path:
        return ""
    path = Path(feedback_path)
    if not path.exists():
        return ""

    text = read_text(path)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    conclusion = ""
    urgent: list[str] = []
    collecting_urgent = False

    for line in lines:
        if line.startswith("- ") and ":" in line and not conclusion:
            conclusion = line.split(":", 1)[1].strip()
            continue
        if line.startswith("## "):
            collecting_urgent = "urgent" in line.lower() or "最" in line
            if len(urgent) >= 3:
                break
            continue
        if collecting_urgent and re.match(r"^\d+\.\s+", line):
            urgent.append(re.sub(r"^\d+\.\s+", "", line).strip())

    parts: list[str] = []
    if conclusion:
        parts.append(conclusion)
    if urgent:
        parts.append("Priority fixes: " + "; ".join(urgent[:3]))
    return " ".join(parts).strip()


def source_folder_name(teacher_name: str, stage_label: str) -> str:
    teacher = str(teacher_name or "").strip()
    stage = str(stage_label or "").strip() or "初稿"
    return f"{teacher}_{stage}" if teacher else stage


def source_key(teacher_name: str, stage_label: str) -> str:
    return safe_name(source_folder_name(teacher_name, stage_label)) or "source"
