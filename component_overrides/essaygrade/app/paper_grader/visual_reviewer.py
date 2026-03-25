from __future__ import annotations

import base64
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
import json
import msvcrt
import os
from pathlib import Path
import time
from typing import Any

import requests
import win32com.client as win32

from .credential_store import credential_entry_exists, load_credential_entry
from .rubric import OFFICIAL_FORMAT_REQUIREMENTS


try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover - optional dependency at runtime
    fitz = None


OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses"
MOONSHOT_DEFAULT_BASE_URL = "https://api.moonshot.cn/v1"
SILICONFLOW_DEFAULT_BASE_URL = "https://api.siliconflow.cn/v1"
WORD_FORMAT_PDF = 17
VISUAL_REVIEW_TIMEOUT_SECONDS = 300
DEFAULT_VISUAL_MODEL = "gpt-5.4"
DEFAULT_MOONSHOT_MODEL = "kimi-thinking-preview"
DEFAULT_SILICONFLOW_VISUAL_MODEL = "Pro/moonshotai/Kimi-K2.5"
DEFAULT_SILICONFLOW_SECONDARY_VISUAL_MODEL = "zai-org/GLM-4.6V"
MAX_VISUAL_PAGES = 4
VISUAL_REVIEW_RETRYABLE_STATUS_CODES = {408, 409, 425, 429, 500, 502, 503, 504}
WORD_EXPORT_LOCK_DEFAULT_PATH = Path(os.getenv("TEMP", ".")) / "essaygrade_word_export.lock"


@dataclass
class VisualReviewResult:
    mode: str
    model: str | None
    pdf_path: str | None
    overall_verdict: str
    visual_order_score: float | None
    confidence: float | None
    major_issues: list[str]
    minor_issues: list[str]
    evidence: list[str]
    page_observations: list[str]
    notes: list[str]
    timing: dict[str, int] | None = None
    raw_response_path: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


def has_moonshot_visual_credentials() -> bool:
    if os.getenv("MOONSHOT_API_KEY"):
        return True
    return credential_entry_exists("moonshot_kimi")


def has_siliconflow_visual_credentials() -> bool:
    if os.getenv("SILICONFLOW_API_KEY"):
        return True
    return credential_entry_exists("siliconflow")


def resolve_moonshot_visual_model_name(model: str | None = None) -> str:
    metadata: dict[str, Any] = {}
    if credential_entry_exists("moonshot_kimi"):
        entry = load_credential_entry("moonshot_kimi")
        metadata = entry.get("metadata") or {}
    if metadata.get("default_model") and (not model or model == DEFAULT_MOONSHOT_MODEL):
        return str(metadata["default_model"])
    return model or os.getenv("MOONSHOT_VISUAL_MODEL") or DEFAULT_MOONSHOT_MODEL


def resolve_siliconflow_visual_model_name(model: str | None = None) -> str:
    metadata: dict[str, Any] = {}
    if credential_entry_exists("siliconflow"):
        entry = load_credential_entry("siliconflow")
        metadata = entry.get("metadata") or {}
    if metadata.get("default_model") and (not model or model == DEFAULT_SILICONFLOW_VISUAL_MODEL):
        return str(metadata["default_model"])
    return model or os.getenv("SILICONFLOW_VISUAL_MODEL") or DEFAULT_SILICONFLOW_VISUAL_MODEL


def resolve_siliconflow_secondary_visual_model_name(model: str | None = None) -> str:
    return model or os.getenv("SILICONFLOW_SECONDARY_VISUAL_MODEL") or DEFAULT_SILICONFLOW_SECONDARY_VISUAL_MODEL


def review_document_with_openai(
    document_path: str,
    output_dir: str | None = None,
    model: str = DEFAULT_VISUAL_MODEL,
    api_key: str | None = None,
    stage: str = "initial_draft",
) -> VisualReviewResult:
    total_started_at = time.perf_counter()
    api_key = api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY 未配置，无法启用大模型视觉审稿。")

    resolved_document = Path(document_path).expanduser().resolve()
    review_dir = _ensure_output_dir(output_dir, resolved_document)
    export_started_at = time.perf_counter()
    pdf_path = export_document_to_pdf(str(resolved_document), str(review_dir))
    export_pdf_ms = int(round((time.perf_counter() - export_started_at) * 1000))
    payload = _build_openai_visual_review_payload(pdf_path, model, stage)
    api_started_at = time.perf_counter()
    response_json = _post_json_request(
        OPENAI_RESPONSES_URL,
        payload,
        api_key,
        extra_headers=None,
    )
    api_request_ms = int(round((time.perf_counter() - api_started_at) * 1000))

    raw_response_path = review_dir / "openai_visual_response.json"
    raw_response_path.write_text(json.dumps(response_json, ensure_ascii=False, indent=2), encoding="utf-8")

    parse_started_at = time.perf_counter()
    parsed = _parse_openai_visual_response(response_json)
    parse_response_ms = int(round((time.perf_counter() - parse_started_at) * 1000))
    total_ms = int(round((time.perf_counter() - total_started_at) * 1000))
    return VisualReviewResult(
        mode="openai",
        model=model,
        pdf_path=str(pdf_path),
        overall_verdict=parsed["overall_verdict"],
        visual_order_score=parsed["visual_order_score"],
        confidence=parsed["confidence"],
        major_issues=parsed["major_issues"],
        minor_issues=parsed["minor_issues"],
        evidence=parsed["evidence"],
        page_observations=parsed["page_observations"],
        notes=parsed["notes"],
        timing={
            "export_pdf_ms": export_pdf_ms,
            "api_request_ms": api_request_ms,
            "parse_response_ms": parse_response_ms,
            "total_ms": total_ms,
        },
        raw_response_path=str(raw_response_path),
    )


def review_document_with_moonshot(
    document_path: str,
    output_dir: str | None = None,
    model: str = DEFAULT_MOONSHOT_MODEL,
    api_key: str | None = None,
    api_base_url: str | None = None,
    stage: str = "initial_draft",
    max_pages: int = MAX_VISUAL_PAGES,
) -> VisualReviewResult:
    config = _resolve_moonshot_config(api_key=api_key, api_base_url=api_base_url, model=model)
    return _review_document_with_chat_provider(
        document_path=document_path,
        output_dir=output_dir,
        mode="moonshot",
        model_name=config["model"],
        api_key=config["api_key"],
        api_base_url=config["api_base_url"],
        stage=stage,
        max_pages=max_pages,
        raw_filename="moonshot_visual_response.json",
        note_prefix="Moonshot/Kimi visual review",
    )
    resolved_document = Path(document_path).expanduser().resolve()
    review_dir = _ensure_output_dir(output_dir, resolved_document)
    pdf_path = export_document_to_pdf(str(resolved_document), str(review_dir))
    config = _resolve_moonshot_config(api_key=api_key, api_base_url=api_base_url, model=model)
    image_payloads = _render_pdf_pages_to_data_urls(pdf_path, review_dir, max_pages=max_pages)
    payload = _build_chat_visual_review_payload(image_payloads, config["model"], stage)
    url = config["api_base_url"].rstrip("/") + "/chat/completions"
    response_json = _post_json_request(url, payload, config["api_key"], extra_headers=None)

    raw_response_path = review_dir / "moonshot_visual_response.json"
    raw_response_path.write_text(json.dumps(response_json, ensure_ascii=False, indent=2), encoding="utf-8")

    parsed = _parse_chat_json_response(response_json)
    parsed["notes"].insert(0, f"使用 Moonshot/Kimi 视觉接口，共发送 {len(image_payloads)} 页图像。")
    return VisualReviewResult(
        mode="moonshot",
        model=config["model"],
        pdf_path=str(pdf_path),
        overall_verdict=parsed["overall_verdict"],
        visual_order_score=parsed["visual_order_score"],
        confidence=parsed["confidence"],
        major_issues=parsed["major_issues"],
        minor_issues=parsed["minor_issues"],
        evidence=parsed["evidence"],
        page_observations=parsed["page_observations"],
        notes=parsed["notes"],
        raw_response_path=str(raw_response_path),
    )


def review_document_with_siliconflow(
    document_path: str,
    output_dir: str | None = None,
    model: str = DEFAULT_SILICONFLOW_VISUAL_MODEL,
    api_key: str | None = None,
    api_base_url: str | None = None,
    stage: str = "initial_draft",
    max_pages: int = MAX_VISUAL_PAGES,
) -> VisualReviewResult:
    config = _resolve_siliconflow_config(api_key=api_key, api_base_url=api_base_url, model=model)
    return _review_document_with_chat_provider(
        document_path=document_path,
        output_dir=output_dir,
        mode="siliconflow",
        model_name=config["model"],
        api_key=config["api_key"],
        api_base_url=config["api_base_url"],
        stage=stage,
        max_pages=max_pages,
        raw_filename="siliconflow_visual_response.json",
        note_prefix="SiliconFlow visual review",
    )
    resolved_document = Path(document_path).expanduser().resolve()
    review_dir = _ensure_output_dir(output_dir, resolved_document)
    pdf_path = export_document_to_pdf(str(resolved_document), str(review_dir))
    config = _resolve_siliconflow_config(api_key=api_key, api_base_url=api_base_url, model=model)
    image_payloads = _render_pdf_pages_to_data_urls(pdf_path, review_dir, max_pages=max_pages)
    payload = _build_chat_visual_review_payload(image_payloads, config["model"], stage)
    url = config["api_base_url"].rstrip("/") + "/chat/completions"
    response_json = _post_json_request(url, payload, config["api_key"], extra_headers=None)

    raw_response_path = review_dir / "siliconflow_visual_response.json"
    raw_response_path.write_text(json.dumps(response_json, ensure_ascii=False, indent=2), encoding="utf-8")

    parsed = _parse_chat_json_response(response_json)
    parsed["notes"].insert(0, f"使用 SiliconFlow 视觉接口，共发送 {len(image_payloads)} 个代表性页面。")
    return VisualReviewResult(
        mode="siliconflow",
        model=config["model"],
        pdf_path=str(pdf_path),
        overall_verdict=parsed["overall_verdict"],
        visual_order_score=parsed["visual_order_score"],
        confidence=parsed["confidence"],
        major_issues=parsed["major_issues"],
        minor_issues=parsed["minor_issues"],
        evidence=parsed["evidence"],
        page_observations=parsed["page_observations"],
        notes=parsed["notes"],
        raw_response_path=str(raw_response_path),
    )


def _review_document_with_chat_provider(
    document_path: str,
    output_dir: str | None,
    mode: str,
    model_name: str,
    api_key: str,
    api_base_url: str,
    stage: str,
    max_pages: int,
    raw_filename: str,
    note_prefix: str,
) -> VisualReviewResult:
    total_started_at = time.perf_counter()
    resolved_document = Path(document_path).expanduser().resolve()
    review_dir = _ensure_output_dir(output_dir, resolved_document)
    export_started_at = time.perf_counter()
    pdf_path = export_document_to_pdf(str(resolved_document), str(review_dir))
    export_pdf_ms = int(round((time.perf_counter() - export_started_at) * 1000))
    render_started_at = time.perf_counter()
    image_payloads = _render_pdf_pages_to_data_urls(pdf_path, review_dir, max_pages=max_pages)
    render_pages_ms = int(round((time.perf_counter() - render_started_at) * 1000))
    payload = _build_chat_visual_review_payload(image_payloads, model_name, stage)
    url = api_base_url.rstrip("/") + "/chat/completions"
    api_started_at = time.perf_counter()
    response_json = _post_json_request(url, payload, api_key, extra_headers=None)
    api_request_ms = int(round((time.perf_counter() - api_started_at) * 1000))

    raw_response_path = review_dir / raw_filename
    raw_response_path.write_text(json.dumps(response_json, ensure_ascii=False, indent=2), encoding="utf-8")

    parse_started_at = time.perf_counter()
    parsed = _parse_chat_json_response(response_json)
    parse_response_ms = int(round((time.perf_counter() - parse_started_at) * 1000))
    total_ms = int(round((time.perf_counter() - total_started_at) * 1000))
    parsed["notes"].insert(0, f"{note_prefix}, sent {len(image_payloads)} page images.")
    return VisualReviewResult(
        mode=mode,
        model=model_name,
        pdf_path=str(pdf_path),
        overall_verdict=parsed["overall_verdict"],
        visual_order_score=parsed["visual_order_score"],
        confidence=parsed["confidence"],
        major_issues=parsed["major_issues"],
        minor_issues=parsed["minor_issues"],
        evidence=parsed["evidence"],
        page_observations=parsed["page_observations"],
        notes=parsed["notes"],
        timing={
            "export_pdf_ms": export_pdf_ms,
            "render_pages_ms": render_pages_ms,
            "api_request_ms": api_request_ms,
            "parse_response_ms": parse_response_ms,
            "total_ms": total_ms,
        },
        raw_response_path=str(raw_response_path),
    )


def review_document_with_siliconflow_ensemble(
    document_path: str,
    output_dir: str | None = None,
    primary_model: str = DEFAULT_SILICONFLOW_VISUAL_MODEL,
    secondary_model: str = DEFAULT_SILICONFLOW_SECONDARY_VISUAL_MODEL,
    api_key: str | None = None,
    api_base_url: str | None = None,
    stage: str = "initial_draft",
    max_pages: int = MAX_VISUAL_PAGES,
) -> VisualReviewResult:
    total_started_at = time.perf_counter()
    resolved_document = Path(document_path).expanduser().resolve()
    review_dir = _ensure_output_dir(output_dir, resolved_document)
    export_started_at = time.perf_counter()
    pdf_path = export_document_to_pdf(str(resolved_document), str(review_dir))
    export_pdf_ms = int(round((time.perf_counter() - export_started_at) * 1000))
    config = _resolve_siliconflow_config(api_key=api_key, api_base_url=api_base_url, model=primary_model)
    secondary_model = resolve_siliconflow_secondary_visual_model_name(secondary_model)
    render_started_at = time.perf_counter()
    image_payloads = _render_pdf_pages_to_data_urls(pdf_path, review_dir, max_pages=max_pages)
    render_pages_ms = int(round((time.perf_counter() - render_started_at) * 1000))

    primary_payload = _build_chat_visual_review_payload(image_payloads, config["model"], stage)
    secondary_payload = _build_chat_visual_review_payload(image_payloads, secondary_model, stage)
    url = config["api_base_url"].rstrip("/") + "/chat/completions"
    primary_response: dict[str, Any] | None = None
    secondary_response: dict[str, Any] | None = None
    primary_error: Exception | None = None
    secondary_error: Exception | None = None

    parallel_api_started_at = time.perf_counter()
    with ThreadPoolExecutor(max_workers=2) as executor:
        primary_future = executor.submit(_post_json_request, url, primary_payload, config["api_key"], None)
        secondary_future = executor.submit(_post_json_request, url, secondary_payload, config["api_key"], None)
        try:
            primary_response = primary_future.result()
        except Exception as exc:
            primary_error = exc
        try:
            secondary_response = secondary_future.result()
        except Exception as exc:
            secondary_error = exc
    parallel_api_wall_ms = int(round((time.perf_counter() - parallel_api_started_at) * 1000))

    primary_raw_path = review_dir / "siliconflow_primary_visual_response.json"
    secondary_raw_path = review_dir / "siliconflow_secondary_visual_response.json"
    if primary_response is not None:
        primary_raw_path.write_text(json.dumps(primary_response, ensure_ascii=False, indent=2), encoding="utf-8")
    if secondary_response is not None:
        secondary_raw_path.write_text(json.dumps(secondary_response, ensure_ascii=False, indent=2), encoding="utf-8")

    if primary_response is None and secondary_response is None:
        raise RuntimeError(
            "Both SiliconFlow visual expert requests failed: "
            f"primary={primary_error!r}; secondary={secondary_error!r}"
        )
    if primary_response is None:
        secondary_result = _result_from_chat_visual_response(
            secondary_response,
            mode="siliconflow",
            model=secondary_model,
            pdf_path=str(pdf_path),
            raw_response_path=str(secondary_raw_path),
            note=f"SiliconFlow secondary visual expert: {secondary_model}.",
        )
        if primary_error is not None:
            secondary_result.notes.append(f"SiliconFlow primary visual expert failed: {primary_error!r}")
        secondary_result.timing = {
            "export_pdf_ms": export_pdf_ms,
            "render_pages_ms": render_pages_ms,
            "parallel_api_wall_ms": parallel_api_wall_ms,
            "total_ms": int(round((time.perf_counter() - total_started_at) * 1000)),
        }
        return secondary_result
    if secondary_response is None:
        primary_result = _result_from_chat_visual_response(
            primary_response,
            mode="siliconflow",
            model=config["model"],
            pdf_path=str(pdf_path),
            raw_response_path=str(primary_raw_path),
            note=f"SiliconFlow primary visual expert: {config['model']}.",
        )
        if secondary_error is not None:
            primary_result.notes.append(f"SiliconFlow secondary visual expert failed: {secondary_error!r}")
        primary_result.timing = {
            "export_pdf_ms": export_pdf_ms,
            "render_pages_ms": render_pages_ms,
            "parallel_api_wall_ms": parallel_api_wall_ms,
            "total_ms": int(round((time.perf_counter() - total_started_at) * 1000)),
        }
        return primary_result

    primary_result = _result_from_chat_visual_response(
        primary_response,
        mode="siliconflow",
        model=config["model"],
        pdf_path=str(pdf_path),
        raw_response_path=str(primary_raw_path),
        note=f"SiliconFlow 主视觉专家：{config['model']}。",
    )
    secondary_result = _result_from_chat_visual_response(
        secondary_response,
        mode="siliconflow",
        model=secondary_model,
        pdf_path=str(pdf_path),
        raw_response_path=str(secondary_raw_path),
        note=f"SiliconFlow 次视觉专家：{secondary_model}。",
    )

    fuse_started_at = time.perf_counter()
    fused = _fuse_visual_reviews(
        primary_result=primary_result,
        secondary_result=secondary_result,
        pdf_path=str(pdf_path),
    )
    fuse_ms = int(round((time.perf_counter() - fuse_started_at) * 1000))
    fused_path = review_dir / "siliconflow_visual_fusion.json"
    fused_path.write_text(json.dumps(fused.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    fused.raw_response_path = str(fused_path)
    fused.timing = {
        "export_pdf_ms": export_pdf_ms,
        "render_pages_ms": render_pages_ms,
        "parallel_api_wall_ms": parallel_api_wall_ms,
        "fuse_ms": fuse_ms,
        "total_ms": int(round((time.perf_counter() - total_started_at) * 1000)),
    }
    return fused


@contextmanager
def _acquire_word_export_lock() -> Any:
    lock_path = Path(os.getenv("ESSAYGRADE_WORD_LOCK_FILE", str(WORD_EXPORT_LOCK_DEFAULT_PATH))).expanduser().resolve()
    timeout_seconds = max(30.0, float(os.getenv("ESSAYGRADE_WORD_LOCK_TIMEOUT_SECONDS", "1800")))
    poll_seconds = max(0.1, float(os.getenv("ESSAYGRADE_WORD_LOCK_POLL_SECONDS", "0.5")))
    deadline = time.monotonic() + timeout_seconds

    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()

        while True:
            handle.seek(0)
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError:
                if time.monotonic() >= deadline:
                    raise RuntimeError(
                        f"Timed out waiting for Word export lock: {lock_path} (>{timeout_seconds:.0f}s)"
                    )
                time.sleep(poll_seconds)

        handle.seek(0)
        handle.truncate()
        handle.write(f"{os.getpid()}\n".encode("ascii", "ignore"))
        handle.flush()
        try:
            yield
        finally:
            try:
                handle.seek(0)
                handle.truncate()
                handle.write(b"\0")
                handle.flush()
                handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass


def export_document_to_pdf(document_path: str, output_dir: str) -> Path:
    source = str(Path(document_path).expanduser().resolve())
    target_dir = _ensure_output_dir(output_dir, Path(source))
    pdf_path = target_dir / (Path(source).stem + ".visual_review.pdf")

    with _acquire_word_export_lock():
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            document = word.Documents.Open(
                FileName=source,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            try:
                document.SaveAs(str(pdf_path), FileFormat=WORD_FORMAT_PDF)
            finally:
                document.Close(False)
        finally:
            word.Quit()
    return pdf_path


def _ensure_output_dir(output_dir: str | None, document_path: Path) -> Path:
    base_dir = Path(output_dir).expanduser().resolve() if output_dir else document_path.parent
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir


def _resolve_moonshot_config(
    api_key: str | None,
    api_base_url: str | None,
    model: str | None,
) -> dict[str, str]:
    metadata: dict[str, Any] = {}
    if credential_entry_exists("moonshot_kimi"):
        entry = load_credential_entry("moonshot_kimi")
        metadata = entry.get("metadata") or {}
        if not api_key:
            api_key = entry.get("fields", {}).get("api_key")
        if not api_base_url:
            api_base_url = metadata.get("api_base_url")
        if metadata.get("default_model") and (not model or model == DEFAULT_MOONSHOT_MODEL):
            model = metadata.get("default_model")

    api_key = api_key or os.getenv("MOONSHOT_API_KEY")
    api_base_url = api_base_url or os.getenv("MOONSHOT_BASE_URL") or MOONSHOT_DEFAULT_BASE_URL
    model = resolve_moonshot_visual_model_name(model)

    if not api_key:
        raise RuntimeError("Moonshot/Kimi API Key 未配置，无法启用 Moonshot 视觉审稿。")

    return {
        "api_key": api_key,
        "api_base_url": api_base_url,
        "model": model,
    }


def _resolve_siliconflow_config(
    api_key: str | None,
    api_base_url: str | None,
    model: str | None,
) -> dict[str, str]:
    metadata: dict[str, Any] = {}
    if credential_entry_exists("siliconflow"):
        entry = load_credential_entry("siliconflow")
        metadata = entry.get("metadata") or {}
        if not api_key:
            api_key = entry.get("fields", {}).get("api_key")
        if not api_base_url:
            api_base_url = metadata.get("api_base_url")
        if metadata.get("default_model") and (not model or model == DEFAULT_SILICONFLOW_VISUAL_MODEL):
            model = metadata.get("default_model")

    api_key = api_key or os.getenv("SILICONFLOW_API_KEY")
    api_base_url = api_base_url or os.getenv("SILICONFLOW_BASE_URL") or SILICONFLOW_DEFAULT_BASE_URL
    model = resolve_siliconflow_visual_model_name(model)

    if not api_key:
        raise RuntimeError("SiliconFlow API Key 未配置，无法启用 SiliconFlow 视觉审稿。")

    return {
        "api_key": api_key,
        "api_base_url": api_base_url,
        "model": model,
    }


def _render_pdf_pages_to_data_urls(pdf_path: Path, output_dir: Path, max_pages: int) -> list[dict[str, str]]:
    if fitz is None:
        raise RuntimeError("缺少 PyMuPDF（fitz），无法把 PDF 渲染成图像供 Kimi 视觉接口使用。")

    pages_dir = output_dir / "moonshot_pages"
    pages_dir.mkdir(parents=True, exist_ok=True)
    rendered: list[dict[str, str]] = []
    with fitz.open(pdf_path) as document:
        selected_indexes = _select_representative_page_indexes(document.page_count, max_pages)
        for index in selected_indexes:
            page = document.load_page(index)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
            image_path = pages_dir / f"page_{index + 1:02d}.png"
            pix.save(str(image_path))
            data_url = "data:image/png;base64," + base64.b64encode(image_path.read_bytes()).decode("ascii")
            rendered.append(
                {
                    "page_label": f"第{index + 1}页",
                    "image_path": str(image_path),
                    "data_url": data_url,
                }
            )
    if not rendered:
        raise RuntimeError("PDF 未渲染出任何页面图像。")
    return rendered


def _select_representative_page_indexes(page_count: int, max_pages: int) -> list[int]:
    if page_count <= 0:
        return []
    if page_count <= max_pages:
        return list(range(page_count))

    indexes = {0, min(1, page_count - 1), page_count // 2, page_count - 1}
    step_count = max(max_pages - len(indexes), 0)
    for step in range(step_count):
        candidate = round((step + 1) * (page_count - 1) / (step_count + 1))
        indexes.add(int(candidate))
    return sorted(indexes)[:max_pages]


def _build_openai_visual_review_payload(pdf_path: Path, model: str, stage: str) -> dict:
    prompt = _visual_review_prompt(stage) + "\n" + _official_format_requirements_prompt()
    encoded_pdf = base64.b64encode(pdf_path.read_bytes()).decode("ascii")
    return {
        "model": model,
        "input": [
            {
                "role": "developer",
                "content": [
                    {
                        "type": "input_text",
                        "text": _visual_system_prompt(),
                    }
                ],
            },
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": prompt},
                    {
                        "type": "input_file",
                        "filename": pdf_path.name,
                        "file_data": encoded_pdf,
                    },
                ],
            },
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "paper_visual_review",
                "description": "Chinese thesis visual formatting review.",
                "strict": True,
                "schema": _visual_review_schema(),
            }
        },
        "reasoning": {"effort": "low"},
        "max_output_tokens": 2200,
    }


def _build_chat_visual_review_payload(pages: list[dict[str, str]], model: str, stage: str) -> dict:
    user_content: list[dict[str, Any]] = [
        {
            "type": "text",
            "text": _visual_review_prompt(stage) + "\n" + _official_format_requirements_prompt()
            + "\n请严格返回一个 JSON 对象，字段必须包含 overall_verdict、visual_order_score、confidence、major_issues、minor_issues、evidence、page_observations、notes。",
        }
    ]
    for page in pages:
        user_content.append({"type": "text", "text": page["page_label"]})
        user_content.append({"type": "image_url", "image_url": {"url": page["data_url"]}})

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _visual_system_prompt()},
            {"role": "user", "content": user_content},
        ],
        "temperature": 0.1,
    }
    if "glm-4.6v" not in model.lower():
        payload["response_format"] = {"type": "json_object"}
    return payload


def _post_json_request(
    url: str,
    payload: dict,
    api_key: str,
    extra_headers: dict[str, str] | None,
) -> dict:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    if extra_headers:
        headers.update(extra_headers)
    max_retries = max(1, int(os.getenv("VISUAL_REVIEW_MAX_RETRIES", "4")))
    backoff_seconds = max(1.0, float(os.getenv("VISUAL_REVIEW_RETRY_BACKOFF_SECONDS", "8")))
    last_error: Exception | None = None

    for attempt in range(max_retries):
        try:
            response = requests.post(
                url,
                headers=headers,
                json=payload,
                timeout=VISUAL_REVIEW_TIMEOUT_SECONDS,
            )
        except (requests.Timeout, requests.ConnectionError) as exc:
            last_error = exc
            if attempt < max_retries - 1:
                time.sleep(backoff_seconds * (attempt + 1))
                continue
            raise RuntimeError(f"Visual review request failed after {max_retries} retries: {exc}") from exc

        if response.status_code >= 400:
            retryable = response.status_code in VISUAL_REVIEW_RETRYABLE_STATUS_CODES
            if retryable and attempt < max_retries - 1:
                time.sleep(backoff_seconds * (attempt + 1))
                continue
            raise RuntimeError(_format_provider_error(response, url))

        return response.json()

    raise RuntimeError(f"Visual review request failed after {max_retries} retries: {last_error}")


def _format_provider_error(response: requests.Response, url: str) -> str:
    detail = f"HTTP {response.status_code}"
    error_type = ""
    try:
        payload = response.json()
    except ValueError:
        payload = None

    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict):
            error_type = str(error.get("type") or "").strip()
            message = str(error.get("message") or "").strip()
            if message:
                detail = f"{detail}: {message}"

    normalized_url = url.lower()
    if "api.kimi.com/coding" in normalized_url and error_type == "access_terminated_error":
        return (
            "Kimi Code 当前只允许官方支持的 Coding Agent 调用，"
            "不能直接作为论文视觉审稿后端使用。"
        )

    return detail


def _parse_openai_visual_response(response_json: dict) -> dict:
    output_text = response_json.get("output_text") or _collect_openai_output_text(response_json)
    if not output_text:
        raise RuntimeError("OpenAI 视觉审稿未返回可解析文本。")
    return _normalize_visual_review_json(output_text)


def _parse_chat_json_response(response_json: dict) -> dict:
    choices = response_json.get("choices") or []
    if not choices:
        raise RuntimeError("Moonshot/Kimi 未返回 choices。")
    content = choices[0].get("message", {}).get("content")
    if isinstance(content, list):
        content = "\n".join(
            str(item.get("text") or "")
            for item in content
            if isinstance(item, dict)
        ).strip()
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("Moonshot/Kimi 未返回可解析文本。")
    return _normalize_visual_review_json(content)


def _normalize_visual_review_json(raw_text: str) -> dict:
    parsed = json.loads(_extract_json_object(raw_text))
    verdict = _normalize_visual_verdict(parsed)
    score = _coerce_float(parsed.get("visual_order_score"))
    if score is None:
        score = _infer_visual_order_score(verdict, parsed)
    confidence = _coerce_float(parsed.get("confidence"))
    if confidence is None:
        confidence = 0.65
    return {
        "overall_verdict": verdict,
        "visual_order_score": max(0.0, min(10.0, float(score))),
        "confidence": max(0.0, min(1.0, float(confidence))),
        "major_issues": _coerce_string_list(parsed.get("major_issues")),
        "minor_issues": _coerce_string_list(parsed.get("minor_issues")),
        "evidence": _coerce_string_list(parsed.get("evidence")),
        "page_observations": _coerce_string_list(parsed.get("page_observations")),
        "notes": _coerce_string_list(parsed.get("notes")),
    }


def _coerce_float(value: Any) -> float | None:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def _coerce_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, dict):
        return [f"{key}: {val}".strip() for key, val in value.items() if str(val).strip()]
    if isinstance(value, str):
        stripped = value.strip()
        return [stripped] if stripped else []
    if isinstance(value, list):
        items: list[str] = []
        for item in value:
            if isinstance(item, dict):
                items.extend(_coerce_string_list(item))
                continue
            text = str(item).strip()
            if text:
                items.append(text)
        return items
    text = str(value).strip()
    return [text] if text else []


def _normalize_visual_verdict(parsed: dict[str, Any]) -> str:
    verdict = parsed.get("overall_verdict")
    if isinstance(verdict, str) and verdict in {"pass", "minor_issues", "major_revision", "rewrite"}:
        return verdict

    major_issues = _coerce_string_list(parsed.get("major_issues"))
    minor_issues = _coerce_string_list(parsed.get("minor_issues"))
    score = _coerce_float(parsed.get("visual_order_score"))
    numeric_verdict = _coerce_float(verdict)

    if major_issues:
        if len(major_issues) >= 2 or (score is not None and score <= 4.5):
            return "rewrite"
        return "major_revision"
    if minor_issues:
        return "minor_issues"
    if score is not None:
        if score >= 8.2:
            return "pass"
        if score >= 6.0:
            return "minor_issues"
        if score >= 4.2:
            return "major_revision"
        return "rewrite"
    if numeric_verdict is not None:
        if numeric_verdict <= 0:
            return "pass"
        if numeric_verdict <= 2:
            return "minor_issues"
        if numeric_verdict <= 3:
            return "major_revision"
        return "rewrite"
    return "minor_issues"


def _infer_visual_order_score(verdict: str, parsed: dict[str, Any]) -> float:
    major_count = len(_coerce_string_list(parsed.get("major_issues")))
    minor_count = len(_coerce_string_list(parsed.get("minor_issues")))
    if verdict == "pass":
        return 8.8
    if verdict == "minor_issues":
        return max(6.2, 8.0 - 0.35 * minor_count)
    if verdict == "major_revision":
        return max(4.2, 6.0 - 0.45 * major_count - 0.15 * minor_count)
    return max(2.0, 4.0 - 0.5 * major_count)


def _result_from_chat_visual_response(
    response_json: dict,
    mode: str,
    model: str,
    pdf_path: str,
    raw_response_path: str,
    note: str,
) -> VisualReviewResult:
    parsed = _parse_chat_json_response(response_json)
    notes = list(parsed["notes"])
    notes.insert(0, note)
    return VisualReviewResult(
        mode=mode,
        model=model,
        pdf_path=pdf_path,
        overall_verdict=parsed["overall_verdict"],
        visual_order_score=parsed["visual_order_score"],
        confidence=parsed["confidence"],
        major_issues=parsed["major_issues"],
        minor_issues=parsed["minor_issues"],
        evidence=parsed["evidence"],
        page_observations=parsed["page_observations"],
        notes=notes,
        raw_response_path=raw_response_path,
    )


def _fuse_visual_reviews(
    primary_result: VisualReviewResult,
    secondary_result: VisualReviewResult,
    pdf_path: str,
) -> VisualReviewResult:
    primary_score = primary_result.visual_order_score if primary_result.visual_order_score is not None else _infer_visual_order_score(primary_result.overall_verdict, {})
    secondary_score = secondary_result.visual_order_score if secondary_result.visual_order_score is not None else _infer_visual_order_score(secondary_result.overall_verdict, {})
    primary_confidence = min(max(primary_result.confidence or 0.7, 0.35), 0.95)
    secondary_confidence = min(max(secondary_result.confidence or 0.6, 0.35), 0.95) * 0.55

    fused_score = (
        primary_score * primary_confidence * 0.88
        + secondary_score * secondary_confidence * 0.12
    ) / max(primary_confidence * 0.88 + secondary_confidence * 0.12, 1e-6)

    major_issues = list(primary_result.major_issues)
    if primary_result.overall_verdict in {"major_revision", "rewrite"}:
        major_issues = _merge_issue_lists(primary_result.major_issues, secondary_result.major_issues[:2])

    minor_issues = _merge_issue_lists(primary_result.minor_issues, secondary_result.minor_issues[:2])
    evidence = _merge_issue_lists(primary_result.evidence, secondary_result.evidence[:2])
    page_observations = _merge_issue_lists(primary_result.page_observations, secondary_result.page_observations[:2])

    if primary_result.overall_verdict in {"rewrite", "major_revision"}:
        verdict = primary_result.overall_verdict
    elif primary_result.overall_verdict == "minor_issues":
        verdict = "minor_issues"
    elif secondary_result.overall_verdict in {"minor_issues", "major_revision", "rewrite"} and secondary_score < 8.0:
        verdict = "minor_issues"
    else:
        verdict = "pass"

    confidence = round(min(0.92, primary_confidence * 0.8 + secondary_confidence * 0.2), 3)
    notes = [
        f"SiliconFlow 双专家融合：主专家 {primary_result.model}，次专家 {secondary_result.model}。",
        "融合策略为确定性加权，且以主专家为准，次专家只做保守补充。",
    ]
    notes.extend(primary_result.notes[:2])
    notes.extend(secondary_result.notes[:2])

    return VisualReviewResult(
        mode="expert",
        model=f"{primary_result.model} + {secondary_result.model}",
        pdf_path=pdf_path,
        overall_verdict=verdict,
        visual_order_score=round(fused_score, 2),
        confidence=confidence,
        major_issues=major_issues,
        minor_issues=minor_issues,
        evidence=evidence,
        page_observations=page_observations,
        notes=_merge_issue_lists(notes),
        raw_response_path=None,
    )


def _merge_issue_lists(*collections: list[str]) -> list[str]:
    merged: list[str] = []
    for collection in collections:
        for item in collection:
            text = str(item).strip()
            if text and text not in merged:
                merged.append(text)
    return merged


def _extract_json_object(raw_text: str) -> str:
    raw_text = raw_text.strip()
    if raw_text.startswith("{") and raw_text.endswith("}"):
        return raw_text
    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise RuntimeError("视觉审稿响应中没有找到 JSON 对象。")
    return raw_text[start : end + 1]


def _collect_openai_output_text(response_json: dict) -> str:
    texts: list[str] = []
    for item in response_json.get("output", []):
        for content in item.get("content", []):
            if content.get("type") == "output_text" and content.get("text"):
                texts.append(content["text"])
    return "\n".join(texts).strip()


def _visual_system_prompt() -> str:
    return (
        "你是中文毕业论文格式审稿员。你只审页面视觉与排版，不审论文内容对错。"
        "请坚持“抓大放小”：整体观感和装订可读性优先，小误差可以记为小问题但不要夸大；"
        "若文档一眼看上去里出外进、空白失控、正文发黑发重、标题层级混乱、页码页眉视觉失衡，必须直接指出。"
    )


def _visual_review_prompt(stage: str) -> str:
    stage_label = "初稿" if stage == "initial_draft" else "终稿"
    return (
        f"请直接审看这份{stage_label}毕业论文的页面视觉呈现，像老师翻阅装订稿一样判断。"
        "重点看：\n"
        "1. 整体排版是否整齐，是否像标准论文成稿；\n"
        "2. 是否存在大块空白、连续空行、章节起落突兀、页码/页眉显得别扭；\n"
        "3. 正文字体是否视觉上稳定，宋体/仿宋等相近正文风格可视为接近，不要把这类小差异当成大错；\n"
        "4. 若正文大量使用黑体、无衬线重字族，或页内疏密严重失衡，要视为大问题；\n"
        "5. 请区分“大问题”和“小问题”：大问题是会破坏第一眼观感、影响装订稿质量的问题；小问题是还能接受但应修改的细节。"
    )


def _official_format_requirements_prompt() -> str:
    paper = OFFICIAL_FORMAT_REQUIREMENTS.get("paper_and_layout", {})
    margins = paper.get("margins_cm") or {}
    title_abstract = OFFICIAL_FORMAT_REQUIREMENTS.get("title_abstract_keywords", {})
    body = OFFICIAL_FORMAT_REQUIREMENTS.get("body_and_structure", {})

    keyword_range = title_abstract.get("keyword_range") or (3, 5)
    abstract_range = title_abstract.get("abstract_strict_range") or (120, 220)
    mandatory_sections = "、".join(str(item) for item in (body.get("mandatory_sections") or []))

    return (
        "【学校格式硬要求】\n"
        f"1) 纸张与页边距：A4，上{margins.get('top', 2.54)}cm，下{margins.get('bottom', 2.54)}cm，左{margins.get('left', 3.17)}cm，右{margins.get('right', 3.17)}cm，页眉{margins.get('header_distance', 1.5)}cm，页脚{margins.get('footer_distance', 1.75)}cm。\n"
        f"2) 页眉页码：页眉文本应接近“{paper.get('header_text', '')}”；前置部分（摘要、目录）用罗马数字，正文页码为“第M页”。\n"
        f"3) 标题/摘要/关键词：标题不超过{title_abstract.get('title_max_chars', 20)}字；摘要约{title_abstract.get('abstract_target_chars', 150)}字（严格区间{abstract_range[0]}-{abstract_range[1]}）；关键词{keyword_range[0]}-{keyword_range[1]}个。\n"
        f"4) 正文与结构：章节题目{body.get('heading_font', '小四黑体')}，正文{body.get('body_font', '五号宋体')}；必备部分包括：{mandatory_sections}。\n"
        "请先按这些硬要求判断，再做视觉层面的“像不像标准论文”判断。"
    )


def _visual_review_schema() -> dict:
    return {
        "type": "object",
        "properties": {
            "overall_verdict": {
                "type": "string",
                "enum": ["pass", "minor_issues", "major_revision", "rewrite"],
            },
            "visual_order_score": {"type": "number"},
            "confidence": {"type": "number"},
            "major_issues": {"type": "array", "items": {"type": "string"}},
            "minor_issues": {"type": "array", "items": {"type": "string"}},
            "evidence": {"type": "array", "items": {"type": "string"}},
            "page_observations": {"type": "array", "items": {"type": "string"}},
            "notes": {"type": "array", "items": {"type": "string"}},
        },
        "required": [
            "overall_verdict",
            "visual_order_score",
            "confidence",
            "major_issues",
            "minor_issues",
            "evidence",
            "page_observations",
            "notes",
        ],
        "additionalProperties": False,
    }
