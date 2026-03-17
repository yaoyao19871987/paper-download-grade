from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import win32crypt


def repo_root() -> Path:
    return Path(__file__).resolve().parents[4]


def credential_store_root() -> Path:
    return repo_root() / ".credential_store"


def credential_entry_path(service: str) -> Path:
    return credential_store_root() / f"{service}.json"


def credential_entry_exists(service: str) -> bool:
    return credential_entry_path(service).exists()


def load_credential_entry(service: str) -> dict[str, Any]:
    path = credential_entry_path(service)
    if not path.exists():
        raise FileNotFoundError(f"Credential entry not found: {path}")

    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if payload.get("protection") not in {None, "DPAPI_CURRENT_USER"}:
        raise RuntimeError(f"Unsupported credential protection mode: {payload.get('protection')}")

    decrypted_fields: dict[str, str] = {}
    for key, value in (payload.get("fields") or {}).items():
        cipher_bytes = bytes.fromhex(value)
        plain_bytes = win32crypt.CryptUnprotectData(cipher_bytes, None, None, None, 0)[1]
        decrypted_fields[key] = plain_bytes.decode("utf-16-le").rstrip("\x00")

    return {
        "service": payload.get("service") or service,
        "saved_at": payload.get("saved_at"),
        "metadata": payload.get("metadata") or {},
        "fields": decrypted_fields,
        "path": str(path.resolve()),
    }
