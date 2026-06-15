"""Persistent file storage: Supabase Storage in production, local filesystem in dev.

Detects mode via st.secrets["supabase"]. When absent (local dev), every
function falls back to plain filesystem I/O using the path as-is.

Supabase setup: create a private bucket in your project's Storage dashboard
and put the bucket name in st.secrets["supabase"]["bucket"].
"""

from __future__ import annotations

from pathlib import Path


def _cfg():
    import streamlit as st
    return st.secrets.get("supabase", {})


def is_cloud() -> bool:
    try:
        return bool(_cfg())
    except Exception:
        return False


def _client():
    from supabase import create_client
    cfg = _cfg()
    return create_client(cfg["url"], cfg["key"])


def _bucket() -> str:
    return _cfg()["bucket"]


def upload(local_path: str | Path, key: str) -> str:
    """Upload local_path to Supabase Storage at key. Returns key.
    No-op in local dev (returns local_path unchanged)."""
    if not is_cloud():
        return str(local_path)
    data = Path(local_path).read_bytes()
    _client().storage.from_(_bucket()).upload(
        key, data,
        file_options={"content-type": "application/octet-stream", "upsert": "true"},
    )
    return key


def read_bytes(path_or_key: str) -> bytes:
    """Read a file from Supabase Storage (cloud) or local filesystem (dev)."""
    if is_cloud():
        return bytes(_client().storage.from_(_bucket()).download(path_or_key))
    return Path(path_or_key).read_bytes()
