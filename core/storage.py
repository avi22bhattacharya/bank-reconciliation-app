"""Persistent file storage: Supabase Storage in production, local filesystem in dev.

Uses the Supabase Storage REST API directly via httpx to avoid initialisation
issues in the supabase-py client. Detects mode via st.secrets["supabase"].
"""

from __future__ import annotations

from pathlib import Path


def _cfg() -> dict:
    import streamlit as st
    return st.secrets.get("supabase", {})


def is_cloud() -> bool:
    try:
        return bool(_cfg())
    except Exception:
        return False


def _headers() -> dict:
    key = _cfg()["key"]
    return {"apikey": key, "Authorization": f"Bearer {key}"}


def _object_url(key: str) -> str:
    cfg = _cfg()
    return f"{cfg['url']}/storage/v1/object/{cfg['bucket']}/{key}"


def upload(local_path: str | Path, key: str) -> str:
    """Upload local_path to Supabase Storage at key. Returns key.
    No-op in local dev (returns local_path unchanged)."""
    if not is_cloud():
        return str(local_path)
    import httpx
    data = Path(local_path).read_bytes()
    headers = {
        **_headers(),
        "Content-Type": "application/octet-stream",
        "x-upsert": "true",
    }
    with httpx.Client(timeout=60) as client:
        r = client.post(_object_url(key), content=data, headers=headers)
        r.raise_for_status()
    return key


def read_bytes(path_or_key: str) -> bytes:
    """Read a file from Supabase Storage (cloud) or local filesystem (dev)."""
    if is_cloud():
        import httpx
        with httpx.Client(timeout=60) as client:
            r = client.get(_object_url(path_or_key), headers=_headers())
            r.raise_for_status()
        return r.content
    return Path(path_or_key).read_bytes()
