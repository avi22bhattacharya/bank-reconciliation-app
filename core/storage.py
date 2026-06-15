"""Persistent file storage: S3 in production, local filesystem in dev.

Detects mode via st.secrets["aws"]. When aws secrets are absent (local dev),
every function falls back to plain filesystem I/O using the path as-is.
"""

from __future__ import annotations

from pathlib import Path


def _cfg():
    import streamlit as st
    return st.secrets.get("aws", {})


def is_cloud() -> bool:
    try:
        return bool(_cfg())
    except Exception:
        return False


def _client():
    import boto3
    cfg = _cfg()
    return boto3.client(
        "s3",
        aws_access_key_id=cfg["access_key_id"],
        aws_secret_access_key=cfg["secret_access_key"],
        region_name=cfg["region"],
    )


def _bucket() -> str:
    return _cfg()["bucket"]


def upload(local_path: str | Path, key: str) -> str:
    """Upload local_path to S3 key. Returns key.
    No-op in local dev (returns local_path unchanged)."""
    if not is_cloud():
        return str(local_path)
    _client().upload_file(str(local_path), _bucket(), key)
    return key


def read_bytes(path_or_key: str) -> bytes:
    """Read a file from S3 (cloud) or local filesystem (dev)."""
    if is_cloud():
        obj = _client().get_object(Bucket=_bucket(), Key=path_or_key)
        return obj["Body"].read()
    return Path(path_or_key).read_bytes()
