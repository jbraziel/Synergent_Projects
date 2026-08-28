"""Persistent file storage abstraction.

Local mode keeps the current filesystem behavior. Cloud mode stores proposal artifacts in
Supabase Storage and returns stable references of the form supabase://bucket/path.
"""
from __future__ import annotations

import mimetypes
import os
import shutil
from pathlib import Path
from functools import lru_cache
from urllib.parse import quote

from config import get_file_mode, get_setting


CLOUD_PREFIX = "supabase://"


def is_cloud_mode() -> bool:
    return get_file_mode() == "cloud"


def _client():
    from cloud_client import get_supabase_client
    return get_supabase_client()


def _bucket_name() -> str:
    return str(get_setting("SUPABASE_BUCKET", "proposal-files"))


def _parse_ref(ref: str):
    if not ref or not ref.startswith(CLOUD_PREFIX):
        return None, None
    remainder = ref[len(CLOUD_PREFIX):]
    bucket, _, path = remainder.partition("/")
    return bucket, path


@lru_cache(maxsize=1)
def ensure_cloud_bucket():
    if not is_cloud_mode():
        return
    bucket = _bucket_name()
    client = _client()
    try:
        client.storage.get_bucket(bucket)
    except Exception:
        try:
            client.storage.create_bucket(bucket, options={"public": False})
        except Exception as exc:
            raise RuntimeError(
                f"Unable to access or create the private Supabase Storage bucket '{bucket}'. "
                "Create it in Supabase Storage or use a server-side key with storage permissions."
            ) from exc


def make_object_path(credit_union: str, folder: str, file_name: str) -> str:
    safe_cu = "".join(c for c in credit_union if c.isalnum() or c in " -_").strip() or "Unknown Credit Union"
    safe_name = "".join(c for c in file_name if c.isalnum() or c in " .-_()[]").strip()
    return f"{safe_cu}/{folder}/{safe_name}"


def store_file(local_path: str, object_path: str | None = None) -> str:
    """Persist a local file and return the reference that should be saved in the proposal."""
    if not is_cloud_mode():
        return str(local_path)

    ensure_cloud_bucket()
    local_path = str(local_path)
    if object_path is None:
        object_path = Path(local_path).name
    object_path = object_path.replace("\\", "/").lstrip("/")
    content_type = mimetypes.guess_type(local_path)[0] or "application/octet-stream"
    bucket = _bucket_name()
    with open(local_path, "rb") as f:
        _client().storage.from_(bucket).upload(
            path=object_path,
            file=f,
            file_options={"content-type": content_type, "upsert": "true"},
        )
    return f"{CLOUD_PREFIX}{bucket}/{object_path}"


def store_bytes(data: bytes, object_path: str, content_type: str = "application/octet-stream") -> str:
    if not is_cloud_mode():
        local_root = Path(os.environ.get("PROPOSAL_OUTPUT_ROOT", "generated_proposals"))
        local_path = local_root / object_path
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(data)
        return str(local_path)

    ensure_cloud_bucket()
    object_path = object_path.replace("\\", "/").lstrip("/")
    bucket = _bucket_name()
    _client().storage.from_(bucket).upload(
        path=object_path,
        file=data,
        file_options={"content-type": content_type, "upsert": "true"},
    )
    return f"{CLOUD_PREFIX}{bucket}/{object_path}"


def read_bytes(ref: str) -> bytes:
    if not ref:
        raise FileNotFoundError("No stored file reference was provided.")
    if ref.startswith(CLOUD_PREFIX):
        bucket, path = _parse_ref(ref)
        return _client().storage.from_(bucket).download(path)
    return Path(ref).read_bytes()


def stored_file_exists(ref: str) -> bool:
    if not ref:
        return False
    if ref.startswith(CLOUD_PREFIX):
        # A cloud reference is only persisted after a successful upload. Avoid downloading every
        # file on every Proposal Library rerun simply to check its existence.
        return True
    return os.path.exists(ref)



def get_signed_download_urls(refs, expires_in: int = 3600) -> dict[str, str]:
    """Create time-limited browser download URLs for many private cloud files at once.

    Proposal Library used to download every visible PPTX/PDF/CSV into the Streamlit
    process just to render its download buttons. This batches URL signing instead, so
    the actual file bytes move only when the user clicks a file button.
    """
    refs = [ref for ref in dict.fromkeys(refs or []) if ref]
    if not refs:
        return {}

    # Local files still use Streamlit's normal download_button flow.
    if not is_cloud_mode():
        return {}

    grouped: dict[str, list[tuple[str, str]]] = {}
    for ref in refs:
        bucket, path = _parse_ref(ref)
        if bucket and path:
            grouped.setdefault(bucket, []).append((ref, path))

    result: dict[str, str] = {}
    client = _client()

    for bucket, items in grouped.items():
        paths = [path for _, path in items]
        proxy = client.storage.from_(bucket)

        try:
            response = proxy.create_signed_urls(paths, int(expires_in))
            data = getattr(response, "data", response)
            if isinstance(data, dict):
                data = (
                    data.get("data")
                    or data.get("signedUrls")
                    or data.get("signed_urls")
                    or []
                )
            if not isinstance(data, list):
                data = []
        except Exception:
            # Compatibility fallback for storage client versions without batch signing.
            data = []
            for path in paths:
                one = proxy.create_signed_url(path, int(expires_in))
                one_data = getattr(one, "data", one)
                if isinstance(one_data, dict) and "data" in one_data and isinstance(one_data["data"], dict):
                    one_data = one_data["data"]
                data.append(one_data if isinstance(one_data, dict) else {"signedURL": one_data, "path": path})

        by_path = {}
        for index, item in enumerate(data):
            if isinstance(item, str):
                item = {"signedURL": item}
            if not isinstance(item, dict):
                continue
            signed = (
                item.get("signedURL")
                or item.get("signedUrl")
                or item.get("signed_url")
                or item.get("url")
            )
            item_path = item.get("path") or (paths[index] if index < len(paths) else None)
            if signed and item_path:
                by_path[item_path] = signed

        for ref, path in items:
            signed = by_path.get(path)
            if not signed:
                continue
            separator = "&" if "?" in signed else "?"
            result[ref] = f"{signed}{separator}download={quote(Path(path).name)}"

    return result


def display_name(ref: str) -> str:
    if not ref:
        return ""
    if ref.startswith(CLOUD_PREFIX):
        _, path = _parse_ref(ref)
        return Path(path).name
    return os.path.basename(ref)


def copy_stored_file(source_ref: str, destination_local_path: str | None = None, destination_object_path: str | None = None) -> str:
    if not source_ref:
        raise FileNotFoundError("No source file is available to copy.")

    if source_ref.startswith(CLOUD_PREFIX):
        bucket, source_path = _parse_ref(source_ref)
        if not destination_object_path:
            raise ValueError("destination_object_path is required when copying a cloud file.")

        # Supabase Storage's server-side copy operation can fail depending on the
        # storage API/client version and will also fail when the destination object
        # already exists.  For proposal lifecycle snapshots, reliability is more
        # important than avoiding a small download/upload, so materialize the source
        # bytes and upload the destination with upsert enabled.
        destination_object_path = destination_object_path.replace("\\", "/").lstrip("/")
        client = _client()
        data = client.storage.from_(bucket).download(source_path)
        content_type = mimetypes.guess_type(destination_object_path)[0] or "application/octet-stream"
        client.storage.from_(bucket).upload(
            path=destination_object_path,
            file=data,
            file_options={"content-type": content_type, "upsert": "true"},
        )
        return f"{CLOUD_PREFIX}{bucket}/{destination_object_path}"

    if destination_local_path:
        Path(destination_local_path).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_ref, destination_local_path)
        return str(destination_local_path)

    if destination_object_path and is_cloud_mode():
        return store_file(source_ref, destination_object_path)

    raise ValueError("A destination path is required.")


def materialize_to_temp(ref: str, target_path: str) -> str:
    """Write a stored file to a local staging path for tools that require a real filesystem path."""
    Path(target_path).parent.mkdir(parents=True, exist_ok=True)
    Path(target_path).write_bytes(read_bytes(ref))
    return target_path


def get_storage_status():
    return {
        "file_mode": "Supabase Storage" if is_cloud_mode() else "Local filesystem",
        "bucket": _bucket_name() if is_cloud_mode() else None,
    }
