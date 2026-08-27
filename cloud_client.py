"""Lazy Supabase client shared by the database and file-storage layers."""
from __future__ import annotations

from functools import lru_cache
from config import get_setting


@lru_cache(maxsize=1)
def get_supabase_client():
    url = get_setting("SUPABASE_URL")
    key = get_setting("SUPABASE_KEY")
    if not url or not key:
        raise RuntimeError(
            "Cloud mode is enabled but SUPABASE_URL / SUPABASE_KEY are missing. "
            "Add them to Streamlit Secrets or switch PROPOSAL_DATA_MODE to local."
        )

    try:
        from supabase import create_client
    except ImportError as exc:
        raise RuntimeError(
            "Cloud mode requires the 'supabase' Python package. Install requirements.txt."
        ) from exc

    return create_client(url, key)
