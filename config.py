"""Configuration helpers for local development and Streamlit Community Cloud.

Values can come from environment variables or Streamlit secrets. Secrets always stay
outside GitHub; see secrets.example.toml.
"""
from __future__ import annotations

import os


def get_setting(name: str, default=None):
    value = os.environ.get(name)
    if value not in (None, ""):
        return value

    try:
        import streamlit as st
        try:
            value = st.secrets.get(name, default)
        except Exception:
            value = default
        if value not in (None, ""):
            return value
    except Exception:
        pass

    return default


def get_bool_setting(name: str, default: bool = False) -> bool:
    value = get_setting(name, default)
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}


def cloud_configured() -> bool:
    return bool(get_setting("SUPABASE_URL")) and bool(get_setting("SUPABASE_KEY"))


def get_data_mode() -> str:
    explicit = str(get_setting("PROPOSAL_DATA_MODE", "")).strip().lower()
    if explicit in {"local", "cloud"}:
        return explicit
    return "cloud" if cloud_configured() else "local"


def get_file_mode() -> str:
    explicit = str(get_setting("PROPOSAL_FILE_MODE", "")).strip().lower()
    if explicit in {"local", "cloud"}:
        return explicit
    return "cloud" if cloud_configured() else "local"


def get_admin_users() -> list[str]:
    raw = get_setting("ADMIN_USERS", "Jen Braziel,Melanie Moore")
    if isinstance(raw, (list, tuple)):
        return [str(x).strip() for x in raw if str(x).strip()]
    return [x.strip() for x in str(raw).split(",") if x.strip()]
