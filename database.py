from __future__ import annotations

import json
import os
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path

from config import get_data_mode


DB_NAME = Path(os.environ.get("PROPOSAL_DB_PATH", str(Path(__file__).parent / "proposals.db")))
DB_NAME.parent.mkdir(parents=True, exist_ok=True)


DEFAULT_PRICING_SETTINGS = [
    # Labor
    dict(key="creative_hourly_rate", category="Labor Rates", label="Creative Concept & Design", value=115.0, value_type="currency", unit="per hour", description="Hourly creative/design rate", sort_order=10, is_repeating=False, active=True),
    dict(key="programming_hourly_rate", category="Labor Rates", label="Programming / Data Mining", value=200.0, value_type="currency", unit="per hour", description="Hourly programming and targeted data-mining rate", sort_order=20, is_repeating=False, active=True),
    dict(key="email_development_hourly_rate", category="Labor Rates", label="Email Development", value=115.0, value_type="currency", unit="per hour", description="Hourly email development rate", sort_order=30, is_repeating=False, active=True),
    # Production / sends
    dict(key="list_markup_pct", category="Production & Markups", label="List Procurement Markup", value=35.0, value_type="percent", unit="markup", description="Markup applied to raw purchased-list costs", sort_order=10, is_repeating=True, active=True),
    dict(key="print_markup_pct", category="Production & Markups", label="Variable Print Production Markup", value=35.0, value_type="percent", unit="markup", description="Markup applied to raw print production costs", sort_order=20, is_repeating=True, active=True),
    dict(key="email_send_fee", category="Production & Markups", label="Email Send Fee", value=100.0, value_type="currency", unit="per send", description="Internal fee per email send", sort_order=30, is_repeating=True, active=True),
    dict(key="four_campaign_discount_pct", category="Production & Markups", label="Four-Campaign Discount", value=10.0, value_type="percent", unit="discount", description="Discount applied to the calculated four-campaign package", sort_order=40, is_repeating=False, active=True),
    # Fixed costs
    dict(key="fixed_preliminary_data_analysis", category="Fixed Costs", label="Preliminary Data Analysis", value=100.0, value_type="currency", unit="fixed", description="", sort_order=10, is_repeating=False, active=True),
    dict(key="fixed_strategy_concept", category="Fixed Costs", label="Strategy & Concept", value=300.0, value_type="currency", unit="fixed", description="", sort_order=20, is_repeating=False, active=True),
    dict(key="fixed_copy_writing", category="Fixed Costs", label="Copy Writing", value=100.0, value_type="currency", unit="fixed", description="", sort_order=30, is_repeating=False, active=True),
    dict(key="fixed_unique_url", category="Fixed Costs", label="Unique URL", value=150.0, value_type="currency", unit="fixed", description="", sort_order=40, is_repeating=False, active=True),
    dict(key="fixed_tracking_reporting", category="Fixed Costs", label="Tracking, Monitoring & Reporting", value=100.0, value_type="currency", unit="fixed", description="", sort_order=50, is_repeating=False, active=True),
    dict(key="fixed_campaign_management", category="Fixed Costs", label="Campaign Management", value=100.0, value_type="currency", unit="fixed", description="", sort_order=60, is_repeating=False, active=True),
    dict(key="fixed_consultative_implementation", category="Fixed Costs", label="Consultative Implementation", value=500.0, value_type="currency", unit="fixed", description="", sort_order=70, is_repeating=False, active=True),
    dict(key="fixed_qr_code", category="Fixed Costs", label="QR Code", value=100.0, value_type="currency", unit="fixed", description="", sort_order=80, is_repeating=False, active=True),
    dict(key="fixed_ach_program_fee", category="Fixed Costs", label="ACH Program Fee", value=2500.0, value_type="currency", unit="fixed", description="Repeats for each campaign in multi-campaign packages", sort_order=90, is_repeating=True, active=True),
    # EMP base tiers
    dict(key="emp_tier_1_base", category="EMP Monthly Pricing", label="Tier 1 Base (2,500–4,999 subscribers)", value=59.54, value_type="currency", unit="monthly", description="", sort_order=10, is_repeating=False, active=True),
    dict(key="emp_tier_2_base", category="EMP Monthly Pricing", label="Tier 2 Base (5,000–9,999 subscribers)", value=108.14, value_type="currency", unit="monthly", description="", sort_order=20, is_repeating=False, active=True),
    dict(key="emp_tier_3_base", category="EMP Monthly Pricing", label="Tier 3 Base (10,000–14,999 subscribers)", value=156.74, value_type="currency", unit="monthly", description="", sort_order=30, is_repeating=False, active=True),
    dict(key="emp_tier_4_base", category="EMP Monthly Pricing", label="Tier 4 Base (15,000–24,999 subscribers)", value=241.79, value_type="currency", unit="monthly", description="", sort_order=40, is_repeating=False, active=True),
    dict(key="emp_tier_5_base", category="EMP Monthly Pricing", label="Tier 5 Base (25,000–49,999 subscribers)", value=363.29, value_type="currency", unit="monthly", description="", sort_order=50, is_repeating=False, active=True),
    dict(key="emp_tier_6_base", category="EMP Monthly Pricing", label="Tier 6 Base (50,000–74,999 subscribers)", value=545.54, value_type="currency", unit="monthly", description="", sort_order=60, is_repeating=False, active=True),
    dict(key="emp_essentials_addon", category="EMP Monthly Pricing", label="Essentials Monthly Add-On", value=100.0, value_type="currency", unit="monthly", description="Added to the tier base", sort_order=70, is_repeating=False, active=True),
    dict(key="emp_premium_addon", category="EMP Monthly Pricing", label="Premium Monthly Add-On", value=200.0, value_type="currency", unit="monthly", description="Added to the tier base", sort_order=80, is_repeating=False, active=True),
    dict(key="emp_elite_addon", category="EMP Monthly Pricing", label="Elite Monthly Add-On", value=200.0, value_type="currency", unit="monthly", description="Added to the tier base", sort_order=90, is_repeating=False, active=True),
    # EMP implementation
    dict(key="emp_essentials_implementation", category="EMP Implementation", label="Essentials Implementation", value=5500.0, value_type="currency", unit="one-time", description="", sort_order=10, is_repeating=False, active=True),
    dict(key="emp_premium_implementation", category="EMP Implementation", label="Premium Implementation", value=8000.0, value_type="currency", unit="one-time", description="", sort_order=20, is_repeating=False, active=True),
    dict(key="emp_elite_implementation", category="EMP Implementation", label="Elite Implementation", value=10500.0, value_type="currency", unit="one-time", description="", sort_order=30, is_repeating=False, active=True),
]


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_cloud_mode() -> bool:
    return get_data_mode() == "cloud"


def get_connection():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout = 30000;")
    return conn


def add_column_if_missing(cursor, table, column, definition):
    cursor.execute(f"PRAGMA table_info({table})")
    columns = [row[1] for row in cursor.fetchall()]
    if column not in columns:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def _cloud_client():
    from cloud_client import get_supabase_client
    return get_supabase_client()


def initialize_database():
    if is_cloud_mode():
        try:
            client = _cloud_client()
            # These selects double as a friendly schema check.
            client.table("proposals").select("id").limit(1).execute()
            client.table("pricing_settings").select("key").limit(1).execute()
            client.table("pricing_history").select("id").limit(1).execute()
            client.table("proposal_pricing_snapshots").select("id").limit(1).execute()
            _ensure_default_pricing_cloud(client)
        except Exception as exc:
            raise RuntimeError(
                "Supabase is configured, but the Proposal Generator tables are not ready. "
                "Run supabase_schema.sql in the Supabase SQL Editor, then restart the app. "
                f"Technical detail: {exc}"
            ) from exc
        return

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proposals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposal_name TEXT NOT NULL,
            credit_union TEXT NOT NULL,
            proposal_type TEXT NOT NULL,
            msr TEXT DEFAULT '',
            status TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            saved_data_json TEXT NOT NULL,
            updated_by TEXT,
            locked_by TEXT,
            locked_at TEXT,
            copied_from_proposal_id INTEGER
        )
    """)
    add_column_if_missing(cursor, "proposals", "updated_by", "TEXT")
    add_column_if_missing(cursor, "proposals", "locked_by", "TEXT")
    add_column_if_missing(cursor, "proposals", "locked_at", "TEXT")
    add_column_if_missing(cursor, "proposals", "copied_from_proposal_id", "INTEGER")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pricing_settings (
            key TEXT PRIMARY KEY,
            category TEXT NOT NULL,
            label TEXT NOT NULL,
            value REAL NOT NULL,
            value_type TEXT NOT NULL DEFAULT 'currency',
            unit TEXT DEFAULT '',
            description TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,
            is_repeating INTEGER NOT NULL DEFAULT 0,
            active INTEGER NOT NULL DEFAULT 1,
            updated_at TEXT,
            updated_by TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pricing_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            setting_key TEXT NOT NULL,
            old_value REAL,
            new_value REAL NOT NULL,
            old_is_repeating INTEGER,
            new_is_repeating INTEGER,
            old_active INTEGER,
            new_active INTEGER,
            changed_at TEXT NOT NULL,
            changed_by TEXT
        )
    """)
    add_column_if_missing(cursor, "pricing_history", "old_active", "INTEGER")
    add_column_if_missing(cursor, "pricing_history", "new_active", "INTEGER")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proposal_pricing_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposal_id INTEGER NOT NULL,
            generated_at TEXT NOT NULL,
            generated_by TEXT,
            pricing_json TEXT NOT NULL
        )
    """)

    for item in DEFAULT_PRICING_SETTINGS:
        cursor.execute("""
            INSERT OR IGNORE INTO pricing_settings
            (key, category, label, value, value_type, unit, description, sort_order, is_repeating, active, updated_at, updated_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item["key"], item["category"], item["label"], item["value"], item["value_type"],
            item["unit"], item["description"], item["sort_order"], int(item["is_repeating"]),
            int(item["active"]), _now(), "System Default"
        ))
    conn.commit()
    conn.close()


def _ensure_default_pricing_cloud(client):
    existing = client.table("pricing_settings").select("key").execute().data or []
    existing_keys = {row["key"] for row in existing}
    missing = []
    now = _now()
    for item in DEFAULT_PRICING_SETTINGS:
        if item["key"] not in existing_keys:
            row = dict(item)
            row["updated_at"] = now
            row["updated_by"] = "System Default"
            missing.append(row)
    if missing:
        client.table("pricing_settings").insert(missing).execute()


def save_proposal(proposal_id, proposal_name, credit_union, proposal_type, status, saved_data, msr="", updated_by="", copied_from_proposal_id=None):
    now = _now()
    if is_cloud_mode():
        client = _cloud_client()
        payload = {
            "proposal_name": proposal_name,
            "credit_union": credit_union,
            "proposal_type": proposal_type,
            "msr": msr,
            "status": status,
            "updated_at": now,
            "saved_data_json": saved_data,
            "updated_by": updated_by,
        }
        if copied_from_proposal_id is not None:
            payload["copied_from_proposal_id"] = copied_from_proposal_id
        if proposal_id:
            response = client.table("proposals").update(payload).eq("id", int(proposal_id)).execute()
            return int(proposal_id)
        payload["created_at"] = now
        payload["copied_from_proposal_id"] = copied_from_proposal_id
        response = client.table("proposals").insert(payload).execute()
        if not response.data:
            raise RuntimeError("Supabase did not return the newly created proposal ID.")
        return int(response.data[0]["id"])

    saved_json = json.dumps(saved_data)
    for attempt in range(5):
        conn = get_connection()
        cursor = conn.cursor()
        try:
            if proposal_id:
                cursor.execute("""
                    UPDATE proposals
                    SET proposal_name = ?, credit_union = ?, proposal_type = ?, msr = ?, status = ?,
                        updated_at = ?, saved_data_json = ?, updated_by = ?,
                        copied_from_proposal_id = COALESCE(?, copied_from_proposal_id)
                    WHERE id = ?
                """, (proposal_name, credit_union, proposal_type, msr, status, now, saved_json,
                      updated_by, copied_from_proposal_id, proposal_id))
            else:
                cursor.execute("""
                    INSERT INTO proposals
                    (proposal_name, credit_union, proposal_type, msr, status, created_at, updated_at,
                     saved_data_json, updated_by, copied_from_proposal_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (proposal_name, credit_union, proposal_type, msr, status, now, now, saved_json,
                      updated_by, copied_from_proposal_id))
                proposal_id = cursor.lastrowid
            conn.commit()
            conn.close()
            return proposal_id
        except sqlite3.OperationalError as exc:
            conn.close()
            if "locked" in str(exc).lower() and attempt < 4:
                time.sleep(0.5)
                continue
            raise


def search_proposals(search_text="", status_filter="All", msr_filter="All"):
    if is_cloud_mode():
        rows = (_cloud_client().table("proposals")
                .select("id,proposal_name,credit_union,proposal_type,status,updated_at,msr,updated_by,locked_by,locked_at")
                .order("updated_at", desc=True).execute().data or [])
        text = search_text.strip().lower()
        legacy = {"Shannan Heacock": "Shannan", "Erica Vachon": "Erica"}
        allowed_msr = {msr_filter, legacy.get(msr_filter)} if msr_filter != "All" else None
        filtered = []
        for row in rows:
            if text and not any(text in str(row.get(k, "")).lower() for k in ("proposal_name", "credit_union", "proposal_type", "msr")):
                continue
            if status_filter != "All" and row.get("status") != status_filter:
                continue
            if allowed_msr and row.get("msr") not in allowed_msr:
                continue
            filtered.append((row.get("id"), row.get("proposal_name"), row.get("credit_union"),
                             row.get("proposal_type"), row.get("status"), row.get("updated_at"),
                             row.get("msr"), row.get("updated_by"), row.get("locked_by"), row.get("locked_at")))
        return filtered

    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT id, proposal_name, credit_union, proposal_type, status, updated_at, msr, updated_by, locked_by, locked_at
        FROM proposals WHERE 1=1
    """
    params = []
    if search_text.strip():
        query += " AND (proposal_name LIKE ? OR credit_union LIKE ? OR proposal_type LIKE ? OR msr LIKE ?)"
        value = f"%{search_text.strip()}%"
        params.extend([value, value, value, value])
    if status_filter != "All":
        query += " AND status = ?"
        params.append(status_filter)
    if msr_filter != "All":
        legacy_msr_names = {"Shannan Heacock": "Shannan", "Erica Vachon": "Erica"}
        legacy_name = legacy_msr_names.get(msr_filter)
        if legacy_name:
            query += " AND msr IN (?, ?)"
            params.extend([msr_filter, legacy_name])
        else:
            query += " AND msr = ?"
            params.append(msr_filter)
    query += " ORDER BY updated_at DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


def load_proposal(proposal_id):
    if is_cloud_mode():
        rows = (_cloud_client().table("proposals").select("saved_data_json")
                .eq("id", int(proposal_id)).limit(1).execute().data or [])
        if not rows:
            return None
        value = rows[0].get("saved_data_json")
        if isinstance(value, dict):
            return value
        try:
            return json.loads(value)
        except (TypeError, json.JSONDecodeError):
            return None

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT saved_data_json FROM proposals WHERE id = ?", (proposal_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    try:
        return json.loads(row[0])
    except json.JSONDecodeError:
        return None


def delete_proposal(proposal_id):
    if is_cloud_mode():
        _cloud_client().table("proposals").delete().eq("id", int(proposal_id)).execute()
        return
    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("DELETE FROM proposals WHERE id = ?", (proposal_id,))
    conn.commit(); conn.close()


def update_proposal_status(proposal_id, status):
    now = _now()
    if is_cloud_mode():
        _cloud_client().table("proposals").update({"status": status, "updated_at": now}).eq("id", int(proposal_id)).execute()
        return
    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("UPDATE proposals SET status = ?, updated_at = ? WHERE id = ?", (status, now, proposal_id))
    conn.commit(); conn.close()


def lock_proposal(proposal_id, user_name):
    now = _now()
    if is_cloud_mode():
        _cloud_client().table("proposals").update({"locked_by": user_name, "locked_at": now}).eq("id", int(proposal_id)).execute()
        return
    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("UPDATE proposals SET locked_by = ?, locked_at = ? WHERE id = ?", (user_name, now, proposal_id))
    conn.commit(); conn.close()


def unlock_proposal(proposal_id, user_name):
    if is_cloud_mode():
        (_cloud_client().table("proposals").update({"locked_by": None, "locked_at": None})
         .eq("id", int(proposal_id)).eq("locked_by", user_name).execute())
        return
    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("UPDATE proposals SET locked_by = NULL, locked_at = NULL WHERE id = ? AND locked_by = ?", (proposal_id, user_name))
    conn.commit(); conn.close()


def get_pricing_settings(include_inactive: bool = True):
    if is_cloud_mode():
        rows = (_cloud_client().table("pricing_settings").select("*")
                .order("category").order("sort_order").execute().data or [])
    else:
        conn = get_connection(); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
        cursor.execute("SELECT * FROM pricing_settings ORDER BY category, sort_order, label")
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()
    normalized = []
    for row in rows:
        item = dict(row)
        item["value"] = float(item.get("value", 0) or 0)
        item["is_repeating"] = bool(item.get("is_repeating", False))
        item["active"] = bool(item.get("active", True))
        if include_inactive or item["active"]:
            normalized.append(item)
    return normalized


def get_pricing_snapshot():
    """Return a JSON-safe full pricing schedule for freezing into a proposal."""
    return {item["key"]: item for item in get_pricing_settings(include_inactive=True)}


def get_default_pricing_snapshot():
    """Legacy schedule used for proposals created before database-backed pricing existed."""
    return {item["key"]: dict(item) for item in DEFAULT_PRICING_SETTINGS}


def update_pricing_setting(setting_key, new_value, changed_by, is_repeating=None, active=None):
    settings = {x["key"]: x for x in get_pricing_settings(include_inactive=True)}
    current = settings.get(setting_key)
    if not current:
        raise KeyError(f"Unknown pricing setting: {setting_key}")

    old_value = float(current["value"])
    old_repeat = bool(current.get("is_repeating", False))
    new_repeat = old_repeat if is_repeating is None else bool(is_repeating)
    new_active = bool(current.get("active", True)) if active is None else bool(active)
    new_value = float(new_value)
    now = _now()

    changed = old_value != new_value or old_repeat != new_repeat or bool(current.get("active", True)) != new_active
    if not changed:
        return False

    if is_cloud_mode():
        client = _cloud_client()
        client.table("pricing_history").insert({
            "setting_key": setting_key,
            "old_value": old_value,
            "new_value": new_value,
            "old_is_repeating": old_repeat,
            "new_is_repeating": new_repeat,
            "old_active": bool(current.get("active", True)),
            "new_active": new_active,
            "changed_at": now,
            "changed_by": changed_by,
        }).execute()
        client.table("pricing_settings").update({
            "value": new_value,
            "is_repeating": new_repeat,
            "active": new_active,
            "updated_at": now,
            "updated_by": changed_by,
        }).eq("key", setting_key).execute()
        return True

    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO pricing_history
        (setting_key, old_value, new_value, old_is_repeating, new_is_repeating, old_active, new_active, changed_at, changed_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (setting_key, old_value, new_value, int(old_repeat), int(new_repeat),
          int(bool(current.get("active", True))), int(new_active), now, changed_by))
    cursor.execute("""
        UPDATE pricing_settings SET value = ?, is_repeating = ?, active = ?, updated_at = ?, updated_by = ?
        WHERE key = ?
    """, (new_value, int(new_repeat), int(new_active), now, changed_by, setting_key))
    conn.commit(); conn.close()
    return True


def add_fixed_cost(label, value, changed_by, is_repeating=False):
    base = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_") or "custom"
    key = f"fixed_{base}"
    existing = {x["key"] for x in get_pricing_settings(include_inactive=True)}
    suffix = 2
    candidate = key
    while candidate in existing:
        candidate = f"{key}_{suffix}"
        suffix += 1
    key = candidate
    now = _now()
    sort_order = 1000 + len([x for x in get_pricing_settings() if x["category"] == "Fixed Costs"])
    row = {
        "key": key, "category": "Fixed Costs", "label": label.strip(), "value": float(value),
        "value_type": "currency", "unit": "fixed", "description": "Admin-added fixed cost",
        "sort_order": sort_order, "is_repeating": bool(is_repeating), "active": True,
        "updated_at": now, "updated_by": changed_by,
    }
    if is_cloud_mode():
        _cloud_client().table("pricing_settings").insert(row).execute()
    else:
        conn = get_connection(); cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO pricing_settings
            (key, category, label, value, value_type, unit, description, sort_order, is_repeating, active, updated_at, updated_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (row["key"], row["category"], row["label"], row["value"], row["value_type"], row["unit"],
              row["description"], row["sort_order"], int(row["is_repeating"]), 1, now, changed_by))
        conn.commit(); conn.close()
    return key


def get_pricing_history(limit=100):
    if is_cloud_mode():
        return (_cloud_client().table("pricing_history").select("*")
                .order("changed_at", desc=True).order("id", desc=True).limit(int(limit)).execute().data or [])
    conn = get_connection(); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
    cursor.execute("SELECT * FROM pricing_history ORDER BY changed_at DESC, id DESC LIMIT ?", (int(limit),))
    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows


def save_pricing_snapshot(proposal_id, pricing_data, generated_by):
    if not proposal_id:
        return None
    now = _now()
    if is_cloud_mode():
        response = _cloud_client().table("proposal_pricing_snapshots").insert({
            "proposal_id": int(proposal_id), "generated_at": now, "generated_by": generated_by,
            "pricing_json": pricing_data,
        }).execute()
        return response.data[0].get("id") if response.data else None
    conn = get_connection(); cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO proposal_pricing_snapshots (proposal_id, generated_at, generated_by, pricing_json)
        VALUES (?, ?, ?, ?)
    """, (proposal_id, now, generated_by, json.dumps(pricing_data)))
    snapshot_id = cursor.lastrowid
    conn.commit(); conn.close()
    return snapshot_id


def get_backend_status():
    return {
        "data_mode": "Supabase Cloud" if is_cloud_mode() else "Local SQLite",
        "database_location": "Supabase PostgreSQL" if is_cloud_mode() else str(DB_NAME),
    }
