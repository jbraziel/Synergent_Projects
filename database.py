import sqlite3
import json
from datetime import datetime
from pathlib import Path
import time


DB_NAME = Path(__file__).parent / "proposals.db"


def get_connection():
    conn = sqlite3.connect(DB_NAME, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA busy_timeout = 30000;")
    return conn


def initialize_database():
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
            locked_at TEXT
        )
    """)
    add_column_if_missing(cursor, "proposals", "updated_by", "TEXT")
    add_column_if_missing(cursor, "proposals", "locked_by", "TEXT")
    add_column_if_missing(cursor, "proposals", "locked_at", "TEXT")
    conn.commit()
    conn.close()

def add_column_if_missing(cursor, table, column, definition):
    cursor.execute(f"PRAGMA table_info({table})")
    columns = [row[1] for row in cursor.fetchall()]

    if column not in columns:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def save_proposal(proposal_id, proposal_name, credit_union, proposal_type, status, saved_data, msr="", updated_by=""):
    import time
    import sqlite3

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    saved_json = json.dumps(saved_data)

    for attempt in range(5):
        conn = get_connection()
        cursor = conn.cursor()

        try:
            if proposal_id:
                cursor.execute("""
                    UPDATE proposals
                    SET proposal_name = ?,
                        credit_union = ?,
                        proposal_type = ?,
                        msr = ?,
                        status = ?,
                        updated_at = ?,
                        saved_data_json = ?,
                        updated_by = ?
                    WHERE id = ?
                """, (
                    proposal_name,
                    credit_union,
                    proposal_type,
                    msr,
                    status,
                    now,
                    saved_json,
                    updated_by,
                    proposal_id
                ))
            else:
                cursor.execute("""
                    INSERT INTO proposals (
                        proposal_name,
                        credit_union,
                        proposal_type,
                        msr,
                        status,
                        created_at,
                        updated_at,
                        saved_data_json,
                        updated_by
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    proposal_name,
                    credit_union,
                    proposal_type,
                    msr,
                    status,
                    now,
                    now,
                    saved_json,
                    updated_by
                ))

                proposal_id = cursor.lastrowid

            conn.commit()
            conn.close()
            return proposal_id

        except sqlite3.OperationalError as e:
            conn.close()

            if "locked" in str(e).lower() and attempt < 4:
                time.sleep(0.5)
                continue

            raise


def search_proposals(search_text="", status_filter="All", msr_filter="All"):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT id, proposal_name, credit_union, proposal_type, status, updated_at, msr, updated_by, locked_by, locked_at
        FROM proposals
        WHERE 1=1
    """

    params = []

    if search_text.strip():
        query += """
            AND (
                proposal_name LIKE ?
                OR credit_union LIKE ?
                OR proposal_type LIKE ?
                OR msr LIKE ?
            )
        """
        search_value = f"%{search_text.strip()}%"
        params.extend([search_value, search_value, search_value, search_value])

    if status_filter != "All":
        query += " AND status = ?"
        params.append(status_filter)

    if msr_filter != "All":
        query += " AND msr = ?"
        params.append(msr_filter)

    query += " ORDER BY updated_at DESC"

    cursor.execute(query, params)
    rows = cursor.fetchall()

    conn.close()
    return rows


def load_proposal(proposal_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT saved_data_json
        FROM proposals
        WHERE id = ?
    """, (proposal_id,))

    row = cursor.fetchone()
    conn.close()

    if not row:
        return None

    try:
        return json.loads(row[0])
    except json.JSONDecodeError:
        return None


def delete_proposal(proposal_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        DELETE FROM proposals
        WHERE id = ?
    """, (proposal_id,))

    conn.commit()
    conn.close()

def update_proposal_status(proposal_id, status):
    conn = get_connection()
    cursor = conn.cursor()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        UPDATE proposals
        SET status = ?,
            updated_at = ?
        WHERE id = ?
    """, (status, now, proposal_id))

    conn.commit()
    conn.close()

def lock_proposal(proposal_id, user_name):
    conn = get_connection()
    cursor = conn.cursor()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        UPDATE proposals
        SET locked_by = ?,
            locked_at = ?
        WHERE id = ?
    """, (user_name, now, proposal_id))

    conn.commit()
    conn.close()


def unlock_proposal(proposal_id, user_name):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE proposals
        SET locked_by = NULL,
            locked_at = NULL
        WHERE id = ?
          AND locked_by = ?
    """, (proposal_id, user_name))

    conn.commit()
    conn.close()
