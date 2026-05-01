import sqlite3
import json
from datetime import datetime
from pathlib import Path

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
            saved_data_json TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


def save_proposal(proposal_id, proposal_name, credit_union, proposal_type, status, saved_data, msr=""):
    conn = get_connection()
    cursor = conn.cursor()

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    saved_json = json.dumps(saved_data)

    if proposal_id:
        cursor.execute("""
            UPDATE proposals
            SET proposal_name = ?,
                credit_union = ?,
                proposal_type = ?,
                msr = ?,
                status = ?,
                updated_at = ?,
                saved_data_json = ?
            WHERE id = ?
        """, (
            proposal_name,
            credit_union,
            proposal_type,
            msr,
            status,
            now,
            saved_json,
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
                saved_data_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            proposal_name,
            credit_union,
            proposal_type,
            msr,
            status,
            now,
            now,
            saved_json
        ))

        proposal_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return proposal_id


def search_proposals(search_text="", status_filter="All", msr_filter="All"):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT id, proposal_name, credit_union, proposal_type, status, updated_at, msr
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
