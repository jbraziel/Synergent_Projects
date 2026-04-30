import sqlite3
import json
from datetime import datetime

DB_NAME = "proposals.db"


def get_connection():
    return sqlite3.connect(DB_NAME)


def initialize_database():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proposals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            proposal_name TEXT NOT NULL,
            credit_union TEXT NOT NULL,
            proposal_type TEXT NOT NULL,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            saved_data_json TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


def save_proposal(proposal_id, proposal_name, credit_union, proposal_type, status, saved_data):
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
                status = ?,
                updated_at = ?,
                saved_data_json = ?
            WHERE id = ?
        """, (
            proposal_name,
            credit_union,
            proposal_type,
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
                status,
                created_at,
                updated_at,
                saved_data_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            proposal_name,
            credit_union,
            proposal_type,
            status,
            now,
            now,
            saved_json
        ))

        proposal_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return proposal_id


def search_proposals(search_text="", status_filter="All"):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT id, proposal_name, credit_union, proposal_type, status, updated_at
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
            )
        """
        search_value = f"%{search_text.strip()}%"
        params.extend([search_value, search_value, search_value])

    if status_filter != "All":
        query += " AND status = ?"
        params.append(status_filter)

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

    return json.loads(row[0])
