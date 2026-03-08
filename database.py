"""
SQLite база данных: пользователи и доступы
"""
import sqlite3
import os
from datetime import datetime
from typing import Optional

DB_PATH = os.getenv("DB_PATH", "evk_bot.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                telegram_id   INTEGER PRIMARY KEY,
                username      TEXT,
                full_name     TEXT,
                status        TEXT DEFAULT 'pending',
                requested_at  TEXT,
                approved_at   TEXT,
                note          TEXT
            );

            CREATE TABLE IF NOT EXISTS user_projects (
                telegram_id  INTEGER,
                project_key  TEXT,
                PRIMARY KEY (telegram_id, project_key)
            );
        """)


# --- Пользователи ---

def get_user(telegram_id: int) -> Optional[dict]:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        return dict(row) if row else None


def create_user(telegram_id: int, username: str, full_name: str, note: str = ""):
    with get_conn() as conn:
        conn.execute(
            """INSERT OR REPLACE INTO users
               (telegram_id, username, full_name, status, requested_at, note)
               VALUES (?, ?, ?, 'pending', ?, ?)""",
            (telegram_id, username or "", full_name or "", datetime.now().isoformat(), note)
        )


def approve_user(telegram_id: int, projects: list):
    with get_conn() as conn:
        conn.execute(
            "UPDATE users SET status='approved', approved_at=? WHERE telegram_id=?",
            (datetime.now().isoformat(), telegram_id)
        )
        conn.execute("DELETE FROM user_projects WHERE telegram_id=?", (telegram_id,))
        for p in projects:
            conn.execute(
                "INSERT INTO user_projects (telegram_id, project_key) VALUES (?,?)",
                (telegram_id, p)
            )


def reject_user(telegram_id: int):
    with get_conn() as conn:
        conn.execute(
            "UPDATE users SET status='rejected' WHERE telegram_id=?", (telegram_id,)
        )


def get_user_projects(telegram_id: int) -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT project_key FROM user_projects WHERE telegram_id=?", (telegram_id,)
        ).fetchall()
        return [r["project_key"] for r in rows]


def get_approved_users() -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM users WHERE status='approved'"
        ).fetchall()
        return [dict(r) for r in rows]


def get_pending_users() -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM users WHERE status='pending'"
        ).fetchall()
        return [dict(r) for r in rows]


def is_approved(telegram_id: int) -> bool:
    user = get_user(telegram_id)
    return user is not None and user["status"] == "approved"


def is_owner(telegram_id: int) -> bool:
    from config import OWNER_ID
    return telegram_id == OWNER_ID
