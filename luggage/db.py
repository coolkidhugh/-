"""SQLite：照片 + 卡联标注 + 位置 + 备注。"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Any, Iterator
from uuid import uuid4

from luggage.config import DB_PATH, DATA_DIR, PHOTOS_DIR


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PHOTOS_DIR.mkdir(parents=True, exist_ok=True)


@contextmanager
def connect() -> Iterator[sqlite3.Connection]:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    ensure_dirs()
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS bags (
                id TEXT PRIMARY KEY,
                card_tag TEXT NOT NULL,
                location TEXT NOT NULL,
                note TEXT,
                bag_color TEXT,
                photo_path TEXT NOT NULL,
                phash TEXT,
                dhash TEXT,
                colorhash TEXT,
                status TEXT NOT NULL DEFAULT 'stored',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                retrieved_at TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_bags_status ON bags(status);
            CREATE INDEX IF NOT EXISTS idx_bags_card ON bags(card_tag);
            CREATE INDEX IF NOT EXISTS idx_bags_location ON bags(location);
            """
        )


def new_id() -> str:
    return uuid4().hex


def insert_bag(fields: dict[str, Any]) -> dict[str, Any]:
    init_db()
    bag_id = fields.get("id") or new_id()
    now = _utc_now()
    row = {
        "id": bag_id,
        "card_tag": (fields.get("card_tag") or "").strip() or f"未编号-{bag_id[:6].upper()}",
        "location": fields["location"].strip(),
        "note": (fields.get("note") or "").strip(),
        "bag_color": (fields.get("bag_color") or "").strip(),
        "photo_path": fields["photo_path"],
        "phash": fields.get("phash") or "",
        "dhash": fields.get("dhash") or "",
        "colorhash": fields.get("colorhash") or "",
        "status": "stored",
        "created_at": now,
        "updated_at": now,
        "retrieved_at": None,
    }
    with connect() as conn:
        conn.execute(
            """
            INSERT INTO bags (
                id, card_tag, location, note, bag_color, photo_path,
                phash, dhash, colorhash, status, created_at, updated_at, retrieved_at
            ) VALUES (
                :id, :card_tag, :location, :note, :bag_color, :photo_path,
                :phash, :dhash, :colorhash, :status, :created_at, :updated_at, :retrieved_at
            )
            """,
            row,
        )
    return get_by_id(bag_id)  # type: ignore[return-value]


def get_by_id(bag_id: str) -> dict[str, Any] | None:
    init_db()
    with connect() as conn:
        row = conn.execute("SELECT * FROM bags WHERE id = ?", (bag_id,)).fetchone()
    return dict(row) if row else None


def get_by_card_tag(card_tag: str) -> list[dict[str, Any]]:
    init_db()
    tag = card_tag.strip()
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM bags
            WHERE card_tag LIKE ? AND status = 'stored'
            ORDER BY datetime(created_at) DESC
            """,
            (f"%{tag}%",),
        ).fetchall()
    return [dict(r) for r in rows]


def list_bags(
    status: str | None = "stored",
    card_tag: str | None = None,
    limit: int = 300,
) -> list[dict[str, Any]]:
    init_db()
    clauses: list[str] = []
    params: list[Any] = []
    if status:
        clauses.append("status = ?")
        params.append(status)
    if card_tag:
        clauses.append("card_tag LIKE ?")
        params.append(f"%{card_tag.strip()}%")
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(limit)
    with connect() as conn:
        rows = conn.execute(
            f"""
            SELECT * FROM bags
            {where}
            ORDER BY datetime(created_at) DESC
            LIMIT ?
            """,
            params,
        ).fetchall()
    return [dict(r) for r in rows]


def list_stored_with_hashes() -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM bags
            WHERE status = 'stored'
              AND photo_path IS NOT NULL
              AND photo_path != ''
            ORDER BY datetime(created_at) DESC
            """
        ).fetchall()
    return [dict(r) for r in rows]


def mark_retrieved(bag_id: str) -> dict[str, Any] | None:
    init_db()
    now = _utc_now()
    with connect() as conn:
        conn.execute(
            """
            UPDATE bags
            SET status = 'retrieved', updated_at = ?, retrieved_at = ?
            WHERE id = ? AND status = 'stored'
            """,
            (now, now, bag_id),
        )
    return get_by_id(bag_id)


def update_note_or_location(
    bag_id: str,
    *,
    note: str | None = None,
    location: str | None = None,
    card_tag: str | None = None,
) -> dict[str, Any] | None:
    init_db()
    bag = get_by_id(bag_id)
    if not bag:
        return None
    new_note = bag["note"] if note is None else note.strip()
    new_loc = bag["location"] if location is None else location.strip()
    new_tag = bag["card_tag"] if card_tag is None else card_tag.strip()
    with connect() as conn:
        conn.execute(
            """
            UPDATE bags
            SET note = ?, location = ?, card_tag = ?, updated_at = ?
            WHERE id = ?
            """,
            (new_note, new_loc, new_tag, _utc_now(), bag_id),
        )
    return get_by_id(bag_id)


def count_by_status() -> dict[str, int]:
    init_db()
    with connect() as conn:
        rows = conn.execute(
            "SELECT status, COUNT(*) AS c FROM bags GROUP BY status"
        ).fetchall()
    return {r["status"]: r["c"] for r in rows}
