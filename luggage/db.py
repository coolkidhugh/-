"""SQLite 行李寄存数据层。"""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Any, Iterator
from uuid import uuid4

from luggage.config import DB_PATH, DATA_DIR, PHOTOS_DIR, QR_DIR


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PHOTOS_DIR.mkdir(parents=True, exist_ok=True)
    QR_DIR.mkdir(parents=True, exist_ok=True)


@contextmanager
def connect() -> Iterator[sqlite3.Connection]:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
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
            CREATE TABLE IF NOT EXISTS luggage_items (
                id TEXT PRIMARY KEY,
                ticket_code TEXT NOT NULL UNIQUE,
                guest_name TEXT NOT NULL,
                room_no TEXT NOT NULL,
                phone TEXT,
                piece_count INTEGER NOT NULL DEFAULT 1,
                bag_type TEXT,
                bag_color TEXT,
                brand_note TEXT,
                location TEXT NOT NULL,
                note TEXT,
                photo_path TEXT NOT NULL,
                qr_path TEXT,
                phash TEXT,
                dhash TEXT,
                colorhash TEXT,
                status TEXT NOT NULL DEFAULT 'stored',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                retrieved_at TEXT
            );

            CREATE INDEX IF NOT EXISTS idx_luggage_status ON luggage_items(status);
            CREATE INDEX IF NOT EXISTS idx_luggage_room ON luggage_items(room_no);
            CREATE INDEX IF NOT EXISTS idx_luggage_ticket ON luggage_items(ticket_code);
            CREATE INDEX IF NOT EXISTS idx_luggage_guest ON luggage_items(guest_name);
            """
        )


def new_ticket_code() -> str:
    """短可读票号：L + 日期 + 4位随机。"""
    day = datetime.now().strftime("%m%d")
    return f"L{day}{uuid4().hex[:4].upper()}"


def insert_item(fields: dict[str, Any]) -> dict[str, Any]:
    init_db()
    item_id = fields.get("id") or uuid4().hex
    now = _utc_now()
    ticket = fields.get("ticket_code") or new_ticket_code()
    row = {
        "id": item_id,
        "ticket_code": ticket,
        "guest_name": fields["guest_name"].strip(),
        "room_no": fields["room_no"].strip(),
        "phone": (fields.get("phone") or "").strip(),
        "piece_count": int(fields.get("piece_count") or 1),
        "bag_type": fields.get("bag_type") or "",
        "bag_color": fields.get("bag_color") or "",
        "brand_note": (fields.get("brand_note") or "").strip(),
        "location": fields["location"],
        "note": (fields.get("note") or "").strip(),
        "photo_path": fields["photo_path"],
        "qr_path": fields.get("qr_path") or "",
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
            INSERT INTO luggage_items (
                id, ticket_code, guest_name, room_no, phone, piece_count,
                bag_type, bag_color, brand_note, location, note,
                photo_path, qr_path, phash, dhash, colorhash,
                status, created_at, updated_at, retrieved_at
            ) VALUES (
                :id, :ticket_code, :guest_name, :room_no, :phone, :piece_count,
                :bag_type, :bag_color, :brand_note, :location, :note,
                :photo_path, :qr_path, :phash, :dhash, :colorhash,
                :status, :created_at, :updated_at, :retrieved_at
            )
            """,
            row,
        )
    return get_by_id(item_id)


def _rows_to_dicts(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    return [dict(r) for r in rows]


def get_by_id(item_id: str) -> dict[str, Any] | None:
    init_db()
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM luggage_items WHERE id = ?", (item_id,)
        ).fetchone()
    return dict(row) if row else None


def get_by_ticket(ticket_code: str) -> dict[str, Any] | None:
    init_db()
    code = ticket_code.strip().upper()
    with connect() as conn:
        row = conn.execute(
            "SELECT * FROM luggage_items WHERE upper(ticket_code) = ?", (code,)
        ).fetchone()
    return dict(row) if row else None


def list_items(
    status: str | None = "stored",
    room_no: str | None = None,
    guest_name: str | None = None,
    limit: int = 200,
) -> list[dict[str, Any]]:
    init_db()
    clauses: list[str] = []
    params: list[Any] = []
    if status:
        clauses.append("status = ?")
        params.append(status)
    if room_no:
        clauses.append("room_no LIKE ?")
        params.append(f"%{room_no.strip()}%")
    if guest_name:
        clauses.append("guest_name LIKE ?")
        params.append(f"%{guest_name.strip()}%")
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    sql = f"""
        SELECT * FROM luggage_items
        {where}
        ORDER BY datetime(created_at) DESC
        LIMIT ?
    """
    params.append(limit)
    with connect() as conn:
        rows = conn.execute(sql, params).fetchall()
    return _rows_to_dicts(rows)


def list_stored_with_hashes() -> list[dict[str, Any]]:
    init_db()
    with connect() as conn:
        rows = conn.execute(
            """
            SELECT * FROM luggage_items
            WHERE status = 'stored' AND photo_path IS NOT NULL AND photo_path != ''
            ORDER BY datetime(created_at) DESC
            """
        ).fetchall()
    return _rows_to_dicts(rows)


def mark_retrieved(item_id: str) -> dict[str, Any] | None:
    init_db()
    now = _utc_now()
    with connect() as conn:
        conn.execute(
            """
            UPDATE luggage_items
            SET status = 'retrieved', updated_at = ?, retrieved_at = ?
            WHERE id = ? AND status = 'stored'
            """,
            (now, now, item_id),
        )
    return get_by_id(item_id)


def update_location(item_id: str, location: str) -> dict[str, Any] | None:
    init_db()
    with connect() as conn:
        conn.execute(
            """
            UPDATE luggage_items
            SET location = ?, updated_at = ?
            WHERE id = ?
            """,
            (location, _utc_now(), item_id),
        )
    return get_by_id(item_id)


def count_by_status() -> dict[str, int]:
    init_db()
    with connect() as conn:
        rows = conn.execute(
            "SELECT status, COUNT(*) AS c FROM luggage_items GROUP BY status"
        ).fetchall()
    return {r["status"]: r["c"] for r in rows}
