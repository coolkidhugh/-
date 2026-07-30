"""寄存业务编排：存入 / 查找 / 取出。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from luggage.config import PHOTOS_DIR
from luggage import db
from luggage.photo_search import compute_hashes, save_photo_bytes, search_by_photo
from luggage.qrutil import make_qr_image, parse_ticket_payload


def deposit(
    *,
    guest_name: str,
    room_no: str,
    location: str,
    photo_bytes: bytes,
    phone: str = "",
    piece_count: int = 1,
    bag_type: str = "",
    bag_color: str = "",
    brand_note: str = "",
    note: str = "",
) -> dict[str, Any]:
    if not guest_name.strip():
        raise ValueError("请填写客人姓名")
    if not room_no.strip():
        raise ValueError("请填写房号")
    if not location.strip():
        raise ValueError("请选择存放位置")
    if not photo_bytes:
        raise ValueError("请拍摄或上传行李照片")

    db.init_db()
    ticket = db.new_ticket_code()
    photo_path = save_photo_bytes(photo_bytes, ticket, PHOTOS_DIR)
    hashes = compute_hashes(photo_path)
    _, qr_path = make_qr_image(ticket, save=True)

    return db.insert_item(
        {
            "ticket_code": ticket,
            "guest_name": guest_name,
            "room_no": room_no,
            "phone": phone,
            "piece_count": piece_count,
            "bag_type": bag_type,
            "bag_color": bag_color,
            "brand_note": brand_note,
            "location": location,
            "note": note,
            "photo_path": str(photo_path),
            "qr_path": str(qr_path) if qr_path else "",
            "phash": hashes.phash,
            "dhash": hashes.dhash,
            "colorhash": hashes.colorhash,
        }
    )


def find_by_ticket_or_scan(raw: str) -> dict[str, Any] | None:
    code = parse_ticket_payload(raw)
    if not code:
        return None
    return db.get_by_ticket(code)


def find_by_photo(
    photo_bytes: bytes,
    *,
    bag_color: str | None = None,
    top_k: int = 8,
) -> list[dict[str, Any]]:
    return search_by_photo(photo_bytes, bag_color=bag_color, top_k=top_k)


def retrieve(item_id: str) -> dict[str, Any] | None:
    item = db.get_by_id(item_id)
    if not item:
        return None
    if item["status"] != "stored":
        return item
    return db.mark_retrieved(item_id)


def photo_exists(item: dict[str, Any]) -> bool:
    path = item.get("photo_path") or ""
    return bool(path) and Path(path).is_file()


def qr_exists(item: dict[str, Any]) -> bool:
    path = item.get("qr_path") or ""
    return bool(path) and Path(path).is_file()
