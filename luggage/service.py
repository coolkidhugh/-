"""业务：拍照存档 / 拍照查找 / 改备注 / 取出。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from luggage.config import PHOTOS_DIR
from luggage import db
from luggage.photo_search import compute_hashes, save_photo_bytes, search_by_photo


def deposit(
    *,
    photo_bytes: bytes,
    location: str,
    card_tag: str = "",
    note: str = "",
    bag_color: str = "",
) -> dict[str, Any]:
    if not photo_bytes:
        raise ValueError("请先拍照或上传行李照片")
    if not location.strip():
        raise ValueError("请填写存放位置")

    db.init_db()
    bag_id = db.new_id()
    photo_path = save_photo_bytes(photo_bytes, bag_id, PHOTOS_DIR)
    hashes = compute_hashes(photo_path)
    color = "" if bag_color in ("", "不标注") else bag_color

    return db.insert_bag(
        {
            "id": bag_id,
            "card_tag": card_tag,
            "location": location,
            "note": note,
            "bag_color": color,
            "photo_path": str(photo_path),
            "phash": hashes.phash,
            "dhash": hashes.dhash,
            "colorhash": hashes.colorhash,
        }
    )


def find_by_photo(
    photo_bytes: bytes,
    *,
    bag_color: str | None = None,
    top_k: int = 6,
) -> list[dict[str, Any]]:
    return search_by_photo(photo_bytes, bag_color=bag_color, top_k=top_k)


def find_by_card(card_tag: str) -> list[dict[str, Any]]:
    return db.get_by_card_tag(card_tag)


def save_remark(
    bag_id: str,
    *,
    note: str | None = None,
    location: str | None = None,
    card_tag: str | None = None,
) -> dict[str, Any] | None:
    return db.update_note_or_location(
        bag_id, note=note, location=location, card_tag=card_tag
    )


def retrieve(bag_id: str) -> dict[str, Any] | None:
    bag = db.get_by_id(bag_id)
    if not bag:
        return None
    if bag["status"] != "stored":
        return bag
    return db.mark_retrieved(bag_id)


def photo_exists(bag: dict[str, Any]) -> bool:
    path = bag.get("photo_path") or ""
    return bool(path) and Path(path).is_file()
