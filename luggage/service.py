"""业务：拍照存档 / 按编号取实体图 / 拍照查找。"""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any

from luggage.config import PHOTOS_DIR, ROOT
from luggage import db
from luggage.photo_search import compute_hashes, save_photo_bytes, search_by_photo

RECORDS_DIR = ROOT / "luggage_records"


def _normalize_card_tag(card_tag: str) -> str:
    tag = (card_tag or "").strip()
    digits = "".join(ch for ch in tag if ch.isdigit())
    # 纯数字卡联号：统一存成至少 7 位（0056469）
    if digits and digits == tag.replace(" ", ""):
        return digits.zfill(7) if len(digits) <= 7 else digits
    return tag or digits


def _persist_record_photo(card_tag: str, photo_bytes: bytes, bag: dict[str, Any]) -> Path:
    """实体图落到仓库目录，查编号时可直接带出。"""
    folder = RECORDS_DIR / _normalize_card_tag(card_tag)
    folder.mkdir(parents=True, exist_ok=True)
    photo_path = folder / "photo.jpg"
    photo_path.write_bytes(photo_bytes)
    meta = {
        "id": bag.get("id"),
        "card_tag": bag.get("card_tag"),
        "location": bag.get("location"),
        "note": bag.get("note"),
        "bag_color": bag.get("bag_color"),
        "status": bag.get("status"),
        "created_at": bag.get("created_at"),
        "photo": "photo.jpg",
        "has_real_photo": True,
    }
    (folder / "record.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (folder / "README.md").write_text(
        f"""# 卡联 {bag.get('card_tag')}

- 位置：{bag.get('location')}
- 备注：{bag.get('note') or '（无）'}
- 实体图：`photo.jpg`

查编号时必须带出 `photo.jpg`。
""",
        encoding="utf-8",
    )
    return photo_path


def deposit(
    *,
    photo_bytes: bytes,
    location: str,
    card_tag: str = "",
    note: str = "",
    bag_color: str = "",
) -> dict[str, Any]:
    if not photo_bytes:
        raise ValueError("必须上传实体照片（现场图），不能只有文字标注")
    if not location.strip():
        raise ValueError("请填写存放位置")

    db.init_db()
    bag_id = db.new_id()
    tag = _normalize_card_tag(card_tag)
    photo_path = save_photo_bytes(photo_bytes, bag_id, PHOTOS_DIR)
    hashes = compute_hashes(photo_path)
    color = "" if bag_color in ("", "不标注") else bag_color

    bag = db.insert_bag(
        {
            "id": bag_id,
            "card_tag": tag,
            "location": location,
            "note": note,
            "bag_color": color,
            "photo_path": str(photo_path),
            "phash": hashes.phash,
            "dhash": hashes.dhash,
            "colorhash": hashes.colorhash,
        }
    )
    # 同步一份实体图到 luggage_records/<编号>/photo.jpg
    archived = _persist_record_photo(tag, photo_bytes, bag)
    # 展示路径优先用归档实体图（持久、可进仓库）
    updated = db.update_photo(
        bag_id,
        photo_path=str(archived),
        phash=hashes.phash,
        dhash=hashes.dhash,
        colorhash=hashes.colorhash,
    )
    return updated or bag


def replace_photo(card_tag: str, photo_bytes: bytes) -> dict[str, Any]:
    """给已有编号补上/替换实体图。"""
    if not photo_bytes:
        raise ValueError("请提供实体照片")
    rows = find_by_card(card_tag)
    if not rows:
        raise ValueError(f"找不到卡联 {card_tag}，请先存档或核对编号")
    bag = rows[0]
    tag = bag["card_tag"]
    hashes = compute_hashes(photo_bytes)
    archived = _persist_record_photo(tag, photo_bytes, bag)
    # 也写一份到 runtime photos
    runtime = save_photo_bytes(photo_bytes, bag["id"], PHOTOS_DIR)
    updated = db.update_photo(
        bag["id"],
        photo_path=str(archived),
        phash=hashes.phash,
        dhash=hashes.dhash,
        colorhash=hashes.colorhash,
    )
    if runtime and runtime != Path(archived):
        # runtime 备份即可
        pass
    return updated or bag


def find_by_photo(
    photo_bytes: bytes,
    *,
    bag_color: str | None = None,
    top_k: int = 6,
) -> list[dict[str, Any]]:
    return search_by_photo(photo_bytes, bag_color=bag_color, top_k=top_k)


def find_by_card(card_tag: str) -> list[dict[str, Any]]:
    return db.get_by_card_tag(card_tag)


def get_photo_path(card_tag: str) -> Path | None:
    """查编号 → 实体图路径。"""
    rows = find_by_card(card_tag)
    if not rows:
        # 直接看归档目录
        folder = RECORDS_DIR / _normalize_card_tag(card_tag)
        candidate = folder / "photo.jpg"
        return candidate if candidate.is_file() else None
    path = Path(rows[0].get("photo_path") or "")
    if path.is_file():
        return path
    folder = RECORDS_DIR / _normalize_card_tag(rows[0].get("card_tag") or card_tag)
    candidate = folder / "photo.jpg"
    return candidate if candidate.is_file() else None


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
