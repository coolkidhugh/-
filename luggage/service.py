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
    """保留用户输入的数字编号（9619 / 0056469），不强制补零。"""
    tag = (card_tag or "").strip()
    digits = "".join(ch for ch in tag if ch.isdigit())
    if digits and digits == tag.replace(" ", ""):
        return digits
    return tag or digits


def _write_record_meta(folder: Path, bag: dict[str, Any], *, has_real_photo: bool) -> None:
    meta = {
        "id": bag.get("id"),
        "card_tag": bag.get("card_tag"),
        "location": bag.get("location"),
        "note": bag.get("note"),
        "bag_color": bag.get("bag_color"),
        "status": bag.get("status"),
        "created_at": bag.get("created_at"),
        "photo": "photo.jpg" if has_real_photo else None,
        "has_real_photo": has_real_photo,
        "needs_real_photo": not has_real_photo,
    }
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "record.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    photo_line = "`photo.jpg`" if has_real_photo else "（待补实体图）"
    (folder / "README.md").write_text(
        f"""# 卡联 {bag.get('card_tag')}

- 位置：{bag.get('location')}
- 备注：{bag.get('note') or '（无）'}
- 实体图：{photo_line}
""",
        encoding="utf-8",
    )


def _persist_record_photo(card_tag: str, photo_bytes: bytes, bag: dict[str, Any]) -> Path:
    """实体图落到仓库目录，查编号时可直接带出。"""
    folder = RECORDS_DIR / _normalize_card_tag(card_tag)
    folder.mkdir(parents=True, exist_ok=True)
    photo_path = folder / "photo.jpg"
    photo_path.write_bytes(photo_bytes)
    _write_record_meta(folder, bag, has_real_photo=True)
    return photo_path


def _pending_placeholder(card_tag: str, location: str) -> bytes:
    """无实体图时的占位图（查编号能看到待补状态，不冒充现场图）。"""
    from PIL import Image, ImageDraw, ImageFont

    img = Image.new("RGB", (800, 1000), (236, 232, 224))
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 40
        )
        font_s = ImageFont.truetype(
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 28
        )
    except Exception:
        font = font_s = ImageFont.load_default()
    d.rectangle([40, 40, 760, 960], outline=(80, 80, 80), width=3)
    d.text((80, 120), "待补实体图", fill=(160, 40, 40), font=font)
    d.text((80, 220), f"卡联 {card_tag}", fill=(30, 30, 30), font=font)
    d.text((80, 300), f"位置 {location}", fill=(30, 30, 30), font=font_s)
    d.text((80, 380), "补传现场照片后即可按编号出真图", fill=(80, 80, 80), font=font_s)
    import io

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    return buf.getvalue()


def register_slot(
    *,
    card_tag: str,
    location: str,
    note: str = "",
    bag_color: str = "",
    photo_bytes: bytes | None = None,
) -> dict[str, Any]:
    """按货架位登记编号。有实体图就存真图，没有则占位待补。"""
    if not card_tag.strip():
        raise ValueError("请填写卡联号")
    if not location.strip():
        raise ValueError("请填写存放位置")

    if photo_bytes:
        return deposit(
            photo_bytes=photo_bytes,
            location=location,
            card_tag=card_tag,
            note=note,
            bag_color=bag_color,
        )

    db.init_db()
    tag = _normalize_card_tag(card_tag)
    # 同编号已在库则更新位置/备注，不重复插
    existing = db.get_by_card_tag(tag)
    if existing:
        bag = existing[0]
        updated = db.update_note_or_location(
            bag["id"],
            note=note if note else None,
            location=location,
            card_tag=tag,
        )
        folder = RECORDS_DIR / tag
        _write_record_meta(folder, updated or bag, has_real_photo=False)
        return updated or bag

    bag_id = db.new_id()
    placeholder = _pending_placeholder(tag, location)
    photo_path = save_photo_bytes(placeholder, bag_id, PHOTOS_DIR)
    color = "" if bag_color in ("", "不标注") else bag_color
    bag = db.insert_bag(
        {
            "id": bag_id,
            "card_tag": tag,
            "location": location,
            "note": (note or "待补实体图").strip(),
            "bag_color": color,
            "photo_path": str(photo_path),
            "phash": "",
            "dhash": "",
            "colorhash": "",
        }
    )
    folder = RECORDS_DIR / tag
    folder.mkdir(parents=True, exist_ok=True)
    # 占位图不叫 photo.jpg，避免冒充实体图
    pending_path = folder / "pending.jpg"
    pending_path.write_bytes(placeholder)
    _write_record_meta(folder, bag, has_real_photo=False)
    # DB 仍指向 runtime 占位，查图时 get_photo_path 会区分
    return bag


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
    """查编号 → 实体图路径。仅返回真正的 photo.jpg（不含占位图）。"""
    tag = _normalize_card_tag(card_tag)
    folder = RECORDS_DIR / tag
    real = folder / "photo.jpg"
    if real.is_file():
        return real
    rows = find_by_card(card_tag)
    if rows:
        path = Path(rows[0].get("photo_path") or "")
        # runtime 占位图不算实体图
        meta_path = (RECORDS_DIR / _normalize_card_tag(rows[0].get("card_tag") or tag) / "record.json")
        if meta_path.is_file():
            try:
                meta = json.loads(meta_path.read_text(encoding="utf-8"))
                if meta.get("has_real_photo") and path.is_file():
                    return path
            except Exception:
                pass
        if path.is_file() and path.name == "photo.jpg":
            return path
    return None


def has_real_photo(card_tag: str) -> bool:
    return get_photo_path(card_tag) is not None


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
