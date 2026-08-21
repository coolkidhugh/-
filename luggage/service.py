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
ARCHIVE_DIR = RECORDS_DIR / "_archive"


def _normalize_card_tag(card_tag: str) -> str:
    """保留用户输入的数字编号（9619 / 0056469），不强制补零。"""
    tag = (card_tag or "").strip()
    digits = "".join(ch for ch in tag if ch.isdigit())
    if digits and digits == tag.replace(" ", ""):
        return digits
    return tag or digits


def _batch_dir(batch: str) -> Path:
    return RECORDS_DIR / batch if batch else RECORDS_DIR


def _record_folder(card_tag: str, batch: str = "") -> Path:
    tag = _normalize_card_tag(card_tag)
    if batch:
        return _batch_dir(batch) / tag
    return RECORDS_DIR / tag


def _write_record_meta(folder: Path, bag: dict[str, Any], *, has_real_photo: bool) -> None:
    meta = {
        "id": bag.get("id"),
        "card_tag": bag.get("card_tag"),
        "batch": bag.get("batch") or "",
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
    batch_line = bag.get("batch") or "（无批次）"
    (folder / "README.md").write_text(
        f"""# 卡联 {bag.get('card_tag')}

- 批次：{batch_line}
- 位置：{bag.get('location')}
- 备注：{bag.get('note') or '（无）'}
- 实体图：{photo_line}
""",
        encoding="utf-8",
    )


def _persist_record_photo(card_tag: str, photo_bytes: bytes, bag: dict[str, Any]) -> Path:
    """实体图落到仓库目录，查编号时可直接带出。"""
    batch = bag.get("batch") or ""
    folder = _record_folder(card_tag, batch)
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
    batch: str = "",
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
            batch=batch,
        )

    db.init_db()
    tag = _normalize_card_tag(card_tag)
    # 同编号+同批次已在库则更新位置/备注
    existing = [
        b
        for b in db.get_by_card_tag(tag)
        if (not batch) or (b.get("batch") or "") == batch
    ]
    if existing:
        bag = existing[0]
        updated = db.update_note_or_location(
            bag["id"],
            note=note if note else None,
            location=location,
            card_tag=tag,
        )
        if batch:
            updated = db.set_batch(bag["id"], batch) or updated
        folder = _record_folder(tag, batch or (updated or bag).get("batch") or "")
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
            "batch": batch,
        }
    )
    folder = _record_folder(tag, batch)
    folder.mkdir(parents=True, exist_ok=True)
    _write_record_meta(folder, bag, has_real_photo=False)
    return bag


def deposit(
    *,
    photo_bytes: bytes,
    location: str,
    card_tag: str = "",
    note: str = "",
    bag_color: str = "",
    batch: str = "",
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
            "batch": batch,
        }
    )
    archived = _persist_record_photo(tag, photo_bytes, bag)
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


def get_photo_path(card_tag: str, batch: str | None = None) -> Path | None:
    """查编号 → 实体图路径。仅返回真正的 photo.jpg（不含占位图）。"""
    tag = _normalize_card_tag(card_tag)
    rows = find_by_card(card_tag)
    if batch:
        rows = [r for r in rows if (r.get("batch") or "") == batch]
    candidates: list[Path] = []
    if rows:
        b = rows[0]
        candidates.append(_record_folder(b.get("card_tag") or tag, b.get("batch") or "") / "photo.jpg")
        path = Path(b.get("photo_path") or "")
        if path.name == "photo.jpg":
            candidates.append(path)
    # 也扫当前批次目录与根目录
    from luggage.layout import CURRENT_BATCH

    for bname in filter(None, [batch, CURRENT_BATCH, ""]):
        candidates.append(_record_folder(tag, bname) / "photo.jpg")
    seen: set[str] = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        if path.is_file():
            return path
    return None


def has_real_photo(card_tag: str, batch: str | None = None) -> bool:
    return get_photo_path(card_tag, batch=batch) is not None


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
