"""拍照找行李：感知哈希 + 颜色哈希相似度。"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

import imagehash
from PIL import Image

from luggage.config import PHOTO_MATCH_MIN_SCORE, PHOTO_SEARCH_TOP_K
from luggage.db import list_stored_with_hashes


@dataclass
class HashBundle:
    phash: str
    dhash: str
    colorhash: str


def _open_image(source: str | Path | bytes | BinaryIO | Image.Image) -> Image.Image:
    if isinstance(source, Image.Image):
        img = source
    elif isinstance(source, (str, Path)):
        img = Image.open(source)
    elif isinstance(source, bytes):
        img = Image.open(BytesIO(source))
    else:
        img = Image.open(source)
    return img.convert("RGB")


def compute_hashes(source: str | Path | bytes | BinaryIO | Image.Image) -> HashBundle:
    img = _open_image(source)
    return HashBundle(
        phash=str(imagehash.phash(img)),
        dhash=str(imagehash.dhash(img)),
        colorhash=str(imagehash.colorhash(img)),
    )


def _hamming(a: str, b: str) -> int | None:
    if not a or not b:
        return None
    try:
        return imagehash.hex_to_hash(a) - imagehash.hex_to_hash(b)
    except Exception:
        return None


def _color_distance(a: str, b: str) -> int | None:
    if not a or not b:
        return None
    try:
        ha = imagehash.hex_to_flathash(a, hashsize=3)
        hb = imagehash.hex_to_flathash(b, hashsize=3)
        return ha - hb
    except Exception:
        try:
            return abs(imagehash.hex_to_hash(a) - imagehash.hex_to_hash(b))
        except Exception:
            return None


def similarity_score(query: HashBundle, candidate: dict[str, Any]) -> float:
    """
    综合相似度 0~1。
    phash/dhash 权重高（外形轮廓），colorhash 辅助颜色。
    """
    scores: list[tuple[float, float]] = []

    pd = _hamming(query.phash, candidate.get("phash") or "")
    if pd is not None:
        # phash 64bit，距离 0 最好；>22 基本不像
        scores.append((max(0.0, 1.0 - pd / 22.0), 0.40))

    dd = _hamming(query.dhash, candidate.get("dhash") or "")
    if dd is not None:
        scores.append((max(0.0, 1.0 - dd / 22.0), 0.30))

    cd = _color_distance(query.colorhash, candidate.get("colorhash") or "")
    if cd is not None:
        # 行李同外形不同色很常见，颜色权重略高
        scores.append((max(0.0, 1.0 - cd / 28.0), 0.30))

    if not scores:
        return 0.0
    total_w = sum(w for _, w in scores)
    return float(sum(s * w for s, w in scores) / total_w)


def search_by_photo(
    source: str | Path | bytes | BinaryIO | Image.Image,
    *,
    top_k: int = PHOTO_SEARCH_TOP_K,
    min_score: float = PHOTO_MATCH_MIN_SCORE,
    bag_color: str | None = None,
) -> list[dict[str, Any]]:
    query = compute_hashes(source)
    candidates = list_stored_with_hashes()
    if bag_color:
        color = bag_color.strip()
        filtered = [c for c in candidates if (c.get("bag_color") or "") == color]
        if filtered:
            candidates = filtered

    ranked: list[dict[str, Any]] = []
    for item in candidates:
        score = similarity_score(query, item)
        if score < min_score:
            continue
        hit = dict(item)
        hit["match_score"] = round(score, 4)
        ranked.append(hit)

    ranked.sort(key=lambda x: x["match_score"], reverse=True)
    return ranked[:top_k]


def save_photo_bytes(data: bytes, ticket_code: str, photos_dir: Path) -> Path:
    photos_dir.mkdir(parents=True, exist_ok=True)
    img = _open_image(data)
    # 统一存 JPEG，控制体积
    path = photos_dir / f"{ticket_code.upper()}.jpg"
    img.save(path, format="JPEG", quality=85, optimize=True)
    return path
