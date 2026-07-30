"""路径与存放区。"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "luggage"
DB_PATH = DATA_DIR / "luggage.db"
PHOTOS_DIR = DATA_DIR / "photos"

# 现场货架位置（可按平面图改）
STORAGE_ZONES = [
    "货架A-上",
    "货架A-下",
    "货架B-上",
    "货架B-下",
    "货架C",
    "墙边靠墙区",
    "办公桌旁",
    "临时区/门口",
]

BAG_COLORS = ["不标注", "黑", "灰", "蓝", "红", "粉", "绿", "棕", "花色/图案", "其他"]

PHOTO_SEARCH_TOP_K = 6
PHOTO_MATCH_MIN_SCORE = 0.30
