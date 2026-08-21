"""路径与货架行位。"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "luggage"
DB_PATH = DATA_DIR / "luggage.db"
PHOTOS_DIR = DATA_DIR / "photos"

# 从里往外共 6 行（最里面 = 第1行）
STORAGE_ROWS = [
    "第1行-最里面",
    "第2行",
    "第3行",
    "第4行",
    "第5行",
    "第6行",
]

STORAGE_ZONES = list(STORAGE_ROWS) + [
    "绿地板堆放区",
    "临时区/门口",
]

BAG_COLORS = ["不标注", "黑", "灰", "蓝", "红", "粉", "绿", "棕", "花色/图案", "其他"]

PHOTO_SEARCH_TOP_K = 6
PHOTO_MATCH_MIN_SCORE = 0.30
