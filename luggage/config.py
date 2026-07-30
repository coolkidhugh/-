"""路径与寄存区配置。"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" / "luggage"
DB_PATH = DATA_DIR / "luggage.db"
PHOTOS_DIR = DATA_DIR / "photos"
QR_DIR = DATA_DIR / "qrcodes"

# 与现场行李间货架对应的存放区（可后续按平面图微调）
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

BAG_TYPES = ["登机箱", "托运箱", "双肩包", "手提包", "其他"]
BAG_COLORS = ["黑", "灰", "蓝", "红", "粉", "绿", "棕", "花色/图案", "其他"]

# 拍照找行李：返回前 N 个候选
PHOTO_SEARCH_TOP_K = 8
PHOTO_MATCH_MIN_SCORE = 0.35
