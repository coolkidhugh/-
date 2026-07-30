"""核心逻辑冒烟测试（无 Streamlit UI）。"""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


def _bag(color: tuple[int, int, int], shape: str, mark: str = "") -> bytes:
    img = Image.new("RGB", (320, 420), (235, 235, 235))
    draw = ImageDraw.Draw(img)
    if shape == "rect":
        draw.rectangle([50, 40, 270, 380], fill=color)
        draw.rectangle([50, 40, 270, 380], outline=(255, 255, 255), width=4)
    else:
        draw.ellipse([40, 80, 280, 360], fill=color)
        draw.ellipse([40, 80, 280, 360], outline=(255, 255, 255), width=4)
    if mark:
        draw.text((70, 190), mark, fill=(255, 255, 255))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()


def main() -> int:
    import luggage.config as cfg
    import luggage.db as db
    import luggage.qrutil as qrutil
    import luggage.service as service

    tmp = Path(tempfile.mkdtemp(prefix="luggage_test_"))
    for mod in (cfg, db, qrutil):
        mod.DATA_DIR = tmp  # type: ignore[attr-defined]
        mod.DB_PATH = tmp / "luggage.db"  # type: ignore[attr-defined]
        mod.PHOTOS_DIR = tmp / "photos"  # type: ignore[attr-defined]
        mod.QR_DIR = tmp / "qrcodes"  # type: ignore[attr-defined]
    service.PHOTOS_DIR = cfg.PHOTOS_DIR

    try:
        black = _bag((20, 20, 20), "rect", "BLK")
        blue = _bag((30, 90, 200), "ellipse", "BLU")
        black_query = _bag((25, 25, 28), "rect", "Q")

        a = service.deposit(
            guest_name="张三",
            room_no="1208",
            location="货架A-上",
            photo_bytes=black,
            bag_color="黑",
            bag_type="登机箱",
        )
        b = service.deposit(
            guest_name="李四",
            room_no="1501",
            location="货架B-下",
            photo_bytes=blue,
            bag_color="蓝",
            bag_type="托运箱",
        )
        assert a["ticket_code"] and Path(a["qr_path"]).is_file()
        assert b["status"] == "stored"

        found = service.find_by_ticket_or_scan(f"LUGGAGE:{a['ticket_code']}")
        assert found and found["id"] == a["id"]

        hits = service.find_by_photo(black_query, top_k=5)
        assert hits, "拍照找行李应至少返回 1 个候选"
        assert hits[0]["ticket_code"] == a["ticket_code"], (
            f"最相似应是黑箱，实际 {hits[0]['ticket_code']} "
            f"score={hits[0]['match_score']}"
        )

        retrieved = service.retrieve(a["id"])
        assert retrieved and retrieved["status"] == "retrieved"

        print("OK", a["ticket_code"], b["ticket_code"], "top_score", hits[0]["match_score"])
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
