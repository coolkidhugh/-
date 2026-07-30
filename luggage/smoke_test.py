"""冒烟：存几件 → 拍照找回位置。"""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


def _bag(color: tuple[int, int, int], shape: str, mark: str) -> bytes:
    img = Image.new("RGB", (320, 420), (235, 235, 235))
    draw = ImageDraw.Draw(img)
    if shape == "rect":
        draw.rectangle([50, 40, 270, 380], fill=color)
    else:
        draw.ellipse([40, 80, 280, 360], fill=color)
    draw.text((70, 190), mark, fill=(255, 255, 255))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    return buf.getvalue()


def main() -> int:
    import luggage.config as cfg
    import luggage.db as db
    import luggage.service as service

    tmp = Path(tempfile.mkdtemp(prefix="bag_test_"))
    cfg.DATA_DIR = tmp
    cfg.DB_PATH = tmp / "luggage.db"
    cfg.PHOTOS_DIR = tmp / "photos"
    db.DATA_DIR = tmp
    db.DB_PATH = cfg.DB_PATH
    db.PHOTOS_DIR = cfg.PHOTOS_DIR
    service.PHOTOS_DIR = cfg.PHOTOS_DIR

    try:
        a = service.deposit(
            photo_bytes=_bag((20, 20, 20), "rect", "A"),
            location="货架A-上",
            card_tag="卡联37",
            note="演示黑箱",
            bag_color="黑",
        )
        b = service.deposit(
            photo_bytes=_bag((30, 90, 200), "ellipse", "B"),
            location="货架B-下",
            card_tag="卡联88",
            note="演示蓝包",
            bag_color="蓝",
        )
        # 再存 8 件凑够「存了10个」
        for i in range(8):
            service.deposit(
                photo_bytes=_bag((80 + i * 15, 40, 40 + i * 10), "ellipse", str(i)),
                location=f"临时区/{i}",
                card_tag=f"卡联{i}",
                note=f"填充{i}",
            )

        assert db.count_by_status().get("stored") == 10

        hits = service.find_by_photo(_bag((25, 25, 28), "rect", "Q"), top_k=3)
        assert hits, "应能找回"
        assert hits[0]["card_tag"] == "卡联37"
        assert hits[0]["location"] == "货架A-上"

        service.save_remark(a["id"], note="改过的备注")
        assert db.get_by_id(a["id"])["note"] == "改过的备注"

        by_card = service.find_by_card("卡联88")
        assert by_card and by_card[0]["id"] == b["id"]

        print("OK stored=10 top=", hits[0]["card_tag"], hits[0]["location"], hits[0]["match_score"])
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
