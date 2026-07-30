"""冒烟：存实体图 → 按短编号取出实体图路径。"""

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
    cfg.DATA_DIR = tmp / "data"
    cfg.DB_PATH = cfg.DATA_DIR / "luggage.db"
    cfg.PHOTOS_DIR = cfg.DATA_DIR / "photos"
    db.DATA_DIR = cfg.DATA_DIR
    db.DB_PATH = cfg.DB_PATH
    db.PHOTOS_DIR = cfg.PHOTOS_DIR
    service.PHOTOS_DIR = cfg.PHOTOS_DIR
    service.RECORDS_DIR = tmp / "luggage_records"

    try:
        a = service.deposit(
            photo_bytes=_bag((20, 20, 20), "rect", "A"),
            location="货架A-上",
            card_tag="56469",
            note="实体黑箱",
            bag_color="黑",
        )
        assert a["card_tag"] == "0056469"
        photo = service.get_photo_path("56469")
        assert photo and photo.is_file(), "查编号必须带出实体图文件"

        hits = service.find_by_card("56469")
        assert hits and hits[0]["card_tag"] == "0056469"

        # 补传替换
        service.replace_photo("56469", _bag((30, 90, 200), "ellipse", "B"))
        photo2 = service.get_photo_path("0056469")
        assert photo2 and photo2.is_file()

        print("OK", a["card_tag"], photo)
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
