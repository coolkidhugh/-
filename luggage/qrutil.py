"""二维码生成：票号 → 可扫描内容。"""

from __future__ import annotations

from pathlib import Path

import qrcode
from qrcode.constants import ERROR_CORRECT_M

from luggage.config import QR_DIR
from luggage.db import ensure_dirs


def ticket_payload(ticket_code: str) -> str:
    """二维码内容：协议前缀 + 票号，便于手机扫码后识别。"""
    return f"LUGGAGE:{ticket_code.strip().upper()}"


def parse_ticket_payload(raw: str) -> str | None:
    text = (raw or "").strip()
    if not text:
        return None
    upper = text.upper()
    if upper.startswith("LUGGAGE:"):
        return upper.split(":", 1)[1].strip()
    # 也允许直接扫到纯票号
    if upper.startswith("L") and len(upper) >= 6:
        return upper
    return upper if text else None


def make_qr_image(ticket_code: str, save: bool = True) -> tuple[object, Path | None]:
    ensure_dirs()
    payload = ticket_payload(ticket_code)
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_M,
        box_size=8,
        border=2,
    )
    qr.add_data(payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    path: Path | None = None
    if save:
        path = QR_DIR / f"{ticket_code.strip().upper()}.png"
        img.save(path)
    return img, path
