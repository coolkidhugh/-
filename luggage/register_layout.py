#!/usr/bin/env python3
"""按 layout.py 批量登记货架位（可重复执行，同编号更新位置）。"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from luggage.layout import LAYOUT  # noqa: E402
from luggage import service  # noqa: E402


def main() -> int:
    created = []
    for row_name, items in LAYOUT:
        for idx, (card, extra) in enumerate(items, start=1):
            location = f"{row_name} · 位{idx}"
            notes = []
            if extra.get("pieces"):
                notes.append(f"{extra['pieces']}件")
            if extra.get("note"):
                notes.append(str(extra["note"]))
            notes.append("待补实体图")
            bag = service.register_slot(
                card_tag=str(card),
                location=location,
                note="；".join(notes),
                bag_color=extra.get("color") or "",
            )
            created.append((bag["card_tag"], bag["location"], bag.get("bag_color") or ""))
            print(f"OK {bag['card_tag']:>8}  {bag['location']}  {bag.get('bag_color') or '-'}")
    print(f"\n合计 {len(created)} 件已登记")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
