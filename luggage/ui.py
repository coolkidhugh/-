"""Streamlit 展示辅助。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import streamlit as st


STATUS_LABEL = {
    "stored": "在库",
    "retrieved": "已取出",
}


def show_item_card(item: dict[str, Any], *, show_score: bool = False) -> None:
    left, right = st.columns([1, 1.2])
    with left:
        photo = item.get("photo_path") or ""
        if photo and Path(photo).is_file():
            st.image(photo, use_container_width=True, caption="行李照片")
        else:
            st.warning("照片文件缺失")
        qr = item.get("qr_path") or ""
        if qr and Path(qr).is_file():
            st.image(qr, width=160, caption=f"QR · {item.get('ticket_code', '')}")
    with right:
        status = STATUS_LABEL.get(item.get("status", ""), item.get("status", ""))
        st.markdown(f"### 票号 `{item.get('ticket_code', '')}`")
        if show_score and "match_score" in item:
            pct = int(round(float(item["match_score"]) * 100))
            st.metric("相似度", f"{pct}%")
        st.write(
            f"**客人**：{item.get('guest_name', '')}　"
            f"**房号**：{item.get('room_no', '')}"
        )
        if item.get("phone"):
            st.write(f"**电话**：{item['phone']}")
        st.write(
            f"**件数**：{item.get('piece_count', 1)}　"
            f"**类型**：{item.get('bag_type') or '—'}　"
            f"**颜色**：{item.get('bag_color') or '—'}"
        )
        if item.get("brand_note"):
            st.write(f"**特征**：{item['brand_note']}")
        st.success(f"存放位置：**{item.get('location', '')}**")
        st.caption(f"状态：{status} · 存入：{item.get('created_at', '')}")
        if item.get("note"):
            st.info(item["note"])


def inject_base_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { max-width: 980px; padding-top: 1.2rem; }
        div[data-testid="stMetricValue"] { font-size: 1.6rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )
