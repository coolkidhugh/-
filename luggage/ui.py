"""展示卡片。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import streamlit as st


def show_bag(bag: dict[str, Any], *, show_score: bool = False) -> None:
    left, right = st.columns([1, 1.1])
    with left:
        photo = bag.get("photo_path") or ""
        if photo and Path(photo).is_file():
            st.image(photo, use_container_width=True, caption="存档照片")
        else:
            st.warning("照片文件缺失")
    with right:
        if show_score and "match_score" in bag:
            pct = int(round(float(bag["match_score"]) * 100))
            st.metric("相似度", f"{pct}%")
        st.markdown(f"### 卡联：`{bag.get('card_tag', '')}`")
        st.success(f"位置：**{bag.get('location', '')}**")
        color = bag.get("bag_color") or ""
        if color:
            st.write(f"颜色：{color}")
        note = bag.get("note") or ""
        if note:
            st.info(f"备注：{note}")
        else:
            st.caption("暂无备注")
        st.caption(f"存入时间：{bag.get('created_at', '')}")


def inject_base_css() -> None:
    st.markdown(
        """
        <style>
        .block-container { max-width: 920px; padding-top: 1.1rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )
