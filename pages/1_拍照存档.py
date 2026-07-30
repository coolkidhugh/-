"""拍照存档：照片 + 卡联标注 + 位置 + 备注。"""

from __future__ import annotations

import streamlit as st

from luggage.config import BAG_COLORS, STORAGE_ZONES
from luggage import service
from luggage.ui import inject_base_css, show_bag

st.set_page_config(page_title="拍照存档", page_icon="📷", layout="wide")
inject_base_css()

st.title("拍照存档")
st.caption("拍一张行李照 → 标卡联号 → 写位置和备注。以后用照片就能找回。")

with st.form("deposit"):
    photo = st.camera_input("拍行李照片（推荐）")
    upload = st.file_uploader("或上传照片", type=["jpg", "jpeg", "png", "webp"])
    card_tag = st.text_input("卡联号 / 标签", placeholder="例如 卡联 37、红绳、客人名")
    location = st.selectbox("存放位置", STORAGE_ZONES)
    custom_loc = st.text_input("或手写位置（优先）", placeholder="例如 货架A第3格")
    bag_color = st.selectbox("颜色（可选，方便筛选）", BAG_COLORS)
    note = st.text_area("备注", placeholder="房号、客人特征、特殊交代…", height=80)
    ok = st.form_submit_button("保存", type="primary", use_container_width=True)

if ok:
    raw = None
    if photo is not None:
        raw = photo.getvalue()
    elif upload is not None:
        raw = upload.getvalue()
    loc = custom_loc.strip() or location
    try:
        bag = service.deposit(
            photo_bytes=raw or b"",
            location=loc,
            card_tag=card_tag,
            note=note,
            bag_color=bag_color,
        )
        st.success(f"已存档 · 卡联 **{bag['card_tag']}** · 位置 **{bag['location']}**")
        show_bag(bag)
    except ValueError as exc:
        st.error(str(exc))
    except Exception as exc:  # noqa: BLE001
        st.exception(exc)
