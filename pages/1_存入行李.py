"""行李寄存 · 存入（拍照 + 生成 QR）"""

from __future__ import annotations

import streamlit as st

from luggage.config import BAG_COLORS, BAG_TYPES, STORAGE_ZONES
from luggage import service
from luggage.ui import inject_base_css, show_item_card

st.set_page_config(page_title="存入行李", page_icon="🧳", layout="wide")
inject_base_css()

st.title("存入行李")
st.caption("拍照建档 → 选存放位置 → 生成 QR 票号（可贴行李 / 给客人）")

with st.form("deposit_form", clear_on_submit=False):
    c1, c2 = st.columns(2)
    with c1:
        guest_name = st.text_input("客人姓名 *")
        room_no = st.text_input("房号 *")
        phone = st.text_input("电话（可选）")
        piece_count = st.number_input("件数", min_value=1, max_value=20, value=1)
    with c2:
        bag_type = st.selectbox("行李类型", BAG_TYPES)
        bag_color = st.selectbox("主色", BAG_COLORS)
        brand_note = st.text_input("品牌/贴纸/特征（可选）", placeholder="如：Rimowa、绑红丝带")
        location = st.selectbox("存放位置 *", STORAGE_ZONES)

    note = st.text_area("备注", height=70)
    photo = st.camera_input("拍摄行李照片 *（推荐）")
    upload = st.file_uploader("或上传照片", type=["jpg", "jpeg", "png", "webp"])
    submitted = st.form_submit_button("确认存入并生成 QR", type="primary", use_container_width=True)

if submitted:
    raw = None
    if photo is not None:
        raw = photo.getvalue()
    elif upload is not None:
        raw = upload.getvalue()

    try:
        item = service.deposit(
            guest_name=guest_name,
            room_no=room_no,
            location=location,
            photo_bytes=raw or b"",
            phone=phone,
            piece_count=int(piece_count),
            bag_type=bag_type,
            bag_color=bag_color,
            brand_note=brand_note,
            note=note,
        )
        st.success(f"已存入 · 票号 **{item['ticket_code']}** · 位置 **{item['location']}**")
        show_item_card(item)
        st.download_button(
            "下载 QR 图片",
            data=open(item["qr_path"], "rb").read() if item.get("qr_path") else b"",
            file_name=f"{item['ticket_code']}.png",
            mime="image/png",
            disabled=not item.get("qr_path"),
        )
    except ValueError as exc:
        st.error(str(exc))
    except Exception as exc:  # noqa: BLE001
        st.exception(exc)
