"""拍照查找：再拍一张 → 告诉你位置，并展示原图。"""

from __future__ import annotations

import streamlit as st

from luggage.config import BAG_COLORS
from luggage import db, service
from luggage.ui import inject_base_css, show_bag

st.set_page_config(page_title="拍照查找", page_icon="🔎", layout="wide")
inject_base_css()

st.title("拍照查找")
st.caption("存了多少件都行。对着要找的行李再拍一张，系统告诉你位置并给出存档原图。")

stored = db.count_by_status().get("stored", 0)
st.write(f"当前在库：**{stored}** 件")

tab_photo, tab_card = st.tabs(["拍照找", "按卡联号找"])

with tab_photo:
    photo = st.camera_input("拍要找的行李")
    upload = st.file_uploader("或上传照片", type=["jpg", "jpeg", "png", "webp"])
    color = st.selectbox("可选：按颜色缩小范围", BAG_COLORS)
    if st.button("开始找", type="primary", use_container_width=True):
        raw = None
        if photo is not None:
            raw = photo.getvalue()
        elif upload is not None:
            raw = upload.getvalue()
        if not raw:
            st.error("请先拍照或上传")
        elif stored == 0:
            st.warning("还没有存档，先去「拍照存档」。")
        else:
            with st.spinner("比对中…"):
                hits = service.find_by_photo(
                    raw,
                    bag_color=None if color == "不标注" else color,
                )
            if not hits:
                st.warning("没找到够像的。换个角度再拍，或用卡联号查。")
            else:
                best = hits[0]
                st.success(
                    f"最可能是卡联 **{best['card_tag']}** → 位置 **{best['location']}**"
                )
                for hit in hits:
                    with st.container(border=True):
                        show_bag(hit, show_score=True)
                        c1, c2 = st.columns(2)
                        with c1:
                            if hit.get("status") == "stored" and st.button(
                                "取出这件", key=f"out_{hit['id']}"
                            ):
                                service.retrieve(hit["id"])
                                st.rerun()
                        with c2:
                            new_note = st.text_input(
                                "改备注",
                                value=hit.get("note") or "",
                                key=f"note_{hit['id']}",
                            )
                            if st.button("保存备注", key=f"save_{hit['id']}"):
                                service.save_remark(hit["id"], note=new_note)
                                st.rerun()

with tab_card:
    tag = st.text_input("卡联号")
    if st.button("按卡联查找", type="primary"):
        rows = service.find_by_card(tag) if tag.strip() else []
        if not rows:
            st.warning("没有匹配的在库记录")
        for row in rows:
            with st.container(border=True):
                show_bag(row)
                if st.button("取出", key=f"card_out_{row['id']}"):
                    service.retrieve(row["id"])
                    st.rerun()
