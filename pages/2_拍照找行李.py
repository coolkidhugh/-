"""行李寄存 · 拍照找行李"""

from __future__ import annotations

import streamlit as st

from luggage.config import BAG_COLORS
from luggage import db, service
from luggage.ui import inject_base_css, show_item_card

st.set_page_config(page_title="拍照找行李", page_icon="📷", layout="wide")
inject_base_css()

st.title("拍照找行李")
st.caption("对着行李或相似照片拍一张，系统按外形/颜色在在库行李里找候选。")

stored = db.count_by_status().get("stored", 0)
st.write(f"当前在库：**{stored}** 件")

photo = st.camera_input("拍摄要找的行李")
upload = st.file_uploader("或上传照片", type=["jpg", "jpeg", "png", "webp"])
color_filter = st.selectbox("可选：按颜色缩小范围", ["（不限）"] + BAG_COLORS)
top_k = st.slider("最多显示候选", 3, 12, 6)

if st.button("开始匹配", type="primary", use_container_width=True):
    raw = None
    if photo is not None:
        raw = photo.getvalue()
    elif upload is not None:
        raw = upload.getvalue()

    if not raw:
        st.error("请先拍照或上传图片")
    elif stored == 0:
        st.warning("暂无在库行李，请先到「存入行李」建档。")
    else:
        bag_color = None if color_filter == "（不限）" else color_filter
        with st.spinner("比对中…"):
            hits = service.find_by_photo(raw, bag_color=bag_color, top_k=top_k)

        if not hits:
            st.warning("没有足够相似的在库行李。可换角度再拍，或改用房号/票号查找。")
        else:
            st.success(f"找到 {len(hits)} 个候选（按相似度排序）")
            for hit in hits:
                with st.container(border=True):
                    show_item_card(hit, show_score=True)
                    if hit.get("status") == "stored":
                        if st.button(
                            f"确认取出 {hit['ticket_code']}",
                            key=f"retrieve_{hit['id']}",
                        ):
                            updated = service.retrieve(hit["id"])
                            if updated and updated["status"] == "retrieved":
                                st.success(f"{updated['ticket_code']} 已标记取出")
                                st.rerun()
