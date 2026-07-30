"""按编号看实体图：输入 0056469 / 56469 → 直接带出现场照片。"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from luggage import service
from luggage.ui import inject_base_css, show_bag

st.set_page_config(page_title="按编号看图", page_icon="🎫", layout="wide")
inject_base_css()

st.title("按编号看实体图")
st.caption("输入卡联号（如 56469 / 0056469），直接带出现场实体照片和位置。")

tag = st.text_input("卡联号", placeholder="56469 或 0056469", value=st.session_state.get("q_tag", ""))
c1, c2 = st.columns([1, 1])
with c1:
    lookup = st.button("查看实体图", type="primary", use_container_width=True)
with c2:
    st.write("")

if lookup or (tag and st.session_state.get("_auto_lookup")):
    rows = service.find_by_card(tag) if tag.strip() else []
    photo = service.get_photo_path(tag) if tag.strip() else None

    if not rows and not photo:
        st.error("没有这个编号的存档。先去「拍照存档」存实体图。")
    else:
        if rows:
            bag = rows[0]
            st.success(f"卡联 **{bag['card_tag']}** → 位置 **{bag['location']}**")
            # 大图优先
            path = service.get_photo_path(bag["card_tag"]) or Path(bag.get("photo_path") or "")
            if path and path.is_file():
                st.image(str(path), use_container_width=True, caption=f"实体图 · {bag['card_tag']}")
            else:
                st.warning("该编号还没有实体图，请在下方补传现场照片。")
            show_bag(bag)
        elif photo:
            st.image(str(photo), use_container_width=True, caption=f"实体图 · {tag}")

st.divider()
st.subheader("给该编号补传 / 替换实体图")
st.caption("0056469 若只有标注卡，在这里上传现场原图即可。")
up_tag = st.text_input("要补图的卡联号", value=tag or "0056469")
up_cam = st.camera_input("拍现场实体图")
up_file = st.file_uploader("或上传现场照片", type=["jpg", "jpeg", "png", "webp"])
if st.button("保存实体图到该编号", type="primary"):
    raw = None
    if up_cam is not None:
        raw = up_cam.getvalue()
    elif up_file is not None:
        raw = up_file.getvalue()
    try:
        bag = service.replace_photo(up_tag, raw or b"")
        st.success(f"已挂上实体图 · {bag['card_tag']}")
        path = service.get_photo_path(bag["card_tag"])
        if path:
            st.image(str(path), use_container_width=True)
        st.rerun()
    except ValueError as exc:
        st.error(str(exc))
    except Exception as exc:  # noqa: BLE001
        st.exception(exc)
