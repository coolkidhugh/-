"""行李寄存 · 扫码/票号查找与取出"""

from __future__ import annotations

import streamlit as st

from luggage import db, service
from luggage.ui import inject_base_css, show_item_card

st.set_page_config(page_title="扫码取件", page_icon="🔎", layout="wide")
inject_base_css()

st.title("扫码 / 票号 / 房号查找")
st.caption("扫 QR 内容（或手输票号）、按房号、按姓名快速定位。")

tab_qr, tab_room, tab_name = st.tabs(["票号 / 扫码", "按房号", "按姓名"])

with tab_qr:
    raw = st.text_input(
        "票号或扫码内容",
        placeholder="例如 L0730A1B2 或 LUGGAGE:L0730A1B2",
    )
    if st.button("查询票号", type="primary") or (raw and st.session_state.get("_auto")):
        item = service.find_by_ticket_or_scan(raw)
        if not item:
            st.error("未找到该票号")
        else:
            show_item_card(item)
            if item["status"] == "stored":
                if st.button("确认取出", type="primary", key="retrieve_ticket"):
                    updated = service.retrieve(item["id"])
                    st.success(f"{updated['ticket_code']} 已取出")
                    st.rerun()
            else:
                st.info("该件已取出")

with tab_room:
    room = st.text_input("房号", key="room_q")
    if st.button("按房号搜索", key="btn_room"):
        rows = db.list_items(status="stored", room_no=room)
        if not rows:
            st.warning("无在库记录")
        for row in rows:
            with st.container(border=True):
                show_item_card(row)
                if st.button(f"取出 {row['ticket_code']}", key=f"r_{row['id']}"):
                    service.retrieve(row["id"])
                    st.rerun()

with tab_name:
    name = st.text_input("客人姓名", key="name_q")
    if st.button("按姓名搜索", key="btn_name"):
        rows = db.list_items(status="stored", guest_name=name)
        if not rows:
            st.warning("无在库记录")
        for row in rows:
            with st.container(border=True):
                show_item_card(row)
                if st.button(f"取出 {row['ticket_code']}", key=f"n_{row['id']}"):
                    service.retrieve(row["id"])
                    st.rerun()
