"""存档列表：看全部、改备注、取出。"""

from __future__ import annotations

import streamlit as st
import pandas as pd

from luggage import db, service
from luggage.config import STORAGE_ZONES
from luggage.ui import inject_base_css, show_bag

st.set_page_config(page_title="存档列表", page_icon="📋", layout="wide")
inject_base_css()

st.title("存档列表")

counts = db.count_by_status()
a, b, c = st.columns(3)
a.metric("在库", counts.get("stored", 0))
b.metric("已取出", counts.get("retrieved", 0))
c.metric("合计", sum(counts.values()) if counts else 0)

status = st.radio(
    "状态",
    ["stored", "retrieved", "全部"],
    horizontal=True,
    format_func=lambda x: {"stored": "在库", "retrieved": "已取出", "全部": "全部"}[x],
)
q = st.text_input("筛卡联号", "")

rows = db.list_bags(
    status=None if status == "全部" else status,
    card_tag=q or None,
)

if not rows:
    st.info("暂无记录。去「拍照存档」存几件。")
else:
    df = pd.DataFrame(rows)[["card_tag", "location", "bag_color", "note", "status", "created_at"]]
    df = df.rename(
        columns={
            "card_tag": "卡联",
            "location": "位置",
            "bag_color": "颜色",
            "note": "备注",
            "status": "状态",
            "created_at": "存入时间",
        }
    )
    st.dataframe(df, use_container_width=True, hide_index=True)

    labels = [f"{r['card_tag']} · {r['location']}" for r in rows]
    pick_idx = st.selectbox("打开一件", range(len(rows)), format_func=lambda i: labels[i])
    bag = rows[pick_idx]
    show_bag(bag)

    with st.form(f"edit_{bag['id']}"):
        new_tag = st.text_input("卡联", value=bag.get("card_tag") or "")
        new_loc = st.selectbox(
            "位置",
            STORAGE_ZONES
            if bag.get("location") in STORAGE_ZONES
            else [bag.get("location") or ""] + STORAGE_ZONES,
            index=0,
        )
        custom = st.text_input("或手写位置", value="")
        new_note = st.text_area("备注", value=bag.get("note") or "", height=80)
        save = st.form_submit_button("保存修改")
        if save:
            service.save_remark(
                bag["id"],
                card_tag=new_tag,
                location=custom.strip() or new_loc,
                note=new_note,
            )
            st.success("已保存")
            st.rerun()

    if bag.get("status") == "stored" and st.button("标记取出", type="primary"):
        service.retrieve(bag["id"])
        st.rerun()
