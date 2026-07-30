"""行李寄存 · 在库一览"""

from __future__ import annotations

import streamlit as st
import pandas as pd

from luggage import db
from luggage.ui import inject_base_css, show_item_card

st.set_page_config(page_title="在库一览", page_icon="📋", layout="wide")
inject_base_css()

st.title("在库一览")

counts = db.count_by_status()
c1, c2, c3 = st.columns(3)
c1.metric("在库", counts.get("stored", 0))
c2.metric("已取出", counts.get("retrieved", 0))
c3.metric("合计", sum(counts.values()) if counts else 0)

status = st.radio("状态", ["stored", "retrieved", "全部"], horizontal=True, format_func=lambda x: {"stored": "在库", "retrieved": "已取出", "全部": "全部"}[x])
q_room = st.text_input("筛选房号", "")
q_name = st.text_input("筛选姓名", "")

status_arg = None if status == "全部" else status
rows = db.list_items(
    status=status_arg,
    room_no=q_room or None,
    guest_name=q_name or None,
    limit=300,
)

if not rows:
    st.info("暂无记录")
else:
    df = pd.DataFrame(rows)[
        [
            "ticket_code",
            "guest_name",
            "room_no",
            "bag_color",
            "bag_type",
            "location",
            "piece_count",
            "status",
            "created_at",
            "retrieved_at",
        ]
    ]
    df = df.rename(
        columns={
            "ticket_code": "票号",
            "guest_name": "客人",
            "room_no": "房号",
            "bag_color": "颜色",
            "bag_type": "类型",
            "location": "位置",
            "piece_count": "件数",
            "status": "状态",
            "created_at": "存入时间",
            "retrieved_at": "取出时间",
        }
    )
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.subheader("明细卡片")
    pick = st.selectbox("选择票号查看", [r["ticket_code"] for r in rows])
    item = next(r for r in rows if r["ticket_code"] == pick)
    show_item_card(item)
