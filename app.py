import streamlit as st
import pandas as pd
import os
from analyze_excel import analyze_reports_ultimate

st.set_page_config(layout="wide", page_title="伯爵酒店工具箱")

st.title("伯爵酒店工具箱")
st.caption("Excel 团队报表分析 · 行李寄存 QR / 拍照找行李")

with st.expander("🧳 行李寄存 / 拍照找行李", expanded=True):
    st.markdown(
        """
用左侧页面进入：

1. **存入行李** — 拍照建档，生成 QR 票号，记录货架位置  
2. **拍照找行李** — 再拍一张，按外形/颜色在在库行李里找候选  
3. **扫码取件** — 票号 / 房号 / 姓名查找并取出  
4. **在库一览** — 库存列表

实现说明见 `luggage/SPEC.md`。
        """
    )

st.subheader("Excel 报告分析器")
st.markdown("---伯爵酒店团队报表分析工具---")

uploaded_files = st.file_uploader("请上传您的 Excel 报告文件 (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.subheader("分析结果")
    
    # Create a temporary directory to save uploaded files
    temp_dir = "./temp_uploaded_files"
    os.makedirs(temp_dir, exist_ok=True)

    file_paths = []
    for uploaded_file in uploaded_files:
        # Save the uploaded file to the temporary directory
        temp_file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        file_paths.append(temp_file_path)

    # Define the desired order of keywords
    desired_order = ["次日到达", "次日在住", "次日离店", "后天到达"]

    # Custom sort function
    def sort_key(file_path):
        file_name = os.path.basename(file_path)
        for i, keyword in enumerate(desired_order):
            if keyword in file_name:
                return i
        return len(desired_order) # Files without keywords go to the end

    # Sort the file_paths based on the desired order
    file_paths.sort(key=sort_key)

    if st.button("开始分析"): # Use a button to trigger analysis
        with st.spinner("正在分析中，请稍候..."):
            summaries, unknown_codes = analyze_reports_ultimate(file_paths)
        
        for summary in summaries:
            st.write(summary)

        if unknown_codes:
            st.subheader("侦测到的未知房型代码 (请检查是否需要更新规则)")
            for code, count in unknown_codes.items():
                st.write(f"代码: '{code}' (出现了 {count} 次)")
        
        # Clean up temporary files and directory
        for f_path in file_paths:
            os.remove(f_path)
        os.rmdir(temp_dir)

else:
    st.info("请上传一个或多个 Excel 文件以开始分析。")

st.markdown(
    """
---
#### Excel 使用说明
1. 点击 Browse files 上传 Excel 报告（可多选）。
2. 点击「开始分析」。
3. 结果与未知房型代码会显示在下方。
"""
)
