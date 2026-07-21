import html
import os

import streamlit as st

from analyze_excel import analyze_reports_ultimate

st.set_page_config(
    layout="wide",
    page_title="伯爵酒店 · 团队报表分析",
    page_icon="🏨",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;500;600;700&family=Noto+Serif+SC:wght@500;600;700&display=swap" rel="stylesheet">
    <style>
    :root {
        --ink: #15202b;
        --ink-soft: #3a4a58;
        --mist: #e8eef3;
        --paper: #f4f7fa;
        --brass: #a67c52;
        --brass-deep: #8a6540;
        --river: #4a6d7c;
        --river-soft: #6b8a98;
        --line: rgba(21, 32, 43, 0.10);
        --glow: rgba(74, 109, 124, 0.18);
    }

    html, body, [class*="css"] {
        font-family: "Noto Sans SC", sans-serif;
        color: var(--ink);
    }

    .stApp {
        background:
            radial-gradient(1200px 600px at 12% -10%, rgba(74, 109, 124, 0.22), transparent 55%),
            radial-gradient(900px 500px at 88% 8%, rgba(166, 124, 82, 0.14), transparent 50%),
            linear-gradient(165deg, #d9e4eb 0%, var(--paper) 42%, #eef2f5 100%);
        background-attachment: fixed;
    }

    .stApp::before {
        content: "";
        position: fixed;
        inset: 0;
        pointer-events: none;
        z-index: 0;
        opacity: 0.35;
        background-image:
            linear-gradient(rgba(21, 32, 43, 0.03) 1px, transparent 1px),
            linear-gradient(90deg, rgba(21, 32, 43, 0.03) 1px, transparent 1px);
        background-size: 48px 48px;
        mask-image: radial-gradient(ellipse at center, black 30%, transparent 78%);
    }

    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 980px !important;
        position: relative;
        z-index: 1;
    }

    #MainMenu, footer, header { visibility: hidden; }

    .hero {
        text-align: center;
        padding: 1.4rem 1rem 2rem;
        animation: rise 0.7s ease-out both;
    }

    .hero-brand {
        font-family: "Noto Serif SC", serif;
        font-size: clamp(2.4rem, 5vw, 3.4rem);
        font-weight: 700;
        letter-spacing: 0.12em;
        color: var(--ink);
        margin: 0;
        line-height: 1.15;
    }

    .hero-rule {
        width: 56px;
        height: 2px;
        margin: 1rem auto 0.95rem;
        background: linear-gradient(90deg, transparent, var(--brass), transparent);
        animation: draw 0.9s ease 0.25s both;
    }

    .hero-sub {
        font-family: "Noto Sans SC", sans-serif;
        font-size: 1.02rem;
        font-weight: 500;
        color: var(--ink-soft);
        letter-spacing: 0.18em;
        margin: 0;
    }

    .hero-note {
        margin-top: 0.7rem;
        font-size: 0.92rem;
        color: var(--river-soft);
        letter-spacing: 0.04em;
    }

    .panel {
        background: rgba(255, 255, 255, 0.72);
        backdrop-filter: blur(10px);
        border: 1px solid var(--line);
        border-radius: 18px;
        padding: 1.35rem 1.4rem 1.2rem;
        box-shadow: 0 18px 40px -28px rgba(21, 32, 43, 0.45);
        animation: rise 0.75s ease-out 0.12s both;
    }

    .panel-title {
        font-family: "Noto Serif SC", serif;
        font-size: 1.15rem;
        font-weight: 600;
        color: var(--ink);
        margin: 0 0 0.35rem;
        letter-spacing: 0.06em;
    }

    .panel-desc {
        margin: 0 0 0.9rem;
        color: var(--ink-soft);
        font-size: 0.92rem;
        line-height: 1.55;
    }

    .result-shell {
        margin-top: 1.1rem;
        animation: rise 0.55s ease-out both;
    }

    .result-card {
        background: linear-gradient(135deg, rgba(255,255,255,0.92), rgba(244,247,250,0.88));
        border: 1px solid var(--line);
        border-left: 3px solid var(--brass);
        border-radius: 14px;
        padding: 1rem 1.15rem;
        margin-bottom: 0.75rem;
        box-shadow: 0 10px 28px -22px rgba(21, 32, 43, 0.5);
        transition: transform 0.25s ease, box-shadow 0.25s ease;
    }

    .result-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 16px 32px -20px rgba(21, 32, 43, 0.55);
    }

    .result-label {
        font-size: 0.78rem;
        letter-spacing: 0.14em;
        text-transform: uppercase;
        color: var(--brass-deep);
        font-weight: 600;
        margin-bottom: 0.35rem;
    }

    .result-body {
        color: var(--ink);
        font-size: 0.98rem;
        line-height: 1.65;
        word-break: break-word;
    }

    .warn-card {
        border-left-color: var(--river);
        background: linear-gradient(135deg, rgba(255,255,255,0.9), rgba(232,238,243,0.85));
    }

    .guide {
        margin-top: 1.6rem;
        padding: 1.1rem 1.2rem;
        border-top: 1px solid var(--line);
        animation: rise 0.8s ease-out 0.2s both;
    }

    .guide h3 {
        font-family: "Noto Serif SC", serif;
        font-size: 1rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        color: var(--ink);
        margin: 0 0 0.55rem;
    }

    .guide ol {
        margin: 0;
        padding-left: 1.15rem;
        color: var(--ink-soft);
        line-height: 1.8;
        font-size: 0.92rem;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.45);
        border: 1.5px dashed rgba(74, 109, 124, 0.35);
        border-radius: 14px;
        padding: 0.55rem 0.7rem 0.2rem;
        transition: border-color 0.25s ease, background 0.25s ease, box-shadow 0.25s ease;
    }

    div[data-testid="stFileUploader"]:hover {
        border-color: var(--brass);
        background: rgba(255, 255, 255, 0.7);
        box-shadow: 0 0 0 4px var(--glow);
    }

    div[data-testid="stFileUploader"] section {
        border: none !important;
        background: transparent !important;
    }

    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, var(--ink) 0%, #243647 100%) !important;
        color: #f7f4ef !important;
        border: none !important;
        border-radius: 999px !important;
        padding: 0.72rem 1.4rem !important;
        font-family: "Noto Sans SC", sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: 0.16em !important;
        box-shadow: 0 12px 28px -14px rgba(21, 32, 43, 0.7) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease, filter 0.2s ease !important;
    }

    .stButton > button:hover {
        transform: translateY(-1px);
        filter: brightness(1.06);
        box-shadow: 0 16px 32px -12px rgba(21, 32, 43, 0.75) !important;
    }

    .stButton > button:active {
        transform: translateY(0);
    }

    div[data-testid="stAlert"] {
        border-radius: 12px;
        border: 1px solid var(--line);
        background: rgba(255, 255, 255, 0.75);
    }

    .file-chip-row {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin: 0.75rem 0 0.2rem;
    }

    .file-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.28rem 0.7rem;
        border-radius: 999px;
        background: rgba(74, 109, 124, 0.10);
        color: var(--river);
        font-size: 0.82rem;
        font-weight: 500;
        border: 1px solid rgba(74, 109, 124, 0.18);
    }

    @keyframes rise {
        from { opacity: 0; transform: translateY(14px); }
        to { opacity: 1; transform: translateY(0); }
    }

    @keyframes draw {
        from { transform: scaleX(0); opacity: 0; }
        to { transform: scaleX(1); opacity: 1; }
    }

    @media (max-width: 640px) {
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        .hero {
            padding-top: 0.6rem;
        }
        .hero-brand {
            letter-spacing: 0.08em;
        }
        .panel {
            padding: 1.1rem 1rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <section class="hero">
        <h1 class="hero-brand">伯爵酒店</h1>
        <div class="hero-rule"></div>
        <p class="hero-sub">团队报表分析</p>
        <p class="hero-note">上传次日到达 / 在住 / 离店等 Excel，一键汇总楼栋与团队房量</p>
    </section>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="panel">
        <p class="panel-title">上传报告</p>
        <p class="panel-desc">支持同时选择多个 .xlsx 文件。系统会按「次日到达 → 次日在住 → 次日离店 → 后天到达」自动排序后分析。</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "选择 Excel 文件",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

desired_order = ["次日到达", "次日在住", "次日离店", "后天到达"]


def sort_key(file_path: str) -> int:
    file_name = os.path.basename(file_path)
    for i, keyword in enumerate(desired_order):
        if keyword in file_name:
            return i
    return len(desired_order)


if uploaded_files:
    chips = "".join(
        f'<span class="file-chip">{html.escape(os.path.splitext(f.name)[0])}</span>'
        for f in uploaded_files
    )
    st.markdown(
        f'<div class="file-chip-row">{chips}</div>',
        unsafe_allow_html=True,
    )

    analyze = st.button("开始分析", type="primary", use_container_width=True)

    if analyze:
        temp_dir = "./temp_uploaded_files"
        os.makedirs(temp_dir, exist_ok=True)

        file_paths = []
        try:
            for uploaded_file in uploaded_files:
                temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                file_paths.append(temp_file_path)

            file_paths.sort(key=sort_key)

            with st.spinner("正在解析报表，请稍候…"):
                summaries, unknown_codes = analyze_reports_ultimate(file_paths)

            st.markdown(
                """
                <div class="result-shell">
                    <p class="panel-title">分析结果</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            for idx, summary in enumerate(summaries, start=1):
                st.markdown(
                    f"""
                    <div class="result-card">
                        <div class="result-label">报告 {idx:02d}</div>
                        <div class="result-body">{html.escape(summary)}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            if unknown_codes:
                codes_html = "".join(
                    f'<div class="result-body">代码「{html.escape(str(code))}」出现 {count} 次</div>'
                    for code, count in unknown_codes.items()
                )
                st.markdown(
                    f"""
                    <div class="result-card warn-card">
                        <div class="result-label">未知房型代码</div>
                        <div class="panel-desc" style="margin-bottom:0.55rem;">以下代码未匹配金陵楼 / 亚太楼规则，请核对是否需要更新。</div>
                        {codes_html}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        finally:
            for f_path in file_paths:
                if os.path.exists(f_path):
                    os.remove(f_path)
            if os.path.isdir(temp_dir) and not os.listdir(temp_dir):
                os.rmdir(temp_dir)
else:
    st.info("请先上传一个或多个 Excel 报告文件。")

st.markdown(
    """
    <div class="guide">
        <h3>使用说明</h3>
        <ol>
            <li>点击上方区域上传 Excel 报告，可一次选择多个文件。</li>
            <li>确认文件列表后，点击「开始分析」。</li>
            <li>结果会按报告逐条展示；如有未知房型代码，会单独提示。</li>
        </ol>
    </div>
    """,
    unsafe_allow_html=True,
)
