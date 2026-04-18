import streamlit as st
import pandas as pd
from docx import Document
from PIL import Image
import numpy as np
import difflib
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io

# --- 悬浮办公风格与侧边栏安全区设计 ---
STYLE = """
<style>
    /* 1. 全局背景 */
    .stApp {
        background-color: #f3f2f1;
        font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
    }
    
    /* 2. 主容器宽度 */
    div[data-testid="stAppViewBlockContainer"] {
        max-width: 98% !important;
        padding-top: 6rem !important; /* 给悬浮标题留出更多空间 */
    }

    /* 3. 缩短且悬浮的中心标题栏 (解决覆盖问题) */
    .app-header {
        position: fixed;
        top: 0;
        left: 20%; /* 左右各缩进20%，确保左侧100%安全 */
        right: 20%;
        height: 50px;
        background-color: #ffffff;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        z-index: 999;
        border-radius: 0 0 12px 12px;
        border: 1px solid #0078d4;
        border-top: none;
    }
    .app-title {
        color: #0078d4;
        font-size: 1.3rem;
        font-weight: 700;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* 4. 侧边栏折叠后的蓝色方块 (由于标题栏缩短，此处不再有冲突) */
    div[data-testid="stSidebarCollapsedControl"] {
        background-color: #0078d4 !important;
        top: 10px !important;
        left: 10px !important;
        width: 45px !important;
        height: 45px !important;
        border-radius: 8px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        z-index: 1000 !important;
        box-shadow: 2px 2px 10px rgba(0,120,212,0.3) !important;
    }
    
    /* 内部图标白色加粗 */
    div[data-testid="stSidebarCollapsedControl"] button svg {
        fill: white !important;
        color: white !important;
        transform: scale(1.2);
    }

    /* 5. 其他办公样式保持 */
    .line-num { color: #888; font-size: 0.8rem; text-align: right; padding-right: 8px; border-right: 1px solid #ddd; min-width: 40px; }
    .doc-content { padding: 6px 12px; font-size: 0.95rem; background-color: white; border-bottom: 1px solid #f3f2f1; }
    .diff-removed { background-color: #fff4ce !important; border-left: 5px solid #ffb900; }
    .diff-added { background-color: #dff6dd !important; border-left: 5px solid #107c10; }
    
    .stButton>button { background-color: #0078d4; color: white; border-radius: 4px; border: none; }
</style>
"""

HEADER_HTML = """
<div class="app-header">
    <div class="app-title">
        <svg width="24" height="24" viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M18 2H8C6.9 2 6.01 2.9 6.01 4L6 28C6 29.1 6.89 30 7.99 30H24C25.1 30 26 29.1 26 28V10L18 2Z" fill="#0078D4"/>
            <path d="M18 2V10H26L18 2Z" fill="#C7E0F4"/>
            <path d="M12 16H20V18H12V16ZM12 20H20V22H12V20ZM12 12H16V14H12V12Z" fill="white"/>
        </svg>
        office-compare
    </div>
</div>
"""

# --- Logic ---

def load_docx_lines(file):
    doc = Document(file)
    return [p.text for p in doc.paragraphs]

def save_docx_from_lines(lines):
    doc = Document()
    for line in lines: doc.add_paragraph(line)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def apply_merge(side, i1, i2, j1, j2):
    if side == 'to_right':
        st.session_state.doc_lines_b[j1:j2] = st.session_state.doc_lines_a[i1:i2]
    else:
        st.session_state.doc_lines_a[i1:i2] = st.session_state.doc_lines_b[j1:j2]
    st.toast("✅ 已合并修改")

def blend_overlay(orig_arr, mask_diff, mask_above, alpha):
    h, w, _ = orig_arr.shape
    overlay = np.full((h, w, 3), 128, dtype=np.uint8)
    overlay[mask_diff & ~mask_above] = [0, 120, 212]
    overlay[mask_above] = [232, 17, 35]
    return (orig_arr.astype(float)*(1-alpha) + overlay.astype(float)*alpha).astype(np.uint8)

# --- UI Setup ---

st.set_page_config(page_title="office-compare", layout="wide", initial_sidebar_state="expanded")
st.markdown(STYLE, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

if 'doc_lines_a' not in st.session_state: st.session_state.doc_lines_a = None
if 'doc_lines_b' not in st.session_state: st.session_state.doc_lines_b = None

with st.sidebar:
    st.markdown("## ⚙️ 设置面板")
    st.markdown("---")
    threshold = st.slider("差异阈值", 0, 255, 30)
    alpha = st.slider("叠加强度", 0, 100, 40) / 100.0
    st.markdown("---")
    if st.session_state.doc_lines_a is not None:
        st.download_button("💾 导出文件 A", save_docx_from_lines(st.session_state.doc_lines_a), "A_modified.docx")
        st.download_button("💾 导出文件 B", save_docx_from_lines(st.session_state.doc_lines_b), "B_modified.docx")

tabs = st.tabs(["📄 文档编辑", "🖼️ 图像比对"])

with tabs[0]:
    u1, u2 = st.columns(2)
    with u1: fa = st.file_uploader("文件 A", type=["docx"], key="fa")
    with u2: fb = st.file_uploader("文件 B", type=["docx"], key="fb")
    
    if fa and fb and (st.session_state.doc_lines_a is None or st.button("🔄 重新加载")):
        st.session_state.doc_lines_a, st.session_state.doc_lines_b = load_docx_lines(fa), load_docx_lines(fb)

    if st.session_state.doc_lines_a is not None:
        opcodes = difflib.SequenceMatcher(None, st.session_state.doc_lines_a, st.session_state.doc_lines_b).get_opcodes()
        for tag, i1, i2, j1, j2 in opcodes:
            r = st.columns([0.6, 9, 1.4, 9, 0.6])
            if tag == 'equal':
                for k in range(i2-i1):
                    r = st.columns([0.6, 9, 1.4, 9, 0.6])
                    r[0].markdown(f"<div class='line-num'>{i1+k+1}</div>", unsafe_allow_html=True)
                    r[1].markdown(f"<div class='doc-content'>{st.session_state.doc_lines_a[i1+k]}</div>", unsafe_allow_html=True)
                    r[2].markdown("<center style='color:#ccc'>—</center>", unsafe_allow_html=True)
                    r[3].markdown(f"<div class='doc-content'>{st.session_state.doc_lines_b[j1+k]}</div>", unsafe_allow_html=True)
                    r[4].markdown(f"<div class='line-num'>{j1+k+1}</div>", unsafe_allow_html=True)
            else:
                r = st.columns([0.6, 9, 1.4, 9, 0.6])
                cA = "\n".join(st.session_state.doc_lines_a[i1:i2]); sA = "diff-removed" if i2>i1 else ""
                r[0].markdown(f"<div class='line-num'>{i1+1 if i2>i1 else ''}</div>", unsafe_allow_html=True)
                r[1].markdown(f"<div class='doc-content {sA}'>{cA if cA else '(空)'}</div>", unsafe_allow_html=True)
                with r[2]:
                    c_b1, c_b2 = st.columns(2)
                    if c_b1.button("←", key=f"L{i1}{j1}"): apply_merge('to_left', i1, i2, j1, j2); st.rerun()
                    if c_b2.button("→", key=f"R{i1}{j1}"): apply_merge('to_right', i1, i2, j1, j2); st.rerun()
                cB = "\n".join(st.session_state.doc_lines_b[j1:j2]); sB = "diff-added" if j2>j1 else ""
                r[3].markdown(f"<div class='doc-content {sB}'>{cB if cB else '(空)'}</div>", unsafe_allow_html=True)
                r[4].markdown(f"<div class='line-num'>{j1+1 if j2>j1 else ''}</div>", unsafe_allow_html=True)

with tabs[1]:
    iu1, iu2 = st.columns(2)
    with iu1: ia_f = st.file_uploader("图 A", type=["png","jpg","jpeg","PNG","JPG","JPEG"], key="ia")
    with iu2: ib_f = st.file_uploader("图 B", type=["png","jpg","jpeg","PNG","JPG","JPEG"], key="ib")
    if ia_f and ib_f:
        if st.button("🚀 像素级比对"):
            img1, img2 = Image.open(ia_f).convert("RGB"), Image.open(ib_f).convert("RGB")
            h, w = max(img1.size[1], img2.size[1]), max(img1.size[0], img2.size[0])
            a1, a2 = np.zeros((h, w, 3), dtype=np.uint8), np.zeros((h, w, 3), dtype=np.uint8)
            a1[:img1.size[1], :img1.size[0], :] = np.array(img1)
            a2[:img2.size[1], :img2.size[0], :] = np.array(img2)
            md = np.sum(np.abs(a1.astype(np.int16)-a2.astype(np.int16)), 2)>0
            ma = np.mean(np.abs(a1.astype(np.int16)-a2.astype(np.int16)), 2)>threshold
            va, vb = blend_overlay(a1, md, ma, alpha), blend_overlay(a2, md, ma, alpha)
            MAX_DIM = 1200
            if max(h, w) > MAX_DIM:
                sc = MAX_DIM / max(h, w); ns = (int(w*sc), int(h*sc))
                vad, vbd = np.array(Image.fromarray(va).resize(ns)), np.array(Image.fromarray(vb).resize(ns))
                oad, obd = np.array(Image.fromarray(a1).resize(ns)), np.array(Image.fromarray(a2).resize(ns))
                dx, dy = 1/sc, 1/sc
            else: vad, vbd, oad, obd, dx, dy = va, vb, a1, a2, 1, 1
            comb = np.concatenate([oad, obd], axis=2)
            fig = make_subplots(rows=1, cols=2, horizontal_spacing=0.01)
            tpl = "<b>X:%{x:.0f} Y:%{y:.0f}</b><br>A RGB: (%{customdata[0]},%{customdata[1]},%{customdata[2]})<br>B RGB: (%{customdata[3]},%{customdata[4]},%{customdata[5]})<extra></extra>"
            fig.add_trace(go.Image(z=vad, dx=dx, dy=dy, customdata=comb, hovertemplate=tpl), row=1, col=1)
            fig.add_trace(go.Image(z=vbd, dx=dx, dy=dy, customdata=comb, hovertemplate=tpl), row=1, col=2)
            fig.update_xaxes(matches='x'); fig.update_yaxes(matches='y')
            fig.update_layout(height=800, margin=dict(l=0, r=0, b=0, t=20), hovermode='closest')
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
