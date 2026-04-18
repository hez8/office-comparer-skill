import streamlit as st
import pandas as pd
from docx import Document
from PIL import Image
import numpy as np
import difflib
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import cv2
import time
import base64

# --- 1. 动态主题与高级样式 ---
def apply_custom_style(dark_mode=False):
    theme = {
        "bg": "#1e1e1e" if dark_mode else "#f3f2f1",
        "card": "#2d2d2d" if dark_mode else "rgba(255, 255, 255, 0.9)",
        "text": "#e1e1e1" if dark_mode else "#323130",
        "border": "#404040" if dark_mode else "#edebe9",
        "header_bg": "rgba(45, 45, 45, 0.8)" if dark_mode else "rgba(255, 255, 255, 0.7)",
        "diff_add": "#1e3a1e" if dark_mode else "#f0fdf4",
        "diff_add_border": "#22c55e",
        "diff_del": "#3e2723" if dark_mode else "#fff9e6",
        "diff_del_border": "#ffb900"
    }
    
    style_html = f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500&family=Inter:wght@400;600;700&display=swap');
        
        header[data-testid="stHeader"] {{ visibility: hidden; height: 0px; }}
        
        .stApp {{ 
            background: {theme["bg"]};
            color: {theme["text"]};
            font-family: 'Inter', sans-serif; 
            transition: all 0.3s ease;
        }}

        .app-header {{
            position: fixed; top: 0; left: 0; right: 0; height: 56px;
            background: {theme["header_bg"]}; backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
            display: flex; align-items: center; justify-content: center;
            box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1); z-index: 999999;
            border-bottom: 1px solid {theme["border"]};
        }}
        .app-title {{ 
            background: linear-gradient(90deg, #0078d4, #00bcf2);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            font-size: 1.4rem; font-weight: 700; display: flex; align-items: center; gap: 15px;
        }}

        div[data-testid="stAppViewBlockContainer"] {{ 
            max-width: 96% !important; padding-top: 70px !important;
            animation: slideUp 0.5s ease-out;
        }}

        .doc-content {{ 
            padding: 8px 15px; font-size: 0.9rem; background: {theme["card"]}; 
            border: 1px solid {theme["border"]}; line-height: 1.5; border-radius: 6px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.05); color: {theme["text"]};
            margin-bottom: 2px; white-space: pre-wrap; font-family: 'Fira Code', monospace;
        }}

        .diff-removed {{ background-color: {theme["diff_del"]} !important; border-left: 5px solid {theme["diff_del_border"]} !important; }}
        .diff-added {{ background-color: {theme["diff_add"]} !important; border-left: 5px solid {theme["diff_add_border"]} !important; }}
        
        .minimap-container {{
            position: fixed; right: 10px; top: 80px; width: 12px; height: 80vh;
            background: rgba(128,128,128,0.1); border-radius: 6px; z-index: 99;
        }}
        .minimap-bit {{ width: 100%; height: 2px; margin-bottom: 1px; }}

        section[data-testid="stSidebar"] {{ background: {theme["card"]} !important; border-right: 1px solid {theme["border"]}; }}
    </style>
    """
    st.markdown(style_html, unsafe_allow_html=True)

HEADER_HTML = """
<div class="app-header">
    <div class="app-title">
        <svg width="28" height="28" viewBox="0 0 32 32" fill="none"><path d="M18 2H8C6.9 2 6.01 2.9 6.01 4L6 28C6 29.1 6.89 30 7.99 30H24C25.1 30 26 29.1 26 28V10L18 2Z" fill="url(#grad1)" /><defs><linearGradient id="grad1"><stop offset="0%" stop-color="#0078D4"/><stop offset="100%" stop-color="#00bcf2"/></linearGradient></defs></svg>
        Office-Comparer Pro
    </div>
</div>
"""

# --- 2. 逻辑引擎 ---

def load_document_lines(uploaded_file):
    """支持 Word、代码、文本等多格式加载"""
    if uploaded_file is None: return []
    filename = uploaded_file.name.lower()
    content = uploaded_file.getvalue()
    
    try:
        if filename.endswith('.docx'):
            doc = Document(io.BytesIO(content))
            return [p.text for p in doc.paragraphs]
        else:
            # 尝试作为文本读取（处理 py, c, cpp, txt 等）
            text = content.decode('utf-8')
            return text.splitlines()
    except Exception as e:
        st.error(f"无法读取文件 {filename}: {e}")
        return []

@st.cache_data
def align_images_cv2(img_a_bytes, img_b_bytes):
    nparr_a, nparr_b = np.frombuffer(img_a_bytes, np.uint8), np.frombuffer(img_b_bytes, np.uint8)
    img_a, img_b = cv2.imdecode(nparr_a, cv2.IMREAD_COLOR), cv2.imdecode(nparr_b, cv2.IMREAD_COLOR)
    orb = cv2.ORB_create(2000)
    kp1, des1 = orb.detectAndCompute(img_a, None)
    kp2, des2 = orb.detectAndCompute(img_b, None)
    bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
    matches = sorted(bf.match(des1, des2), key=lambda x: x.distance)
    if len(matches) > 10:
        src_pts = np.float32([kp1[m.queryIdx].pt for m in matches]).reshape(-1, 1, 2)
        dst_pts = np.float32([kp2[m.trainIdx].pt for m in matches]).reshape(-1, 1, 2)
        M, _ = cv2.findHomography(dst_pts, src_pts, cv2.RANSAC, 5.0)
        h, w, _ = img_a.shape
        img_b_aligned = cv2.warpPerspective(img_b, M, (w, h))
        return img_a, img_b_aligned, True
    return img_a, img_b, False

def blend_overlay(orig_arr, mask_diff, mask_above, alpha):
    h, w, _ = orig_arr.shape
    overlay = np.full((h, w, 3), 128, dtype=np.uint8)
    overlay[mask_diff & ~mask_above] = [0, 120, 212]
    overlay[mask_above] = [232, 17, 35]
    return (orig_arr.astype(float)*(1-alpha) + overlay.astype(float)*alpha).astype(np.uint8)

# --- 3. UI 布局 ---
st.set_page_config(page_title="Office-Comparer Pro", layout="wide")
if 'dark_mode' not in st.session_state: st.session_state.dark_mode = False

with st.sidebar:
    st.markdown("### 🎨 界面定制")
    st.session_state.dark_mode = st.toggle("🌙 黑暗模式", value=st.session_state.dark_mode)
    view_mode = st.radio("👀 查看模式", ["左右双栏", "混合视图"])
    st.markdown("---")
    tab_type = st.radio("📁 比对类型", ["文档对比", "Excel 表格", "图像比对"])
    st.markdown("---")
    st.markdown("### ⚙️ 算法参数")
    threshold = st.slider("差异灵敏度", 0, 255, 10)
    alpha = st.slider("差异覆盖强度", 0, 100, 80) / 100.0

apply_custom_style(st.session_state.dark_mode)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

# --- 4. 核心功能 ---
if tab_type == "文档对比":
    u1, u2 = st.columns(2)
    # 扩展支持的文件类型
    allowed_types = ["docx", "txt", "py", "c", "cpp", "h", "java", "js", "html", "css", "md"]
    fa = u1.file_uploader("文件 A", type=allowed_types)
    fb = u2.file_uploader("文件 B", type=allowed_types)
    
    if fa and fb:
        lines_a, lines_b = load_document_lines(fa), load_document_lines(fb)
        matcher = difflib.SequenceMatcher(None, lines_a, lines_b)
        opcodes = matcher.get_opcodes()
        
        # Minimap
        st.markdown('<div class="minimap-container">', unsafe_allow_html=True)
        for tag, i1, i2, j1, j2 in opcodes:
            color = "#22c55e" if tag == 'insert' else ("#ffb900" if tag in ['delete','replace'] else "transparent")
            st.markdown(f'<div class="minimap-bit" style="background:{color}"></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if view_mode == "左右双栏":
            for tag, i1, i2, j1, j2 in opcodes:
                c1, c2 = st.columns(2)
                if tag == 'equal':
                    for k in range(i2-i1):
                        c1.markdown(f"<div class='doc-content'>{lines_a[i1+k]}</div>", unsafe_allow_html=True)
                        c2.markdown(f"<div class='doc-content'>{lines_b[j1+k]}</div>", unsafe_allow_html=True)
                else:
                    if i2>i1: c1.markdown(f"<div class='doc-content diff-removed'>{'<br>'.join(lines_a[i1:i2])}</div>", unsafe_allow_html=True)
                    if j2>j1: c2.markdown(f"<div class='doc-content diff-added'>{'<br>'.join(lines_b[j1:j2])}</div>", unsafe_allow_html=True)
        else:
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == 'equal':
                    for line in lines_a[i1:i2]: st.markdown(f"<div class='doc-content'>{line}</div>", unsafe_allow_html=True)
                elif tag in ['delete','replace']:
                    for line in lines_a[i1:i2]: st.markdown(f"<div class='doc-content diff-removed'>[-] {line}</div>", unsafe_allow_html=True)
                if tag in ['insert','replace']:
                    for line in lines_b[j1:j2]: st.markdown(f"<div class='doc-content diff-added'>[+] {line}</div>", unsafe_allow_html=True)

elif tab_type == "Excel 表格":
    u1, u2 = st.columns(2)
    ea, eb = u1.file_uploader("表格 A"), u2.file_uploader("表格 B")
    if ea and eb:
        df_a, df_b = pd.read_excel(ea), pd.read_excel(eb)
        c1, c2 = st.columns(2)
        c1.dataframe(df_a, use_container_width=True); c2.dataframe(df_b, use_container_width=True)

elif tab_type == "图像比对":
    iu1, iu2 = st.columns(2)
    ia_f, ib_f = iu1.file_uploader("图片 A"), iu2.file_uploader("图片 B")
    if ia_f and ib_f:
        with st.spinner("🔍 正在对齐图像..."):
            a1, a2, _ = align_images_cv2(ia_f.getvalue(), ib_f.getvalue())
            md = np.sum(np.abs(a1.astype(np.int16)-a2.astype(np.int16)), 2)>0
            ma = np.mean(np.abs(a1.astype(np.int16)-a2.astype(np.int16)), 2)>threshold
            va, vb = blend_overlay(a1, md, ma, alpha), blend_overlay(a2, md, ma, alpha)
            fig = make_subplots(rows=1, cols=2)
            fig.add_trace(go.Image(z=va), 1, 1); fig.add_trace(go.Image(z=vb), 1, 2)
            fig.update_layout(height=700, margin=dict(l=0,r=0,b=0,t=10))
            st.plotly_chart(fig, use_container_width=True)
