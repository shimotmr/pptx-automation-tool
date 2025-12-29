import streamlit as st
import os
import uuid
import json
import shutil
import traceback
import requests
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# =========================
# 基本設定
# =========================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"

# =========================
# 全站 CSS（企業版）
# =========================
st.markdown("""
<style>
header, footer { display:none !important; }

.block-container{
  padding-top:1.5rem !important;
}

/* 品牌藍 */
.brand-info{
  background:#EAF3FF;
  border-left:4px solid #0B4F8A;
  padding:12px 16px;
  border-radius:10px;
  color:#0B4F8A;
  font-weight:600;
}

/* Section Card */
.section-card{
  border:1px solid #E0E0E0;
  border-radius:14px;
  padding:16px;
  margin-bottom:18px;
}

/* Result Card */
.result-card{
  border:1px solid #E0E0E0;
  border-radius:12px;
  padding:12px 16px;
  margin-bottom:10px;
  display:flex;
  align-items:center;
  justify-content:space-between;
}

/* FileUploader 精簡 */
[data-testid="stFileUploaderDropzoneInstructions"] > div { display:none !important; }

[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳 PPTX";
  font-weight:700;
  font-size:0.9rem;
}

[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"單一檔案上限 5GB";
  font-size:0.75rem;
  color:#888;
}

section[data-testid="stFileUploaderDropzone"]{
  padding:0.6rem 0.9rem !important;
  border-radius:14px !important;
  background:#F8FAFD !important;
}

/* Dropzone 只留一顆瀏覽檔案 */
section[data-testid="stFileUploaderDropzone"] button{
  display:flex !important;
  align-items:center;
  justify-content:center;
  min-height:42px;
  font-size:0;
}
section[data-testid="stFileUploaderDropzone"] button::after{
  content:"瀏覽檔案";
  font-size:0.9rem;
  font-weight:700;
}

/* 隱藏列表區第二顆按鈕 */
div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# 工具函式
# =========================
def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def load_history(filename):
    if not os.path.exists(HISTORY_FILE):
        return []
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        return json.load(f).get(filename, [])

def save_history(filename, jobs):
    data = {}
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    data[filename] = jobs
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# =========================
# Header
# =========================
st.markdown(f"""
<div style="display:flex;flex-direction:column;align-items:center;margin-bottom:12px;">
  <img src="{LOGO_URL}" style="width:300px;">
  <div style="margin-top:6px;color:#666;font-weight:500;">
    簡報案例自動化發布平台
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="brand-info">
上傳 PPT → 拆分 → 影片雲端化 → 簡報發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# =========================
# 初始化
# =========================
if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []

# =========================
# Step 1
# =========================
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.subheader("步驟一｜選擇檔案來源")

uploaded_file = st.file_uploader("PPTX", type=["pptx"], label_visibility="collapsed")

if uploaded_file:
    cleanup_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")
    with open(source_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    prs = Presentation(source_path)
    st.markdown(
        f"<div class='brand-info'>已讀取：{uploaded_file.name}（共 {len(prs.slides)} 頁）</div>",
        unsafe_allow_html=True
    )
    st.session_state.current_file = uploaded_file.name
    st.session_state.total_slides = len(prs.slides)

st.markdown("</div>", unsafe_allow_html=True)

# =========================
# Step 2（簡化示意）
# =========================
if "current_file" in st.session_state:
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("步驟二｜開始處理")

    auto_clean = st.checkbox("完成後自動清除暫存檔", value=True)

    if st.button("開始處理", use_container_width=True):
        status = st.empty()
        status.markdown("<div class='brand-info'>流程執行中，請稍候…</div>", unsafe_allow_html=True)

        # === 實際流程 ===
        # 此處呼叫你的 execute_automation_logic（略）

        status.markdown("<div class='brand-info'>流程已完成，所有步驟成功執行。</div>", unsafe_allow_html=True)

        st.markdown("### 產出結果")

        for i in range(1):
            link = "https://docs.google.com/presentation"
            st.markdown(f"""
            <div class="result-card">
              <div>案例簡報</div>
              <div>
                <a href="{link}" target="_blank">開啟簡報</a>
                &nbsp;
                <span onclick="navigator.clipboard.writeText('{link}')" style="cursor:pointer;">📋</span>
              </div>
            </div>
            """, unsafe_allow_html=True)

        if auto_clean:
            cleanup_workspace()

        st.divider()

        if st.button("返回並處理新檔", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
