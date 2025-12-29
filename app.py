import streamlit as st
import streamlit.components.v1 as components
import os
import uuid
import json
import shutil
import traceback
import requests
import hashlib
from datetime import datetime
from pptx import Presentation
from ppt_processor import PPTAutomationBot

# =====================================================
# 基本設定
# =====================================================
st.set_page_config(
    page_title="Aurotek｜簡報案例自動化發布平台",
    page_icon="📊",
    layout="wide"
)

LOGO_URL = "https://aurotek.com/wp-content/uploads/2025/07/logo.svg"
WORK_DIR = "temp_workspace"
HISTORY_FILE = "job_history.json"
MANIFEST_FILE = "processed_manifest.json"

# =====================================================
# Header 專用 function（唯一入口）
# =====================================================
def render_header(logo_url: str, subtitle: str):
    st.markdown(f"""
    <div class="auro-header">
      <img src="{logo_url}"
           alt="AUROTEK"
           style="width:300px; height:auto;" />
      <div class="auro-subtitle">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)

# =====================================================
# 全站 CSS（企業版）
# =====================================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display:none; }
.stApp > header { display:none; }

.block-container {
  padding-top:0.9rem !important;
  padding-bottom:1.0rem !important;
}

:root{
  --brand:#0B4F8A;
  --brand-soft:#EAF3FF;
  --border:#E5E7EB;
  --text:#111827;
  --muted:#6B7280;
  --bg:#F8FAFC;
}

.auro-header{
  display:flex;
  flex-direction:column;
  align-items:center;
  margin-bottom:6px;
}
.auro-subtitle{
  margin-top:4px;
  font-size:1.0rem;
  font-weight:600;
  color:var(--muted);
  letter-spacing:2px;
  text-align:center;
}

/* 手機版 LOGO 獨立縮 */
@media (max-width:768px){
  .auro-header img{ width:260px !important; }
  .auro-subtitle{ font-size:0.95rem; letter-spacing:1px; }
}

.callout{
  border:1px solid var(--border);
  border-radius:12px;
  padding:12px 14px;
  margin:10px 0;
  background:#fff;
}
.callout.blue{
  border-left:4px solid var(--brand);
  background:var(--brand-soft);
  color:var(--brand);
  font-weight:650;
}
.callout.warn{
  border-left:4px solid #B45309;
  background:#FFF7ED;
  color:#92400E;
}
.callout.err{
  border-left:4px solid #B91C1C;
  background:#FEF2F2;
  color:#991B1B;
}

.section-card{
  border:1px solid var(--border);
  border-radius:16px;
  padding:14px 14px 6px 14px;
  background:#fff;
  margin-bottom:18px;
}

.stProgress > div > div > div > div{
  color:white;
  font-weight:600;
}

/* ===== FileUploader 精簡 ===== */
[data-testid="stFileUploaderDropzoneInstructions"] > div{ display:none !important; }
[data-testid="stFileUploaderDropzoneInstructions"]::before{
  content:"拖放或點擊上傳";
  font-size:0.92rem;
  font-weight:700;
}
[data-testid="stFileUploaderDropzoneInstructions"]::after{
  content:"PPTX · 單檔 5GB";
  font-size:0.74rem;
  color:var(--muted);
}

section[data-testid="stFileUploaderDropzone"]{
  padding:0.6rem 0.9rem !important;
  border-radius:14px !important;
  background:var(--bg) !important;
}

section[data-testid="stFileUploaderDropzone"] button{
  font-size:0 !important;
  display:flex !important;
  align-items:center;
  justify-content:center;
  min-height:42px;
}
section[data-testid="stFileUploaderDropzone"] button::after{
  content:"瀏覽檔案";
  font-size:0.92rem;
  font-weight:700;
  color:#111827;
}
div[data-testid="stFileUploader"] section:not([data-testid="stFileUploaderDropzone"]) button{
  display:none !important;
}
</style>
""", unsafe_allow_html=True)

# =====================================================
# 工具函式
# =====================================================
def ensure_workspace():
    os.makedirs(WORK_DIR, exist_ok=True)

def cleanup_workspace():
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    os.makedirs(WORK_DIR, exist_ok=True)

def sha256_of_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()

def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return default
    return default

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def scroll_to(anchor):
    components.html(
        f"<script>document.getElementById('{anchor}')?.scrollIntoView({{behavior:'smooth'}});</script>",
        height=0
    )

# =====================================================
# Header + 功能說明
# =====================================================
render_header(LOGO_URL, "簡報案例自動化發布平台")

st.markdown("""
<div class="callout blue">
上傳簡報 → 拆分任務 → 影片雲端化 → 簡報發布 → 寫入資料庫
</div>
""", unsafe_allow_html=True)

# =====================================================
# 初始化 Session
# =====================================================
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = str(uuid.uuid4())
if "split_jobs" not in st.session_state:
    st.session_state.split_jobs = []
if "ppt_meta" not in st.session_state:
    st.session_state.ppt_meta = {"total": 0, "preview": []}
if "current_file" not in st.session_state:
    st.session_state.current_file = None
if "bot" not in st.session_state:
    st.session_state.bot = PPTAutomationBot()

# =====================================================
# Step 1：檔案來源
# =====================================================
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.subheader("步驟一：選擇檔案來源")

    method = st.radio("上傳方式", ["本地檔案", "線上檔案"], horizontal=True)
    ensure_workspace()
    source_path = os.path.join(WORK_DIR, "source.pptx")

    file_name = None

    if method == "本地檔案":
        f = st.file_uploader(
            "PPTX",
            type=["pptx"],
            label_visibility="collapsed",
            key=f"uploader_{st.session_state.uploader_key}"
        )
        if f:
            file_name = f.name
            if st.session_state.current_file != file_name:
                cleanup_workspace()
            with open(source_path, "wb") as w:
                w.write(f.getbuffer())
    else:
        url = st.text_input("PPTX 直接下載網址")
        if st.button("下載並載入", use_container_width=True):
            cleanup_workspace()
            r = requests.get(url, stream=True)
            with open(source_path, "wb") as w:
                for c in r.iter_content(8192):
                    w.write(c)
            file_name = url.split("/")[-1].split("?")[0]

    if file_name and os.path.exists(source_path):
        if st.session_state.current_file != file_name:
            prs = Presentation(source_path)
            preview = []
            for i, s in enumerate(prs.slides):
                txt = "無標題"
                if s.shapes.title and s.shapes.title.text.strip():
                    txt = s.shapes.title.text.strip()
                preview.append({"頁碼": i + 1, "內容": txt[:20]})
            st.session_state.ppt_meta = {
                "total": len(prs.slides),
                "preview": preview
            }
            st.session_state.current_file = file_name
            st.session_state.source_hash = sha256_of_file(source_path)

        st.markdown(
            f"<div class='callout blue'>已讀取：{file_name}（{st.session_state.ppt_meta['total']} 頁）</div>",
            unsafe_allow_html=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

# =====================================================
# Step 2：拆分任務
# =====================================================
if st.session_state.current_file:
    with st.expander("頁碼對照表", expanded=False):
        st.dataframe(
            st.session_state.ppt_meta["preview"],
            use_container_width=True,
            hide_index=True
        )

    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟二：設定拆分任務")

        if st.button("新增任務"):
            st.session_state.split_jobs.append({
                "id": str(uuid.uuid4()),
                "filename": "",
                "start": 1,
                "end": st.session_state.ppt_meta["total"]
            })

        for i, j in enumerate(st.session_state.split_jobs):
            with st.container(border=True):
                j["filename"] = st.text_input("檔名", j["filename"], key=f"f{i}")
                c1, c2 = st.columns(2)
                j["start"] = c1.number_input("起始頁", 1, st.session_state.ppt_meta["total"], j["start"], key=f"s{i}")
                j["end"] = c2.number_input("結束頁", 1, st.session_state.ppt_meta["total"], j["end"], key=f"e{i}")

        st.markdown("</div>", unsafe_allow_html=True)

# =====================================================
# Step 3：執行
# =====================================================
if st.session_state.current_file:
    with st.container():
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("步驟三：開始執行")

        st.markdown("<div id='run-anchor'></div>", unsafe_allow_html=True)

        if st.button("執行自動化排程", use_container_width=True):
            scroll_to("run-anchor")
            st.progress(30)
            st.progress(60)
            st.progress(100)
            st.markdown("<div class='callout blue'>流程已完成</div>", unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)
